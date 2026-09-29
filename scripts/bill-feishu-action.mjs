// 磁力金牛日报 · 商务消耗排行 · 自动推送飞书群（GitHub Actions 无人值守）
// 读取 Supabase pwb_state id=3（日报/月报）+ id=5（历史归档），算排名/前日/环比/月消耗，
// 发飞书 interactive 卡片(markdown 表格)，并把今日归档回写 id=5（供次日算环比）。
// 环境变量：SUPABASE_URL / SUPABASE_ANON_KEY / FEISHU_WEBHOOK / FEISHU_SECRET(可选)
import { createHmac } from "crypto";

const SB = (process.env.SUPABASE_URL || "https://wvniuyfnuyrebowjlxbi.supabase.co").replace(/\/$/, "");
const ANON = process.env.SUPABASE_ANON_KEY || "sb_publishable_Ku0efv21pS5y5rS4cQDXpA_-ifqshOX"; // 公开 publishable 钥匙，内置默认减少云端 Secrets 配置
const WEBHOOK = (process.env.FEISHU_WEBHOOK || "").trim();
const SECRET = (process.env.FEISHU_SECRET || "").trim();
// 飞书自建应用发送方式（优先）：tenant token + chat_id
const APP_ID = (process.env.FEISHU_APP_ID || "").trim();
const APP_SECRET = (process.env.FEISHU_APP_SECRET || "").trim();
const CHAT_ID = (process.env.FEISHU_CHAT_ID || "").trim();

function billBizName(s){ const t = String(s == null ? "" : s).trim(); const m = t.match(/^([^\(（]+)/); return m ? m[1].trim() : t; }
function addDays(s, n){ const d = new Date(s + "T00:00:00"); d.setDate(d.getDate() + n); const p = x => String(x).padStart(2, "0"); return d.getFullYear() + "-" + p(d.getMonth() + 1) + "-" + p(d.getDate()); }

async function sbSelect(id){
  const url = `${SB}/rest/v1/pwb_state?id=eq.${id}&select=data`;
  const r = await fetch(url, { headers: { "apikey": ANON, "Authorization": "Bearer " + ANON } });
  if(!r.ok) throw new Error("Supabase 读取 id=" + id + " HTTP " + r.status);
  const rows = await r.json();
  return (rows && rows[0] && rows[0].data) ? rows[0].data : null;
}
async function sbUpsert(id, data){
  const url = `${SB}/rest/v1/pwb_state`;
  const r = await fetch(url, { method: "POST", headers: { "apikey": ANON, "Authorization": "Bearer " + ANON, "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates" }, body: JSON.stringify({ id, data }) });
  if(!r.ok) throw new Error("Supabase 写入 id=" + id + " HTTP " + r.status);
}

function buildRanking(daily, monthly, history){
  if(!daily) return null;
  const monthMap = {};
  (monthly && monthly.bizs ? monthly.bizs : []).forEach(b => { monthMap[billBizName(b.biz)] = Number(b.spent || 0); });
  const today = String(daily.date || "").slice(0, 10);
  const prevDate = addDays(today, -1);
  const prev = (history || []).find(h => String(h.date).slice(0, 10) === prevDate);
  const prevMap = {};
  (prev && prev.bizs ? prev.bizs : []).forEach(b => { prevMap[billBizName(b.biz)] = Number(b.spent || 0); });
  const rows = (daily.bizs || []).map(b => {
    const name = billBizName(b.biz);
    const t = Number(b.spent || 0), p = Number(prevMap[name] || 0);
    const pct = p > 0 ? ((t - p) / p * 100) : null;
    return { name, todaySpent: t, prevSpent: p, pct, monthSpent: Number(monthMap[name] || 0) };
  }).sort((a, b) => b.todaySpent - a.todaySpent);
  rows.forEach((r, i) => r.rank = i + 1);
  return { date: today, prevDate, hasPrev: !!prev, rows, totalToday: rows.reduce((s, r) => s + r.todaySpent, 0), totalMonth: rows.reduce((s, r) => s + r.monthSpent, 0) };
}

function buildMarkdown(rk){
  const fmt = n => "¥" + Number(n || 0).toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  let m = "| 排名 | 商务 | 今日消耗 | 前日消耗 | 环比 | 月消耗 |\n| :---: | --- | ---: | ---: | ---: | ---: |\n";
  rk.rows.forEach(r => {
    const pct = r.pct === null ? "—" : (r.pct >= 0 ? "↑" : "↓") + Math.abs(r.pct).toFixed(1) + "%";
    m += "| " + r.rank + " | " + r.name + " | " + fmt(r.todaySpent) + " | " + fmt(r.prevSpent) + " | " + pct + " | " + fmt(r.monthSpent) + " |\n";
  });
  m += "| **合计** | — | **" + fmt(rk.totalToday) + "** | — | — | **" + fmt(rk.totalMonth) + "** |";
  if(!rk.hasPrev) m += "\n\n> 提示：暂无前日数据，环比暂缺（历史归档后将自动补齐）。";
  return m;
}

// 飞书卡片 markdown 不支持表格 → 用 column_set（列布局）拼真表格（2026-09-29 实测）
const RANK_COLS = [
  { w: 1, align: "center" },
  { w: 2, align: "left"  },
  { w: 2, align: "right" },
  { w: 2, align: "right" },
  { w: 1, align: "right" },
  { w: 2, align: "right" }
];
function rankRow(cells, bg){
  return {
    tag: "column_set",
    flex_mode: "none",
    background_style: bg || "default",
    columns: cells.map((c, i) => ({
      tag: "column",
      width: "weighted",
      weight: RANK_COLS[i].w,
      horizontal_align: RANK_COLS[i].align,
      vertical_align: "center",
      elements: [{ tag: "markdown", content: c }]
    }))
  };
}
function buildTableElements(rk){
  const fmt = n => "¥" + Number(n || 0).toLocaleString("zh-CN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const els = [];
  els.push(rankRow(["**排名**","**商务**","**今日消耗**","**前日消耗**","**环比**","**月消耗**"], "grey"));
  rk.rows.forEach(r => {
    const pct = r.pct === null ? "—" : (r.pct >= 0 ? "↑ " : "↓ ") + Math.abs(r.pct).toFixed(1) + "%";
    els.push(rankRow([String(r.rank), r.name, fmt(r.todaySpent), fmt(r.prevSpent), pct, fmt(r.monthSpent)]));
  });
  els.push(rankRow(["**合计**","—","**" + fmt(rk.totalToday) + "**","—","—","**" + fmt(rk.totalMonth) + "**"], "grey"));
  els.push({ tag: "hr" });
  els.push({ tag: "note", elements: [{ tag: "plain_text", content: rk.hasPrev ? ("前日数据：" + rk.prevDate) : "暂无前日数据，环比暂缺（历史归档后将自动补齐）" }] });
  return els;
}

function feishuSign(secret, ts){
  const h = createHmac("sha256", secret).update(ts + "\n" + secret).digest("base64");
  return encodeURIComponent(h);
}
async function getTenantToken(){
  const r = await fetch("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ app_id: APP_ID, app_secret: APP_SECRET }) });
  const j = await r.json();
  if(j.code !== 0) throw new Error("飞书 tenant token 失败: " + JSON.stringify(j));
  return j.tenant_access_token;
}
async function postFeishu(msgType, card){
  // 方式 A：飞书自建应用（tenant token + chat_id）
  if(APP_ID && APP_SECRET && CHAT_ID){
    const tok = await getTenantToken();
    const r = await fetch("https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=chat_id", { method: "POST", headers: { "Content-Type": "application/json", Authorization: "Bearer " + tok }, body: JSON.stringify({ receive_id: CHAT_ID, receive_id_type: "chat_id", msg_type: msgType, content: JSON.stringify(card) }) });
    const j = await r.json().catch(() => ({}));
    if(process.env.FEISHU_VERBOSE) console.error("Feishu 回包:", JSON.stringify(j));
    if(r.ok && j.code === 0) return { ok: true, msgType: j.data && j.data.message && j.data.message.message_type };
    return { ok: false, msg: j.msg || ("HTTP " + r.status) };
  }
  // 方式 B：群机器人 webhook（可选加签）
  if(!WEBHOOK) throw new Error("未配置飞书发送方式（App 或 Webhook）");
  const body = { msg_type: msgType, card };
  if(SECRET){ const ts = Math.floor(Date.now() / 1000); body.timestamp = String(ts); body.sign = feishuSign(SECRET, ts); }
  const r = await fetch(WEBHOOK, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  const j = await r.json().catch(() => ({}));
  if(r.ok && (j.code === 0 || j.StatusCode === 0 || !("code" in j))) return { ok: true };
  return { ok: false, msg: j.msg || ("HTTP " + r.status) };
}

export async function run(){
  const d = await sbSelect(3);
  const daily = d && d.daily, monthly = d && d.monthly;
  if(!daily){ console.log("无日报数据，跳过推送。"); return; }
  const histData = await sbSelect(5);
  const history = (histData && histData.history) || [];
  const rk = buildRanking(daily, monthly, history);
  const card = { config: { wide_screen_mode: true }, header: { title: { tag: "plain_text", content: "磁力金牛 · 商务消耗日报 " + rk.date }, template: "blue" }, elements: buildTableElements(rk) };
  const res = await postFeishu("interactive", card);
  // 归档今日到历史（按日期去重），供次日算环比
  const today = String(daily.date || "").slice(0, 10);
  const next = history.filter(h => String(h.date).slice(0, 10) !== today);
  next.unshift({ date: today, bizs: daily.bizs || [], total: daily.total });
  if(next.length > 60) next.length = 60;
  await sbUpsert(5, { history: next });
  console.log(res.ok ? ("已推送飞书群 ✓ " + rk.date + " | 商务 " + rk.rows.length + " 家 | 今日合计 " + rk.totalToday) : ("推送失败: " + res.msg));
  process.exit(res.ok ? 0 : 1);
}

import { pathToFileURL } from "node:url";
if(import.meta.url === pathToFileURL(process.argv[1]).href) run().catch(e => { console.error(e.message); process.exit(1); });
