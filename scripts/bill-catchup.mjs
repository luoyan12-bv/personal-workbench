// 幂等兜底补跑：检测今日日报卡片是否已推送，没发过才补推（run() 内部还有二次去重，双保险）
import { run } from "./bill-feishu-action.mjs";

const SB = (process.env.SUPABASE_URL || "https://wvniuyfnuyrebowjlxbi.supabase.co").replace(/\/$/, "");
const K = process.env.SUPABASE_ANON_KEY || "sb_publishable_Ku0efv21pS5y5rS4cQDXpA_-ifqshOX";

async function sbGet(id){
  const r = await fetch(`${SB}/rest/v1/pwb_state?id=eq.${id}&select=data`, { headers: { apikey: K, Authorization: "Bearer " + K } });
  if(!r.ok) throw new Error("Supabase HTTP " + r.status);
  const rows = await r.json();
  return (rows && rows[0] && rows[0].data) || null;
}

const d = await sbGet(3);
const daily = d && d.daily;
if(!daily){ console.log("无日报数据，无需补推。"); process.exit(0); }
const h = await sbGet(5);
const dates = ((h && h.history) || []).map(x => String(x.date).slice(0, 10));
if(dates.includes(String(daily.date).slice(0, 10))){
  console.log("今日卡片已推送（" + daily.date + "），无需补推。");
  process.exit(0);
}
console.log("检测到 " + daily.date + " 尚未推送，补跑中…");
await run();
