#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
read-bill-mail.py — 独立运行的磁力金牛报表同步脚本（日报/周报/月报）

功能：
  1. 用 IMAP 直接登录邮箱（默认 QQ 邮箱 imap.qq.com:993 SSL），无需任何外部 AI / 连接器 / 人工；
  2. 自动筛选主题含「日报 / 周报 / 月报」且发件人命中白名单（默认快手代理商小助手）的邮件；
  3. 解析邮件附件（.xlsx 优先，.csv / 正文文本兜底），按商务合并消耗；
  4. 按 日报/周报/月报 分类汇总，输出：
       - 本地 JSON 覆盖：latest-{daily|weekly|monthly}.json（与工作台「消耗账单」模块数据结构一致）
       - 本地 CSV 追加：bill-{daily|weekly|monthly}.csv（历史累积，utf-8-sig 兼容 Excel）
       - Supabase pwb_state id=3（工作台前端实时可见，可用 --no-supabase 关闭）
  5. 完善的错误处理（认证失败/网络异常/解析失败分类退出码）+ 双通道日志（控制台 + 滚动文件）。

解析核心（列名自动识别 / 商务合并 / 金额清洗 / 日期范围计算）源自 sync-bill-mail.py，
规则保持一致；本脚本为单文件自包含实现，不依赖该文件存在。

用法示例：
  python read-bill-mail.py --user your-email@qq.com --password 你的16位授权码 --days 1
  python read-bill-mail.py --type daily --days 1 --dry-run     # 只解析预览，不落盘不推送
  python read-bill-mail.py --sample                            # 内置示例数据自测全链路

退出码：
  0 成功（含部分邮件解析失败但至少一类成功）
  1 运行时错误（参数缺失/配置错误/文件写失败等）
  2 参数错误（argparse 自动）
  3 认证失败（授权码错误或过期）
  4 网络错误（IMAP/Supabase 连接失败，重试 1 次后仍失败）
  5 无匹配邮件（日期范围内没有符合条件的邮件）
  6 全部邮件解析失败（至少一封匹配但无一成功解析）
  7 Supabase 推送失败
"""

import argparse
import base64
import csv
import io
import json
import logging
import os
import re
import socket
import sys
import time
import urllib.request
import warnings
from datetime import datetime, timezone, timedelta
from email import policy
from email.parser import BytesParser
from email.utils import parseaddr, parsedate_to_datetime

# ---------------------------------------------------------------------------
# 常量区
# ---------------------------------------------------------------------------

# Supabase 默认值（公开 publishable 钥匙，安全可嵌入；可用环境变量覆盖）
SUPABASE_URL_DEFAULT = "https://wvniuyfnuyrebowjlxbi.supabase.co"
ANON_KEY_DEFAULT = "sb_publishable_Ku0efv21pS5y5rS4cQDXpA_-ifqshOX"

# QQ 邮箱默认 IMAP 配置
IMAP_HOST_DEFAULT = "imap.qq.com"
IMAP_PORT_DEFAULT = 993
MAIL_SENDER_DEFAULT = "agent@contact.kuaishou.com"  # 快手代理商小助手

TYPE_TITLE = {"daily": "日报", "weekly": "周报", "monthly": "月报"}
TYPE_RE = re.compile(r"(日报|周报|月报)")

# 列名自动识别（与 sync-bill-mail.py 保持一致，精确匹配、忽略大小写）
FIELD_ALIASES = {
    "account": ["广告主名称", "账户名称", "账户", "账户名", "广告主", "客户名称", "名称"],
    "id": ["广告主id", "广告主ID", "账户id", "账户ID", "id", "编号", "账户编号"],
    "biz": ["销售责任人", "商务", "商务名称", "销售", "责任人", "商务负责人"],
    "spent": ["总花费(元)(合)", "总花费(元)", "总花费", "总消耗", "消耗", "花费", "消耗金额", "金额", "总金额"],
    "date": ["时间", "日期", "消耗日期", "报表日期"],
}

CSV_COLUMNS = [
    "报表类型", "报表日期", "日期范围", "商务", "账户", "id", "花费(元)",
    "邮件主题", "发件人", "邮件日期", "文件来源", "解析时间",
]

EXIT_OK = 0
EXIT_RUNTIME = 1
EXIT_ARGS = 2
EXIT_AUTH = 3
EXIT_NETWORK = 4
EXIT_NO_MAIL = 5
EXIT_PARSE = 6
EXIT_SUPABASE = 7

BEIJING = timezone(timedelta(hours=8))

logger = logging.getLogger("bill-sync")


# ---------------------------------------------------------------------------
# 自定义异常
# ---------------------------------------------------------------------------

class AuthError(Exception):
    """IMAP 认证失败（授权码错误/过期）"""

class NetworkError(Exception):
    """网络/连接类错误"""

class ParseError(Exception):
    """邮件解析失败"""

class SupabaseError(Exception):
    """Supabase 推送失败"""


# ---------------------------------------------------------------------------
# 配置：环境变量 / .env / CLI 三层合并
# ---------------------------------------------------------------------------

def load_env_file():
    """手写 .env 加载（零额外依赖）：脚本同目录与当前目录下的 .env，已有环境变量优先。"""
    candidates = []
    script_dir = os.path.dirname(os.path.abspath(__file__))
    for d in (script_dir, os.getcwd()):
        p = os.path.join(d, ".env")
        if os.path.isfile(p):
            candidates.append(p)
    seen = set()
    for path in candidates:
        if path in seen:
            continue
        seen.add(path)
        try:
            with open(path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "=" not in line:
                        continue
                    key, _, val = line.partition("=")
                    key = key.strip()
                    val = val.strip().strip('"').strip("'")
                    if key and key not in os.environ:
                        os.environ[key] = val
        except OSError:
            pass


def _read_password(args):
    """密码优先级：环境变量 IMAP_PASS > --password-file > --password。"""
    env_pass = os.environ.get("IMAP_PASS")
    if env_pass:
        return env_pass, "环境变量 IMAP_PASS"
    if args.password_file:
        try:
            with open(args.password_file, "r", encoding="utf-8") as f:
                pw = f.read().strip()
            if not pw:
                raise RuntimeError(f"密码文件 {args.password_file} 为空")
            return pw, f"密码文件 {args.password_file}"
        except OSError as e:
            raise RuntimeError(f"无法读取密码文件 {args.password_file}: {e}")
    if args.password:
        logger.warning("检测到通过命令行传入 --password：密码会出现在进程列表中，建议改用 IMAP_PASS 环境变量或 --password-file")
        return args.password, "命令行 --password"
    raise RuntimeError("未提供邮箱密码/授权码：请设置环境变量 IMAP_PASS，或使用 --password-file / --password")


def _parse_date(s):
    try:
        return datetime.strptime(s, "%Y-%m-%d").date()
    except (ValueError, TypeError):
        raise RuntimeError(f"日期格式错误: {s}（应为 YYYY-MM-DD）")


def load_config(args):
    env = os.environ
    today = datetime.now(BEIJING).date()
    # 日期范围（自然日，北京时间；--days 含今天，--start/--end 默认到今天）
    if args.days:
        if args.days < 1:
            raise RuntimeError("--days 必须 >= 1")
        start, end = today - timedelta(days=args.days - 1), today
    else:
        start = _parse_date(args.start) if args.start else None
        end = _parse_date(args.end) if args.end else (today if start else None)
    cfg = {
        "imap_host": args.imap_host or env.get("IMAP_HOST") or IMAP_HOST_DEFAULT,
        "imap_port": args.imap_port or int(env.get("IMAP_PORT") or IMAP_PORT_DEFAULT),
        "ssl": not args.no_ssl,
        "user": args.user or env.get("IMAP_USER") or "",
        "folder": args.folder or env.get("IMAP_FOLDER") or "INBOX",
        "sender": args.sender if args.sender is not None else env.get("MAIL_SENDER") or MAIL_SENDER_DEFAULT,
        "extra_keywords": [k.strip() for k in (args.subject_keywords or env.get("SUBJECT_KEYWORDS") or "").split(",") if k.strip()],
        "type_filter": None if args.type == "all" else args.type,
        "start": datetime.combine(start, datetime.min.time()).replace(tzinfo=BEIJING) if start else None,
        "end": datetime.combine(end, datetime.min.time()).replace(tzinfo=BEIJING) if end else None,
        "out_dir": args.out_dir or env.get("OUT_DIR") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "bill_out"),
        "supabase": not args.no_supabase,
        "supabase_url": env.get("SUPABASE_URL") or SUPABASE_URL_DEFAULT,
        "anon_key": env.get("SUPABASE_ANON_KEY") or ANON_KEY_DEFAULT,
        "dry_run": args.dry_run,
        "sample": args.sample,
        "log_file": args.log_file,
    }
    cfg["log_file"] = cfg["log_file"] or os.path.join(cfg["out_dir"], "bill-sync.log")
    if not cfg["sample"]:
        if not cfg["user"]:
            raise RuntimeError("未提供邮箱账号：请设置环境变量 IMAP_USER 或使用 --user")
        cfg["password"], cfg["password_src"] = _read_password(args)
    return cfg


# ---------------------------------------------------------------------------
# 日志
# ---------------------------------------------------------------------------

def setup_logging(log_file, verbose=False):
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    sh = logging.StreamHandler(sys.stderr)
    sh.setFormatter(fmt)
    sh.setLevel(logging.DEBUG if verbose else logging.INFO)
    logger.addHandler(sh)
    try:
        os.makedirs(os.path.dirname(os.path.abspath(log_file)), exist_ok=True)
        from logging.handlers import RotatingFileHandler
        fh = RotatingFileHandler(log_file, maxBytes=1024 * 1024, backupCount=3, encoding="utf-8")
        fh.setFormatter(fmt)
        fh.setLevel(logging.DEBUG)
        logger.addHandler(fh)
    except OSError as e:
        logger.warning("无法创建文件日志 %s：%s（仅输出到控制台）", log_file, e)


# ---------------------------------------------------------------------------
# 报表解析核心（源自 sync-bill-mail.py，规则保持一致）
# ---------------------------------------------------------------------------

def find_col(headers, aliases):
    for a in aliases:
        for i, h in enumerate(headers):
            if str(h).strip().lower() == a.lower():
                return i
    return -1


def num(v):
    if v is None or v == "":
        return 0.0
    s = str(v).replace("¥", "").replace(",", "").replace("，", "").strip()
    try:
        return float(s)
    except Exception:
        return 0.0


def biz_name(sales):
    s = str(sales or "").strip()
    m = re.match(r"^([^\(（]+)", s)
    return m.group(1).strip() if m else s


def calc_range(rtype, date):
    """报表覆盖的日期范围（周报=上一自然周，月报=当月），返回中文友好格式，daily 返回空串。"""
    try:
        d = datetime.strptime(date, "%Y-%m-%d")
        if rtype == "weekly":
            monday = d - timedelta(days=d.weekday() + 7)  # 上一周周一
            sunday = monday + timedelta(days=6)
            return f"{monday.month}月{monday.day}日 ~ {sunday.month}月{sunday.day}日"
        if rtype == "monthly":
            start = d.replace(day=1)
            if d.month == 12:
                end = datetime(d.year + 1, 1, 1) - timedelta(days=1)
            else:
                end = datetime(d.year, d.month + 1, 1) - timedelta(days=1)
            return f"{start.month}月{start.day}日 ~ {end.month}月{end.day}日"
        return ""
    except Exception:
        return ""


def parse_rows(headers, rows, fallback_biz=""):
    """统一的按行解析 + 商务合并（xlsx 与 csv/正文共用）。"""
    ci = {k: find_col(headers, v) for k, v in FIELD_ALIASES.items()}
    if ci["biz"] < 0 or ci["spent"] < 0:
        raise ParseError("未识别到 商务/销售责任人 或 消耗/总花费 列: " + json.dumps(headers, ensure_ascii=False))
    accounts = []
    for r in rows:
        if not r:
            continue
        biz = str(r[ci["biz"]] or "").strip() if ci["biz"] < len(r) else ""
        spent = num(r[ci["spent"]]) if ci["spent"] < len(r) else 0.0
        account = str(r[ci["account"]] or "").strip() if ci["account"] >= 0 and ci["account"] < len(r) else ""
        aid = str(r[ci["id"]] or "").strip() if ci["id"] >= 0 and ci["id"] < len(r) else ""
        if not biz and not account and not spent:
            continue
        if not biz:
            biz = fallback_biz or "(未填商务)"
        accounts.append({"account": account, "id": aid, "spent": spent, "biz": biz})
    biz_agg = {}
    for a in accounts:
        key = a["biz"]
        if key not in biz_agg:
            biz_agg[key] = {"biz": key, "spent": 0.0, "accounts": []}
        biz_agg[key]["spent"] += a["spent"]
        biz_agg[key]["accounts"].append({"account": a["account"], "id": a["id"], "spent": round(a["spent"], 2)})
    bizs = sorted(biz_agg.values(), key=lambda x: -x["spent"])
    for b in bizs:
        b["spent"] = round(b["spent"], 2)
    total = round(sum(a["spent"] for a in accounts), 2)  # 用原始账户值求和，避免累积误差
    return {"total": total, "bizs": bizs, "count": len(accounts)}


def parse_xlsx_bytes(blob, fallback_biz=""):
    try:
        import openpyxl
        warnings.filterwarnings("ignore", message="Workbook contains no default style")
    except ImportError:
        raise ParseError("需要 openpyxl 才能解析 .xlsx 附件，请执行: pip install openpyxl")
    try:
        wb = openpyxl.load_workbook(io.BytesIO(blob), data_only=True)
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
    except Exception as e:
        raise ParseError(f"xlsx 解析失败: {e}")
    if not rows:
        raise ParseError("报表为空（xlsx 无数据行）")
    headers = [str(h) if h is not None else "" for h in rows[0]]
    return parse_rows(headers, rows[1:], fallback_biz)


def decode_text(blob):
    """字节流转文本：utf-8 → gbk → latin-1 兜底。"""
    for enc in ("utf-8-sig", "gbk", "latin-1"):
        try:
            return blob.decode(enc)
        except (UnicodeDecodeError, AttributeError):
            continue
    return blob.decode("latin-1", errors="replace")


def parse_csv_text(text, fallback_biz=""):
    """csv/正文文本按 CSV 行解析；自动定位表头行（含商务/花费列的那一行）。"""
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        raise ParseError("报表为空（无文本行）")
    rows = list(csv.reader(lines))
    hdr_idx = 0
    for i, row in enumerate(rows[:10]):
        cells = [str(c or "").strip().lower() for c in row]
        if find_col(cells, FIELD_ALIASES["biz"]) >= 0 and find_col(cells, FIELD_ALIASES["spent"]) >= 0:
            hdr_idx = i
            break
    headers = [str(c) if c is not None else "" for c in rows[hdr_idx]]
    return parse_rows(headers, rows[hdr_idx + 1:], fallback_biz)


def html_to_text(html):
    html = re.sub(r"<br\s*/?>", "\n", html, flags=re.IGNORECASE)
    html = re.sub(r"<[^>]+>", "", html)
    import html as _html
    return _html.unescape(html).strip()


# ---------------------------------------------------------------------------
# IMAP 层
# ---------------------------------------------------------------------------

def imap_login(cfg):
    import imaplib
    last_err = None
    for attempt in (1, 2):
        try:
            if cfg["ssl"]:
                conn = imaplib.IMAP4_SSL(cfg["imap_host"], cfg["imap_port"], timeout=30)
            else:
                conn = imaplib.IMAP4(cfg["imap_host"], cfg["imap_port"])
            conn.login(cfg["user"], cfg["password"])
            logger.info("IMAP 登录成功：%s:%s%s", cfg["imap_host"], cfg["imap_port"], "" if cfg["ssl"] else "（无 SSL）")
            return conn
        except imaplib.IMAP4.error as e:
            msg = str(e)
            if "login" in msg.lower() or "authentication" in msg.lower() or "auth" in msg.lower():
                raise AuthError(f"IMAP 认证失败（{msg}）：请检查账号与授权码（非登录密码），QQ 邮箱需在设置中开启 IMAP/SMTP 并生成 16 位授权码")
            last_err = e
        except (socket.timeout, TimeoutError, ConnectionError, OSError) as e:
            last_err = e
        if attempt == 1:
            logger.warning("IMAP 连接失败（%s），3 秒后重试...", last_err)
            time.sleep(3)
    raise NetworkError(f"IMAP 连接失败（重试后仍失败）: {last_err}")


def _match_type(subject, cfg):
    m = TYPE_RE.search(subject or "")
    if m:
        rtype = {"日报": "daily", "周报": "weekly", "月报": "monthly"}[m.group(1)]
        if cfg["type_filter"] and rtype != cfg["type_filter"]:
            return None
        return rtype
    # 附加关键词兜底
    for kw in cfg["extra_keywords"]:
        if kw and kw in (subject or ""):
            for t, label in TYPE_TITLE.items():
                if label in kw:
                    if cfg["type_filter"] and t != cfg["type_filter"]:
                        return None
                    return t
    return None


def _sender_addr(from_hdr):
    return parseaddr(from_hdr or "")[1].strip().lower()


def _imap_date(d):
    return d.strftime("%d-%b-%Y")  # IMAP SINCE 格式


def search_mails(conn, cfg):
    """SELECT 文件夹 → 拉取头部 → 本地过滤（类型/发件人/日期范围）。返回 MailMeta 列表。"""
    import imaplib
    try:
        typ, data = conn.select(cfg["folder"], readonly=True)
        if typ != "OK":
            raise RuntimeError(f"无法打开文件夹 {cfg['folder']}: {data}")
        crit = ["ALL"]
        if cfg["start"]:
            crit = ["(SINCE %s)" % _imap_date(cfg["start"])]
        typ, data = conn.uid("search", None, *crit)
        if typ != "OK" or not data or not data[0]:
            logger.info("文件夹 %s 中没有候选邮件", cfg["folder"])
            return []
        uids = data[0].split()
    except imaplib.IMAP4.error as e:
        raise NetworkError(f"IMAP 搜索失败: {e}")

    start = cfg["start"]
    end_excl = (cfg["end"] + timedelta(days=1)) if cfg["end"] else None
    mails = []
    for uid in uids:
        try:
            typ, d = conn.uid("fetch", uid, "(BODY.PEEK[HEADER])")
            if typ != "OK" or not d or not d[0]:
                continue
            raw = d[0][1]
            msg = BytesParser(policy=policy.default).parsebytes(raw)
            subject = str(msg.get("Subject", "") or "")
            from_hdr = str(msg.get("From", "") or "")
            date_hdr = str(msg.get("Date", "") or "")
        except Exception as e:
            logger.warning("读取邮件头部失败 uid=%s: %s", uid, e)
            continue

        # 发件人白名单（空=不限；大小写不敏感，按邮箱地址匹配）
        if cfg["sender"]:
            whitelist = {s.strip().lower() for s in cfg["sender"].split(",") if s.strip()}
            if whitelist and _sender_addr(from_hdr) not in whitelist:
                continue

        rtype = _match_type(subject, cfg)
        if rtype is None:
            continue

        # 邮件日期（时区容错，统一转北京时间取日期）
        dt = None
        if date_hdr:
            try:
                dt = parsedate_to_datetime(date_hdr)
            except Exception:
                dt = None
        if dt is None:
            logger.warning("跳过无法解析邮件日期的邮件 uid=%s 主题=%s", uid, subject)
            continue
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        local = dt.astimezone(BEIJING)
        mdate = local.date()

        # 日期范围过滤（北京时间自然日）
        if start and mdate < start.date():
            continue
        if end_excl and mdate >= end_excl.date():
            continue

        mails.append({
            "uid": uid.decode(errors="replace"), "subject": subject, "from": from_hdr,
            "date": dt, "date_local": local, "mail_date": local.strftime("%Y-%m-%d"),
            "type": rtype,
        })
    return mails


def fetch_attachments(conn, meta):
    """拉取整封邮件，提取 xlsx/csv/txt 附件；无附件时用正文文本兜底。"""
    import imaplib
    try:
        typ, d = conn.uid("fetch", meta["uid"], "(RFC822)")
        if typ != "OK" or not d or not d[0]:
            raise NetworkError(f"拉取邮件正文失败 uid={meta['uid']}")
        msg = BytesParser(policy=policy.default).parsebytes(d[0][1])
    except imaplib.IMAP4.error as e:
        raise NetworkError(f"IMAP 拉取失败: {e}")

    atts = []
    for part in msg.walk():
        fn = part.get_filename()
        if fn:
            payload = part.get_payload(decode=True)
            if payload is None:
                continue
            lower = fn.lower()
            if lower.endswith((".xlsx", ".xls", ".csv", ".txt")):
                atts.append({"name": fn, "payload": payload, "is_body": False})
        else:
            ct = part.get_content_type()
            if ct == "text/plain":
                try:
                    body = part.get_content().strip()
                except Exception:
                    body = ""
                if body and len(body) > 5:
                    atts.append({"name": "(正文文本)", "payload": body.encode("utf-8"), "is_body": True})
            elif ct == "text/html":
                try:
                    body = html_to_text(part.get_content()).strip()
                except Exception:
                    body = ""
                if body and len(body) > 5:
                    atts.append({"name": "(正文HTML)", "payload": body.encode("utf-8"), "is_body": True})
    return atts


def parse_attachment(att, rtype, meta):
    """按附件类型分发解析，返回 {total, bizs, count}。"""
    name = att["name"].lower()
    if name.endswith(".xlsx") or name.endswith(".xls"):
        parsed = parse_xlsx_bytes(att["payload"])
    else:
        text = decode_text(att["payload"])
        parsed = parse_csv_text(text)
    if parsed["count"] == 0:
        raise ParseError(f"附件 {att['name']} 无有效数据行")
    return parsed


# ---------------------------------------------------------------------------
# 本地输出
# ---------------------------------------------------------------------------

class _FileLock:
    """跨平台文件锁（Windows: msvcrt；Unix: fcntl），防止多任务并发追加写坏 CSV。"""

    def __init__(self, path, timeout=15):
        self.path = path
        self.timeout = timeout
        self.fh = None

    def __enter__(self):
        self.fh = open(self.path, "a+b")
        if self.fh.tell() == 0:  # msvcrt.locking 需锁定区域有字节
            self.fh.write(b"\x00")
            self.fh.flush()
        self.fh.seek(0)
        try:
            import msvcrt  # Windows
            msvcrt.locking(self.fh.fileno(), msvcrt.LK_LOCK, 1)
        except ImportError:
            import fcntl  # Linux / macOS（GitHub Actions）
            fcntl.flock(self.fh.fileno(), fcntl.LOCK_EX)
        return self

    def __exit__(self, *exc):
        if not self.fh:
            return
        try:
            import msvcrt
            try:
                self.fh.seek(0)
                msvcrt.locking(self.fh.fileno(), msvcrt.LK_UNLCK, 1)
            except Exception:
                pass
        except ImportError:
            import fcntl
            try:
                fcntl.flock(self.fh.fileno(), fcntl.LOCK_UN)
            except Exception:
                pass
        self.fh.close()


def append_csv(out_dir, rtype, rep):
    """bill-{type}.csv 追加历史；首次写表头。"""
    path = os.path.join(out_dir, f"bill-{rtype}.csv")
    meta = rep.get("_meta", {})
    rows = []
    for b in rep.get("bizs", []):
        if b.get("accounts"):
            for a in b["accounts"]:
                rows.append({
                    "报表类型": TYPE_TITLE[rtype],
                    "报表日期": rep.get("date", ""),
                    "日期范围": rep.get("dateRange", ""),
                    "商务": b.get("biz", ""),
                    "账户": a.get("account", ""),
                    "id": a.get("id", ""),
                    "花费(元)": a.get("spent", 0),
                    "邮件主题": meta.get("subject", ""),
                    "发件人": meta.get("sender", ""),
                    "邮件日期": meta.get("mailDate", ""),
                    "文件来源": meta.get("source", ""),
                    "解析时间": meta.get("parsedAt", ""),
                })
    if not rows:
        logger.info("类型 %s 无账户明细，跳过 CSV 追加", rtype)
        return
    new_file = not os.path.exists(path)
    lock_path = os.path.join(out_dir, ".bill-sync.lock")
    try:
        with _FileLock(lock_path):
            with open(path, "a", encoding="utf-8-sig", newline="") as f:
                w = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
                if new_file:
                    w.writeheader()
                w.writerows(rows)
    except Exception as e:
        raise RuntimeError(f"CSV 写入失败 {path}: {e}")
    logger.info("CSV 追加 %d 行 -> %s", len(rows), path)


def write_local(cfg, reports):
    os.makedirs(cfg["out_dir"], exist_ok=True)
    for rtype, rep in reports.items():
        jpath = os.path.join(cfg["out_dir"], f"latest-{rtype}.json")
        try:
            with open(jpath, "w", encoding="utf-8") as f:
                json.dump(rep, f, ensure_ascii=False, indent=2)
            logger.info("JSON 写入 -> %s", jpath)
        except OSError as e:
            raise RuntimeError(f"JSON 写入失败 {jpath}: {e}")
        append_csv(cfg["out_dir"], rtype, rep)


# ---------------------------------------------------------------------------
# Supabase 推送（复用 sync-bill-mail.py 模式）
# ---------------------------------------------------------------------------

def _supabase_get(cfg):
    url = cfg["supabase_url"].rstrip("/") + "/rest/v1/pwb_state?id=eq.3&select=data"
    req = urllib.request.Request(url, headers={
        "apikey": cfg["anon_key"], "Authorization": "Bearer " + cfg["anon_key"],
    })
    with urllib.request.urlopen(req, timeout=20) as r:
        rows = json.loads(r.read().decode("utf-8"))
    return (rows[0]["data"] if rows and rows[0].get("data") else {}), r.status


def _supabase_post(cfg, data):
    body = {"id": 3, "data": data, "updated_at": datetime.now(timezone.utc).isoformat()}
    req = urllib.request.Request(
        cfg["supabase_url"].rstrip("/") + "/rest/v1/pwb_state",
        data=json.dumps(body, ensure_ascii=False).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "apikey": cfg["anon_key"],
            "Authorization": "Bearer " + cfg["anon_key"],
            "Prefer": "resolution=merge-duplicates",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as r:
        return r.status


def _push_supabase_locked(cfg, reports):
    """在持有文件锁的前提下执行「读旧状态 -> 合并 -> 回写整行」。"""
    state = {}
    try:
        state, _ = _supabase_get(cfg)
        logger.info("读取云端旧状态成功（已有 %d 个报表字段）", len(state or {}))
    except Exception as e:
        logger.warning("读取云端旧状态失败（将作为空状态覆盖）: %s", e)
    state = state or {}
    for rtype, rep in reports.items():
        state[rtype] = rep
    state["updatedAt"] = datetime.now(timezone.utc).isoformat()

    last_err = None
    for attempt in (1, 2):
        try:
            status = _supabase_post(cfg, state)
            if status in (200, 201, 204):
                logger.info("Supabase 推送成功 HTTP %s", status)
                return status
            last_err = RuntimeError(f"HTTP {status}")
        except Exception as e:
            last_err = e
        if attempt == 1:
            logger.warning("Supabase 推送失败（%s），3 秒后重试...", last_err)
            time.sleep(3)
    raise SupabaseError(f"Supabase 推送失败（重试后仍失败）: {last_err}")


def push_supabase(cfg, reports):
    """推送到 Supabase pwb_state id=3。

    必须整段加锁：日报/周报/月报三个计划任务在同一时刻（每天 10:00）触发，
    若「读旧状态 -> 合并 -> 回写整行」不加锁，后完成的进程会用自己读到的旧快照
    覆盖先完成进程的写入，导致其它报表字段被回滚（如 daily 停在前一天）。
    """
    lock_path = os.path.join(cfg["out_dir"], ".bill-sync.lock")
    try:
        with _FileLock(lock_path, timeout=60):
            return _push_supabase_locked(cfg, reports)
    except SupabaseError:
        raise
    except Exception as e:
        logger.warning("获取文件锁失败（%s），降级为无锁推送", e)
        return _push_supabase_locked(cfg, reports)


# ---------------------------------------------------------------------------
# 示例数据（--sample 自测，绕过 IMAP）
# ---------------------------------------------------------------------------

SAMPLE_CSV = """广告主名称,销售责任人,总花费(元)
推广账户A,张三,100.50
推广账户B,张三,200.00
推广账户C,李四,300.00
"""


def run_sample(cfg):
    today = datetime.now(BEIJING).date()
    reports = {}
    for rtype in ("daily", "weekly", "monthly"):
        if cfg["type_filter"] and rtype != cfg["type_filter"]:
            continue
        parsed = parse_csv_text(SAMPLE_CSV)
        rep = build_report(rtype, parsed, {
            "subject": f"【{TYPE_TITLE[rtype]}】测试报表",
            "sender": MAIL_SENDER_DEFAULT,
            "mailDate": today.strftime("%Y-%m-%d"),
            "source": "(内置示例)",
        })
        reports[rtype] = rep
    return reports


def report_biz_date(rtype, mail_date):
    """报表业务日期：日报=T-1（北京时间昨天），周报/月报=邮件日期。"""
    if rtype == "daily":
        return (datetime.now(BEIJING).date() - timedelta(days=1)).strftime("%Y-%m-%d")
    return mail_date


def build_report(rtype, parsed, meta):
    date = report_biz_date(rtype, meta["mailDate"])
    now = datetime.now(BEIJING).strftime("%Y-%m-%d %H:%M:%S")
    rep = {
        "title": TYPE_TITLE[rtype] + " · " + date,
        "date": date,
        "dateRange": calc_range(rtype, date),
        "total": parsed["total"],
        "count": parsed["count"],
        "bizs": parsed["bizs"],
        "updatedAt": now,
        "_meta": {
            "subject": meta.get("subject", ""),
            "sender": meta.get("sender", ""),
            "mailDate": meta["mailDate"],  # 保留邮件到达日期（CSV「邮件日期」列），报表日期另行计算
            "source": meta.get("source", ""),
            "parsedAt": now,
        },
    }
    return rep


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def build_argparser():
    p = argparse.ArgumentParser(
        description="独立运行：IMAP 读取日报/周报/月报邮件并汇总到本地 + Supabase",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="示例：\n"
               "  python read-bill-mail.py --user 2080572504@qq.com --password 16位授权码 --days 1\n"
               "  python read-bill-mail.py --type daily --days 1 --dry-run\n"
               "  python read-bill-mail.py --sample",
    )
    g = p.add_argument_group("邮箱服务器（默认 QQ 邮箱；均可由环境变量覆盖）")
    g.add_argument("--imap-host", help="IMAP 服务器（默认 imap.qq.com，环境变量 IMAP_HOST）")
    g.add_argument("--imap-port", type=int, help="IMAP 端口（默认 993，环境变量 IMAP_PORT）")
    sslg = g.add_mutually_exclusive_group()
    sslg.add_argument("--ssl", dest="ssl", action="store_true", help="使用 SSL（默认开启）")
    sslg.add_argument("--no-ssl", dest="no_ssl", action="store_true", help="不使用 SSL")
    g.add_argument("--user", help="邮箱账号（环境变量 IMAP_USER）")
    g.add_argument("--password", help="邮箱授权码/密码（不推荐命令行，优先 IMAP_PASS 环境变量）")
    g.add_argument("--password-file", help="从文件读取授权码（首行）")
    g.add_argument("--folder", help="收件箱文件夹（默认 INBOX，环境变量 IMAP_FOLDER）")
    g.add_argument("--sender", help="发件人白名单（逗号分隔，空串=不限；默认快手代理商小助手，环境变量 MAIL_SENDER）")
    g.add_argument("--subject-keywords", help="附加主题关键词（逗号分隔；默认识别 日报/周报/月报）")

    g2 = p.add_argument_group("筛选与日期范围")
    g2.add_argument("--type", choices=["daily", "weekly", "monthly", "all"], default="all", help="只处理某类报表（默认 all）")
    dg = g2.add_mutually_exclusive_group()
    dg.add_argument("--days", type=int, help="近 N 天（含今天；与 --start/--end 互斥）")
    dg.add_argument("--start", help="开始日期 YYYY-MM-DD")
    dg.add_argument("--end", help="结束日期 YYYY-MM-DD（默认今天）")

    g3 = p.add_argument_group("输出")
    g3.add_argument("--out-dir", help="输出目录（默认脚本目录下 bill_out，环境变量 OUT_DIR）")
    g3.add_argument("--no-supabase", action="store_true", help="不推送 Supabase，仅本地输出（默认推送）")
    g3.add_argument("--log-file", help="文件日志路径（默认 <out-dir>/bill-sync.log）")

    g4 = p.add_argument_group("运行模式")
    g4.add_argument("--verbose", action="store_true", help="输出 DEBUG 日志")
    g4.add_argument("--dry-run", action="store_true", help="只解析并打印结果，不写盘、不推送")
    g4.add_argument("--sample", action="store_true", help="用内置示例数据自测全链路（跳过 IMAP）")
    return p


def main(argv=None):
    args = build_argparser().parse_args(argv)
    load_env_file()
    try:
        cfg = load_config(args)
    except RuntimeError as e:
        print(f"配置错误: {e}", file=sys.stderr)
        return EXIT_RUNTIME

    setup_logging(cfg["log_file"], args.verbose)
    logger.info("========== 开始执行（类型=%s, 输出=%s%s）==========",
                args.type, cfg["out_dir"], ", 仅本地" if not cfg["supabase"] else ", 本地+Supabase")

    reports = {}
    try:
        if cfg["sample"]:
            logger.info("运行模式：内置示例数据（跳过 IMAP）")
            reports = run_sample(cfg)
        else:
            conn = imap_login(cfg)
            try:
                metas = search_mails(conn, cfg)
                logger.info("筛选到 %d 封候选邮件", len(metas))
                if not metas:
                    print("无匹配邮件：日期范围内没有主题/发件人符合的报表邮件", file=sys.stderr)
                    return EXIT_NO_MAIL
                ok_types = set()
                for meta in sorted(metas, key=lambda m: m["date"], reverse=True):
                    rtype = meta["type"]
                    if rtype in reports:  # 每类只保留最新一封
                        continue
                    logger.info("处理 uid=%s [%s] 主题=%s 日期=%s", meta["uid"], TYPE_TITLE[rtype], meta["subject"], meta["mail_date"])
                    try:
                        atts = fetch_attachments(conn, meta)
                        if not atts:
                            logger.warning("uid=%s 无附件且无正文，跳过", meta["uid"])
                            continue
                        last_parse_err = None
                        parsed = None
                        for att in atts:
                            try:
                                parsed = parse_attachment(att, rtype, meta)
                                src = att["name"]
                                break
                            except ParseError as e:
                                last_parse_err = e
                                logger.warning("uid=%s 附件 %s 解析失败: %s（尝试下一个）", meta["uid"], att["name"], e)
                        if parsed is None:
                            raise ParseError(f"全部附件解析失败: {last_parse_err}")
                        reports[rtype] = build_report(rtype, parsed, {
                            "subject": meta["subject"], "sender": meta["from"],
                            "mailDate": meta["mail_date"], "source": src,
                        })
                        ok_types.add(rtype)
                    except ParseError as e:
                        logger.error("uid=%s 解析失败: %s（继续处理其余邮件）", meta["uid"], e)
            finally:
                try:
                    conn.logout()
                except Exception:
                    pass
            if not reports:
                print("全部候选邮件解析失败（详见日志）", file=sys.stderr)
                return EXIT_PARSE
    except AuthError as e:
        logger.error("认证失败: %s", e)
        return EXIT_AUTH
    except NetworkError as e:
        logger.error("网络错误: %s", e)
        return EXIT_NETWORK
    except RuntimeError as e:
        logger.error("运行时错误: %s", e)
        return EXIT_RUNTIME

    # 输出
    summary = {"ok": True, "reports": {}}
    if cfg["dry_run"]:
        print("【DRY-RUN 预览，不写盘不推送】")
        for rtype, rep in sorted(reports.items()):
            print(f"  [{TYPE_TITLE[rtype]}] {rep['date']} 总消耗 ¥{rep['total']}  商务 {len(rep['bizs'])} 个  账户 {rep['count']} 个")
            for b in rep["bizs"]:
                print(f"      {b['biz']}: ¥{b['spent']}（{len(b['accounts'])} 个账户）")
        return EXIT_OK

    try:
        write_local(cfg, reports)
    except RuntimeError as e:
        logger.error("本地写入失败: %s", e)
        return EXIT_RUNTIME

    if cfg["supabase"]:
        try:
            push_supabase(cfg, reports)
        except SupabaseError as e:
            logger.error("%s", e)
            return EXIT_SUPABASE

    for rtype, rep in sorted(reports.items()):
        summary["reports"][rtype] = {
            "date": rep["date"], "dateRange": rep["dateRange"],
            "total": rep["total"], "count": rep["count"], "bizs": len(rep["bizs"]),
        }
        logger.info("[%s] %s 总消耗 ¥%s 商务 %d 个 账户 %d 个 -> latest-%s.json + bill-%s.csv",
                    TYPE_TITLE[rtype], rep["date"], rep["total"], len(rep["bizs"]), rep["count"], rtype, rtype)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    logger.info("========== 执行完成 ==========")
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())
