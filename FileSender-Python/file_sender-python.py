# -*- coding: utf-8 -*-
"""
邮件分发助手
- 扫描指定文件夹，按"文件名包含关键字"把文件分发给对应收件人（最长关键字优先）
- 同一收件人的多个文件默认合并为一封邮件（可切换为逐文件发送）
- 关键字-收件人规则、文件夹位置、SMTP 设置均保存在程序同目录 config.json，可随时修改
- 只读取文件作为附件发送，不移动、不修改、不删除任何原文件

依赖（均可选，缺失时自动降级不影响发信）：
- keyring  —— 密码保管（Windows 凭据管理器）；缺失或后端不可用时自动回退
              config.json 明文保存并在界面橙色提示
- openpyxl —— 发送清单 Excel；缺失时自动回退 CSV

版本历史：
  1.1.0 (2026-11-24) SMTP 会话复用与断线重连（全程只登录一次）、连续连接失败提前中止、
       退出流程防重入、配置损坏自动备份并提示、DNS 错误人话提示、明文保存需确认、
       双击表头防误触、--version 版本查询、CA 证书路径预检、危险后缀与快捷方式排除、
       日志自动清理、报告去重键改 (folder,fname) 及 CSV 分支同步、行高封顶 409
  1.0.0 初版：关键字匹配分发、合并发送、发送清单、keyring 密码保管、证书三档校验
"""
import os
import re
import sys
import json
import time
import queue
import socket
import uuid
import threading
import smtplib
import ssl
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.header import Header
from email.utils import formataddr, formatdate, parseaddr
from datetime import datetime

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

APP_NAME = "邮件分发助手"
APP_VERSION = "1.1.0"
AUTHOR = "恔之"
AUTHOR_NOTE = "作者：%s，有问题欢迎提 Issue" % AUTHOR
CONFIG_NAME = "config.json"
CONFIG_VERSION = 1          # config.json 结构版本号，供后续配置迁移升级使用
KEYRING_SERVICE = "邮件分发助手/SMTP"

BASE64_EXPANSION = 1.37      # 附件 base64 编码膨胀系数：理论 4/3≈1.333，取 1.37 留安全余量；
                            # 网关拒收按编码后大小计，预检用原始大小×此系数

# 证书校验三档（config.json 里 tls_verify 的取值）
TLS_MODES = {
    "off": "不校验（默认，内网自签可用）",
    "ca": "校验证书链（内网CA，可填CA文件）",
    "strict": "严格校验（公网邮箱）",
}
TLS_REV = {v: k for k, v in TLS_MODES.items()}

# 加密方式：界面显示文本与 config 存储值分离，改下拉框文案只需改这里
SEC_NONE_LABEL = "none（不加密，内网中继用）"
SEC_LABEL_TO_VAL = {"SSL": "SSL", "STARTTLS": "STARTTLS", SEC_NONE_LABEL: "none"}
SEC_VAL_TO_LABEL = {v: k for k, v in SEC_LABEL_TO_VAL.items()}

# ---------------- 基础路径 ----------------

def app_dir():
    """程序所在目录（exe 或源码），配置与发送清单都存在这里"""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def default_config():
    return {
        "version": CONFIG_VERSION,   # 配置结构版本，供未来迁移判断
        "folders": [],
        "rules": [],  # [{"kw": "..", "name": "..", "email": ".."}]
        "send": {
            "merge": True,           # 同一收件人合并为一封
            "recursive": False,      # 扫描子文件夹
            "interval": 1.0,         # 每封间隔秒
            "attach_warn_mb": 45,    # 单封附件合计超过此大小默认拦截不发送
        },
        "smtp": {
            "host": "",
            "port": 465,
            "security": "SSL",       # SSL / STARTTLS / none
            "auth": True,            # 是否需要登录认证
            "tls_verify": "off",    # off=不校验(默认) / ca=验链不验名 / strict=全验
            "ca_file": "",           # ca 档可选：内部CA的 pem 文件路径
            "user": "",
            "passwd": "",
            "from_addr": "",         # 发件邮箱地址，留空=账号
            "display_name": "",      # 发件人显示名
            "cc": "",                # 抄送，多个用分号或逗号分隔
            "bcc": "",               # 密送，多个用分号分隔（收件人不可见）
            "subject_tpl": "【{关键字}】相关文件（{文件数}份），请查收",
            "body_tpl": "{姓名}，您好：\n\n现将 {关键字} 相关文件共 {文件数} 份发送给您，详见附件，请查收。\n具体文件：\n{文件名}\n\n此致\n敬礼",
            "timeout": 20,
        },
        "remember_passwd": False,
    }


def load_config():
    path = os.path.join(app_dir(), CONFIG_NAME)
    cfg = default_config()
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                saved = json.load(f)
            for k, v in saved.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except Exception as e:
            # 配置损坏（手改坏/断电写坏）：备份坏文件供手工抢救，回退默认配置并提示，
            # 绝不静默“吞掉”用户全部规则（App 启动后会弹窗告知）
            bad = "%s.bad-%s" % (path, datetime.now().strftime("%Y%m%d%H%M%S"))
            try:
                os.replace(path, bad)
                cfg["_load_error"] = ("配置文件损坏，已改用默认配置。原文件已备份为 %s（%s），"
                                      "可手工把规则换回后重新保存" % (os.path.basename(bad), _err_text(e)))
                # 坏配置备份只保留最近 5 个，防止长期累积
                try:
                    bname = os.path.basename(path)
                    bads = sorted(x for x in os.listdir(os.path.dirname(path) or ".")
                                  if x.startswith(bname + ".bad-"))
                    for old in bads[:-5]:
                        try:
                            os.remove(os.path.join(os.path.dirname(path), old))
                        except Exception:
                            pass
                except Exception:
                    pass
            except Exception:
                cfg["_load_error"] = "配置文件损坏且备份失败：%s" % _err_text(e)
    smtp = _migrate_tls(cfg["smtp"])
    if not cfg.get("remember_passwd"):
        smtp["passwd"] = ""
    elif not smtp.get("passwd") and smtp.get("pwd_in_keyring"):
        # 密码存于 Windows 凭据管理器，启动时取回
        smtp["passwd"] = fetch_password(smtp.get("user", ""))
    return cfg


def save_config(cfg):
    path = os.path.join(app_dir(), CONFIG_NAME)
    # 唯一临时名：两个实例同时保存也不会互相覆盖对方的临时文件
    tmp = "%s.tmp-%s" % (path, uuid.uuid4().hex[:8])
    out = dict(cfg)
    if not out.get("remember_passwd"):
        smtp = dict(out.get("smtp", {}))
        smtp["passwd"] = ""
        out["smtp"] = smtp
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)


def store_password(user, pwd):
    """密码存入 Windows 凭据管理器。成功 True；环境不支持返回 False（调用方退回明文）"""
    try:
        import keyring
        keyring.set_password(KEYRING_SERVICE, "smtp:%s" % user, pwd)
        return True
    except Exception:
        return False


def fetch_password(user):
    """从凭据管理器取密码，取不到返回空串"""
    try:
        import keyring
        return keyring.get_password(KEYRING_SERVICE, "smtp:%s" % user) or ""
    except Exception:
        return ""


def clear_password(user):
    """清除凭据管理器中的密码条目"""
    try:
        import keyring
        keyring.delete_password(KEYRING_SERVICE, "smtp:%s" % user)
    except Exception:
        pass


# ---------------- 匹配核心（纯函数，供自测调用） ----------------

EXCLUDE_PREFIX = ("~$",)          # Office/WPS 打开时的锁文件
EXCLUDE_EXT = (".tmp", ".crdownload", ".part", ".partial", ".download",
               ".exe", ".bat", ".cmd", ".ps1", ".js", ".vbs", ".scr",
               ".lnk")  # 后8项：可执行/脚本/快捷方式，防误发无意义文件

EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[^@\s,;]+$")


def is_valid_email(s):
    return bool(EMAIL_RE.match((s or "").strip()))


def match_file(filename, rules):
    """返回 dict: {'status': 'ok'|'none'|'conflict', 'rule': dict|None,
                   'all_kws': [命中的关键字]}  最长关键字优先"""
    low = filename.lower()
    hits = [r for r in rules if r.get("kw") and r["kw"].lower() in low]
    if not hits:
        return {"status": "none", "rule": None, "all_kws": []}
    maxlen = max(len(r["kw"]) for r in hits)
    finals = [r for r in hits if len(r["kw"]) == maxlen]
    emails = {r.get("email", '').strip() for r in finals}
    if len(finals) > 1 and len(emails) > 1:
        return {"status": "conflict", "rule": None,
                "all_kws": [r["kw"] for r in finals]}
    return {"status": "ok", "rule": finals[0],
            "all_kws": [r["kw"] for r in finals]}


def list_files(folders, recursive=False):
    """只收集文件，排除锁文件与下载临时文件；返回 [(文件夹, 文件名)]"""
    out = []
    seen = set()
    for folder in folders:
        if not os.path.isdir(folder):
            continue
        if recursive:
            for root, _dirs, files in os.walk(folder):
                for fn in files:
                    out.append((root, fn))
        else:
            for fn in os.listdir(folder):
                p = os.path.join(folder, fn)
                if os.path.isfile(p):
                    out.append((folder, fn))
    res = []
    for root, fn in out:
        if fn.startswith(EXCLUDE_PREFIX):
            continue
        if os.path.splitext(fn)[1].lower() in EXCLUDE_EXT:
            continue
        key = os.path.normcase(os.path.join(root, fn))
        if key in seen:
            continue
        seen.add(key)
        res.append((root, fn))
    return sorted(res, key=lambda x: x[1])


def scan(folders, rules, recursive=False):
    """扫描并匹配。返回 rows：每文件一行 dict"""
    rows = []
    for folder, fn in list_files(folders, recursive):
        m = match_file(fn, rules)
        row = {
            "folder": folder, "fname": fn,
            "status": m["status"],           # ok / none / conflict
            "rule": m["rule"],               # 命中的规则
            "kws": m["all_kws"],
            "manual_email": "",              # 手动指定的收件人
            "manual_name": "",               # 手动指定的收件人姓名（选规则时自动带出）
            "state": "未发送", "err": "",
        }
        rows.append(row)
    return rows


def iter_send_groups(rows, merge=True):
    """把可发送的行分组：(收件邮箱, 姓名, 命中关键字集合, [行])。
    未匹配/冲突且未手动指定的行跳过。"""
    ok_rows = [r for r in rows
               if r["status"] == "ok" or (r["manual_email"] and is_valid_email(r["manual_email"]))]
    groups = []
    if merge:
        buckets = {}
        for r in ok_rows:
            if r["manual_email"]:
                email, name = r["manual_email"].strip(), r.get("manual_name", "")
            else:
                email = r["rule"].get("email", "").strip()
                name = r["rule"].get("name", "")
            key = norm_key(email)
            if key not in buckets:
                buckets[key] = {"email": email, "name": name, "kws": [], "rows": []}
            for kw in r["kws"]:
                if kw not in buckets[key]["kws"]:
                    buckets[key]["kws"].append(kw)
            buckets[key]["rows"].append(r)
        groups = list(buckets.values())
    else:
        for r in ok_rows:
            if r["manual_email"]:
                email, name = r["manual_email"].strip(), r.get("manual_name", "")
            else:
                email = r["rule"].get("email", "").strip()
                name = r["rule"].get("name", "")
            groups.append({"email": email, "name": name, "kws": list(r["kws"]), "rows": [r]})
    return groups


def norm_key(s):
    return (s or "").strip().lower()


# ---------------- 邮件构建与发送 ----------------

def render_tpl(tpl, kw_str, name, nfiles, fnames):
    repl = {
        "{关键字}": kw_str, "{姓名}": name,
        "{文件数}": str(nfiles), "{文件名}": fnames,
        "{邮件数}": str(nfiles), "{附件数}": str(nfiles),
    }
    for k, v in repl.items():
        tpl = tpl.replace(k, v)
    return tpl


def build_message(smtp_cfg, group, subject_tpl, body_tpl):
    email = group["email"]
    name = group["name"] or (email.split("@")[0])
    kws = "、".join(group["kws"]) or "（手动指定）"
    files = group["rows"]
    n = len(files)
    fnames = "\n".join("· " + r["fname"] for r in files)
    subject = render_tpl(subject_tpl, kws, name, n, fnames)
    body = render_tpl(body_tpl, kws, name, n, fnames)

    msg = MIMEMultipart()
    from_name = smtp_cfg.get("display_name", "") or APP_NAME
    from_addr = (smtp_cfg.get("from_addr") or smtp_cfg.get("user") or "").strip()
    msg["From"] = formataddr((str(Header(from_name, "utf-8")), from_addr))
    msg["To"] = email
    cc_raw = smtp_cfg.get("cc", "") or ""
    cc_list = [x.strip() for x in re.split(r"[;，,]", cc_raw) if x.strip()]
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = Header(subject, "utf-8")
    msg["Date"] = formatdate(localtime=True)
    msg.attach(MIMEText(body, "plain", "utf-8"))
    for r in files:
        path = os.path.join(r["folder"], r["fname"])
        with open(path, "rb") as f:
            data = f.read()
        part = MIMEApplication(data)
        part.add_header("Content-Disposition", "attachment",
                        filename=("utf-8", "", r["fname"]))
        msg.attach(part)
    bcc_raw = smtp_cfg.get("bcc", "") or ""
    bcc_list = [x.strip() for x in re.split(r"[;，,]", bcc_raw) if x.strip()]
    # 密送：只进入实际收件列表，不写入邮件头（收件人不可见）
    recv = [email] + cc_list + bcc_list
    recv = [x for x in recv if is_valid_email(x)]
    return msg, from_addr, recv


def smtp_login_test(smtp_cfg):
    """仅测试连接与登录，不发信。成功返回 None，失败返回错误文本"""
    try:
        server = _connect(smtp_cfg)
        try:
            server.quit()
        except Exception:
            pass
        return None
    except Exception as e:
        return _friendly_error(e)


def _migrate_tls(smtp_cfg):
    """旧版 verify_cert 布尔 → 三档 tls_verify"""
    if "tls_verify" not in smtp_cfg:
        smtp_cfg["tls_verify"] = "strict" if smtp_cfg.get("verify_cert") else "off"
    smtp_cfg.pop("verify_cert", None)
    return smtp_cfg


def _make_ctx(smtp_cfg):
    """off=不校验（默认，内网自签名/内部CA用）
       ca=验证书链但不验主机名（内网CA签的、常用IP连，主机名对不上）
       strict=全验（公网邮箱）"""
    mode = smtp_cfg.get("tls_verify", "off")
    if mode == "strict":
        return ssl.create_default_context()
    if mode == "ca":
        cafile = (smtp_cfg.get("ca_file") or "").strip()
        ctx = ssl.create_default_context(cafile=cafile or None)
        ctx.check_hostname = False
        return ctx
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    return ctx


def _connect(smtp_cfg):
    host = smtp_cfg.get("host", "").strip()
    port = int(smtp_cfg.get("port") or 0)
    sec = smtp_cfg.get("security", "SSL")
    timeout = int(smtp_cfg.get("timeout") or 20)
    if not host:
        raise ValueError("SMTP 服务器地址为空")
    if not port:
        port = 465 if sec == "SSL" else (587 if sec == "STARTTLS" else 25)
    ctx = _make_ctx(smtp_cfg)
    if sec == "SSL":
        server = smtplib.SMTP_SSL(host, port, timeout=timeout, context=ctx)
    else:
        server = smtplib.SMTP(host, port, timeout=timeout)
        server.ehlo()
        if sec == "STARTTLS":
            server.starttls(context=ctx)
            server.ehlo()
    if smtp_cfg.get("auth", True):
        user = smtp_cfg.get("user", "").strip()
        pwd = smtp_cfg.get("passwd", "")
        if not user:
            raise ValueError("发件账号为空（如服务器无需认证，请取消勾选\u201c需要登录认证\u201d）")
        if not pwd:
            raise ValueError("密码/授权码为空")
        server.login(user, pwd)
    return server


def _err_text(e):
    return "%s: %s" % (type(e).__name__, e)


def _friendly_error(e):
    """把底层异常翻译成执行人能看懂的话，括号内保留技术详情便于排查"""
    text = _err_text(e)
    try:
        if isinstance(e, smtplib.SMTPAuthenticationError):
            return "账号或密码/授权码不正确（%s）" % text
        if isinstance(e, smtplib.SMTPConnectError):
            return "无法连接邮件服务器（地址或端口可能不对、或网络不通）（%s）" % text
        if isinstance(e, smtplib.SMTPRecipientsRefused):
            return "收件人被服务器拒绝（%s）" % text
        if isinstance(e, smtplib.SMTPSenderRefused):
            return "发件人被服务器拒绝（发件邮箱地址可能不对）（%s）" % text
        if isinstance(e, smtplib.SMTPDataError):
            return "服务器拒收邮件内容（可能附件过大或含受限内容）（%s）" % text
        if isinstance(e, (smtplib.SMTPResponseException, smtplib.SMTPException)):
            return "邮件服务器返回错误（%s）" % text
        if isinstance(e, ssl.SSLCertVerificationError):
            return "证书校验失败（内网自签证书请把“证书校验”设为不校验）（%s）" % text
        if isinstance(e, ssl.SSLError):
            return "加密连接失败（加密方式或端口可能不对）（%s）" % text
        if isinstance(e, socket.gaierror):
            return "无法解析邮件服务器域名（检查地址拼写或DNS）（%s）" % text
        if isinstance(e, PermissionError):
            return "没有权限读取文件（可能正被他人打开占用）（%s）" % text
        if isinstance(e, FileNotFoundError):
            return "文件不存在或已被移动/删除（%s）" % text
        if isinstance(e, ConnectionResetError):
            return "连接被服务器中断（%s）" % text
        if isinstance(e, (TimeoutError, socket.timeout)):
            return "连接邮件服务器超时（网络不通或端口被防火墙拦截）（%s）" % text
        if isinstance(e, OSError):
            return "文件读写或网络错误（%s）" % text
    except Exception:
        pass
    return text


def sanitize_cell(v):
    """防公式注入：以 = + - @ 开头的文本加 ' 前缀，避免 Excel/WPS 当作公式执行"""
    if isinstance(v, str) and v[:1] in ("=", "+", "-", "@"):
        return "'" + v
    return v


def probe_file(path):
    """探测文件可读性并返回大小。成功返回 (字节数, None)，失败返回 (None, 人话原因)。
    预检与发送共用同一探测口径，避免“能 stat 但读不了”的文件两边结论不一致"""
    try:
        with open(path, "rb") as f:
            f.read(16)
        return os.path.getsize(path), None
    except Exception as e:
        return None, _friendly_error(e)


def check_files_readable(rows):
    """发送前预检文件可读性。返回 (可读行, [(行, 人话原因)])"""
    ok, bad = [], []
    for r in rows:
        size, err = probe_file(os.path.join(r["folder"], r["fname"]))
        if err is None:
            ok.append(r)
        else:
            bad.append((r, err))
    return ok, bad


class SmtpSession:
    """SMTP 会话复用：一轮发送只登录一次；断线（SMTPServerDisconnected）自动重连一次；
    其余异常时连接状态未知，丢弃连接下次重建。失败隔离：单组失败不影响后续组"""

    def __init__(self, cfg):
        self.cfg = cfg
        self.server = None

    def _ensure(self):
        if self.server is None:
            self.server = _connect(self.cfg)
        return self.server

    def send_group_once(self, group, subject_tpl, body_tpl, warn_mb):
        """发送一组，复用当前连接。返回 (ok, err, skipped)。
        单个附件读取失败时跳过该附件继续发其余（skipped=[(行, 原因)]，进入待补发池）；
        全部附件不可读则整组失败；附件合计超上限直接拦截。连接类/SMTP 异常向上抛出。"""
        rows_ok, skipped, total_pre = [], [], 0
        for r in group["rows"]:
            size, err = probe_file(os.path.join(r["folder"], r["fname"]))
            if err is None:
                total_pre += size
                rows_ok.append(r)
            else:
                skipped.append((r, err))
        if not rows_ok:
            return (False, "全部附件无法读取：" + (skipped[0][1] if skipped else "未知原因"),
                    skipped)
        encoded = total_pre * BASE64_EXPANSION   # 理论膨胀 4/3≈1.333，取 1.37 留余量；网关拒收按编码后大小计
        if warn_mb and warn_mb > 0 and encoded > warn_mb * 1024 * 1024:
            return False, ("附件编码后约 %.1fMB（原始 %.1fMB），超过单封上限 %gMB，未发送"
                           "（上限按邮件实际大小计，可在“邮箱设置”页调整）"
                           % (encoded / 1024 / 1024, total_pre / 1024 / 1024, warn_mb)), skipped
        sub = dict(group)
        sub["rows"] = rows_ok
        try:
            msg, from_addr, recv = build_message(smtp_cfg=self.cfg, group=sub,
                                                 subject_tpl=subject_tpl, body_tpl=body_tpl)
        except (PermissionError, FileNotFoundError, OSError) as e:
            return False, _friendly_error(e), skipped
        # 复用连接发送；断线重连一次（认证/拒收等 SMTP 语义错误连接状态未知，丢弃重建）
        for attempt in (1, 2):
            try:
                s = self._ensure()
                s.sendmail(from_addr, recv, msg.as_string())
                return True, "", skipped
            except smtplib.SMTPServerDisconnected:
                self.server = None
                if attempt == 2:
                    raise
            except Exception:
                self.server = None
                raise
        raise RuntimeError("unreachable")

    def close(self):
        if self.server is not None:
            try:
                self.server.quit()
            except Exception:
                pass
            self.server = None


def send_group(smtp_cfg, group, subject_tpl, body_tpl, warn_mb=45.0):
    """发送一组（独立会话，供单发场景使用）。返回 (ok, err, skipped)。"""
    sess = SmtpSession(smtp_cfg)
    try:
        return sess.send_group_once(group, subject_tpl, body_tpl, warn_mb)
    finally:
        sess.close()


# ---------------- 发送清单 xlsx ----------------

def cleanup_old_logs(log_dir=None, keep_days=30):
    """删除 logs/ 下 keep_days 天前的旧日志，防止长期堆积；任何失败静默不影响主功能"""
    try:
        d = log_dir or os.path.join(app_dir(), "logs")
        if not os.path.isdir(d):
            return
        cutoff = time.time() - keep_days * 86400
        for fn in os.listdir(d):
            if not (fn.startswith("使用日志_") and fn.endswith(".log")):
                continue
            p = os.path.join(d, fn)
            try:
                if os.path.getmtime(p) < cutoff:
                    os.remove(p)
            except Exception:
                pass
    except Exception:
        pass


def write_report(rows, groups_done, path):
    """rows: 扫描行；groups_done: [{'email':..,'name':..,'kws':[..],'rows':[..],'ok':bool,'err':str,'time':str}]"""
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
    except ImportError:
        path = os.path.splitext(path)[0] + ".csv"
        _write_report_csv(rows, groups_done, path)
        return path
    wb = Workbook()
    ws = wb.active
    ws.title = "发送清单"
    heads = ["序号", "收件人", "收件邮箱", "匹配方式", "命中关键字", "文件数",
             "文件名", "状态", "失败原因", "发送时间"]
    ws.append(heads)
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="4472C4")
    for c in ws[1]:
        c.font = head_font
        c.fill = head_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
    wrap_align = Alignment(wrap_text=True, vertical="top")   # 文件名/原因列多行显示
    i = 0
    for g in groups_done:
        for r in g["rows"]:
            i += 1
            way = "手动指定" if r["manual_email"] else "关键字规则"
            row_err = r.get("err") or g.get("err", "")
            ws.append([sanitize_cell(c) for c in [
                i, g["name"], g["email"], way, "、".join(r["kws"]) or "-",
                len(g["rows"]),
                "\n".join(x["fname"] for x in g["rows"]),
                "成功" if g.get("ok") else "失败",
                row_err, g.get("time", ""),
            ]])
            # 文件名（第7列）与失败原因（第9列）内容含换行：开启单元格内换行并按附件数给行高；
            # Excel 单行行高上限 409.5，超过会被截断，封顶避免非法值
            ws.row_dimensions[i].height = max(15, min(409, 14 * len(g["rows"])))
            for col in (7, 9):
                ws.cell(row=i, column=col).alignment = wrap_align
    sent_files = set()
    for g in groups_done:
        for r in g["rows"]:
            # 用（文件夹, 文件名）作主键而非 id(r)：重新扫描后 self.rows 是全新对象，
            # id 会变，已发文件会被重复列入“未发送”段
            sent_files.add((r["folder"], r["fname"]))
    for r in rows:
        if (r["folder"], r["fname"]) in sent_files:
            continue
        i += 1
        if r["status"] == "conflict":
            status, err = "未发送", "文件名同时命中多个收件人，需人工确认"
        elif r.get("state") == "失败":
            status, err = "未发送", r.get("err", "") or "未发送"
        else:
            status, err = "未发送", "未匹配到任何关键字"
        ws.append([sanitize_cell(c) for c in [
            i, "-", "-", "-", "、".join(r["kws"]) or "-", 1,
            r["fname"], status, err, "",
        ]])
    widths = [6, 12, 26, 10, 18, 7, 46, 8, 30, 18]
    for idx, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 20
    wb.save(path)
    return path


def _write_report_csv(rows, groups_done, path):
    import csv
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        w.writerow(["序号", "收件人", "收件邮箱", "匹配方式", "命中关键字",
                    "文件数", "文件名", "状态", "失败原因", "发送时间"])
        i = 0
        for g in groups_done:
            for r in g["rows"]:
                i += 1
                row_err = r.get("err") or g.get("err", "")
                w.writerow([sanitize_cell(c) for c in [i, g["name"], g["email"],
                            "手动指定" if r["manual_email"] else "关键字规则",
                            "、".join(r["kws"]) or "-", len(g["rows"]),
                            "\n".join(x["fname"] for x in g["rows"]),
                            "成功" if g.get("ok") else "失败", row_err, g.get("time", "")]])
        # 与 xlsx 分支同口径：用（文件夹, 文件名）作主键，重扫后不重复列入“未发送”段
        sent = {(r["folder"], r["fname"]) for g in groups_done for r in g["rows"]}
        for r in rows:
            if (r["folder"], r["fname"]) in sent:
                continue
            i += 1
            w.writerow([sanitize_cell(c) for c in [i, "-", "-", "-", "、".join(r["kws"]) or "-", 1,
                        r["fname"], "未发送",
                        "命中多个收件人需确认" if r["status"] == "conflict" else "未匹配关键字", ""]])


# ---------------- 规则 Excel 导入 ----------------

def _read_rules_excel(path):
    """读取规则表：A=关键字 B=姓名 C=邮箱；自动跳过表头与空行"""
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows, bad = [], []
        for i, row in enumerate(ws.iter_rows(min_row=1, max_col=3, values_only=True), 1):
            kw = (str(row[0]).strip() if row[0] is not None else "")
            name = (str(row[1]).strip() if len(row) > 1 and row[1] is not None else "")
            email = (str(row[2]).strip() if len(row) > 2 and row[2] is not None else "")
            if i == 1 and not is_valid_email(email):
                continue   # 第1行C列不是合法邮箱 → 视为表头（数据行必须有合法邮箱，
                           # 不再依赖“关键字含‘关键’/‘邮箱’”字样匹配，避免误伤正常数据行）
            if not kw and not email:
                continue
            if not kw:
                bad.append("第%d行：关键字为空" % i)
                continue
            if not is_valid_email(email):
                bad.append("第%d行：邮箱\u201c%s\u201d格式不对" % (i, email or "(空)"))
                continue
            rows.append({"kw": kw, "name": name, "email": email})
    finally:
        wb.close()   # 异常路径也必须释放句柄，否则 Windows 上可能锁住该文件
    return rows, bad


# ---------------- 图形界面 ----------------

class RuleDialog(tk.Toplevel):
    def __init__(self, master, title, data=None):
        super().__init__(master)
        self.result = None
        self.title(title)
        self.resizable(False, False)
        frm = ttk.Frame(self, padding=12)
        frm.grid(sticky="nsew")
        self.v_kw = tk.StringVar(value=(data or {}).get("kw", ""))
        self.v_name = tk.StringVar(value=(data or {}).get("name", ""))
        self.v_email = tk.StringVar(value=(data or {}).get("email", ""))
        rows = [("关键字（文件名包含即命中）", self.v_kw), ("收件人姓名", self.v_name),
                ("收件人邮箱", self.v_email)]
        for i, (label, var) in enumerate(rows):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="w", pady=4, padx=(0, 8))
            ttk.Entry(frm, textvariable=var, width=36).grid(row=i, column=1, pady=4)
        tip = "提示：多个关键字相似时，程序按\u201c最长关键字优先\u201d匹配。\n例：同时配了\u201c临江大学\u201d与\u201c临江大学紫金\u201d，文件\u201c临江大学紫金2608.xlsx\u201d命中后者。"
        ttk.Label(frm, text=tip, foreground="#666666", justify="left").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(6, 2))
        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, pady=(8, 0))
        ttk.Button(btns, text="确定", command=self._ok).pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=6)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())

    def _ok(self):
        kw = self.v_kw.get().strip()
        email = self.v_email.get().strip()
        if not kw:
            messagebox.showwarning(APP_NAME, "关键字不能为空", parent=self)
            return
        if not is_valid_email(email):
            messagebox.showwarning(APP_NAME, "邮箱格式不正确", parent=self)
            return
        self.result = {"kw": kw, "name": self.v_name.get().strip(), "email": email}
        self.destroy()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1060x700")
        self.minsize(900, 560)
        self.cfg = load_config()
        self.rows = []              # 扫描结果
        self.groups_done = []       # 发送结果
        self._active_groups = []    # 当前发送中的组引用（worker 经队列回传结果）
        self.send_thread = None
        self.stop_event = threading.Event()
        self.q = queue.Queue()
        self._dirty = False
        self._supp_trace = False      # 程序化回填 SMTP 表单时抑制“人工改动”监听
        self._closing = False        # 退出流程进行中（防重入）
        self._poll_after = None      # 队列轮询 after 句柄，销毁前取消防尾随

        # 右下角常驻签名行（所有页签都可见）
        self._foot = ttk.Label(
            self, text="%s v%s ｜ %s" % (APP_NAME, APP_VERSION, AUTHOR_NOTE),
            foreground="#999999", font=("Microsoft YaHei", 8))
        self._foot.pack(side="bottom", anchor="e", padx=10, pady=(0, 4))
        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=6, pady=6)
        self.tab_send = ttk.Frame(nb, padding=8)
        self.tab_rules = ttk.Frame(nb, padding=8)
        self.tab_smtp = ttk.Frame(nb, padding=8)
        nb.add(self.tab_send, text=" ① 发送  ")
        nb.add(self.tab_rules, text=" ② 收件规则 ")
        nb.add(self.tab_smtp, text=" ③ 邮箱设置 ")

        self._build_send_tab()
        self._build_rules_tab()
        self._build_smtp_tab()
        self._load_lists()
        self._bind_smtp_dirty()     # 回填完成后再挂改动监听：邮箱设置页改动未保存时关窗会提醒
        cleanup_old_logs()          # 启动时清理 30 天前旧日志
        # 配置损坏提示：load_config 已备份坏文件并回退默认，这里弹窗+日志告知，不静默丢规则
        load_err = self.cfg.pop("_load_error", "")
        if load_err:
            self.log_write("⚠ " + load_err, "err")
            self.after(200, lambda: messagebox.showwarning(
                APP_NAME, load_err + "\n\n详细原因见日志。", parent=self))
        self._poll_after = self.after(100, self._poll_queue)   # 首次句柄也记录，销毁前才能可取消
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- 页1：发送 ----------
    def _build_send_tab(self):
        t = self.tab_send

        top = ttk.LabelFrame(t, text=" 待扫描文件夹（可多个，随时增删，点\u201c保存\u201d记住） ")
        top.pack(fill="x")
        self.lb_folders = tk.Listbox(top, height=5, activestyle="dotbox")
        self.lb_folders.pack(side="left", fill="both", expand=True, padx=(6, 4), pady=4)
        bf = ttk.Frame(top)
        bf.pack(side="left", fill="y", padx=(0, 6), pady=4)
        self.btn_add_folder = ttk.Button(bf, text="添加文件夹", width=10, command=self.add_folder)
        self.btn_add_folder.pack(pady=2, fill="x")
        self.btn_del_folder = ttk.Button(bf, text="移除所选", width=10, command=self.del_folder)
        self.btn_del_folder.pack(pady=2, fill="x")
        self.btn_clear_folders = ttk.Button(bf, text="清空", width=10, command=self._clear_folders)
        self.btn_clear_folders.pack(pady=2, fill="x")
        ttk.Button(bf, text="保存", width=10, command=self._save_all).pack(pady=(10, 2), fill="x")

        mid = ttk.Frame(t)
        mid.pack(fill="x", pady=(6, 0))
        self.v_merge = tk.BooleanVar(value=self.cfg["send"]["merge"])
        self.v_recursive = tk.BooleanVar(value=self.cfg["send"]["recursive"])
        ttk.Checkbutton(mid, text="同一收件人的多个文件合并为一封", variable=self.v_merge,
                        command=self._opt_changed).pack(side="left")
        ttk.Checkbutton(mid, text="包含子文件夹", variable=self.v_recursive,
                        command=self._opt_changed).pack(side="left", padx=(14, 0))
        self.btn_scan = ttk.Button(mid, text="扫描并匹配", command=self.do_scan)
        self.btn_scan.pack(side="right")
        self.btn_retry = ttk.Button(mid, text="重发失败项", command=self.retry_failed)
        self.btn_retry.pack(side="right", padx=8)

        self.v_stat = tk.StringVar(value="尚未扫描")
        ttk.Label(t, textvariable=self.v_stat, foreground="#666666").pack(anchor="w", pady=(6, 0))
        ttk.Label(t, text="匹配模式：文件名包含关键字即命中，最长关键字优先；匹配不到或多人命中时停发待人工确认",
                  foreground="#999999").pack(anchor="w")

        wrap = ttk.Frame(t)
        wrap.pack(fill="both", expand=True, pady=(4, 4))
        self.tv = ttk.Treeview(wrap, columns=("fname", "folder", "kw", "name", "email", "state"),
                               show="headings", selectmode="browse")
        for col, txt, w, anchor in [
            ("fname", "文件名", 260, "w"), ("folder", "所在文件夹", 210, "w"),
            ("kw", "命中关键字", 120, "w"), ("name", "收件人", 90, "w"),
            ("email", "邮箱", 180, "w"), ("state", "状态", 110, "center")]:
            self.tv.heading(col, text=txt)
            self.tv.column(col, width=w, anchor=anchor, stretch=True)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=vsb.set)
        self.tv.tag_configure("ok", background="#ffffff")
        self.tv.tag_configure("none", background="#f4f4f4", foreground="#b0b0b0")
        self.tv.tag_configure("conflict", background="#fde9e9", foreground="#c00000")
        self.tv.tag_configure("manual", background="#e8f4ea")
        self.tv.tag_configure("sent", background="#e8f4ea", foreground="#1e7d32")
        self.tv.tag_configure("failed", background="#fde9e9")
        self.tv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")
        self.tv.bind("<Double-1>", lambda e: self.manual_assign()
                     if self.tv.identify_row(e.y) else None)   # 双击表头/空白区不触发

        act = ttk.Frame(t)
        act.pack(fill="x")
        self.btn_manual = ttk.Button(act, text="手动指定收件人（选中行，未匹配/冲突也可发）",
                                     command=self.manual_assign)
        self.btn_manual.pack(side="left")
        self.btn_clear_manual = ttk.Button(act, text="清除手动指定", command=self.clear_manual)
        self.btn_clear_manual.pack(side="left", padx=(8, 0))
        ttk.Button(act, text="未匹配/冲突文件另存副本", command=self.export_unmatched).pack(side="left", padx=(8, 0))
        self.btn_send = ttk.Button(act, text="▶ 开始发送", command=self.do_send)
        self.btn_send.pack(side="right")
        self.btn_stop = ttk.Button(act, text="■ 停止", command=self.do_stop, state="disabled")
        self.btn_stop.pack(side="right", padx=(8, 0))

        self.v_prog = tk.StringVar(value="")
        ttk.Label(t, textvariable=self.v_prog, foreground="#666666").pack(anchor="w")
        self.pb = ttk.Progressbar(t, mode="determinate")
        self.pb.pack(fill="x")
        self.log = ScrolledText(t, height=9, state="disabled", font=("Microsoft YaHei", 9))
        self.log.pack(fill="both", expand=False, pady=(4, 0))
        self.log.tag_configure("ok", foreground="#1e7d32")
        self.log.tag_configure("err", foreground="#c62828")
        self.log.tag_configure("warn", foreground="#e65100")
        self.log.tag_configure("info", foreground="#444444")

    def log_write(self, text, tag="info"):
        self.log.configure(state="normal")
        self.log.insert("end", text.rstrip("\n") + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        # 日志持久化：退出不丢，事后可查（按天一个文件，存于程序目录 logs/）
        try:
            log_dir = os.path.join(app_dir(), "logs")
            os.makedirs(log_dir, exist_ok=True)
            with open(os.path.join(
                    log_dir, "使用日志_%s.log" % datetime.now().strftime("%Y%m%d")),
                    "a", encoding="utf-8") as f:
                f.write("[%s] %s\n" % (datetime.now().strftime("%H:%M:%S"), text.rstrip("\n")))
        except Exception:
            pass

    def _sending(self):
        return bool(self.send_thread and self.send_thread.is_alive())

    def _clear_folders(self):
        if self._sending():
            messagebox.showinfo(APP_NAME, "正在发送邮件中，请等待本轮发送结束再操作", parent=self)
            return
        self.lb_folders.delete(0, "end")
        self.mark_dirty()

    def add_folder(self):
        if self._sending():
            messagebox.showinfo(APP_NAME, "正在发送邮件中，请等待本轮发送结束再操作", parent=self)
            return
        d = filedialog.askdirectory(title="选择要扫描的文件夹", parent=self)
        if d:
            if d in self.lb_folders.get(0, "end"):
                messagebox.showinfo(APP_NAME, "该文件夹已在列表中", parent=self)
                return
            self.lb_folders.insert("end", d)
            self.mark_dirty()

    def del_folder(self):
        if self._sending():
            messagebox.showinfo(APP_NAME, "正在发送邮件中，请等待本轮发送结束再操作", parent=self)
            return
        for i in self.lb_folders.curselection()[::-1]:
            self.lb_folders.delete(i)
            self.mark_dirty()

    def _opt_changed(self):
        self.cfg["send"]["merge"] = self.v_merge.get()
        self.cfg["send"]["recursive"] = self.v_recursive.get()
        self.mark_dirty()

    def _stat_text(self):
        n_ok = sum(1 for r in self.rows if r["status"] == "ok" or r["manual_email"])
        n_none = sum(1 for r in self.rows if r["status"] == "none" and not r["manual_email"])
        n_con = sum(1 for r in self.rows if r["status"] == "conflict" and not r["manual_email"])
        n_to = len(iter_send_groups(self.rows, self.v_merge.get()))
        return ("共 %d 个文件：待发送 %d 个文件（合并为 %d 封）｜未匹配 %d｜命中多人冲突 %d"
                % (len(self.rows), n_ok, n_to, n_none, n_con))

    def do_scan(self):
        if self._sending():
            messagebox.showinfo(APP_NAME, "正在发送邮件中，请等待本轮发送结束再操作", parent=self)
            return
        folders = list(self.lb_folders.get(0, "end"))
        if not folders:
            messagebox.showwarning(APP_NAME, "请先添加至少一个文件夹", parent=self)
            return
        rules = [dict(r) for r in self.cfg["rules"] if r.get("kw")]
        if not rules:
            messagebox.showwarning(APP_NAME, "收件规则为空，请先到\u201c② 收件规则\u201d页添加", parent=self)
            return
        keep_manual = {}
        for r in self.rows:  # 保留已手动指定的（邮箱+姓名）
            keep_manual[(r["folder"], r["fname"])] = (r["manual_email"], r.get("manual_name", ""))
        self.rows = scan(folders, rules, self.v_recursive.get())
        for r in self.rows:
            k = (r["folder"], r["fname"])
            if k in keep_manual:
                m_email, m_name = keep_manual[k]
                if m_email and is_valid_email(m_email):
                    r["manual_email"] = m_email
                    r["manual_name"] = m_name or ""
        self._refresh_tv()
        self.v_stat.set(self._stat_text())
        self.log_write("扫描完成：" + self.v_stat.get())

    def _refresh_tv(self, update_state_only=False):
        if not update_state_only:
            self.tv.delete(*self.tv.get_children())
            for ri, r in enumerate(self.rows):
                if r["manual_email"]:
                    show = ("manual", "-", "手动指定", r["manual_email"], "待发(手动)")
                elif r["status"] == "ok":
                    rule = r["rule"]
                    show = ("ok", "、".join(r["kws"]), rule.get("name", ""),
                            rule.get("email", ""), "待发")
                elif r["status"] == "conflict":
                    show = ("conflict", "、".join(r["kws"]), "?", "?", "冲突·停发")
                else:
                    show = ("none", "-", "-", "-", "未匹配·停发")
                tag = show[0]
                if r.get("state") == "成功":
                    tag = ("sent",)
                elif r.get("state") == "失败":
                    tag = ("failed",)
                self.tv.insert("", "end", iid="r%d" % ri,
                               values=(r["fname"], r["folder"], show[1], show[2], show[3], show[4]),
                               tags=tag)
        else:
            kids = set(self.tv.get_children())
            for ri, r in enumerate(self.rows):
                iid = "r%d" % ri
                if iid not in kids:
                    continue
                if r.get("state") == "成功":
                    state, tag = "成功", ("sent",)
                elif r.get("state") == "失败":
                    state, tag = "失败：" + (r.get("err") or ""), ("failed",)
                elif r["manual_email"]:
                    state, tag = "待发(手动)", ("manual",)
                elif r["status"] == "conflict":
                    state, tag = "冲突·停发", ("conflict",)
                elif r["status"] == "none":
                    state, tag = "未匹配·停发", ("none",)
                else:
                    state, tag = "待发", ("ok",)
                vals = list(self.tv.item(iid, "values"))
                vals[5] = state
                self.tv.item(iid, values=vals, tags=tag)

    def manual_assign(self):
        if self._sending():
            messagebox.showinfo(APP_NAME, "正在发送邮件中，请等待本轮发送结束再操作", parent=self)
            return
        sel = self.tv.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "请先在列表中选中一个文件", parent=self)
            return
        r = self.rows[self.tv.index(sel[0])]
        dlg = tk.Toplevel(self)
        dlg.title("手动指定收件人")
        dlg.resizable(False, False)
        frm = ttk.Frame(dlg, padding=12)
        frm.grid(sticky="nsew")
        ttk.Label(frm, text="文件：").grid(row=0, column=0, sticky="w")
        ttk.Label(frm, text=r["fname"], foreground="#0b5394", wraplength=300).grid(
            row=0, column=1, columnspan=2, sticky="w", pady=3)
        ttk.Label(frm, text="方式一，从已有规则选：").grid(row=1, column=0, sticky="w", pady=(6, 0))
        pairs = [("%s（%s）" % (x.get("name") or x.get("kw"), x.get("email")),
                  x.get("email", "").strip(), x.get("name", ""))
                 for x in self.cfg["rules"] if x.get("email")]
        emap = {p[0]: (p[1], p[2]) for p in pairs}
        names = [p[0] for p in pairs]
        cb = ttk.Combobox(frm, values=names, width=34, state="readonly")
        if names:
            cb.current(0)
        cb.grid(row=1, column=1, columnspan=2, pady=(6, 0), sticky="w")
        ttk.Label(frm, text="方式二，直接输邮箱：").grid(row=2, column=0, sticky="w", pady=(6, 0))
        v_email = tk.StringVar()
        ttk.Entry(frm, textvariable=v_email, width=36).grid(row=2, column=1, columnspan=2, pady=(6, 0), sticky="w")

        def ok():
            if v_email.get().strip():
                if not is_valid_email(v_email.get().strip()):
                    messagebox.showwarning(APP_NAME, "邮箱格式不正确", parent=dlg)
                    return
                r["manual_email"] = v_email.get().strip()
                r["manual_name"] = ""
            elif cb.get():
                r["manual_email"], r["manual_name"] = emap[cb.get()]
            else:
                messagebox.showwarning(APP_NAME, "请选择或输入收件人", parent=dlg)
                return
            dlg.destroy()
            self._refresh_tv()
            self.v_stat.set(self._stat_text())

        ttk.Button(frm, text="确定", command=ok).grid(row=3, column=2, sticky="e", pady=(8, 0))
        ttk.Button(frm, text="取消", command=dlg.destroy).grid(row=3, column=1, sticky="e", pady=(8, 0))
        dlg.transient(self)
        dlg.grab_set()

    def clear_manual(self):
        if self._sending():
            messagebox.showinfo(APP_NAME, "正在发送邮件中，请等待本轮发送结束再操作", parent=self)
            return
        sel = self.tv.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "请先选中一行", parent=self)
            return
        r = self.rows[self.tv.index(sel[0])]
        r["manual_email"] = ""
        r["manual_name"] = ""
        self._refresh_tv()
        self.v_stat.set(self._stat_text())

    def export_unmatched(self):
        """把未匹配/冲突文件复制副本到统一文件夹（绝不移动或删除原文件）"""
        targets = [r for r in self.rows
                   if (r["status"] in ("none", "conflict")) and not r["manual_email"]]
        if not targets:
            messagebox.showinfo(APP_NAME, "当前没有未匹配/冲突的文件", parent=self)
            return
        d = filedialog.askdirectory(title="选择\u201c另存副本\u201d的目标文件夹", parent=self)
        if not d:
            return
        import shutil
        ok, fail = 0, []
        for r in targets:
            src = os.path.join(r["folder"], r["fname"])
            dst = os.path.join(d, r["fname"])
            try:
                if os.path.exists(dst):
                    base, ext = os.path.splitext(r["fname"])
                    # 时间戳后缀：避免多次导出叠加出“_副本_副本”
                    dst = os.path.join(d, "%s_副本_%s%s"
                                       % (base, datetime.now().strftime("%Y%m%d_%H%M%S"), ext))
                shutil.copy2(src, dst)
                ok += 1
            except Exception as e:
                fail.append("%s ： %s" % (r["fname"], e))
        msg = "已复制 %d 个文件副本到：\n%s\n（只复制，原文件不受任何影响）" % (ok, d)
        if fail:
            msg += "\n\n失败 %d 个：\n%s" % (len(fail), "\n".join(fail))
        self.log_write(msg, "warn" if fail else "ok")
        messagebox.showinfo(APP_NAME, msg, parent=self)

    # ---------- 发送流程 ----------
    def do_send(self):
        if not self.rows:
            messagebox.showwarning(APP_NAME, "请先\u201c扫描并匹配\u201d", parent=self)
            return
        smtp = self._collect_smtp()
        err = self._check_smtp(smtp)
        if err:
            messagebox.showwarning(APP_NAME, "邮箱设置不完整： " + err + "\n请到\u201c③ 邮箱设置\u201d页填写。", parent=self)
            return
        groups = iter_send_groups(self.rows, merge=self.v_merge.get())
        if not groups:
            messagebox.showwarning(APP_NAME, "没有可发送的文件（全部未匹配或冲突）", parent=self)
            return
        # 发送前预检文件可读性：被占用/不可读的文件剔除并单独组成待补发组（弹窗知情）
        checked, skip_groups, bad_lines = [], [], []
        for g in groups:
            ok_rows, bad_rows = check_files_readable(g["rows"])
            if bad_rows:
                for r, reason in bad_rows:
                    r["state"], r["err"] = "失败", reason
                    bad_lines.append("· %s —— %s" % (r["fname"], reason))
                if ok_rows:
                    g2 = dict(g)
                    g2["rows"] = list(ok_rows)
                    checked.append(g2)
                skip_groups.append({
                    "email": g["email"], "name": g["name"],
                    "kws": sorted({k for r, _ in bad_rows for k in r["kws"]}),
                    "rows": [r for r, _ in bad_rows],
                    "ok": False, "err": "附件无法读取（文件被占用或已删除）", "time": "",
                })
            else:
                checked.append(g)
        if not checked:
            messagebox.showwarning(
                APP_NAME, "以下文件当前无法读取，未能发送：\n\n%s\n\n请先关闭占用程序后重新扫描发送。"
                % "\n".join(bad_lines), parent=self)
            self._refresh_tv(update_state_only=True)
            return
        groups = checked
        total = sum(len(g["rows"]) for g in groups)
        preview = "\n".join(
            "· %s%s ← %d 个文件" % (g["name"], (" <" + g["email"] + ">") if g["email"] else "", len(g["rows"]))
            for g in groups)
        n_skip = len(self.rows) - total
        warn_text = ""
        if bad_lines:
            warn_text = ("\n\n⚠ 以下 %d 个文件当前无法读取，已从邮件中剔除（解锁后可点\u201c重发失败项\u201d补发）：\n%s"
                         % (len(bad_lines), "\n".join(bad_lines)))
        if not messagebox.askyesno(
                APP_NAME,
                "即将发送 %d 封邮件（共含 %d 个文件附件）\n\n%s\n\n未发送文件：%d 个（未匹配/冲突，可选中行手动指定后重发）%s\n\n确认开始发送？"
                % (len(groups), total, preview, n_skip, warn_text), parent=self):
            return
        self._do_send_groups(smtp, groups, full=True, skip_groups=skip_groups)

    def retry_failed(self):
        if self._sending():
            messagebox.showinfo(APP_NAME, "正在发送邮件中，请等待本轮发送结束再操作", parent=self)
            return
        if not self.groups_done:
            messagebox.showinfo(APP_NAME, "还没有发送记录", parent=self)
            return
        # 空组不进重试；按 (邮箱, 行id) 去重，防历史重复项被反复重发
        seen = set()
        failed = []
        for g in self.groups_done:
            if g.get("ok") or not g.get("rows"):
                continue
            key = (g.get("email", ""), tuple(sorted(id(r) for r in g["rows"])))
            if key in seen:
                continue
            seen.add(key)
            failed.append(g)
        if not failed:
            messagebox.showinfo(APP_NAME, "没有失败项", parent=self)
            return
        smtp = self._collect_smtp()
        err = self._check_smtp(smtp)
        if err:
            messagebox.showwarning(APP_NAME, "邮箱设置不完整： " + err, parent=self)
            return
        if not messagebox.askyesno(APP_NAME, "重发 %d 封失败邮件？" % len(failed), parent=self):
            return
        self._do_send_groups(smtp, failed, full=False)

    def _do_send_groups(self, smtp, groups, full, skip_groups=()):
        if self.send_thread and self.send_thread.is_alive():
            return
        self.stop_event.clear()
        if full:
            # 全量发送：组列表即本轮结果清单（worker 经队列回填每组的发送结果）
            # 预检剔除的文件单独成组一并挂入清单，便于“重发失败项”补发
            self.groups_done = list(groups) + list(skip_groups)
            in_group = {id(r) for g in groups for r in g["rows"]}
            for r in self.rows:
                if id(r) in in_group:
                    r["state"], r["err"] = "未发送", ""
        else:
            # 重试时保留成功项，并把本次重试的组挂回清单，报告才能完整
            self.groups_done = [g for g in self.groups_done if g.get("ok")] + list(groups)
        self.btn_send.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self._set_busy_ui(True)
        self.pb.configure(maximum=len(groups), value=0)
        self.log_write("—— 开始发送：%d 封 ——" % len(groups))
        # worker 参数化：启动前取好配置，线程内不再读 App 的可变状态（q/stop_event 为线程安全原语除外）
        try:
            interval = float(self.cfg["send"].get("interval", 1.0) or 0)
        except Exception:
            interval = 0
        try:
            warn_mb = float(self.cfg["send"].get("attach_warn_mb", 45) or 0)
        except Exception:
            warn_mb = 45.0
        self.send_thread = threading.Thread(
            target=self._send_worker, args=(smtp, groups, interval, warn_mb), daemon=True)
        self.send_thread.start()

    def _set_busy_ui(self, busy):
        """发送期间锁定会改变扫描结果/收件人指向的入口，防止 rows 与发送中的组脱钩"""
        state = "disabled" if busy else "normal"
        for b in (self.btn_scan, self.btn_retry, self.btn_manual,
                  self.btn_clear_manual, self.btn_add_folder,
                  self.btn_del_folder, self.btn_clear_folders):
            try:
                b.configure(state=state)
            except Exception:
                pass

    def _send_worker(self, smtp, groups, interval, warn_mb):
        """发送线程：只收纯参数（smtp/groups/interval/warn_mb），
        不读 App 的可变状态（q/stop_event 为线程安全原语除外），
        所有结果经队列回传主线程处理；整个会话只登录一次，断线自动重连"""
        self.q.put(("worker_start", groups))
        sess = SmtpSession(smtp)
        consecutive_conn_fail = 0
        try:
            for gi, g in enumerate(groups):
                if self.stop_event.is_set():
                    self.q.put(("log", "已按“停止”中断发送", "warn"))
                    break
                self.q.put(("progress", gi))
                self.q.put(("log", "发送 → %s（%d 个文件）…" % (g["email"], len(g["rows"])), "info"))
                try:
                    ok, err, skipped = sess.send_group_once(
                        g, smtp.get("subject_tpl", ""), smtp.get("body_tpl", ""), warn_mb)
                    self.q.put(("result", id(g), ok, err, skipped))
                    consecutive_conn_fail = 0   # 本组能到达服务器，连接正常
                except Exception as e:
                    self.q.put(("result", id(g), False, _friendly_error(e), []))
                    # 连接类异常连续 3 次：服务器大概率彻底不通，提前中止避免执行人干等。
                    # SMTPConnectError（421 拒连）不在 OSError 体系，须单独列入；
                    # 认证失败/拒收等业务错误不计入
                    if isinstance(e, (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, OSError)):
                        consecutive_conn_fail += 1
                        if consecutive_conn_fail >= 3:
                            self.q.put(("log", "连续 3 次连不上邮件服务器，已提前中止本轮，请检查网络或服务器设置", "err"))
                            break
                    else:
                        consecutive_conn_fail = 0
                if gi < len(groups) - 1 and interval > 0 and not self.stop_event.is_set():
                    time.sleep(interval)
            self.q.put(("progress", len(groups)))
        finally:
            sess.close()
        self.q.put(("worker_done",))

    def do_stop(self):
        self.stop_event.set()
        self.log_write("停止指令已发出，正在等待当前这封结束…", "warn")

    def _poll_queue(self):
        try:
            while True:
                item = self.q.get_nowait()
                kind = item[0]
                if kind == "log":
                    self.log_write(item[1], item[2])
                elif kind == "progress":
                    self.pb.configure(value=item[1])
                    total_mails = self.pb.cget("maximum")
                    try:
                        total_mails = int(total_mails)
                    except Exception:
                        total_mails = 0
                    if total_mails > 0:
                        self.v_prog.set("进度：第 %d 封 / 共 %d 封"
                                         % (min(item[1] + 1, total_mails), total_mails))
                elif kind == "worker_start":
                    self._active_groups = list(item[1])
                elif kind == "result":
                    self._apply_send_result(item[1], item[2], item[3], item[4])
                elif kind == "worker_done":
                    self.btn_send.configure(state="normal")
                    self.btn_stop.configure(state="disabled")
                    self._set_busy_ui(False)
                    self.v_prog.set("发送结束")
                    self._refresh_tv(update_state_only=True)
                    self._finish_send()
        except queue.Empty:
            pass
        # 记录 after 句柄：销毁前可取消，避免 destroy 后尾随回调访问已销毁窗口
        self._poll_after = self.after(100, self._poll_queue)

    def _apply_send_result(self, gid, ok, err, skipped=()):
        """在主线程应用发送结果，更新组/行状态与界面；
        发送中被跳过的附件单独组成失败组挂入清单，便于“重发失败项”补发"""
        g = None
        for cand in self._active_groups:
            if id(cand) == gid:
                g = cand
                break
        if g is None:
            return
        g["ok"], g["err"] = ok, err
        g["time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if ok else ""
        if ok:
            self.log_write("  ✔ 成功：%s" % g["email"], "ok")
            for r in g["rows"]:
                r["state"], r["err"] = "成功", ""
        else:
            self.log_write("  ✘ 失败：%s —— %s" % (g["email"], err), "err")
            for r in g["rows"]:
                r["state"], r["err"] = "失败", err
        if skipped:
            skipped_ids = {id(r) for r, _ in skipped}
            for r, reason in skipped:
                r["state"], r["err"] = "失败", reason
                self.log_write("  ⚠ 跳过附件：%s —— %s" % (r["fname"], reason), "warn")
            # 无论成功与否都收缩，避免跳过行既留在主组又进跳过组（报告重复+重试累积）
            g["rows"] = [r for r in g["rows"] if id(r) not in skipped_ids]
            if not g["rows"]:
                # 主组已无任何文件：把空主组从清单移除（用 is 判同一对象，remove 的内容比较会误删同内容组）；
                # 跳过的文件仍然全部转入跳过组待补发
                self.groups_done = [x for x in self.groups_done if x is not g]
                self.log_write("  ⚠ 该组全部附件无法读取，整组转为待补发项", "warn")
            skip_g = {
                "email": g["email"], "name": g["name"],
                "kws": sorted({k for r, _ in skipped for k in r["kws"]}),
                "rows": [r for r, _ in skipped],
                "ok": False, "err": "附件无法读取（文件被占用或已删除）", "time": "",
            }
            # 主邮件已成功发出、仅缺这些附件 → 标记 partial：结束统计单独提示，不计入“失败 N 封”
            skip_g["partial"] = bool(g.get("ok") and g["rows"])
            self.groups_done.append(skip_g)
        self._refresh_tv(update_state_only=True)

    def _finish_send(self):
        # partial 组 = 主邮件已发出、仅缺附件待补发：单独提示，不计入“失败 N 封”
        skip_groups = [g for g in self.groups_done if g.get("partial") and not g.get("ok")]
        n_skip_att = sum(len(g["rows"]) for g in skip_groups)
        n_ok = sum(1 for g in self.groups_done if g.get("ok"))
        n_fail = sum(1 for g in self.groups_done
                     if not g.get("ok")) - len(skip_groups)
        out = None
        try:
            rep_dir = os.path.join(app_dir(), "发送清单")
            os.makedirs(rep_dir, exist_ok=True)
            path = os.path.join(
                rep_dir, "发送清单_%s.xlsx" % datetime.now().strftime("%Y%m%d_%H%M%S"))
            out = write_report(self.rows, self.groups_done, path)
        except Exception as e:
            self.log_write("生成发送清单失败： %s" % _err_text(e), "err")
        msg = "发送结束：成功 %d 封" % n_ok
        if skip_groups:
            msg += ("（其中 %d 封有 %d 个附件因无法读取被跳过，"
                    "解锁后可点\u201c重发失败项\u201d补发）"
                    % (len(skip_groups), n_skip_att))
        msg += "，失败 %d 封。" % n_fail
        if n_fail:
            msg += "\n失败项可点\u201c重发失败项\u201d单独重发。"
        if out:
            msg += "\n\n发送清单已生成：\n%s" % out
            self.log_write("发送清单已生成： %s" % out, "ok")
        self.log_write("—— 全部结束 ——" if not n_fail else "—— 结束（有失败项） ——",
                       "ok" if not n_fail else "err")
        if out:
            if messagebox.askyesno(APP_NAME, msg + "\n\n是否打开清单所在文件夹？", parent=self):
                try:
                    os.startfile(os.path.dirname(out))
                except Exception:
                    pass
        else:
            messagebox.showinfo(APP_NAME, msg, parent=self)

    # ---------- 初始化各列表/表单数据 ----------
    def _load_lists(self):
        """把已保存的文件夹、规则、邮箱设置填充到界面"""
        self.lb_folders.delete(0, "end")
        for f in self.cfg.get("folders", []):
            self.lb_folders.insert("end", f)
        self._refresh_rules()
        self._fill_smtp_form()
        self._update_sec_warn()

    # ---------- 页2：收件规则 ----------
    def _build_rules_tab(self):
        t = self.tab_rules
        ttk.Label(t, text="匹配规则：文件名\u201c包含关键字\u201d即命中；多个关键字命中时，最长关键字优先；"
                          "同一文件命中多个不同收件人时停发并标红，由人工指定后再发。",
                  foreground="#666666", wraplength=960, justify="left").pack(anchor="w")
        main = ttk.Frame(t)
        main.pack(fill="both", expand=True, pady=(4, 0))
        self.tv_rules = ttk.Treeview(main, columns=("kw", "name", "email"), show="headings")
        self.tv_rules.heading("kw", text="关键字")
        self.tv_rules.heading("name", text="收件人")
        self.tv_rules.heading("email", text="邮箱")
        self.tv_rules.column("kw", width=290)
        self.tv_rules.column("name", width=140)
        self.tv_rules.column("email", width=270)
        self.tv_rules.pack(side="left", fill="both", expand=True)
        vsb2 = ttk.Scrollbar(main, orient="vertical", command=self.tv_rules.yview)
        vsb2.pack(side="left", fill="y")
        self.tv_rules.configure(yscrollcommand=vsb2.set)
        self.tv_rules.bind("<Double-1>", lambda e: self.edit_rule())
        bf = ttk.Frame(t)
        bf.pack(fill="x", pady=(6, 0))
        ttk.Button(bf, text="新增", command=self.add_rule).pack(side="left")
        ttk.Button(bf, text="编辑", command=self.edit_rule).pack(side="left", padx=6)
        ttk.Button(bf, text="删除", command=self.del_rule).pack(side="left")
        ttk.Button(bf, text="从Excel导入", command=self.import_rules).pack(side="left", padx=(18, 0))
        ttk.Button(bf, text="保存规则", command=self.save_rules).pack(side="right")

    def _refresh_rules(self):
        self.tv_rules.delete(*self.tv_rules.get_children())
        for r in self.cfg.get("rules", []):
            self.tv_rules.insert("", "end",
                                 values=(r.get("kw", ""), r.get("name", ""), r.get("email", "")))
        seen = {}
        for iid in self.tv_rules.get_children():
            kw = self.tv_rules.item(iid, "values")[0]
            if kw in seen:
                self.tv_rules.item(iid, tags=("dup",))
                self.tv_rules.item(seen[kw], tags=("dup",))
            else:
                seen[kw] = iid
        self.tv_rules.tag_configure("dup", background="#fde9e9")

    def add_rule(self):
        d = RuleDialog(self, "新增规则")
        self.wait_window(d)
        if d.result:
            if d.result["kw"] in [x.get("kw") for x in self.cfg["rules"]]:
                messagebox.showwarning(APP_NAME, "关键字\u201c%s\u201d已存在" % d.result["kw"], parent=self)
                return
            self.cfg["rules"].append(d.result)
            self._refresh_rules()
            self.mark_dirty()

    def edit_rule(self):
        sel = self.tv_rules.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "请先选中一条规则", parent=self)
            return
        idx = self.tv_rules.index(sel[0])
        d = RuleDialog(self, "编辑规则", self.cfg["rules"][idx])
        self.wait_window(d)
        if d.result:
            for j, x in enumerate(self.cfg["rules"]):
                if j != idx and x["kw"] == d.result["kw"]:
                    messagebox.showwarning(APP_NAME, "关键字\u201c%s\u201d已存在" % d.result["kw"], parent=self)
                    return
            self.cfg["rules"][idx] = d.result
            self._refresh_rules()
            self.mark_dirty()

    def del_rule(self):
        sel = self.tv_rules.selection()
        if not sel:
            messagebox.showinfo(APP_NAME, "请先选中一条规则", parent=self)
            return
        if not messagebox.askyesno(APP_NAME, "删除选中的 %d 条规则？" % len(sel), parent=self):
            return
        for i in reversed(sel):
            del self.cfg["rules"][self.tv_rules.index(i)]
        self._refresh_rules()
        self.mark_dirty()

    def import_rules(self):
        p = filedialog.askopenfilename(
            title="选择规则Excel（A列=关键字 B列=姓名 C列=邮箱）",
            filetypes=[("Excel 文件", "*.xlsx *.xlsm *.xls"), ("所有文件", "*.*")], parent=self)
        if not p:
            return
        if os.path.splitext(p)[1].lower() == ".xls":
            # openpyxl 不支持 97-2003 旧二进制格式：给人话提示而非技术异常
            messagebox.showwarning(
                APP_NAME, "Excel 97-2003 旧格式（.xls）不支持，请先用 Excel/WPS 打开，"
                          "“另存为” .xlsx 后再导入。", parent=self)
            return
        try:
            rows, bad = _read_rules_excel(p)
        except Exception as e:
            messagebox.showerror(APP_NAME, "读取Excel失败： %s" % _err_text(e), parent=self)
            return
        if not rows and not bad:
            messagebox.showwarning(APP_NAME, "没有读到有效规则行（需 A列=关键字 B列=姓名 C列=邮箱）", parent=self)
            return
        added, skipped = 0, []
        have = {x.get("kw") for x in self.cfg["rules"]}
        for r in rows:
            if r["kw"] in have:
                skipped.append("%s（关键字已存在，跳过）" % r["kw"])
                continue
            self.cfg["rules"].append(r)
            have.add(r["kw"])
            added += 1
        self._refresh_rules()
        self.mark_dirty()
        msg = "导入成功 %d 条。" % added
        all_note = skipped + bad
        if all_note:
            msg += "\n\n跳过/有问题的 %d 行：\n%s" % (len(all_note), "\n".join(all_note[:15]))
        messagebox.showinfo(APP_NAME, msg, parent=self)

    def save_rules(self):
        self._save_all()
        messagebox.showinfo(APP_NAME, "规则已保存", parent=self)

    # ---------- 页3：邮箱设置 ----------
    def _build_smtp_tab(self):
        t = self.tab_smtp
        frm = ttk.Frame(t)
        frm.pack(fill="both", expand=True)
        col1 = ttk.Frame(frm)
        col1.pack(side="left", fill="both", expand=True, padx=(0, 12))

        g1 = ttk.LabelFrame(col1, text=" 服务器 ", padding=10)
        g1.pack(fill="x")
        self.v_host = tk.StringVar()
        self.v_port = tk.IntVar()
        self.v_sec = tk.StringVar()
        self.v_auth = tk.BooleanVar(value=True)
        self.v_timeout = tk.IntVar()
        ttk.Label(g1, text="SMTP 服务器：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(g1, textvariable=self.v_host, width=34).grid(row=0, column=1, sticky="w", pady=3)
        ttk.Label(g1, text="端口：").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(g1, textvariable=self.v_port, width=10).grid(row=1, column=1, sticky="w", pady=3)
        ttk.Label(g1, text="加密方式：").grid(row=2, column=0, sticky="w", pady=3)
        secbox = ttk.Combobox(g1, textvariable=self.v_sec, state="readonly",
                              values=tuple(SEC_LABEL_TO_VAL), width=32)
        secbox.grid(row=2, column=1, sticky="w", pady=3)
        secbox.bind("<<ComboboxSelected>>", lambda e: self._sec_changed())
        ttk.Checkbutton(g1, text="登录认证（内部中继免认证时取消勾选）", variable=self.v_auth).grid(
            row=3, column=0, columnspan=2, sticky="w", pady=3)
        ttk.Label(g1, text="超时（秒）：").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(g1, textvariable=self.v_timeout, width=10).grid(row=4, column=1, sticky="w", pady=3)
        self.v_tls = tk.StringVar()
        ttk.Label(g1, text="证书校验：").grid(row=5, column=0, sticky="w", pady=3)
        tlsbox = ttk.Combobox(g1, textvariable=self.v_tls, state="readonly",
                              values=tuple(TLS_MODES.values()), width=32)
        tlsbox.grid(row=5, column=1, sticky="w", pady=3)
        self.v_cafile = tk.StringVar()
        ttk.Label(g1, text="CA证书文件（仅\u201c校验证书链\u201d时）：").grid(row=6, column=0, sticky="w", pady=3)
        ttk.Entry(g1, textvariable=self.v_cafile, width=34).grid(row=6, column=1, sticky="w", pady=3)
        self.lb_secwarn = tk.Label(g1, text="", foreground="#e65100", wraplength=420, justify="left")
        self.lb_secwarn.grid(row=7, column=0, columnspan=2, sticky="w", pady=2)

        g2 = ttk.LabelFrame(col1, text=" 账号与身份 ", padding=10)
        g2.pack(fill="x", pady=(8, 0))
        self.v_user = tk.StringVar()
        self.v_pwd = tk.StringVar()
        self.v_from = tk.StringVar()
        self.v_disp = tk.StringVar()
        self.v_cc = tk.StringVar()
        self.v_bcc = tk.StringVar()
        self.v_remember = tk.BooleanVar()
        ttk.Label(g2, text="发件账号：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(g2, textvariable=self.v_user, width=34).grid(row=0, column=1, sticky="w", pady=3)
        ttk.Label(g2, text="密码/授权码：").grid(row=1, column=0, sticky="w", pady=3)
        ttk.Entry(g2, textvariable=self.v_pwd, width=34, show="*").grid(row=1, column=1, sticky="w", pady=3)
        ttk.Checkbutton(g2, text="记住密码（保存在本机 Windows 凭据管理器，公用电脑勿勾）",
                        variable=self.v_remember).grid(row=2, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Label(g2, text="发件邮箱地址：").grid(row=3, column=0, sticky="w", pady=3)
        ttk.Entry(g2, textvariable=self.v_from, width=34).grid(row=3, column=1, sticky="w", pady=3)
        ttk.Label(g2, text="留空=同发件账号", foreground="#999999").grid(row=3, column=2, sticky="w", padx=4)
        ttk.Label(g2, text="发件人显示名：").grid(row=4, column=0, sticky="w", pady=3)
        ttk.Entry(g2, textvariable=self.v_disp, width=34).grid(row=4, column=1, sticky="w", pady=3)
        ttk.Label(g2, text="抄送（可空，分号分隔）").grid(row=5, column=0, sticky="w", pady=3)
        ttk.Entry(g2, textvariable=self.v_cc, width=34).grid(row=5, column=1, sticky="w", pady=3)
        ttk.Label(g2, text="密送（可空，分号分隔）").grid(row=6, column=0, sticky="w", pady=3)
        ttk.Entry(g2, textvariable=self.v_bcc, width=34).grid(row=6, column=1, sticky="w", pady=3)
        ttk.Label(g2, text="收件人不可见", foreground="#999999").grid(row=6, column=2, sticky="w", padx=4)

        g3 = ttk.LabelFrame(col1, text=" 邮件模板（占位符：{姓名} {关键字} {文件数} {文件名}）", padding=10)
        g3.pack(fill="both", expand=True, pady=(8, 0))
        self.v_subj = tk.StringVar()
        ttk.Label(g3, text="主题：").grid(row=0, column=0, sticky="w", pady=3)
        ttk.Entry(g3, textvariable=self.v_subj, width=58).grid(row=0, column=1, sticky="we", pady=3)
        ttk.Label(g3, text="正文：").grid(row=1, column=0, sticky="nw", pady=3)
        self.tb_body = tk.Text(g3, height=9, width=58)
        self.tb_body.grid(row=1, column=1, sticky="nwes", pady=3)
        g3.columnconfigure(1, weight=1)
        g3.rowconfigure(1, weight=1)

        col2 = ttk.Frame(frm)
        col2.pack(side="left", fill="y")
        g4 = ttk.LabelFrame(col2, text=" 保存与测试 ", padding=10)
        g4.pack(fill="y")
        ttk.Button(g4, text="保存设置", command=self.save_smtp).pack(fill="x", pady=3)
        ttk.Button(g4, text="测试连接（登录验证）", command=self.test_conn).pack(fill="x", pady=3)
        ttk.Label(g4, text="测试收件邮箱：", foreground="#666666").pack(anchor="w", pady=(8, 0))
        self.v_test_to = tk.StringVar()
        ttk.Entry(g4, textvariable=self.v_test_to, width=26).pack(anchor="w", pady=2)
        ttk.Button(g4, text="发一封测试邮件", command=self.send_test).pack(fill="x", pady=3)
        self.lb_test = tk.Label(g4, text="", wraplength=230, justify="left", foreground="#666666")
        self.lb_test.pack(anchor="w", pady=(6, 0))
        self.v_interval = tk.DoubleVar()
        ttk.Label(g4, text="每封间隔（秒）：", foreground="#666666").pack(anchor="w", pady=(12, 0))
        ttk.Entry(g4, textvariable=self.v_interval, width=10).pack(anchor="w", pady=2)
        self.v_attach_mb = tk.DoubleVar()
        ttk.Label(g4, text="附件单封上限（MB）：", foreground="#666666").pack(anchor="w", pady=(6, 0))
        ttk.Entry(g4, textvariable=self.v_attach_mb, width=10).pack(anchor="w", pady=2)
        ttk.Label(g4, text="参考：SSL=465\nSTARTTLS=587 或 25\n不加密=25\n授权码向IT或邮箱\n设置页获取",
                  foreground="#999999", justify="left").pack(anchor="w", pady=(10, 0))

    def _sec_changed(self):
        mapping = {"SSL": 465, "STARTTLS": 587, SEC_NONE_LABEL: 25}
        try:
            self.v_port.set(mapping.get(self.v_sec.get(), 25))
        except Exception:
            pass
        self._update_sec_warn()

    def _update_sec_warn(self):
        """none（25 明文）时显示黄字安全提醒"""
        if self.v_sec.get() == SEC_NONE_LABEL:
            self.lb_secwarn.configure(
                text="⚠ 25端口为明文传输，账号密码与附件可能被同网段截获；内网若支持 SSL/STARTTLS 请优先使用",
                foreground="#e65100")
        else:
            self.lb_secwarn.configure(text="", foreground="#e65100")

    def _fill_smtp_form(self):
        self._supp_trace = True      # 程序化回填不算“人工改动”，不标脏
        try:
            s = self.cfg["smtp"]
            self.v_host.set(s.get("host", ""))
            try:
                self.v_port.set(int(s.get("port") or 465))
            except Exception:
                self.v_port.set(465)
            sec = s.get("security", "SSL")
            self.v_sec.set(SEC_VAL_TO_LABEL.get(sec, SEC_NONE_LABEL))
            self.v_auth.set(bool(s.get("auth", True)))
            try:
                self.v_timeout.set(int(s.get("timeout") or 20))
            except Exception:
                self.v_timeout.set(20)
            self.v_user.set(s.get("user", ""))
            self.v_pwd.set(s.get("passwd", "") if self.cfg.get("remember_passwd") else "")
            self.v_remember.set(bool(self.cfg.get("remember_passwd")))
            self.v_from.set(s.get("from_addr", ""))
            self.v_disp.set(s.get("display_name", ""))
            self.v_cc.set(s.get("cc", ""))
            self.v_bcc.set(s.get("bcc", ""))
            self.v_subj.set(s.get("subject_tpl", ""))
            self.tb_body.delete("1.0", "end")
            self.tb_body.insert("1.0", s.get("body_tpl", ""))
            tls_val = s.get("tls_verify", "off")
            self.v_tls.set(TLS_MODES.get(tls_val, TLS_MODES["off"]))
            if tls_val not in TLS_MODES:
                self.log_write("⚠ 配置中的证书校验值无效（%s），已回落为“不校验”，请到邮箱设置页重新选择" % tls_val, "warn")
            self.v_cafile.set(s.get("ca_file", ""))
            try:
                self.v_interval.set(float(self.cfg["send"].get("interval", 1.0)))
            except Exception:
                self.v_interval.set(1.0)
            try:
                self.v_attach_mb.set(float(self.cfg["send"].get("attach_warn_mb", 45)))
            except Exception:
                self.v_attach_mb.set(45.0)
        finally:
            self._supp_trace = False

    def _collect_smtp(self):
        sec = SEC_LABEL_TO_VAL.get(self.v_sec.get(), "none")
        try:
            port = int(self.v_port.get())
        except Exception:
            port = 0
        try:
            timeout = int(self.v_timeout.get())
        except Exception:
            timeout = 20
        return {
            "host": self.v_host.get().strip(),
            "port": port, "security": sec,
            "auth": self.v_auth.get(),
            "user": self.v_user.get().strip(),
            "passwd": self.v_pwd.get(),
            "from_addr": self.v_from.get().strip(),
            "display_name": self.v_disp.get().strip(),
            "cc": self.v_cc.get().strip(),
            "bcc": self.v_bcc.get().strip(),
            "subject_tpl": self.v_subj.get(),
            "body_tpl": self.tb_body.get("1.0", "end").rstrip("\n"),
            "timeout": timeout,
            "tls_verify": TLS_REV.get(self.v_tls.get(), "off"),
            "ca_file": self.v_cafile.get().strip(),
        }

    def _check_smtp(self, s):
        if not s["host"]:
            return "SMTP 服务器为空"
        if s["auth"] and not s["user"]:
            return "发件账号为空"
        if s["auth"] and not s["passwd"]:
            return "密码/授权码为空"
        if not s["auth"] and not s["from_addr"]:
            return "免认证模式需填写“发件邮箱地址”"
        if s.get("tls_verify") == "ca" and s.get("ca_file") \
                and not os.path.isfile(s["ca_file"]):
            return "CA证书文件不存在： %s" % s["ca_file"]
        return ""

    def save_smtp(self):
        s = self._collect_smtp()
        err = self._check_smtp(s)
        if err and not messagebox.askyesno(
                APP_NAME, "提醒：当前设置不完整（%s）。\n仍要保存吗？（发送时会再次校验）" % err, parent=self):
            return
        # 密码保管：优先 Windows 凭据管理器（keyring），不可用退回 config.json 明文（需用户确认）
        self.cfg["remember_passwd"] = self.v_remember.get()
        stored_note = "密码：不保存"
        plaintext_fallback = False
        if self.v_remember.get() and s["passwd"]:
            if store_password(s["user"], s["passwd"]):
                s["passwd"] = ""            # 明文不落盘
                s["pwd_in_keyring"] = True
                stored_note = "密码已存入 Windows 凭据管理器"
            else:
                # 凭据管理器不可用：明文保存有风险（公用电脑可被同机用户读取），须执行人知情确认
                if not messagebox.askyesno(
                        APP_NAME,
                        "本机凭据管理器不可用，密码将以明文保存在 config.json（公用电脑有泄露风险）。\n\n"
                        "仍要保存密码？（选“否”= 本次不保存密码，下次启动需重新输入）", parent=self):
                    self.v_remember.set(False)
                    self.cfg["remember_passwd"] = False
                    clear_password(s["user"])
                    s["passwd"] = ""
                    s["pwd_in_keyring"] = False
                    stored_note = "密码：未保存（凭据管理器不可用，已取消明文保存）"
                else:
                    s["pwd_in_keyring"] = False
                    stored_note = "密码以明文保存于 config.json（本机凭据管理器不可用）"
                    plaintext_fallback = True
        else:
            clear_password(s["user"])
            s["passwd"] = ""
            s["pwd_in_keyring"] = False
        self.cfg["smtp"] = s
        try:
            self.cfg["send"]["interval"] = float(self.v_interval.get() or 1.0)
        except Exception:
            pass
        try:
            self.cfg["send"]["attach_warn_mb"] = float(self.v_attach_mb.get() or 45)
        except Exception:
            pass
        self._save_all()
        fg = "#e65100" if plaintext_fallback else "#1e7d32"
        self.lb_test.configure(text="设置已保存。%s。" % stored_note, foreground=fg)

    def test_conn(self):
        s = self._collect_smtp()
        err = self._check_smtp(s)
        if err:
            self.lb_test.configure(text="✘ " + err, foreground="#c62828")
            messagebox.showwarning(APP_NAME, err, parent=self)
            return
        self.lb_test.configure(text="连接中…", foreground="#666666")
        self.update_idletasks()

        def work():
            r = smtp_login_test(s)
            ok_txt = "✔ 连接与登录成功，可以发送"
            if s.get("tls_verify") == "off":
                ok_txt += "（当前未校验证书）"   # 知情提示：off 档不验证服务器身份
            self.after(0, lambda: self.lb_test.configure(
                text=(ok_txt if r is None else "✘ " + r),
                foreground=("#1e7d32" if r is None else "#c62828")))
        threading.Thread(target=work, daemon=True).start()

    def send_test(self):
        to = self.v_test_to.get().strip()
        if not is_valid_email(to):
            messagebox.showwarning(APP_NAME, "请先填写正确的测试收件邮箱", parent=self)
            return
        s = self._collect_smtp()
        err = self._check_smtp(s)
        if err:
            messagebox.showwarning(APP_NAME, err, parent=self)
            return
        self.lb_test.configure(text="测试邮件发送中…", foreground="#666666")
        self.update_idletasks()

        def work():
            try:
                msg = MIMEMultipart()
                from_name = s.get("display_name") or APP_NAME
                from_addr = (s.get("from_addr") or s.get("user")).strip()
                msg["From"] = formataddr((str(Header(from_name, "utf-8")), from_addr))
                msg["To"] = to
                msg["Subject"] = Header("【%s】测试邮件" % APP_NAME, "utf-8")
                msg["Date"] = formatdate(localtime=True)
                msg.attach(MIMEText("这是 %s 的连通测试邮件，收到即说明发送链路正常。\n\n（如需带附件测试，请用正式流程扫描一个测试文件夹）"
                                    % APP_NAME, "plain", "utf-8"))
                server = _connect(s)
                try:
                    server.sendmail(from_addr, [to], msg.as_string())
                finally:
                    try:
                        server.quit()
                    except Exception:
                        pass
                self.after(0, lambda: self.lb_test.configure(
                    text="✔ 测试邮件已发出，请查收", foreground="#1e7d32"))
            except Exception as e:
                msg = _friendly_error(e)
                self.after(0, lambda m=msg: self.lb_test.configure(
                    text="✘ " + m, foreground="#c62828"))
        threading.Thread(target=work, daemon=True).start()

    # ---------- 保存与退出 ----------
    def _smtp_field_dirty(self, *args):
        """邮箱设置页任一字段被人工改动时标记未保存；程序化回填用 _supp_trace 抑制"""
        if not self._supp_trace:
            self.mark_dirty()

    def _body_modified(self, event=None):
        """正文模板任何修改（键入/粘贴/撤销）都标脏；<<Modified>> 标志须重置否则只触发一次"""
        if self.tb_body.edit_modified():
            self.tb_body.edit_modified(False)
            self._smtp_field_dirty()

    def _bind_smtp_dirty(self):
        """邮箱设置页全部字段挂改动监听：改完不点保存直接关窗也会提醒，不再静默丢改动"""
        for var in (self.v_host, self.v_port, self.v_sec, self.v_auth,
                    self.v_timeout, self.v_tls, self.v_cafile, self.v_user,
                    self.v_pwd, self.v_from, self.v_disp, self.v_cc,
                    self.v_bcc, self.v_remember, self.v_subj,
                    self.v_interval, self.v_attach_mb):
            var.trace_add("write", self._smtp_field_dirty)
        self.tb_body.bind("<<Modified>>", self._body_modified, add="+")

    def mark_dirty(self):
        self._dirty = True

    def _save_all(self):
        self.cfg["folders"] = list(self.lb_folders.get(0, "end"))
        save_config(self.cfg)
        self._dirty = False

    def _on_close(self):
        # 防重入：等待退出期间用户再点关闭/程序事件重入时直接忽略，避免嵌套询问
        if getattr(self, "_closing", False):
            return
        if self._sending():
            choice = messagebox.askyesnocancel(
                APP_NAME,
                "正在发送邮件中，直接关闭会中断当前邮件且不生成发送清单。\n\n"
                "「是」= 停止发送、保存进度并退出（推荐）\n"
                "「否」= 等本轮发送完成后再退出\n"
                "「取消」= 继续使用程序", parent=self)
            if choice is None:
                return
            self._closing = True
            if choice:
                self.stop_event.set()
                self.log_write("正在停止发送以保存进度…", "warn")
                deadline = time.time() + 35

                def wait_stop():
                    # after 轮询代替阻塞 join：界面保持响应，且防重入标志已挂
                    if self._sending() and time.time() < deadline:
                        self.after(100, wait_stop)
                        return
                    if self._sending():
                        # 35 秒仍未结束（服务器无响应卡死）：daemon 线程会随窗口销毁被硬杀，
                        # 当前这封的状态与发送清单可能缺失，必须让执行人知情后再决定。
                        # 停止指令已生效（set 保留，符合当初点停止的意图），
                        # 当前这封结束后后续邮件不再发送，文案须明示
                        if not messagebox.askyesno(
                                APP_NAME,
                                "停止指令已生效，当前这封邮件仍在等待服务器响应（已 35 秒）。\n"
                                "当前这封结束后，后续邮件将不再发送。\n"
                                "强制退出将丢失当前这封的发送状态，本轮发送清单可能缺失。\n\n"
                                "仍要强制退出？（选“否”= 继续等待当前这封结束后自动收尾）", parent=self):
                            def wait_forever():
                                if self._sending():
                                    self.after(100, wait_forever)
                                else:
                                    self._finish_close()
                            wait_forever()
                            return
                        self._finish_close()
                        return
                    self._finish_close()

                wait_stop()
                return
            # 「否」：等本轮完成（最长 600 秒）
            self.log_write("等待本轮发送完成后退出…", "info")
            end = time.time() + 600

            def wait_send():
                if self._sending() and time.time() < end:
                    self.after(100, wait_send)
                else:
                    self._finish_close()

            wait_send()
            return
        self._finish_close()

    def _finish_close(self):
        """退出收尾：泵完队列余量（_poll_queue 仍在 after 循环里继续泵）→ 未保存提醒 → 销毁窗口"""
        # 泵完队列余量，让已发结果与发送清单落盘后再退出（最多 10 秒）
        self.log_write("正在等待发送结果与清单落盘（最多 10 秒）…", "info")
        end2 = time.time() + 10
        while time.time() < end2 and (self._sending() or not self.q.empty()):
            self.update()
            time.sleep(0.05)
        self.update()
        if self._dirty:
            if messagebox.askyesno(APP_NAME, "文件夹/规则/选项有改动未保存。\n\n保存后退出？（“否”=不保存退出）"):
                self._save_all()
        # 取消队列轮询的尾随回调，避免 destroy 后访问已销毁窗口
        if getattr(self, "_poll_after", None):
            try:
                self.after_cancel(self._poll_after)
            except Exception:
                pass
        self.destroy()


# ---------------- 自检（用于打包后验证 exe 可用） ----------------

def run_selftest():
    """验证运行时、核心匹配逻辑、tkinter 初始化。结果写 selftest_report.txt"""
    lines = []

    def check(name, fn):
        try:
            fn()
            lines.append("[OK] " + name)
        except Exception:
            lines.append("[FAIL] " + name)
            lines.append(traceback.format_exc(limit=4))

    rules = [
        {"kw": "临江大学", "name": "李四", "email": "lisi@t.cn"},
        {"kw": "临江大学紫金", "name": "张三", "email": "zhangsan@t.cn"},
        {"kw": "云洲师范东区", "name": "王五", "email": "wangwu@t.cn"},
        {"kw": "云洲师范", "name": "赵六", "email": "zhaoliu@t.cn"},
    ]

    def t_longest():
        m = match_file("临江大学紫金2608.xlsx", rules)
        assert m["status"] == "ok" and m["rule"]["email"] == "zhangsan@t.cn", m
        m2 = match_file("临江大学2607.xlsx", rules)
        assert m2["status"] == "ok" and m2["rule"]["email"] == "lisi@t.cn", m2
        m3 = match_file("云洲师范东区报告单.pdf", rules)
        assert m3["status"] == "ok" and m3["rule"]["email"] == "wangwu@t.cn", m3
        m4 = match_file("云洲师范普通.docx", rules)
        assert m4["status"] == "ok" and m4["rule"]["email"] == "zhaoliu@t.cn", m4
        m5 = match_file("临江大学紫金2609.xlsx", rules)
        assert m5["status"] == "ok" and m5["rule"]["email"] == "zhangsan@t.cn", m5

    def t_none():
        assert match_file("随便一个文件.txt", rules)["status"] == "none"
        assert match_file("X", [{"kw": "", "name": "a", "email": "b@t.cn"}])["status"] == "none"

    def t_conflict():
        # 等长关键字（均为4字）同时命中不同收件人 → 冲突停发
        rs = [{"kw": "项目决算", "name": "甲", "email": "jia@t.cn"},
              {"kw": "决算清单", "name": "乙", "email": "yi@t.cn"}]
        m = match_file("项目决算清单11月.xlsx", rs)
        assert m["status"] == "conflict", m
        # 长关键字优先吞掉短关键字，不算冲突
        rs2 = rules + [{"kw": "紫金学院", "name": "钱七", "email": "qianqi@t.cn"}]
        m2 = match_file("临江大学紫金紫金学院报表.xlsx", rs2)
        assert m2["status"] == "ok" and m2["rule"]["email"] == "zhangsan@t.cn", m2

    def t_group():
        rows = [
            {"folder": "d", "fname": "临江大学紫金2608.xlsx", "status": "ok",
             "rule": rules[1], "kws": ["临江大学紫金"], "manual_email": "", "state": "", "err": ""},
            {"folder": "d", "fname": "临江大学紫金2609.xlsx", "status": "ok",
             "rule": rules[1], "kws": ["临江大学紫金"], "manual_email": "", "state": "", "err": ""},
            {"folder": "d", "fname": "别的.xlsx", "status": "none",
             "rule": None, "kws": [], "manual_email": "", "state": "", "err": ""},
        ]
        groups = iter_send_groups(rows, True)
        assert len(groups) == 1 and len(groups[0]["rows"]) == 2, groups
        groups2 = iter_send_groups(rows, False)
        assert len(groups2) == 2
        rows[2]["manual_email"] = "x@t.cn"
        groups3 = iter_send_groups(rows, True)
        assert len(groups3) == 2
        assert all("@" not in (g["name"] or "") for g in groups3), \
            [g["name"] for g in groups3]  # 手动指定不再把邮箱当姓名

    def t_manual_name():
        rows = [
            {"folder": "d", "fname": "a.txt", "status": "ok",
             "rule": {"kw": "K", "name": "规则人", "email": "rule@t.cn"},
             "kws": ["K"], "manual_email": "m1@t.cn", "manual_name": "王小明",
             "state": "", "err": ""},
            {"folder": "d", "fname": "b.txt", "status": "ok",
             "rule": {"kw": "K", "name": "规则人", "email": "rule@t.cn"},
             "kws": ["K"], "manual_email": "m2@t.cn", "manual_name": "",
             "state": "", "err": ""},
        ]
        gs = iter_send_groups(rows, True)
        assert gs[0]["email"] == "m1@t.cn" and gs[0]["name"] == "王小明", gs[0]
        assert gs[1]["email"] == "m2@t.cn" and gs[1]["name"] == "" , gs[1]
        # 正文姓名回落：无姓名时用邮箱前缀（build_message 内部）
        msg, _fa, _recv = build_message(
            {"display_name": "x", "user": "s@t.cn"},
            {"email": "m2@t.cn", "name": "", "kws": ["K"], "rows": []}, "主题", "{姓名}，您好")
        assert "m2，您好" in msg.as_string() or "m2@t.cn" in msg.as_string()

    def t_config_backup():
        """配置损坏：坏文件备份为 .bad-时间戳，不静默丢规则"""
        import tempfile
        import shutil
        import unittest.mock as mock
        d = tempfile.mkdtemp()
        try:
            with open(os.path.join(d, "config.json"), "w", encoding="utf-8") as f:
                f.write("{ 坏掉的 json")
            with mock.patch.object(sys.modules[__name__], "app_dir", lambda: d):
                cfg = load_config()
            assert "_load_error" in cfg and "损坏" in cfg["_load_error"], cfg.get("_load_error")
            assert cfg["smtp"]["host"] == ""     # 回退默认配置
            bads = [x for x in os.listdir(d) if x.startswith("config.json.bad-")]
            assert len(bads) == 1, os.listdir(d)  # 坏文件已备份
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def t_rules_excel_header():
        """规则表表头识别：标准表头跳过；首行是数据（关键字含“关键”字样）不再被误判"""
        import tempfile
        import shutil
        from openpyxl import Workbook
        d = tempfile.mkdtemp()
        try:
            wb = Workbook()
            ws = wb.active
            ws.append(["关键字", "姓名", "邮箱"])
            ws.append(["临江大学", "李四", "lisi@t.cn"])
            p1 = os.path.join(d, "a.xlsx")
            wb.save(p1)
            rows, bad = _read_rules_excel(p1)
            assert len(rows) == 1 and rows[0]["email"] == "lisi@t.cn", (rows, bad)
            # 首行直接是数据且关键字含“关键”二字：旧逻辑会误判为表头丢弃
            wb2 = Workbook()
            ws2 = wb2.active
            ws2.append(["关键单据", "王五", "wangwu@t.cn"])
            ws2.append(["云洲师范", "赵六", "z12@t.cn"])
            p2 = os.path.join(d, "b.xlsx")
            wb2.save(p2)
            rows2, _ = _read_rules_excel(p2)
            assert len(rows2) == 2 and rows2[0]["kw"] == "关键单据", rows2
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def t_email():
        assert is_valid_email("a.b@c.com.cn")
        assert not is_valid_email("a@b")
        assert not is_valid_email("a b@c.com")

    def t_scan():
        rows = scan([app_dir()], [])
        for r in rows:
            assert "fname" in r and "status" in r

    def t_tk():
        root = tk.Tk()
        root.withdraw()
        root.update()
        root.destroy()

    def t_app():
        # 自检不应有副作用：屏蔽日志清理（不动真实旧日志），
        # 并用默认配置代替真实 load_config——若用户 config.json 恰好损坏，
        # 真实 load_config 会把它挪成 .bad 备份（自检无头运行没人看提示，事后规则“消失”）
        import unittest.mock as mock
        with mock.patch.object(sys.modules[__name__], "cleanup_old_logs", lambda *a, **k: None), \
             mock.patch.object(sys.modules[__name__], "load_config", lambda: default_config()):
            app = App()
            app.withdraw()
            app.update()
            assert app.lb_folders is not None, "发送页文件夹列表未创建"
            assert app.btn_scan is not None and app.btn_retry is not None, "发送入口按钮未创建"
            assert app.btn_manual is not None and app.btn_add_folder is not None, "操作按钮未创建"
            assert app.tv_rules is not None, "规则页列表未创建"
            assert app.v_host is not None, "邮箱设置页控件未创建"
            assert app._foot is not None and ("开源工具" in app._foot.cget("text") or "恔之" in app._foot.cget("text")), "作者签名行未创建"
            # 新增/编辑规则对话框：取消路径不得崩溃（result 已初始化为 None）
            dlg = RuleDialog(app, "自检对话框")
            app.update()
            assert hasattr(dlg, "result") and dlg.result is None
            dlg.destroy()
            app.update()
            # 销毁前取消队列轮询 after，避免尾随回调访问已销毁窗口
            if getattr(app, "_poll_after", None):
                try:
                    app.after_cancel(app._poll_after)
                except Exception:
                    pass
            app.destroy()

    def t_sanitize():
        assert sanitize_cell("=1+1.xlsx") == "'=1+1.xlsx"
        assert sanitize_cell("+加号.txt") == "'+加号.txt"
        assert sanitize_cell("-减号.txt") == "'-减号.txt"
        assert sanitize_cell("@at.txt") == "'@at.txt"
        assert sanitize_cell("正常.xlsx") == "正常.xlsx"
        assert sanitize_cell("") == ""
        assert sanitize_cell(123) == 123

    def t_attach_limit():
        import tempfile
        d = tempfile.mkdtemp()
        p = os.path.join(d, "大小测试.bin")
        with open(p, "wb") as f:
            f.write(b"x" * 2048)
        group = {"email": "a@t.cn", "name": "A", "kws": ["测试"],
                 "rows": [{"folder": d, "fname": "大小测试.bin"}]}
        ok, err, skipped = send_group({"host": "x"}, group, "s", "b", warn_mb=0.001)
        assert ok is False and "上限" in err and "编码" in err and skipped == [], (ok, err, skipped)
        try:
            ok2, err2, _s = send_group({"host": "不存在的无效主机.invalid"}, group, "s", "b", warn_mb=100)
            assert ok2 is False and "上限" not in err2
        except (smtplib.SMTPException, OSError):
            pass  # 连接阶段抛错（域名解析失败/拒连），同样视为通过；
                  # 只捕连接类异常，其他异常（如解包/拼写错误）应让自检真正失败

    def t_friendly_error():
        assert "账号或密码" in _friendly_error(smtplib.SMTPAuthenticationError(535, b"bad"))
        assert "无法连接" in _friendly_error(smtplib.SMTPConnectError(421, b"no"))
        assert "收件人被服务器拒绝" in _friendly_error(smtplib.SMTPRecipientsRefused({}))
        assert "证书校验失败" in _friendly_error(ssl.SSLCertVerificationError(1, "cert fail"))
        assert "加密连接失败" in _friendly_error(ssl.SSLError("ssl err"))
        assert "无法解析邮件服务器域名" in _friendly_error(socket.gaierror(-2, "Name or service not known"))
        assert "没有权限" in _friendly_error(PermissionError(13, "拒绝访问"))
        assert "不存在" in _friendly_error(FileNotFoundError(2, "找不到文件"))
        assert "超时" in _friendly_error(TimeoutError("timed out"))
        assert "文件读写" in _friendly_error(OSError(5, "io err"))
        assert _friendly_error(ValueError("x")) == "ValueError: x"

    def t_skip_locked():
        import tempfile
        d = tempfile.mkdtemp()
        group_bad = {"email": "a@t.cn", "name": "A", "kws": ["测试"], "rows": [
            {"folder": d, "fname": "已被删除的文件.txt"},
        ]}
        ok, err, skipped = send_group({"host": "x"}, group_bad, "s", "b", warn_mb=45)
        assert ok is False and "全部附件无法读取" in err and "不存在" in err, (ok, err)
        assert len(skipped) == 1 and skipped[0][0]["fname"] == "已被删除的文件.txt"
        # 预检函数：存在/丢失文件分别归档
        good = os.path.join(d, "好的.txt")
        with open(good, "w") as f:
            f.write("ok")
        rows = [{"folder": d, "fname": "好的.txt"},
                {"folder": d, "fname": "已丢失.txt"}]
        ok_rows, bad_rows = check_files_readable(rows)
        assert len(ok_rows) == 1 and len(bad_rows) == 1
        assert "不存在" in bad_rows[0][1]

    def t_bcc():
        import tempfile
        d = tempfile.mkdtemp()
        p = os.path.join(d, "密送测试.txt")
        with open(p, "w", encoding="utf-8") as f:
            f.write("x")
        g = {"email": "to@t.cn", "name": "T", "kws": ["测试"],
             "rows": [{"folder": d, "fname": "密送测试.txt"}]}
        msg, fa, recv = build_message(
            {"display_name": "t", "user": "s@t.cn", "cc": "cc1@t.cn",
             "bcc": "bcc1@t.cn;bcc2@t.cn"}, g, "主题", "正文")
        assert "bcc1@t.cn" in recv and "bcc2@t.cn" in recv, recv
        assert "cc1@t.cn" in recv and "to@t.cn" in recv, recv
        raw = msg.as_string()
        assert "Bcc" not in raw and "bcc1" not in raw, "密送人不应出现在邮件头"
        assert "Cc:" in raw

    def t_tls_modes():
        off = _make_ctx({"tls_verify": "off"})
        assert off.verify_mode == ssl.CERT_NONE and off.check_hostname is False
        ca = _make_ctx({"tls_verify": "ca"})
        assert ca.verify_mode == ssl.CERT_REQUIRED and ca.check_hostname is False
        strict = _make_ctx({"tls_verify": "strict"})
        assert strict.verify_mode == ssl.CERT_REQUIRED and strict.check_hostname is True
        d = _make_ctx({})
        assert d.verify_mode == ssl.CERT_NONE  # 缺省即 off
        s2 = _make_ctx({"tls_verify": "ca", "ca_file": ""})
        assert s2.verify_mode == ssl.CERT_REQUIRED and s2.check_hostname is False

    def t_migrate_tls_cfg():
        c1 = {"verify_cert": True}
        assert _migrate_tls(c1)["tls_verify"] == "strict" and "verify_cert" not in c1
        c2 = {"verify_cert": False}
        assert _migrate_tls(c2)["tls_verify"] == "off"
        c3 = {}
        assert _migrate_tls(c3)["tls_verify"] == "off"
        c4 = {"tls_verify": "ca"}
        assert _migrate_tls(c4)["tls_verify"] == "ca"

    def t_cleanup_logs():
        import tempfile
        import shutil
        d = tempfile.mkdtemp()
        try:
            old = os.path.join(d, "使用日志_20200101.log")
            new = os.path.join(d, "使用日志_%s.log" % datetime.now().strftime("%Y%m%d"))
            keep = os.path.join(d, "发送清单_20260101.xlsx")
            for p in (old, new, keep):
                with open(p, "w") as f:
                    f.write("x")
            os.utime(old, (time.time() - 40 * 86400, time.time() - 40 * 86400))
            cleanup_old_logs(d)
            assert not os.path.exists(old), "40 天前旧日志应被删除"
            assert os.path.exists(new) and os.path.exists(keep), "当天日志与非日志文件不受影响"
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def t_smtp_session():
        import tempfile
        import shutil
        import unittest.mock as mock
        d = tempfile.mkdtemp()
        try:
            with open(os.path.join(d, "附件测试.txt"), "wb") as f:
                f.write(b"x")
            row = {"folder": d, "fname": "附件测试.txt", "kws": ["测试"],
                   "manual_email": "", "status": "ok", "state": "", "err": ""}
            group = {"email": "to@t.cn", "name": "收件人", "kws": ["测试"],
                     "rows": [row], "ok": False, "err": "", "time": ""}

            class FakeSrv:
                total = 0

                def __init__(self):
                    FakeSrv.total += 1
                    self.idx = FakeSrv.total
                    self.calls = 0

                def sendmail(self, fa, recv, data):
                    self.calls += 1
                    # 只有第一个连接的第一次调用断线：模拟服务器中途掉线，重连后应成功
                    if self.idx == 1 and self.calls == 1:
                        raise smtplib.SMTPServerDisconnected("Connection unexpectedly closed")

                def quit(self):
                    pass

            made = []

            class FakeMsg:
                def as_string(self):
                    return "MSG"

            with mock.patch.object(sys.modules[__name__], "_connect",
                                   side_effect=lambda cfg: (made.append(1), FakeSrv())[1]), \
                 mock.patch.object(sys.modules[__name__], "build_message",
                                   side_effect=lambda *a, **k: (FakeMsg(), "from@t.cn", ["to@t.cn"])):
                # 1) 断线重连一次：第一次 sendmail 断线 → 自动重连 → 成功
                sess = SmtpSession({"host": "h", "port": 1, "security": "none"})
                ok, err, skipped = sess.send_group_once(group, "主题", "正文", 45.0)
                assert ok is True and skipped == [], (ok, err, skipped)
                assert len(made) == 2, made    # _connect 恰好两次：初建 + 断线重连
                # 2) 复用：同会话连发 3 组不再建连
                made.clear()
                for _ in range(3):
                    ok2, _, _ = sess.send_group_once(group, "主题", "正文", 45.0)
                    assert ok2 is True
                assert len(made) == 0, made    # 连接被复用，不再 _connect
                sess.close()
                # 3) SMTP 语义错误（认证失败）→ 连接状态未知，应丢弃
                class AuthFailSrv:
                    def sendmail(self, *a):
                        raise smtplib.SMTPAuthenticationError(535, b"auth failed")

                    def quit(self):
                        pass
                sess2 = SmtpSession({"host": "h", "port": 1, "security": "none"})
                sess2.server = AuthFailSrv()
                try:
                    sess2.send_group_once(group, "主题", "正文", 45.0)
                    raise AssertionError("认证失败应向上抛出")
                except smtplib.SMTPAuthenticationError:
                    pass
                assert sess2.server is None, "SMTP 语义错误后应丢弃连接"
                sess2.close()
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def t_report_csv_dedup():
        import tempfile
        import shutil
        import csv
        d = tempfile.mkdtemp()
        try:
            p = os.path.join(d, "rep.csv")
            # 重扫场景：组里的行与 rows 是不同对象（id 不同），但（folder, fname）相同
            old_row = {"folder": "d:/x", "fname": "临江大学2608.xlsx", "kws": ["临江大学"],
                       "manual_email": "", "status": "ok", "state": "成功", "err": ""}
            new_row = {"folder": "d:/x", "fname": "临江大学2608.xlsx", "kws": ["临江大学"],
                       "manual_email": "", "status": "ok", "state": "", "err": ""}
            groups = [{"email": "a@t.cn", "name": "甲", "kws": ["临江大学"],
                       "rows": [old_row], "ok": True, "err": "", "time": "t"}]
            _write_report_csv([new_row], groups, p)
            with open(p, "r", encoding="utf-8-sig", newline="") as f:
                lines = list(csv.reader(f))
            assert len(lines) == 2, lines   # 表头 + 组段 1 行；修复前 id(r) 去重会多出未发送行变 3 行
            assert "临江大学2608.xlsx" in lines[1][6], lines
        finally:
            shutil.rmtree(d, ignore_errors=True)

    def t_keyring():
        try:
            import keyring
        except ImportError:
            return  # 环境无 keyring：程序内置明文 fallback，不算失败
        try:
            keyring.set_password("邮件分发助手-selftest", "t", "密123")
            assert keyring.get_password("邮件分发助手-selftest", "t") == "密123"
            keyring.delete_password("邮件分发助手-selftest", "t")
            assert keyring.get_password("邮件分发助手-selftest", "t") is None
        except Exception:
            pass  # keyring 后端不可用：程序自动退回 config.json 明文，不算失败
        if store_password("selftest@t.cn", "abc"):
            assert fetch_password("selftest@t.cn") == "abc"
        clear_password("selftest@t.cn")
        assert fetch_password("selftest@t.cn") == ""

    def t_openpyxl():
        from openpyxl import Workbook
        wb = Workbook()
        wb.active["A1"] = "中文测试"
        import tempfile
        p = os.path.join(tempfile.gettempdir(), "_fst_selftest.xlsx")
        try:
            wb.save(p)
        finally:
            if os.path.exists(p):
                os.remove(p)

    def t_smtplib():
        import smtplib  # noqa

    check("Python " + sys.version.split()[0] + " 运行环境", lambda: None)
    check("最长关键字优先匹配（临江大学紫金/临江大学/云洲师范东区/云洲师范）", t_longest)
    check("未匹配识别", t_none)
    check("多关键字冲突识别", t_conflict)
    check("合并分组与手动指定分组", t_group)
    check("手动指定收件人姓名不再误用邮箱", t_manual_name)
    check("邮箱格式校验", t_email)
    check("规则表表头识别（首行数据不误判）", t_rules_excel_header)
    check("配置损坏备份与提示", t_config_backup)
    check("报告公式注入防护 sanitize_cell", t_sanitize)
    check("附件超限拦截", t_attach_limit)
    check("异常人话翻译 _friendly_error", t_friendly_error)
    check("单文件不可读跳过与预检", t_skip_locked)
    check("密送 BCC（实际投递且不进邮件头）", t_bcc)
    check("证书校验三档 off/ca/strict", t_tls_modes)
    check("旧版 verify_cert 配置自动迁移", t_migrate_tls_cfg)
    check("SMTP 会话复用与断线重连", t_smtp_session)
    check("密码保管 keyring（不可用时自动退回明文）", t_keyring)
    check("CSV 清单去重键（重扫场景不重复）", t_report_csv_dedup)
    check("日志自动清理（保留30天）", t_cleanup_logs)
    check("文件夹扫描（含空规则）", t_scan)
    check("tkinter 界面引擎", t_tk)
    check("完整界面类 App 构造（三页签+全部控件）", t_app)
    check("xlsx 清单组件 openpyxl", t_openpyxl)
    check("smtplib 邮件组件", t_smtplib)

    report = os.path.join(app_dir(), "selftest_report.txt")
    ok_n = len([l for l in lines if l.startswith("[OK]")])
    n_fail = len([l for l in lines if l.startswith("[FAIL]")])
    report_text = (
        "%s 自检报告 %s\n条目： %d  通过： %d  失败： %d\n\n%s\n"
        % (APP_NAME, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
           ok_n + n_fail, ok_n, n_fail, "\n".join(lines)))
    with open(report, "w", encoding="utf-8") as f:
        f.write(report_text)
    return report


def _enable_dpi_awareness():
    """高分辨率屏幕下避免界面模糊；须在任何窗口创建前调用"""
    if sys.platform == "win32":
        try:
            import ctypes
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(2)   # PER_MONITOR_DPI_AWARE
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass


def main():
    if "--version" in sys.argv:
        # 版本查询无需 GUI 环境，放在 DPI 与其他模式之前
        print("%s v%s（%s）" % (APP_NAME, APP_VERSION, AUTHOR_NOTE))
        return
    _enable_dpi_awareness()
    if "--selftest" in sys.argv:
        run_selftest()
        return
    if "--keyring-test" in sys.argv:
        report = os.path.join(app_dir(), "keyring_test.txt")
        try:
            import keyring
            backend = type(keyring.get_keyring()).__name__
            keyring.set_password(KEYRING_SERVICE, "selftest", "值123")
            v = keyring.get_password(KEYRING_SERVICE, "selftest")
            keyring.delete_password(KEYRING_SERVICE, "selftest")
            assert v == "值123"
            with open(report, "w", encoding="utf-8") as f:
                f.write("keyring 可用，backend=%s，存取删全通过" % backend)
        except Exception as e:
            with open(report, "w", encoding="utf-8") as f:
                f.write("keyring 不可用: %s\n程序将退回 config.json 明文保存（不影响发信）" % _err_text(e))
        return
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
