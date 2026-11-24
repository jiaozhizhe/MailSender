# -*- coding: utf-8 -*-
"""UI 级端到端测试：完整用户流程回归（skipped收缩/无重复/无空组/重试去重版）"""
import os
import sys
import json
import time
import socket
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "build-env"))
# 主程序文件名含连字符（file_sender-python.py），不能直接 import，按路径加载
import importlib.util as _ilu
_spec = _ilu.spec_from_file_location(
    "fs_main", os.path.join(HERE, "file_sender-python.py"))
fs = _ilu.module_from_spec(_spec)
_spec.loader.exec_module(fs)

TEST_DIR = os.path.join(HERE, "ui_test_env")
PORT = 25253
HOST = "127.0.0.1"

received = []
conn_count = [0]
lock = threading.Lock()


class FakeSMTP(threading.Thread):
    def run(self):
        srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        srv.bind((HOST, PORT))
        srv.listen(8)
        srv.settimeout(120)
        try:
            while True:
                try:
                    conn, _ = srv.accept()
                    with lock:
                        conn_count[0] += 1
                except socket.timeout:
                    break
                threading.Thread(target=self.handle, args=(conn,), daemon=True).start()
        finally:
            srv.close()

    @staticmethod
    def _line(f):
        return f.readline().decode("utf-8", "replace").rstrip("\r\n")

    def handle(self, conn):
        f = conn.makefile("rb")
        def send(s):
            conn.sendall((s + "\r\n").encode())
        send("220 fake ready")
        mail_from, rcpts = "", []
        try:
            while True:
                line = self._line(f)
                if not line:
                    break
                up = line.upper()
                if up.startswith(("EHLO", "HELO")):
                    send("250-fake"); send("250 OK")
                elif up.startswith("AUTH"):
                    send("235 ok")
                elif up.startswith("MAIL FROM"):
                    mail_from = line; send("250 ok")
                elif up.startswith("RCPT TO"):
                    rcpts.append(line); send("250 ok")
                elif up == "DATA":
                    send("354 go")
                    chunks = []
                    while True:
                        raw = self._line(f)
                        if raw == ".":
                            break
                        chunks.append(raw)
                    with lock:
                        received.append((mail_from, rcpts[:], "\n".join(chunks)))
                    rcpts = []
                    send("250 accepted")
                elif up == "QUIT":
                    send("221 bye"); break
                else:
                    send("250 ok")
        except Exception:
            pass
        finally:
            try:
                conn.close()
            except Exception:
                pass


def pump(app, until=None, timeout=25):
    end = time.time() + timeout
    while time.time() < end:
        if until is not None and until():
            for _ in range(8):
                app.update()
            return True
        app.update()
        time.sleep(0.02)
    return False


def worker_done(app):
    t = app.send_thread
    if t is not None and t.is_alive():
        return False
    if not app.q.empty():
        return False
    return True


def latest_report():
    d = os.path.join(TEST_DIR, "发送清单")
    files = [os.path.join(d, x) for x in os.listdir(d)]
    return max(files, key=os.path.getmtime)


def main():
    os.makedirs(TEST_DIR, exist_ok=True)
    for x in os.listdir(TEST_DIR):
        p = os.path.join(TEST_DIR, x)
        if os.path.isdir(p):
            import shutil
            shutil.rmtree(p)
        else:
            os.remove(p)
    files_dir = os.path.join(TEST_DIR, "待发文件")
    os.makedirs(files_dir, exist_ok=True)
    test_files = {
        "临江大学紫金2608.txt": "内容A",
        "临江大学紫金2609.txt": "内容B",
        "云洲师范东区报告单.pdf": "内容C",
        "临江大学2607.docx": "内容D",
        "无关文件.txt": "不该发",
        "=公式测试.xlsx": "公式注入文件名",
    }
    for fn, c in test_files.items():
        with open(os.path.join(files_dir, fn), "w", encoding="utf-8") as f:
            f.write(c)

    cfg = {
        "folders": [files_dir],
        "rules": [
            {"kw": "临江大学", "name": "李四", "email": "lisi@t.cn"},
            {"kw": "临江大学紫金", "name": "张三", "email": "zhangsan@t.cn"},
            {"kw": "云洲师范东区", "name": "王五", "email": "wangwu@t.cn"},
        ],
        "send": {"merge": True, "recursive": False, "interval": 0, "attach_warn_mb": 45},
        "smtp": {
            "host": HOST, "port": PORT, "security": "none", "auth": False,
            "user": "", "passwd": "", "from_addr": "relay@t.cn",
            "display_name": "测试", "cc": "", "bcc": "",
            "subject_tpl": "【{关键字}】（{文件数}份）",
            "body_tpl": "{姓名}：{关键字} 共{文件数}份\n{文件名}", "timeout": 8,
            "tls_verify": "off", "ca_file": "",
        },
        "remember_passwd": False,
    }
    with open(os.path.join(TEST_DIR, "config.json"), "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)

    fs.app_dir = lambda: TEST_DIR
    fs.messagebox.askyesno = lambda *a, **k: True
    fs.messagebox.askyesnocancel = lambda *a, **k: True
    fs.messagebox.showinfo = lambda *a, **k: None
    fs.messagebox.showwarning = lambda *a, **k: None
    fs.messagebox.showerror = lambda *a, **k: None
    os_startfile = fs.os.startfile
    fs.os.startfile = lambda p: None

    server = FakeSMTP()
    server.daemon = True
    server.start()
    time.sleep(0.3)

    app = fs.App()
    app.withdraw()
    app.update()
    assert app.v_host.get() == HOST

    # ---- 1) 扫描 + 正常发送 ----
    app.do_scan()
    app.update()
    assert len(app.rows) == 6
    by = {r["fname"]: r for r in app.rows}
    app.do_send()
    assert pump(app, until=lambda: worker_done(app))
    assert all(g.get("ok") for g in app.groups_done)
    assert len(app.groups_done) == 3
    time.sleep(0.3)
    with lock:
        assert len(received) == 3, len(received)
        assert conn_count[0] == 1, conn_count[0]   # 会话复用：3 封邮件只登录一次
    print("[OK] 正常发送：3 封成功，全程只建 1 个 SMTP 连接")

    # ---- 2) ok=False + skipped 回归：主组收缩、无重复行、无空组 ----
    app.do_scan()
    app.update()
    by3 = {r["fname"]: r for r in app.rows}
    rowA = by3["临江大学紫金2608.txt"]
    rowB = by3["临江大学紫金2609.txt"]
    rowC = by3["临江大学2607.docx"]
    gA = {"email": "z9@t.cn", "name": "回归九", "kws": ["临江大学紫金"], "rows": [rowA, rowB]}
    app.groups_done.append(gA)   # 模拟真实流程：主组先挂入清单，worker 结果回来再由 _apply 处理
    app._active_groups = [gA]
    app._apply_send_result(id(gA), False, "模拟SMTP错误",
                            [(rowA, "文件被占用（测试）")])
    app.update()
    assert [r["fname"] for r in gA["rows"]] == ["临江大学紫金2609.txt"], \
        [r["fname"] for r in gA["rows"]]   # 修复前被跳过的A仍留在主组
    assert all(g["rows"] for g in app.groups_done), "清单中不允许出现空组"
    ids = [id(r) for g in app.groups_done for r in g["rows"]]
    assert len(ids) == len(set(ids)), "任何行不允许在多个组中重复"
    # 报告无重复：z9 收件人的组段中 2608 恰好 1 次（未收缩会 2 次）
    tmp_rep = os.path.join(TEST_DIR, "回归报告.xlsx")
    fs.write_report(app.rows, app.groups_done, tmp_rep)
    from openpyxl import load_workbook
    wb = load_workbook(tmp_rep); ws = wb.active
    n_2608_z9 = sum(1 for r in list(ws.iter_rows(values_only=True))[1:]
                    if r[6] and "2608" in str(r[6]) and str(r[2]) == "z9@t.cn")
    assert n_2608_z9 == 1, n_2608_z9
    wb.close()
    # 空主组源头移除：gB 全部附件被跳过 → gB 从清单消失、整组转为跳过组
    gB = {"email": "z10@t.cn", "name": "回归十", "kws": ["临江大学"], "rows": [rowC]}
    app.groups_done.append(gB)
    app._active_groups = [gB]
    app._apply_send_result(id(gB), False, "模拟全部失败",
                            [(rowC, "文件被占用（测试）")])
    app.update()
    assert gB["rows"] == []
    assert all(x is not gB for x in app.groups_done), "空主组应从清单移除"
    assert all(g["rows"] for g in app.groups_done)
    print("[OK] skipped 收缩/无重复/无空组：主组与跳过组不交叉、空主组被移除")

    # ---- 3) 重试去重：人为塞重复组与空组，重发应只发 3 封 ----
    app.groups_done.append({
        "email": "z9@t.cn", "name": "回归九", "kws": ["临江大学紫金"],
        "rows": [rowA], "ok": False,
        "err": "附件无法读取（文件被占用或已删除）", "time": ""})   # 与既有跳过组重复
    app.groups_done.append({
        "email": "z9@t.cn", "name": "空组", "kws": [],
        "rows": [], "ok": False, "err": "x", "time": ""})
    before9 = len(received)
    app.retry_failed()
    assert pump(app, until=lambda: worker_done(app) and
                all(g.get("ok") for g in app.groups_done)), "重试超时"
    time.sleep(0.4)
    with lock:
        delta9 = len(received) - before9
    assert delta9 == 3, delta9   # 未去重会发 4 封、未滤空组会更多；gA/skipA/skipB 各 1 封
    assert all(g.get("ok") for g in app.groups_done)
    assert all(g["rows"] for g in app.groups_done)
    ids = [id(r) for g in app.groups_done for r in g["rows"]]
    assert len(ids) == len(set(ids))
    print("[OK] 重试去重与空组过滤：实际补发 %d 封（重复组与空组被排除）" % delta9)

    # ---- 4) 坏端口 → 重发回归 ----
    time.sleep(1.1)
    app.v_port.set(59999)
    app.do_send()
    assert pump(app, until=lambda: worker_done(app))
    failed = [g for g in app.groups_done if not g.get("ok")]
    assert len(failed) == 3 and all(g.get("err") for g in failed)
    time.sleep(1.1)
    app.v_port.set(PORT)
    app.retry_failed()
    assert pump(app, until=lambda: worker_done(app) and
                all(g.get("ok") for g in app.groups_done))
    with lock:
        assert len(received) == 3 + 3 + 3, len(received)
    rep3 = latest_report()
    wb = load_workbook(rep3); ws = wb.active
    st3 = [r[7] for r in list(ws.iter_rows(values_only=True))[1:]]
    assert st3.count("成功") == 4 and st3.count("失败") == 0, st3
    wb.close()
    print("[OK] 坏端口全失败后重发，报告完整")

    # ---- 4.5) 连续连接失败提前中止：4 组中只试 3 组 ----
    time.sleep(1.1)
    app.cfg["rules"].append({"kw": "临江体院", "name": "赵六", "email": "z12@t.cn"})
    with open(os.path.join(files_dir, "临江体院2606.txt"), "w", encoding="utf-8") as f:
        f.write("内容F")
    app.do_scan()
    app.update()
    assert len(app.rows) == 7, len(app.rows)   # 原 6 个测试文件 + 临江体院2606
    app.v_port.set(59999)
    app.do_send()
    assert pump(app, until=lambda: worker_done(app))
    failed45 = [g for g in app.groups_done if not g.get("ok")]
    handled = [g for g in failed45 if g.get("err")]
    assert len(failed45) == 4, len(failed45)      # 4 组都在清单（do_send 全量挂入）
    assert len(handled) == 3, [g.get("err") for g in failed45]  # 但只有前 3 组发了（第 4 组被提前中止）
    log_text = app.log.get("1.0", "end")
    assert "连续 3 次" in log_text, "日志应有提前中止提示"
    time.sleep(1.1)
    app.v_port.set(PORT)
    app.retry_failed()
    assert pump(app, until=lambda: worker_done(app) and
                all(g.get("ok") for g in app.groups_done))
    app.cfg["rules"] = [r for r in app.cfg["rules"] if r.get("kw") != "临江体院"]
    os.remove(os.path.join(files_dir, "临江体院2606.txt"))
    print("[OK] 连续 3 次连接失败提前中止；重试全部成功")

    # ---- 5) 附件超限（编码后口径） ----
    time.sleep(1.1)
    app.do_scan()      # 场景4.5 已删除临江体院文件，重扫让 rows 与磁盘一致
    app.update()
    app.cfg["send"]["attach_warn_mb"] = 0.000001
    app.do_send()
    assert pump(app, until=lambda: worker_done(app))
    errs = [g.get("err", "") for g in app.groups_done]
    assert all("上限" in e and "编码" in e for e in errs), errs
    time.sleep(1.1)
    app.cfg["send"]["attach_warn_mb"] = 45
    app.retry_failed()
    assert pump(app, until=lambda: worker_done(app) and
                all(g.get("ok") for g in app.groups_done))
    print("[OK] 附件超限（编码口径）拦截与恢复")

    # ---- 6) 文件丢失预检剔除与补发闭环 ----
    time.sleep(1.1)
    lost = os.path.join(files_dir, "临江大学紫金2610.txt")
    with open(lost, "w", encoding="utf-8") as f:
        f.write("内容E")
    app.do_scan()
    app.update()
    assert len(app.rows) == 7
    by2 = {r["fname"]: r for r in app.rows}
    os.remove(lost)
    app.do_send()
    assert pump(app, until=lambda: worker_done(app))
    gs_ok = [g for g in app.groups_done if g.get("ok")]
    gs_fail = [g for g in app.groups_done if not g.get("ok")]
    assert len(gs_ok) == 3 and len(gs_fail) == 1
    assert gs_fail[0]["err"].startswith("附件无法读取")
    assert all(g["rows"] for g in app.groups_done)
    time.sleep(1.1)
    with open(lost, "w", encoding="utf-8") as f:
        f.write("内容E-恢复")
    app.retry_failed()
    assert pump(app, until=lambda: worker_done(app) and
                all(g.get("ok") for g in app.groups_done))
    rep7 = latest_report()
    wb = load_workbook(rep7); ws = wb.active
    st7 = [r[7] for r in list(ws.iter_rows(values_only=True))[1:]]
    assert st7.count("成功") == 5 and st7.count("失败") == 0, st7
    wb.close()
    print("[OK] 文件丢失预检剔除与补发闭环")

    # ---- 6.5) partial 统计口径：主邮件成功 + 部分附件被跳过 → 不计入“失败 N 封” ----
    app.do_scan()
    app.update()
    by4 = {r["fname"]: r for r in app.rows}
    gC = {"email": "z11@t.cn", "name": "回归十一", "kws": ["回归"],
          "rows": [by4["临江大学紫金2608.txt"], by4["云洲师范东区报告单.pdf"]],
          "ok": False, "err": "", "time": ""}
    app.groups_done.append(gC)
    app._active_groups = [gC]
    app._apply_send_result(id(gC), True, "",
                            [(by4["云洲师范东区报告单.pdf"], "文件被占用（测试）")])
    app.update()
    assert [r["fname"] for r in gC["rows"]] == ["临江大学紫金2608.txt"]
    skips_c = [g for g in app.groups_done if g.get("partial")]
    assert len(skips_c) == 1 and len(skips_c[0]["rows"]) == 1, \
        [ (g["email"], g.get("partial"), len(g["rows"])) for g in app.groups_done ]
    captured = {}
    orig_ask = fs.messagebox.askyesno
    fs.messagebox.askyesno = lambda *a, **k: (captured.setdefault("msg", a[1]) or True)
    app._finish_send()
    fs.messagebox.askyesno = orig_ask
    m = captured.get("msg", "")
    assert "成功 5 封" in m, m
    assert "其中 1 封有 1 个附件" in m, m
    assert "失败 0 封" in m, m
    # 报告去重：重扫后同一 (folder,fname) 不允许在“未发送”段重复出现
    rep_c = latest_report()
    wb = load_workbook(rep_c); ws = wb.active
    hits = [c for row in list(ws.iter_rows(values_only=True))
            for c in row if c is not None and "临江大学紫金2608" in str(c)]
    # 紫金组 2 个附件各一行（文件名列显示全组名）+ gC 1 行 = 3；
    # 单独 cell（纯文件名）只应 1 个 = gC；修复前 id(r) 去重会让“未发送”段再列 1 次
    assert len(hits) == 3, hits
    n_exact = sum(1 for c in hits if str(c) == "临江大学紫金2608.txt")
    assert n_exact == 1, hits
    wb.close()
    print("[OK] partial 统计口径：成功邮件缺附件不计入失败封数；报告去重键改 (folder,fname)")

    # ---- 7) 规则对话框取消 ----
    d = fs.RuleDialog(app, "取消测试")
    app.update()
    assert d.result is None
    d.destroy()
    app.update()
    print("[OK] 规则对话框取消不崩溃")

    # ---- 8) keyring ----
    fs.clear_password("tester@t.cn")
    app.v_user.set("tester@t.cn")
    app.v_pwd.set("secret123")
    app.v_remember.set(True)
    app.save_smtp()
    app.update()
    assert app.cfg["smtp"]["passwd"] == ""
    app2 = fs.App()
    app2.withdraw()
    app2.update()
    assert app2.v_pwd.get() == "secret123"
    app2.v_remember.set(False)
    app2.save_smtp()
    app2.update()
    assert fs.fetch_password("tester@t.cn") == ""
    app2.destroy()
    app.update()
    print("[OK] keyring 密码保管回归")

    # ---- 8.5) 邮箱设置页改动未保存标记 ----
    app._dirty = False
    app.v_host.set("10.1.2.3")
    app.update()
    assert app._dirty is True, "邮箱设置页字段改动应标记未保存（关窗时会提醒）"
    app._dirty = False
    app.tb_body.insert("1.0", "x")   # 正文模板键入也应标脏
    app.update()
    assert app._dirty is True
    app._dirty = False
    print("[OK] 邮箱设置页改动标记未保存")

    # ---- 8.6) ca 模式预检 CA 文件存在性 ----
    s_bad = {"host": "10.0.0.1", "auth": False, "user": "", "passwd": "",
             "from_addr": "a@t.cn", "tls_verify": "ca",
             "ca_file": "Z:\\不存在_的CA证书.pem"}
    err_c = app._check_smtp(s_bad)
    assert err_c.startswith("CA证书文件不存在"), err_c
    s_ok = dict(s_bad, ca_file="")
    assert app._check_smtp(s_ok) == ""   # 留空（用系统内置证书库）不拦截
    print("[OK] ca 模式预检 CA 文件存在性")

    app.destroy()
    fs.os.startfile = os_startfile
    print("\n===== UI 级端到端测试：全部断言通过 =====")


if __name__ == "__main__":
    main()



