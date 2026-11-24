# MailSender · 邮件分发助手（Python 源码版）

本文件夹为源码版本。**不想碰 Python 的同学请到仓库根的 [`FileSender-EXE/`](../FileSender-EXE/) 文件夹下载免安装 exe，双击即用。**

## 从源码运行 / 打包

依赖：Python 3.8+（开发环境 3.12）；`keyring`、`openpyxl` 为可选依赖（缺失自动降级，不装也能跑）。

```powershell
# 源码直接运行（可选依赖缺失时功能自动降级）
python file_sender-python.py

# 建议在隔离目录安装依赖，不污染全局环境
python -m pip install --target ./build-env keyring openpyxl pyinstaller

# 打包单文件 exe（与本仓库 FileSender-EXE/ 里的完全同源）
$env:PYTHONPATH = (Resolve-Path ./build-env).Path
python -m PyInstaller --onefile --windowed --name FileSender file_sender-python.py
```

产物在 `dist/FileSender.exe`，拷到任意文件夹双击即用（配置与清单生成在 exe 同目录）。

## 测试

```powershell
$env:PYTHONPATH = (Resolve-Path ./build-env).Path
# 内置自检（25 项：运行环境/匹配逻辑/分组/模板/密送/TLS/会话复用/配置备份/清单去重/界面构造…）
python file_sender-python.py --selftest
# UI 级端到端测试（13 场景：FakeSMTP 全流程、重试、熔断、超限拦截、补发闭环…）
python ui_e2e_test.py
```

端到端测试自带本地假 SMTP 服务器（127.0.0.1:25253），不发真实邮件。

## 功能特性

| 特性 | 说明 |
|---|---|
| 最长关键字优先 | 同时配置「临江大学」与「临江大学紫金」时，文件名命中更长者；零命中或等长命中多人自动停发待确认，绝不乱猜 |
| 合并发送 | 同一收件人多个文件默认合并一封（可切为逐文件） |
| 发送前确认 | 每次发送先弹收件清单过目，确认后才发出 |
| 全程留痕 | 每轮发送自动生成「发送清单」Excel（xlsx，缺 openpyxl 时自动降级 csv）：谁收到什么、哪些没发、失败原因 |
| 失败重发 | 只补发失败项；文件被占用/丢失的附件单独进待补发池，解锁后一键补发 |
| SMTP 会话复用 | 一轮发送只登录一次，断线自动重连，连续 3 次连接失败自动熔断（防企业邮箱风控） |
| 密码保管 | Windows 凭据管理器（keyring），不可用时经用户确认后回退配置文件明文 |
| TLS 三档 | 不校验（内网自签）/ 校验证书链 / 严格校验（公网） |
| 多重防护 | 附件按编码后大小预检拦截（理论膨胀 4/3，取 1.37 留余量）、危险后缀（exe/bat/lnk 等）自动跳过、Excel 公式注入防护 |
| 优雅退出 | 发送中关窗有三选一确认，进度与清单不静默丢失 |

## 目录结构

```
file_sender-python.py   # 主程序（单文件，含 GUI / 匹配 / 发送 / 自检；文件名带 -python，标明需要 Python 环境）
ui_e2e_test.py           # UI 级端到端测试
使用说明.html             # 面向执行人的完整使用说明（规则、SMTP 参数、授权码指引、FAQ）
```

## 设计说明（写给想改造的开发者）

- **tk 经典线程模型**：发送 worker 只通过 `queue.Queue` 回传结果，主线程统一更新 UI；`after` 轮询驱动
- **失败隔离**：单组失败不中断整轮；`partial` 标记区分「邮件没发出」与「邮件发出但缺附件」
- **业务主键**：报告去重用 `(folder, fname)` 而非内存 `id()`，跨重扫/重试轮次稳定
- **自检无副作用**：`--selftest` 屏蔽日志清理、用默认配置，不触碰用户真实配置与日志
- **仅依赖标准库即可运行**，第三方包全部可选降级，保证在净 Windows 机器上开箱可用

## License

MIT，见[仓库根 LICENSE](../LICENSE)。

> 作者：恔之 ｜ AI生成

> AI生成
