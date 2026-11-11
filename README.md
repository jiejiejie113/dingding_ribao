# 钉钉生产日报自动化

从 MySQL 导出宜搭业务表，合并生成 Excel 数据透视表及生产简报，通过钉钉发送至指定群，支持 Windows 每日定时任务。

## 功能

- 全库逐表流式导出，五张业务表以中文命名。
- 以跟单为基准合并纸样分配、纸样、裁床、车位数据。
- 业务键去重、超期未分配剔除、处理时长计算及件数异常标色。
- Microsoft Excel 自动刷新模板中的七张透视表和简报。
- 通过 `dinging-message-api` 的 dws 封装发送 Excel 和 Markdown 日报。
- 每天08:50准备，09:00发送；按日期、目标群和文件/简报分别记录发送状态。
- 数据库密码由当前 Windows 用户的凭据管理器保存，不写入源码。

## 环境准备

需要 Windows、Python 3.10+、Microsoft Excel，以及已安装并登录的 dws CLI。

```powershell
git clone --recurse-submodules https://github.com/jiejiejie113/dingding_ribao.git
python -m pip install -r requirements.txt
dws auth login --device --no-browser --recommend
```

如果普通克隆没有下载消息框架：

```powershell
git submodule update --init --recursive
```

## 本机业务规则与模板

当前处理逻辑引用以下本机文件，必须自行准备；模板可能包含业务数据，因此不上传仓库：

```text
C:\Users\caohanjie\Desktop\合并转透视表\数据透视表1.6.py
C:\Users\caohanjie\Desktop\合并转透视表\模板.xlsx
```

在其他电脑运行时，请调整 `process_excel.py` 中的 `REFERENCE`、`TEMPLATE`，并配置对应文件。项目目前依赖该参考程序提供字段清单、日期计算及 Excel COM 重试逻辑。

## 运行入口

| 入口 | 功能 |
|---|---|
| `开始导出.bat` | 导出、合并、生成报表，不发群 |
| `仅合并生成透视表.bat` | 使用当前 Excel 重新生成报表 |
| `发送最新报表到群.bat` | 生成最新报表并发送，同日同群完成后跳过 |
| `python daily_send.py --dry-run` | 完整准备和预览，不真实发送 |
| `python daily_send.py --once-at YYYY-MM-DDTHH:MM` | 当天指定时间额外发送一批 |
| `python verify_report.py` | 只读验证当天标准文件名报表 |

MySQL连接默认值在 `export_mysql.py`，发送群稳定ID在 `daily_send.py` 的 `GROUP`。首次运行导出时输入密码，将保存到 Windows 凭据管理器；双击入口默认用户名为 `yida_ro`。

## 定时任务

```powershell
powershell -ExecutionPolicy Bypass -File .\register_schedule.ps1
```

创建 `YidaExcelDailySend`。需要电脑开机、用户保持登录且内网可达。处理超过09:00会完成后发送。此任务依赖本机 Excel 交互会话，不是无桌面的服务器部署方案。

发送记录及附件保存在 `send-logs`，导出文件保存在 `excel-dingding`，均被 Git 忽略。文件发送和文字消息是两条独立群消息；dws 返回成功代表提交发送成功，实际接收请在群中确认。

## 统计口径

- 日流量取昨日筛选后的透视表总计；周一取周六、周日。
- 待完成纸样仅包括“进行中、未处理”，兼容数字前缀。
- 超过10天：提交纸样部日期至当前统计日超过10天。
- 车办改转态日期使用数据库 `modified_time`，不是车衣完成时间。
- 重复业务键按首条保留；数据库导出未指定排序，未定义最新记录优先规则。

详细操作参见 [使用说明.md](使用说明.md)。
