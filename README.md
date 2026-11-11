# 钉钉生产日报自动化

从 MySQL 导出宜搭业务表，合并生成 Excel 数据透视表及生产简报，通过钉钉发送至指定群，支持 Windows 每日定时任务。

## 功能

- 全库逐表流式导出，五张业务表以中文命名。
- 以跟单为基准合并纸样分配、纸样、裁床、车位数据。
- 业务键去重、超期未分配剔除、处理时长计算及件数异常标色。
- Microsoft Excel 自动刷新月份模板中的11张透视表和简报，包含4个年月字段。
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
项目目录\1.7数据透视表（月份）.py
项目目录\2026-10-16 模板.xlsx
```

项目依赖本目录的1.7参考程序提供40列字段清单、日期计算及 Excel COM 重试逻辑，模板需自行准备。数据按件数异常置顶、提交日期从新到旧排序。

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

- 周一生成上周一至周日的上周周报，其余日期生成昨日生产日报；四张月份表保留全量月份统计，不加入日报/周报日期筛选。手头未完成及处理天数保留当前存量口径，不能还原过去某天的历史状态。
- 待完成纸样仅包括“进行中、未处理”，兼容数字前缀。
- 超过10天：提交纸样部日期至当前统计日超过10天。
- 车办改转态日期使用数据库 `modified_time`，不是车衣完成时间。
- 重复业务键按首条保留；数据库导出未指定排序，未定义最新记录优先规则。

详细操作参见 [使用说明.md](使用说明.md)。
