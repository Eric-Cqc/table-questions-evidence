# 表格问数与逐数证据

[![GitHub stars](https://img.shields.io/github/stars/Eric-Cqc/table-questions-evidence?style=flat)](https://github.com/Eric-Cqc/table-questions-evidence/stargazers) · [反馈问题](https://github.com/Eric-Cqc/table-questions-evidence/issues/new/choose) · [查看全部作品](https://github.com/Eric-Cqc)

**把一个明确的表格问题，变成可复算、有来源单元格的答案。**

例如：“各渠道订单金额合计和平均值是多少，缺失金额别当零？”这个工具读取原生命名 Excel Table，按确定的口径计算，并输出每个数字的来源行、有效样本数和缺失排除记录。原文件改变后，旧计划和旧答案都会失效。

它是供宿主 Agent 调用的离线计算工具：**宿主理解中文、确认业务口径和生成计划；Python 负责计算与证据。** 不含内置模型，也不自动理解任意自然语言。运行仅需 Python 3.10+ 标准库，不需 API、数据库或网络服务。

## 一分钟看懂示例

公开样例是完全构造的五笔订单，不含个人或真实研究数据。

| 渠道 | 订单数 | 已知金额合计 | 已知金额均值 | 有金额 / 全部订单 |
|---|---:|---:|---:|---:|
| Web | 2 | 400 CNY | 200 CNY | 2 / 2 |
| Store | 3 | 200 CNY | 100 CNY | 2 / 3 |

Store 的空金额没有当成零；均值分母是 2，订单数仍为 3。均值证据指向 `Orders!C4` 和 `Orders!C6`，缺失行为第 5 行。

## 运行

从仓库根目录执行；输出目录必须尚不存在。

```sh
python3 src/table_analysis.py inspect examples/orders.xlsx
python3 src/table_analysis.py run examples/plan.json --out demo-result
python3 src/table_analysis.py verify demo-result/result.json --source examples/orders.xlsx
python3 tests/test_workflow.py
```

也可使用 `bash analyze` 替代 `python3 src/table_analysis.py`。计划内的相对路径以运行目录为基准。公开计划可直接运行；自己的问题应交给宿主按 [HOST_WORKFLOW.md](HOST_WORKFLOW.md) 生成计划，不需要终端用户编写 JSON。

输出包含：

- `answer.md`：中文结果、口径、样本数与证据 ID。
- `results.csv`：可排序的结果长表。
- `result.json`：完整精度、源格、每行贡献、排除原因，以及源文件和计划哈希。

运行后先看 `answer.md`，再通过 `result.json` 追踪数字。`verify` 会重新计算，而不是只检查文件存在。

## 做了哪些设计

计算使用 Decimal，明确区分求和、均值、按组与全样本分母。缺失处理必须指定；空均值与零分母显示不可计算。所有计划绑定源文件哈希；源变化必须重新检查数据和口径，不能简单替换哈希继续用旧结论。

支持筛选、最多三个字段分组、行表达式、计数/去重/求和/均值/极值、差分、比例与 Top N。Top N 不缩小全样本分母。公式默认拒绝，只能明确采用已保存缓存并披露限制；工具不执行公式或宏。

## 当前验证与限制

公开版端到端测试使用新生成的合成工作簿，核对计算结果、有效分母和单元格证据，执行独立重算，确认已有输出、未解决口径、错误缺失政策、源变化都会被拒绝，并检查 Markdown 动态文字不会成为链接或 HTML。测试不依赖原研究工作簿。

当前只支持原生命名 Excel Table，不支持任意排版区域、跨表连接、自动日期解释、预测或因果推断。业务含义仍需宿主和使用者确认。输出目录应由单个写入者使用，本版未提供并发作业协调。

工具不联网、不改原表；但输出 JSON 会保留计划路径、原始字段与用到的单元格，**不是自动脱敏包**。分享真实结果前检查权限和内容。Markdown 动态文本已转义，CSV 危险前缀被转义；JSON 保留原文。哈希用于检测变化，不证明数据真实性或决定者身份。

本仓库没有附带再许可声明。公开展示不等于授予任何额外使用许可。

## 使用、反馈与关注

如果这个项目对你有帮助，欢迎点击仓库右上角 **Star**；想跟进版本更新，可用 **Watch → Custom → Releases**。遇到问题请提交复现步骤、预期与实际结果；也欢迎分享希望解决的业务问题。请只使用合成或已脱敏示例，不提交密钥、个人信息或真实客户资料。

查看 [参与说明](CONTRIBUTING.md)。
