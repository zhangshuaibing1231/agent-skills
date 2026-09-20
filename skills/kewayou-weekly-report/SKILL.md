---
name: kewayou-weekly-report
description: 将客无忧销售与售后日报去重汇总，按固定新模板生成并校验 Word 周报。适用于任何支持 Agent Skills 或可运行 Python 脚本的智能体；不用于普通日报润色或无模板的通用工作总结。
---

# 客无忧周报生成

把一周日报整理为可直接提交的 Word 周报，保留新模板的章节、表格、底色、边框和字体层级。

## 输入与判断

- 接受用户直接粘贴的多天日报，也接受包含日报的本地文本文件。
- 优先采用用户明确给出的周报周期；未给出时，根据日报语境和当前日期确定最近一个周一至周五。若日期确实无法判断，再询问用户。
- 不把“准备收款、催打款、做确认单、品牌继续跟进”提前算作成交。只有明确出现“已收款、已付款、已打款、成交完成”等结果，才进入成交客户。
- 同一客户在多天出现时合并为一条，以最新状态为准，同时保留影响决策的关键过程。
- 暂缓、考虑中、未回复不等于流失。只有明确转向其他渠道、拒绝合作或终止项目，才进入流失客户。
- 不推测成交金额、门店数量、负责人或产品能力。来源不明确的字段留作一般描述，不编造细节。

## 工作流

1. 阅读 [周报数据规则](references/report-data.md)，将日报整理为结构化 JSON。
2. 在当前任务的可写工作目录中创建 UTF-8 编码的 `weekly_report_data.json`。
3. 找到本 `SKILL.md` 所在目录，将其作为技能根目录。确认 Python 环境已安装依赖；未安装时，按当前 Agent 的权限规则安装 `requirements.txt`。
4. 运行生成脚本。Linux、macOS 或多数 Agent 环境使用：

   ```bash
   python3 <技能根目录>/scripts/generate_weekly_report.py \
     --input /absolute/path/weekly_report_data.json \
     --output /absolute/path/姓名_本周工作周报_新模板_YYYYMMDD-YYYYMMDD.docx
   ```

   Windows 也可使用 `py -3` 替代 `python3`。
5. 运行结构校验：

   ```bash
   python3 <技能根目录>/scripts/validate_weekly_report.py \
     --input /absolute/path/姓名_本周工作周报_新模板_YYYYMMDD-YYYYMMDD.docx
   ```

6. 如果当前 Agent 能渲染 DOCX，或环境安装了 Word、LibreOffice 等文档工具，再做逐页视觉检查：
   - 表格无截断、重叠或溢出；
   - 跨页卡点客户包含续页表头；
   - 章节标题不单独落在页尾；
   - “需要公司支持的问题”完整、清晰；
   - 没有模板示例、填写说明或上周内容残留。
7. 如排版不合格，调整 JSON 文案长度或脚本分页设置后重新生成并复查。
8. 最终只交付 DOCX；除非用户明确要求，不提供渲染 PDF、PNG 或中间 JSON。

## 内容要求

- 跟进中客户优先保留本周最重要的 5 至 8 个，突出当前阶段、本周变化和下一步动作。
- 成交客户写清购买产品、关键促成点和待执行事项。
- 卡点客户写真实阻碍和可执行的下周动作，风险等级仅使用“高、中、低”。
- 销售技巧、常见问题、行业信息和公司支持事项必须来自本周记录，不写空泛总结。
- 没有明确流失客户时，由脚本填入“本周暂无明确流失客户”，不要虚构流失案例。
- 行业信息通常保留 2 至 3 条，优先写价格反馈、采购偏好、产品能力和对接趋势。

## 模板与脚本

- Word 模板：`assets/weekly_report_template.docx`
- 生成脚本：`scripts/generate_weekly_report.py`
- 结构校验：`scripts/validate_weekly_report.py`
- Python 依赖：`requirements.txt`
- 数据结构及分类口径：`references/report-data.md`
- 跨 Agent 加载方式：`references/compatibility.md`

## 兼容性

本技能不要求任何特定平台的环境变量或专用工具。支持 `SKILL.md` 的 Agent 应读取本文件并按相对路径调用资源；不支持技能协议的 Agent 可把本文件作为工作流提示词，再直接运行脚本。需要安装到其他 Agent 时，阅读 [跨 Agent 兼容说明](references/compatibility.md)。
