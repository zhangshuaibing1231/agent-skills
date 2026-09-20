# 跨 Agent 兼容说明

## 通用目录结构

安装时必须保留完整目录，不要只复制 `SKILL.md`：

```text
kewayou-weekly-report/
├── SKILL.md
├── requirements.txt
├── scripts/
├── references/
└── assets/
```

模板和脚本都使用相对路径，技能目录可以放在任意位置。

## 支持 Agent Skills 的平台

把整个 `kewayou-weekly-report` 文件夹放入目标 Agent 配置的 Skills 目录，然后刷新技能列表或重启 Agent。调用时可直接说明“使用 kewayou-weekly-report 生成本周周报”。

不同平台的技能目录名称可能不同，应以该 Agent 当前配置为准。不要把文件夹内的文件拆散，也不要修改 `SKILL.md` 中的 `name`。

## 不支持 Skills 协议的平台

仍可按以下方式使用：

1. 将 `SKILL.md` 作为系统提示词、项目说明或工作流说明提供给 Agent。
2. 让 Agent 按 `references/report-data.md` 生成 UTF-8 JSON。
3. 安装依赖：

   ```bash
   python3 -m pip install -r requirements.txt
   ```

4. 生成 Word：

   ```bash
   python3 scripts/generate_weekly_report.py --input weekly_report_data.json --output weekly_report.docx
   ```

5. 校验 Word：

   ```bash
   python3 scripts/validate_weekly_report.py --input weekly_report.docx
   ```

Windows 可将 `python3` 替换为 `py -3`。

## 运行要求

- Python 3.9 或更高版本。
- `python-docx`，通过根目录的 `requirements.txt` 安装。
- 生成 DOCX 不需要网络。
- 若目标 Agent 没有 Word 或 LibreOffice，结构校验仍可运行，但无法自动完成视觉排版检查；此时应在 Word 或 WPS 中打开最终文件确认分页。
