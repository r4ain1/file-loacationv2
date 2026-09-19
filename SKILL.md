---
name: file-archive-workflow
description: 企业文件归档工作流——按既定命名规范与目录体系，把散落在桌面/下载区/微信暂存区的文件（PDF/Word/Excel/PPT/zip/图片）读取内容、判断归属类别、重命名并归档到工作文件体系，同时编写背景说明 Note。触发场景：用户说"归档/帮我归类/放在该放的地方/放到XX类/按我的习惯重命名/存到工作文件夹"，并 @ 了一个或多个文件路径。
---

# File Archive Workflow（文件归档工作流）

> **跨机可移植版**：所有机器相关路径均用 `{{占位符}}` 表示，实际值写在同目录的 `config.md` 中（见仓库 README 的安装说明）。使用前先确认 `config.md` 已按本机环境填写。

## 安装

把本目录（含本文件与 `config.md`）复制到你的 Agent 的 skills 目录：

| Harness | 放置路径 |
|---|---|
| Claude Code | `~/.claude/skills/file-archive-workflow/` |
| WorkBuddy / CodeBuddy | `~/.workbuddy/skills/file-archive-workflow/` |
| Cursor | `.cursor/skills/file-archive-workflow/`（或写入 `.cursor/rules/`） |
| Codex | `~/.codex/skills/file-archive-workflow/` |

复制后**必须先按本机环境修改 `config.md`**。

## 环境占位符（对应 `config.md`）

| 占位符 | 含义 | 示例 |
|---|---|---|
| `{{ARCHIVE_ROOT}}` | 工作文件体系根目录 | `D:\HuaweiYun_Rebuild_20260826\华为家庭存储\001_Work` |
| `{{DESKTOP}}` | 桌面/收件目录（源文件所在） | `C:\Users\rg\Desktop` |
| `{{PYTHON}}` | Python 解释器绝对路径 | `...\python.exe` |
| `{{MEMORY_DIR}}` | 工作记忆日志目录 | `C:\Users\rg\WorkBuddy\work\.workbuddy\memory` |
| `{{VISION_PROXY}}` | 视觉识别脚本（可选，扫描件 PDF/图片用） | `C:\Users\rg\.opencode-go\vision_proxy.py` |
| `{{AUTHOR_SELF}}` | 我方文档作者名（助手产出时写此名） | `WorkBuddy` |

## Overview

把用户发来的杂散文件按一套固定规范归档进工作文件体系。**必须严格遵守**：

- **文件命名规范**：`[序号_][文件名]_[YYYYMMDD-HHMM]_[作者/发起人]_[V版本号]`
  - 日期必带具体时间（如 `20260828-1356`），方便追溯
  - 序号（01/02/03…）用于一组关联文件排序，**主件在前**（如协议在前、支撑说明在后）
  - 作者规则：我方/助手产出写实际作者或 `{{AUTHOR_SELF}}`；**对方发来的文件**写对方交接人/经手人姓名
  - 版本迭代：后续修改版递增版本号（V1→V2→V3）
- **项目文件夹命名**：`[YYYY-MM-DD][项目主题]`（如 `2026-08-28某某项目申报`），建在体系对应类别目录下
- **归档配套 Note**：每个新项目文件夹内创建背景说明 Markdown（`[项目名]_背景说明Note_[YYYYMMDD-HHMM]_{{AUTHOR_SELF}}_V1.md`）
- **源文件保留**：归档一律 `cp`（复制），不移动、不删除源文件；临时解压/渲染目录用完即删

## Workflow Decision Tree（工作流决策树）

```
收到文件 + 归档指令
├─ 1. 读取文件内容（判断性质与归属）
│    ├─ Word(.docx) → python-docx 读段落+表格
│    ├─ Word(.doc) → UTF-16LE 读取（老格式）
│    ├─ Excel(.xlsx) → openpyxl 读工作表+前若干行
│    ├─ PPT(.pptx) → python-pptx 读每页文本（含表格）
│    ├─ PDF 有文字层 → pypdf 提取文本
│    ├─ PDF 扫描件 → PyMuPDF(fitz) 渲染PNG + 视觉模型识别关键页
│    ├─ zip → 解压到桌面临时目录后逐个读成员文件
│    └─ 图片 → 视觉模型直接识别
├─ 2. 搜索既有目录（find -iname 关键词）
│    └─ 关键词取文件主题核心词：机构名/项目名/人名
├─ 3. 判定归属类别
│    ├─ 命中既有项目目录 → 并入同主题目录
│    └─ 无命中 → 在正确类别下新建 [YYYY-MM-DD][主题] 目录
│    ★ 拿不准时用 AskUserQuestion 让用户拍板（给 2-3 个候选 + 推荐项）
│    ★ 不要把文件并进别人的个人目录
├─ 4. 归档执行（单条 Bash 命令，勿合并多个操作）
│    ├─ mkdir -p 目标目录
│    ├─ TS=$(date +%Y%m%d-%H%M) 生成时间戳
│    ├─ cp 源文件 → 目标/新命名
│    └─ ★ 沙箱会拒绝合并命令（&& 多个cp/mv），拆成单条命令逐个执行
├─ 5. 编写背景说明 Note（见 Note 模板）
└─ 6. 收尾
     ├─ 追加工作记忆日志（{{MEMORY_DIR}}/YYYY-MM-DD.md）
     ├─ 清理临时目录（rm -rf 仅限自己创建的 _temp_* 目录）
     └─ ★ 立即展示归档结果给用户（不等用户追问）
```

## Step 1：读取文件内容

用 `{{PYTHON}}` 读取，**读取只为判断归属，归档是纯文件复制操作**。

### Excel（openpyxl）
```bash
"{{PYTHON}}" -c "
import openpyxl
wb = openpyxl.load_workbook(r'<源文件路径>', data_only=True)
print('工作表:', wb.sheetnames)
for name in wb.sheetnames:
    ws = wb[name]
    print(f'===== {name} (行{ws.max_row} 列{ws.max_column}) =====')
    for row in ws.iter_rows(min_row=1, max_row=min(15, ws.max_row), values_only=True):
        vals = [str(v)[:25] if v is not None else '' for v in row]
        if any(vals):
            print(' | '.join(vals))
" 2>&1 | head -60
```

### Word（python-docx）
```bash
"{{PYTHON}}" -c "
import docx
d = docx.Document(r'<源文件路径>')
for p in d.paragraphs:
    t = p.text.strip()
    if t: print(t)
for i, tb in enumerate(d.tables):
    print(f'--- 表{i} ---')
    for row in tb.rows[:8]:
        print(' | '.join(c.text.strip()[:30] for c in row.cells))
"
```

### PPT（python-pptx）
```python
from pptx import Presentation
prs = Presentation(r'<源文件路径>')
for i, slide in enumerate(prs.slides):
    print(f'===== 幻灯片{i+1} =====')
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                t = ''.join(run.text for run in para.runs).strip()
                if t: print(t)
        if shape.has_table:
            print('[表格]')
            for row in shape.table.rows:
                print(' | '.join(c.text.strip()[:20] for c in row.cells))
```

### PDF
- **有文字层**：pypdf 提取 `page.extract_text()`
- **扫描件（无文字层）**：PyMuPDF 逐页渲染 PNG（dpi=150），再用视觉模型识别关键页（首页确认协议类型/双方，末页确认签署盖章）：
```bash
# 1) 渲染
"{{PYTHON}}" -c "
import fitz
doc = fitz.open(r'<pdf路径>')
for i in range(len(doc)):
    pix = doc[i].get_pixmap(dpi=150)
    pix.save(rf'{{DESKTOP}}/_temp_xxx/page_{i+1:02d}.png')
"
# 2) 视觉识别某页
"{{PYTHON}}" "{{VISION_PROXY}}" "{{DESKTOP}}/_temp_xxx/page_01.png" 2>/dev/null
```
识别完**必须删除** `_temp_*` 渲染目录。

### zip / 图片
- zip：解压到 `{{DESKTOP}}/_temp_unzip_xxx` 逐个读，归档后删除临时目录
- 图片：交给视觉模型识别（如 `{{VISION_PROXY}}`）

## Step 2-3：搜索既有目录 & 判定归属

```bash
find "{{ARCHIVE_ROOT}}" -maxdepth 4 -type d -iname "*<关键词>*" 2>/dev/null | head -20
# 同时搜文件（有时目录名不含关键词但文件名含）
find "{{ARCHIVE_ROOT}}" -maxdepth 5 -type f -iname "*<关键词>*" 2>/dev/null | head -20
```
- 关键词用文件主题核心词（人名/机构名/项目名），**不要只搜文件名**
- 命中既有项目目录 → 同主题文件并入该目录（保持项目档案聚拢）
- 无命中 → 在正确类别下新建 `[YYYY-MM-DD][主题]` 项目目录
- **拿不准就 AskUserQuestion**，给 2-3 个候选位置 + 推荐项，用户拍板后再动手。宁可不归档，也不要放错

## Step 4：归档执行

```bash
mkdir -p "{{ARCHIVE_ROOT}}/<类别路径>/<YYYY-MM-DD项目主题>"
TS=$(date +%Y%m%d-%H%M) && echo "时间戳: $TS"
cp "{{DESKTOP}}/<源文件>" "{{ARCHIVE_ROOT}}/<类别路径>/<YYYY-MM-DD项目主题>/<新文件名>_${TS}_<作者>_<V版本>.<ext>"
ls -la "{{ARCHIVE_ROOT}}/<类别路径>/<YYYY-MM-DD项目主题>/"
```
注意：
- Git Bash 路径转换：`D:\` → `/d/`，`C:\` → `/c/`，路径含空格必须加双引号
- **时间戳用 `date +%Y%m%d-%H%M` 生成，不要手算**
- 一组关联文件：加序号前缀，主件排 01

## Step 5：背景说明 Note 模板

文件：`[项目名]_背景说明Note_[YYYYMMDD-HHMM]_{{AUTHOR_SELF}}_V1.md`，放在项目文件夹内：

```markdown
# [项目名] 背景说明

## 归档信息
- **归档日期**：YYYY-MM-DD
- **归档人**：{{AUTHOR_SELF}}（经用户确认）
- **归档位置**：[类别路径]

## 文件清单
| 文件 | 说明 |
|---|---|
| `[新文件名]` | [版本/性质说明] |

## 背景与要点
- **性质**：[这是一份什么文件]
- **关键信息**：
  - [甲/乙方、金额、期限、联系人、数据要点等，逐条列出]
- **后续建议**：[如到期提醒、待补空白项、待签署等]

---
*本 Note 自动生成归档备案。*
```
内容**必须来自实际读取的文件内容**，不得虚构数据/人名/日期。

## Step 6：收尾

1. **工作记忆日志**（追加，勿覆盖）：`{{MEMORY_DIR}}/YYYY-MM-DD.md`（不存在则创建）
2. **清理临时目录**：仅限自己创建的 `_temp_*` 目录
3. **★ 立即展示归档结果**：归档完成即把重命名后的文件（含所在文件夹）展示给用户确认，**不等用户追问**（用户明确要求的习惯）

## 常见坑（实战踩过）

- ❌ 合并 Bash 命令被沙箱拒绝（`&&` 多个 cp/mv/解压）→ ✅ 拆成单条命令
- ❌ 想当然把协议归法务 → ✅ 先搜其他类别，用户口径优先
- ❌ PDF 扫描件用 pypdf 提取出空文本就下结论 → ✅ 必须渲染成图 + 视觉识别
- ❌ 把文件并进别人的个人目录 → ✅ 新建本项目目录或并入同主题项目目录
- ❌ 归档后删用户源文件 → ✅ 一律 cp；源文件是否删除由用户决定
- ❌ 命名时间戳手写 → ✅ 一律 `date +%Y%m%d-%H%M` 动态生成
- ❌ 换机器后路径失效 → ✅ 只改 `config.md`，不要改 SKILL.md

## Resources

无附属脚本，流程全部用环境内既有工具（python-docx/openpyxl/python-pptx/pypdf/PyMuPDF + 可选视觉识别脚本）。

## 多电脑与双份归档（强制）

1. 开始前识别当前电脑，并动态定位 华为家庭存储\\001_Work。
2. 工作文件复制两份：项目/职能目录一份，0 总资料对应知识域一份。
3. 无法判断归属时再询问用户；能够从内容和现有目录判断时直接执行。
4. 桌面源文件默认保留。

