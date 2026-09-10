# file-archive-workflow

企业文件归档工作流 —— 一个跨机器、跨 Agent harness 可移植的 Agent Skill。

把散落在桌面 / 下载区 / 微信暂存区的文件（PDF / Word / Excel / PPT / zip / 图片）自动完成：
**读取内容 → 判断归属 → 规范重命名 → 复制归档 → 生成背景说明 Note → 展示结果**。

---

## 为什么需要它

归档这件事的难点不是"复制文件"，而是**每次都要重新交代一遍**：文件放哪、怎么命名、
作者写谁、要不要保留原件、归档完要不要给我看。这套 Skill 把这些固化为可执行流程，
换台电脑、换个 Agent 工具（Claude Code / WorkBuddy / Cursor / Codex）都能直接用。

## 设计原则

- **核心与环境分离**：`SKILL.md` 只含 `{{占位符}}`，机器相关路径全部集中在 `config.md`
- **只复制不移动**：归档一律 `cp`，绝不删除源文件
- **时间戳动态生成**：禁止手写日期，一律 `date +%Y%m%d-%H%M`
- **拿不准就问**：归属不确定时用选项让用户拍板，宁可不归档也不放错
- **归档即反馈**：完成后立刻把结果展示给用户，不等追问

---

## 安装（任选你的 harness）

```bash
# 1. 克隆到任意位置
git clone <本仓库地址> ~/skills-src/file-archive-workflow

# 2. 复制到你的 harness 的 skills 目录
cp -r ~/skills-src/file-archive-workflow ~/.claude/skills/       # Claude Code
cp -r ~/skills-src/file-archive-workflow ~/.workbuddy/skills/    # WorkBuddy / CodeBuddy
cp -r ~/skills-src/file-archive-workflow ~/.codex/skills/        # Codex
cp -r ~/skills-src/file-archive-workflow .cursor/skills/         # Cursor（项目级）
```

## ⚠️ 首次使用必做：填写 config.md

复制完成后，**按本机环境修改 `config.md` 第 1 节的占位符实际值**：

| 占位符 | 填什么 |
|---|---|
| `{{ARCHIVE_ROOT}}` | 你的工作文件体系根目录 |
| `{{DESKTOP}}` | 桌面 / 源文件常驻目录 |
| `{{PYTHON}}` | Python 解释器绝对路径 |
| `{{MEMORY_DIR}}` | 工作记忆日志目录 |
| `{{VISION_PROXY}}` | 视觉识别脚本（可选，处理扫描件 PDF / 图片） |
| `{{AUTHOR_SELF}}` | 我方文档作者名（助手产出时写入文件名） |

`config.md` 其余小节（目录类别地图、已归档项目速查）按你的组织实际情况改写即可；
换公司/换项目体系时**只改这个文件，不用动 `SKILL.md`**。

## Python 依赖

```bash
pip install python-docx openpyxl python-pptx pypdf PyMuPDF Pillow
```

## 文件结构

```
file-archive-workflow/
├── SKILL.md      # 通用核心流程（含占位符，跨机器通用）
├── config.md     # 本机/本组织环境配置（换机器只改这里）
└── README.md     # 本说明
```

## 使用方式

直接对 Agent 说，并 @ 上文件路径：

> @`~/Desktop/某某协议.pdf` 帮我归档  
> 把这几个文件放到该放的地方，按我的习惯重命名  
> 这个材料存到工作文件夹里

Skill 会自动：读内容 → 搜既有目录 → （必要时询问归属）→ 重命名归档 → 生成 Note → 展示结果。

## 命名规范速查

```
[序号_][文件名]_[YYYYMMDD-HHMM]_[作者/发起人]_[V版本号].[扩展名]

示例：
01_设备机时租赁服务协议-西湖大学_20260902-1411_阮观_V1.docx
02_5600+及Micro液相色谱仪机时租赁定价公允性说明_20260902-1411_阮观_V1.docx
```

- 日期必带具体时间，便于追溯
- 序号用于一组关联文件排序，主件在前
- 对方发来的文件，作者位写对方交接人姓名
- 版本迭代递增 V1 → V2 → V3
