# 联考成绩分析流水线（exam-score-analysis）

一个面向中学教研室的 **Agent Skill**，遵循开放 Agent Skills 标准（`SKILL.md`），可直接用于 Codex、WorkBuddy、Claude Code、Cursor、Gemini CLI 等 30+ 宿主。

**它解决什么问题**：把一次考试的数据包，变成一份能逐格回溯的 Excel 分析工作簿 + 一套教师会 PPT。核心不是算得快，而是**口径一次锁定、结论可回溯**——同类任务最常见的失败是口径反复变更导致全量重做，以及表格与结论不同源导致的硬错。

**它能做**：
- 判定本次是联考还是单校，走不同分支（联考做校际对比，单校从班级分析起）
- 班级/年级均分、排名对比、赋分处理、贡献率
- 模拟上线测算（官方名次 / 联考一分一段 / 校内模拟线）
- 学生分层与短板诊断（A 值 = 该生该科得分率 − 本校同组合平均得分率）
- Excel 三册工作簿 + 20–30 页 PPT
- 单元格级独立核对（不复用生成阶段数据，约 1.1 万项/轮）

**安装（30 秒）**：

```bash
# 1) 取回技能目录（本仓库是技能合集，只需这一个子目录）
git clone --depth 1 https://github.com/DaDel7924/Personal-Agent-Skills.git /tmp/pas
cp -R /tmp/pas/exam-score-analysis ~/.agents/skills/

# 2) 装依赖并自检
python3 -m pip install -r ~/.agents/skills/exam-score-analysis/requirements.txt
python3 ~/.agents/skills/exam-score-analysis/scripts/selftest.py
```

指向其他宿主（可选）：

```bash
ln -sfn ~/.agents/skills/exam-score-analysis ~/.workbuddy/skills/exam-score-analysis   # WorkBuddy
ln -sfn ~/.agents/skills/exam-score-analysis ~/.claude/skills/exam-score-analysis      # Claude Code
```

**调用**：Codex 里输入 `$exam-score-analysis`，WorkBuddy 直接说"帮我分析这次考试成绩"，Claude Code 用 `/exam-score-analysis`。

---

本技能遵循开放 Agent Skills 标准（`SKILL.md` + `scripts/` + `references/`），同一份目录可被 Codex、Claude Code、WorkBuddy、Gemini CLI、Copilot/VS Code、Cursor、OpenCode 等 30+ 宿主读取。**内容无需改动，只需放到各宿主扫描的目录。**

## 一、各宿主看哪里

| 宿主 | 用户级目录 | 项目级目录 | 调用方式 |
|---|---|---|---|
| **WorkBuddy** | `~/.workbuddy/skills/<name>/` | `<项目>/.workbuddy/skills/` | 描述匹配自动触发 |
| **Codex CLI / IDE** | `~/.agents/skills/<name>/` | `<仓库>/.agents/skills/` | `$exam-score-analysis` 或 `/skills`；描述匹配自动触发 |
| **Claude Code** | `~/.claude/skills/<name>/` | `<项目>/.claude/skills/` | `/exam-score-analysis`；描述匹配自动触发 |
| **Gemini CLI / Cursor / OpenCode** | 多数读 `~/.agents/skills/` | 同上 | `@` 或描述匹配 |

两个要点：

1. **Claude Code 不读 `.agents/skills`**，Codex 不读 `.claude/skills`。所以"一份真身 + 两处链接"最省事。
2. **`~/.codex/skills/` 已弃用**（仍能加载）。新装请用 `~/.agents/skills/`。

## 二、推荐装法：一份真身 + 符号链接

把真实目录放在 `~/.agents/skills/`，其余宿主用软链指向它。改一处，全体生效。

```bash
# 1) 真身（若已在 ~/.workbuddy/skills，直接移过来）
mkdir -p ~/.agents/skills
mv ~/.workbuddy/skills/exam-score-analysis ~/.agents/skills/

# 2) 链回各宿主
mkdir -p ~/.workbuddy/skills ~/.claude/skills
ln -sfn ~/.agents/skills/exam-score-analysis ~/.workbuddy/skills/exam-score-analysis
ln -sfn ~/.agents/skills/exam-score-analysis ~/.claude/skills/exam-score-analysis
```

不想动现有位置时，反向链接亦可（真身留在原处，另建软链）：

```bash
mkdir -p ~/.agents/skills
ln -sfn ~/.workbuddy/skills/exam-score-analysis ~/.agents/skills/exam-score-analysis
```

Codex 与 Claude Code 都会跟随符号链接。

## 三、项目内分发（多人协作）

把技能目录放进仓库的 `.agents/skills/exam-score-analysis/` 并提交。任何人用 Codex/Cursor/OpenCode 打开该仓库即自动加载；Claude Code 用户需另建 `.claude/skills/` 链接。

代价：技能随仓库走，与项目绑定。适合同一学校的教研组共用。

## 四、依赖

```bash
python3 -m pip install -r requirements.txt   # openpyxl / xlrd / python-pptx / lxml
python3 scripts/selftest.py                 # 必须全部 OK
```

`selftest.py` 检查四件事：依赖齐全、舍入与排名口径正确、联考/单校判定可用、核对与修补脚本可导入。任一 FAIL 都别开工。

四项依赖全部为本地计算，**不联网、不调外部 API**，脚本里没有硬编码绝对路径（读取路径全部由命令行参数或函数入参传入）。

## 五、跨宿主差异（哪些会丢）

| 字段 / 能力 | 说明 |
|---|---|
| `allowed-tools`、`displayName`、`slug`、`namespace`、`tags` | 宿主专有，Codex/Claude Code 会忽略，不影响运行 |
| `agents/openai.yaml` | 仅 Codex/ChatGPT 读取；其他宿主忽略。里面可关掉自动触发（`allow_implicit_invocation: false`） |
| 自动触发依赖 `description` | Codex 在索引超预算时会**先截短描述**，所以触发词必须放在描述开头。本技能的描述已前置「考试成绩分析、联考/单校、班级均分、上线测算、质量分析会 PPT」 |
| 沙箱与审批 | Codex 默认沙箱可能限制写盘位置。运行前把产物目录放在工作区内；批量写入前先 `--dry-run` |
| Windows | 脚本为纯 Python，路径用 `os.path`，中文目录可用。终端编码建议 UTF-8（`chcp 65001`） |

## 六、安全提示

技能本质是"代理会照做的指令 + 可执行脚本"。本技能只做本地表格读写，无网络请求、无外发动作。若从别处获取同类技能，先读 `scripts/` 再启用——这是公开技能生态公认的必要动作。

## 七、目录清单

```
exam-score-analysis/
├── SKILL.md                 # 主文件：Step 0 问询 + 7 阶段 + 11 条硬规则
├── README.md                # 本文件
├── requirements.txt
├── agents/openai.yaml       # 仅 Codex/ChatGPT 使用
├── references/
│   ├── exam-mode.md         # 先读：联考/单校判定、单校口径替换、内部恒等式
│   ├── intake.md            # 开工 6 问 + 确认单模板
│   ├── calibration.md       # 7 项口径锁定
│   ├── data-pitfalls.md     # 12 条数据陷阱
│   ├── audit-method.md      # L1–L5 分层核对
│   ├── workbook-spec.md     # 三册工作簿蓝图
│   └── ppt-spec.md          # PPT：有模板复用 / 无模板从零设计 + 改写规范
└── scripts/
    ├── selftest.py          # 安装自检
    ├── reader.py            # 读取/归一化/模式判定/校数统计
    ├── audit_scaffold.py    # 独立核对骨架
    └── xlsx_patch.py        # 保留图表的单元格级修补
```
