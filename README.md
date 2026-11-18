# Personal-Agent-Skills

个人 Agent 技能合集。每个技能遵循开放 Agent Skills 标准（`SKILL.md` + `scripts/` + `references/`），同一份目录可直接用于 Codex、WorkBuddy、Claude Code、Cursor、Gemini CLI 等 30+ 宿主。

## 技能索引

| 技能 | 用途 | 触发方式 |
|---|---|---|
| [exam-score-analysis](./exam-score-analysis/) | 中学考试成绩分析：联考/单校判定、班级与年级均分、排名对比、模拟上线、学生分层与短板诊断，产出 Excel 工作簿与教师会 PPT，并做单元格级核对 | `$exam-score-analysis` / 直接说"做个成绩分析" |

## 安装

每个技能是一个独立顶层目录，装到宿主扫描路径即可。以 `exam-score-analysis` 为例：

```bash
git clone --depth 1 https://github.com/DaDel7924/Personal-Agent-Skills.git /tmp/pas
cp -R /tmp/pas/exam-score-analysis ~/.agents/skills/
python3 -m pip install -r ~/.agents/skills/exam-score-analysis/requirements.txt
python3 ~/.agents/skills/exam-score-analysis/scripts/selftest.py
```

各宿主的技能目录：

| 宿主 | 用户级目录 | 调用 |
|---|---|---|
| Codex CLI / IDE | `~/.agents/skills/` | `$技能名` 或 `/skills` |
| WorkBuddy | `~/.workbuddy/skills/` | 描述匹配自动触发 |
| Claude Code | `~/.claude/skills/` | `/技能名` |
| Cursor / Gemini CLI / OpenCode | `~/.agents/skills/` | `@` 或描述匹配 |

一份真身 + 软链，可让多个宿主共用同一份：

```bash
ln -sfn ~/.agents/skills/exam-score-analysis ~/.workbuddy/skills/exam-score-analysis
```

## 仓库约定

- **一个技能一个顶层目录**，目录名与 `SKILL.md` 的 `name` 字段一致
- 每个技能自带 `README.md`（安装与跨宿主说明）、`requirements.txt`、`scripts/selftest.py`（装完先跑）
- 技能内不包含任何真实学生数据、学校名称或可识别信息；示例一律用虚构校名
- 许可证按技能粒度声明，见各技能目录内的 `LICENSE`

## 目录结构示例

```
Personal-Agent-Skills/
└── exam-score-analysis/
    ├── SKILL.md              # 主文件：触发描述 + 执行流程
    ├── README.md             # 安装、跨宿主、依赖
    ├── LICENSE               # MIT
    ├── requirements.txt
    ├── agents/openai.yaml    # 仅 Codex/ChatGPT 读取
    ├── references/           # 按需加载的参考文档
    └── scripts/              # 可执行脚本 + 自检
```
