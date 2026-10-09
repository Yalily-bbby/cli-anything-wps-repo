# cli-anything-wps

一个 Codex / Agent Skill：在 Windows 上通过 WPS Office 的 COM 自动化接口，
程序化新建与编辑 Office 文档并导出 DOCX / XLSX / PPTX / PDF / CSV / HTML，
也支持用 JSON 数据批量生成 PPT。

## 仓库结构

```
skills/
└─ cli-anything-wps/
   ├─ SKILL.md              # 技能正文与说明（含 YAML frontmatter）
   └─ agents/
      └─ openai.yaml        # 面向宿主 UI 的清单：显示名、简介、默认提示词
```

`skills/<技能名>/` 是 Codex 约定的技能目录，技能名取路径的最后一段，
因此下面的安装命令装完会得到 `$CODEX_HOME/skills/cli-anything-wps`。

## 安装

### 方式一：官方 skill-installer 脚本（推荐）

```bash
python "$CODEX_HOME/skills/.system/skill-installer/scripts/install-skill-from-github.py" \
  --repo <你的账号>/cli-anything-wps-repo \
  --path skills/cli-anything-wps
```

默认装到 `~/.codex/skills/`；`CODEX_HOME` 环境变量存在时以它为准。
注意：**目标目录已存在时脚本会直接报错退出**，需要先删掉旧的再装。

### 方式二：Codex++ 面板

在「Codex MCP & 插件 → Skills → 新增」里填：

| 字段 | 填什么 |
| --- | --- |
| 仓库 | `<你的账号>/cli-anything-wps-repo` |
| 路径 | `skills/cli-anything-wps` |
| 分支 / ref | `main` |

> 「路径」指仓库内的相对路径，和官方精选技能用的 `skills/.curated/<名字>` 是同一套规则。

### 方式三：手工拷贝

把 `skills/cli-anything-wps/` 整个目录复制到 `%USERPROFILE%\.codex\skills\` 下即可。
技能会在**新开的对话**里生效——技能清单在会话启动时确定，老对话不会热加载。

## 运行前置

| 依赖 | 说明 |
| --- | --- |
| Windows | COM 自动化仅 Windows 可用 |
| WPS Office | 技能基于 WPS 的 COM 接口，ProgID 为 `KWPP.Application` |
| Python 3.10+ 与 pywin32 | 用于调用 COM |

## 已知限制

- `SKILL.md` 里记录的是**原机器上的绝对路径**（例如 `D:\python.exe`、
  `D:\Scripts\cli-anything-wps.exe`、仓库源码目录）。换机器使用时，
  需要把这些路径改成目标机器上实际的解释器与 CLI 位置。
- 技能正文提到上游文档里那句 `taskkill /F /IM wps.exe /T` **不要执行**——
  它会强制关闭用户正在编辑的 WPS 文档，正常导出不需要它。
- `SKILL.md` 正文是中文。

## 许可

本仓库未附带许可证文件。若要公开分发，请自行补充。
