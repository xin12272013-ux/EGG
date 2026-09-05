# EGG 头马俱乐部专属技能集

上海 EGG「诞」头马演讲俱乐部（D85 大区 L 中区）的 WorkBuddy 技能包，覆盖**会议物料 → 演讲训练 → 内容传播**的完整闭环。

> 愿景：企业家精神 · 利他思维 · 成长型思维
> 金句：万物皆有裂痕，那是光照进来的地方
> 标语：让你开口，诞生无限

## 技能一览

| 目录 | 技能 | 干什么 | 触发场景 |
|---|---|---|---|
| `egg-toastmasters-speech-mentor/` | **蛋导**（演讲教练 v4.1） | 八维能力全覆盖的演讲全流程教练：从零生成、比赛级评估定级、反馈改稿、六维磨稿；备赛侧含「大区冠军工程蓝图」（十二道工序 + 28 项验收清单） | 改稿 / 磨稿 / 备稿 / 备赛 / 个评 / 即兴 / 演讲评估 |
| `egg-meeting-generator/` | **会议文件生成器** | 贴一段会议角色接龙 → 自动生成 Agenda Excel（4 个 sheet）+ PPT 背景板（30+ 页），基于 258/261/262 次实战模板 | 接龙 / 摘花 / 生成 Agenda / 生成 PPT |
| `egg-poster-generator/` | **会议海报生成器** | 内建 EGG 全部固定信息，只需主题 + 届次 → 出成品 PNG 海报（AI 画背景 + PIL 叠中文字，汉字 100% 正确） | EGG 海报 / 例会海报 / 嘉宾海报 / 接龙配图 |
| `egg-xhs-content-series/` | **小红书内容系列** | 从选题 → 文案 → 封面图 → 长图 → 发布排期 → 引流承接（入会清单 / 私信话术）→ 资产总览的一站式工作流 | 小红书日更 / 系列内容 / 引流清单 / 排发布表 |

## 安装

把需要的技能目录整个复制到 WorkBuddy 技能目录即可：

```bash
# Windows
cp -r egg-toastmasters-speech-mentor %USERPROFILE%\.workbuddy\skills\
cp -r egg-meeting-generator        %USERPROFILE%\.workbuddy\skills\
cp -r egg-poster-generator         %USERPROFILE%\.workbuddy\skills\
cp -r egg-xhs-content-series       %USERPROFILE%\.workbuddy\skills\
```

```bash
# macOS / Linux
cp -r egg-* ~/.workbuddy/skills/
```

重启 WorkBuddy 后技能生效。

## 依赖

| 技能 | 依赖 |
|---|---|
| 蛋导 | 评估/改稿默认交付排版 PDF，需本机 Chrome（headless 打印） |
| 会议文件生成器 | Python `python-pptx` `openpyxl` |
| 海报生成器 | Python `Pillow`；底层脚本复用通用技能 `toastmasters-poster-prompt` 的 `overlay_text.py` / `composite_guest.py` |
| 小红书系列 | 无额外依赖 |

> 提示：`egg-poster-generator` 是**薄封装**，实际调用通用技能 `toastmasters-poster-prompt` 里的绘图与叠字脚本，两者需同时安装。

## 目录结构

```
EGG/
├── egg-toastmasters-speech-mentor/   # 蛋导 v4.1
│   ├── SKILL.md
│   ├── assets/report-template.html
│   └── references/                   # 13 份方法论（结构逻辑/听众分析/交付工坊/
│                                     #   冠军蓝图/评估手册/即兴手册/辩论应答……）
├── egg-meeting-generator/
│   ├── SKILL.md
│   ├── scripts/                      # generate_*.py / patch_262.py / fix_pptx_text.py
│   ├── references/template_info.md
│   └── templates/                    # 258 / 261 / 262 次 PPT 背景板 + Agenda（约 93 MB）
├── egg-poster-generator/
│   ├── SKILL.md
│   ├── scripts/make_egg_poster.py
│   └── references/                   # egg-profile.md / egg-theme-history.md（防意象撞车）
└── egg-xhs-content-series/
    └── SKILL.md
```

## 备注

- `egg-meeting-generator/templates/` 含三个大体积 PPT 模板（合计约 93 MB），是技能运行的基础素材，建议保留。
- `egg-poster-generator/references/egg-theme-history.md` 登记了历次海报意象，出新海报前先查，避免与历史撞车。
- 其他头马俱乐部（GTD 等）请使用通用版技能，不要套用本仓库的 EGG 专属信息。
