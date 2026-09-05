---
name: egg-poster-generator
description: |
  上海 EGG「诞」头马演讲俱乐部**专属**会议海报生成器。内建 EGG 全部固定信息（俱乐部全称、
  例会地址、常规时段、愿景三词、金句、标语、D85 大区 L 中区），出海报时这些**不用再问用户**，
  只需主题 + 届次就能出成品。
  流程：解析活动信息/角色接龙 → 用 EGG 默认值补全 → 内置绘图模型出背景图（默认金蓝嘉宾风）
  → PIL 叠加中文会议信息（主题大字由 PIL 渲染，汉字 100% 正确）→ 合成嘉宾照片 → 交付成品 PNG。
  默认出 1 张，默认交付图片本身，不输出即梦提示词文本。
  当文本中出现 EGG、「诞」、蛋、或快乐哥为自己俱乐部做海报时使用。
  触发词：EGG 海报、EGG 例会海报、诞俱乐部海报、EGG 宣传图、EGG 配图、EGG 接龙配图、
  EGG 活动海报、EGG 嘉宾海报。
  ⚠️ 若用户是**其他**头马俱乐部（GTD、其它中文/英文俱乐部）或明确说"通用"，改用
  `toastmasters-poster-prompt` 技能——本技能只在 EGG 场景下使用。
agent_created: true
---

# EGG「诞」头马演讲俱乐部 · 会议海报生成器

EGG 专属。与通用技能 `toastmasters-poster-prompt` 的关系：

| | 本技能（EGG 专属） | 通用技能 |
|---|---|---|
| 适用 | **仅 EGG「诞」** | 所有其它头马俱乐部 |
| 俱乐部信息 | **内建**，只需确认 | 一律问用户，不猜 |
| 品牌元素 | 金句/标语/愿景**默认带上** | 视用户提供 |
| 主题历史 | **有**，避免意象撞车 | 无 |
| 脚本 | 薄封装 `make_egg_poster.py`，**底层复用通用技能脚本** | 原始脚本 |

> 脚本不复制——直接调用通用技能目录里的 `overlay_text.py` / `composite_guest.py` /
> `parse_activity.py`。**通用技能改脚本，本技能自动受益**，不会出现两份分叉。
> 若通用技能缺失，报错并提示先修复，不要自己另写一份。

## EGG 固定信息（直接取用，仅向用户做一次确认）

| 字段 | 值 |
|---|---|
| 俱乐部全称 | **上海 EGG「诞」头马演讲俱乐部** |
| 例会地址 | 上海市徐汇区云锦路 181 号 龙华街道水岸党群服务中心 214 室（龙华地铁站附近） |
| 常规时段 | **周日上午**（实测 9:40–12:30；以当次接龙为准，接龙有就用接龙的） |
| 费用 | 所有来宾和角色**全免费** |
| 大区 | D85 大区 L 中区 |
| 愿景三词 | 企业家精神 · 利他思维 · 成长型思维（接龙版本：创业-提供增长 / 给予帮助 / 持续成长） |
| 金句 | 万物皆有裂痕，那是光照进来的地方 |
| 标语 | 让你开口，诞生无限 |

⚠️ **唯一例外**：地址与时间**以当次接龙原文为准**。接龙写了就用接龙的，没写才用上表默认值，
并在交付时标注「地址/时间用的是 EGG 默认值，请确认」。改了场地却不改海报，比没海报更糟。

## 核心原则

继承通用技能全部 7 条，另加 3 条 EGG 专属：

8. **固定信息不问第二遍**。俱乐部名、地址、大区、费用默认带出，只在交付核对表里让用户扫一眼。
9. **标语位优先用 EGG 金句或愿景三词**。用户没指定 `--slogan` 时，默认填
   `企业家精神 · 给予帮助 · 持续成长`；主题调性偏内省/疗愈时改用金句
   `万物皆有裂痕，那是光照进来的地方`。
10. **出完图登记主题历史**。往 `references/egg-theme-history.md` 追加一行（届次/主题/意象/风格/日期），
   下次出海报先查这个表**避免意象撞车**——连着两期都用"赛道"会很尴尬。

## 工作流

### Step 0 — 解析输入

EGG 的输入**两类都有**，判断方式同通用技能：

| 类型 | 特征 | 处理 |
|---|---|---|
| A. 活动信息（多数） | 届次 + 主题 + 嘉宾姓名/头衔/介绍/照片 | 跑通用解析器 |
| B. 角色接龙 | 一串角色（会议经理/主持人/时间官/哼哈官/备稿1…） | 正则抽取 |

```bash
python "C:/Users/Administrator/.workbuddy/skills/toastmasters-poster-prompt/scripts/parse_activity.py" \
  --file activity.txt --out parsed.json
```

EGG 接龙的两个特征，解析时留意：

- 届次写作「第 N 次**会议**」（不是"例会"），如「EGG"诞"头马演讲俱乐部第263次会议」
- 结尾常有愿景行 `创业-提供增长 / 企业家精神 - 给予帮助 - 持续成长`，**不要误当嘉宾块**
- 角色名常带 emoji 占位（`[烟花]`/`[玫瑰]` = 待招），这些行直接忽略

### Step 1 — 只问真正缺的

**不问**：俱乐部名、地址、大区、费用、常规时段（内建）。
**必问**：会议主题（没有主题就没有意象锚点，不许猜）、届次（解析不出时）。
**确认一次**：地址/时间用的是默认值还是接龙值。

嘉宾信息逐位追问缺失项（照片要文件路径或 URL）。

### Step 2 — 风格卡（EGG 默认：金蓝嘉宾风）

读通用技能 `references/style-library.md`。EGG 的默认选择：

| 场合 | 风格 |
|---|---|
| 单场例会 / 嘉宾分享（**默认**） | **风格 A 金蓝嘉宾风**——深海军蓝渐变底 + 香槟金点缀 + 圆形嘉宾头像金环 |
| 传统文化 / 疗愈 / 节气 / 阅读主题 | 风格 B 新国潮水墨 |
| 多时段大型活动（开放日、联合活动） | 风格 C 全景活动日 |

### Step 3 — 查主题意象（先查 EGG 历史，避免撞车）

1. **先读 `references/egg-theme-history.md`**——同主题或近义主题用过的意象不要再选
2. 再读通用技能 `references/theme-visual-map.md` 匹配意象
3. 无命中走回退：主题拆成可画名词，套「金色几何层叠 + 渐变光晕 + 细腻噪点」（风格 A 下用金色材质）

### Step 4–5 — 填六段骨架 + 过自检

严格按通用技能 `references/prompt-skeleton.md` 的六段式与长度预算（150–300 中文字）。
自检清单同通用技能 Step 5，**额外确认**：本场意象未出现在 EGG 主题历史里。

### Step 6 — 出图（默认 1 张）

1. 生成前告知积分消耗（单张约 5–10 积分）
2. 调 `connect_cloud_service` 取 `clientTempToken`（有效期约 15 分钟，每轮重取）
3. 传给 `ImageGen`：`size` **直接填目标画布尺寸** `1080x1440`，`quality="high"`
   - 模型会对齐 64 倍数（实测输出 1072×1440），比例几乎不变
   - ⚠️ **别生成 1024×1536 再 cover 到 3:4**：纵向裁 180px 会让留白区上移、文字压到意象（实测事故）
4. 默认只出 1 张，用户要多版挑选时才出 3 张
5. **垂直元素（光柱/台阶/上升光束）天生画超高**，顶部侵入留白区就改水平带状构图重写

### Step 7 — 叠字（用 EGG 一键脚本）

```bash
python "C:/Users/Administrator/.workbuddy/skills/egg-poster-generator/scripts/make_egg_poster.py" \
  --bg bg.png --out EGG263_poster.png \
  --theme "备稿马拉松" --session "第 263 次会议"
```

只需主题 + 届次，其余（俱乐部名、地址、标语、logo）**全部自动带上**。覆盖用 `--time` / `--address` / `--slogan`。
脚本内部调通用技能 `overlay_text.py`，参数透传。

### Step 8 — 合成嘉宾照片（有嘉宾时）

调通用技能 `scripts/composite_guest.py`，按 `references/guest-compositing.md` 的「不突兀六原则」。
多位嘉宾串联跑（上一张输出当下一张的 `--base`），两人 `--center-x` 0.32/0.68，三人 0.24/0.50/0.76。

### Step 9 — 文字越界自动检测（必做）

用**差分法**（比单行中位数法准，后者会把背景元素误判成文字）：

```bash
# 先生成不叠字的纯背景
python overlay_text.py --bg bg.png --out co.png --canvas 1080x1440
```

```python
diff = np.abs(fg - bg) > 30                 # 成品 − 纯背景 = 真正的文字区域
# 文字区背景复杂度 < 3% 才算通过（字压在纯净留白上）
```

越界就重生成，不硬交付。

### Step 10 — 交付 + 登记历史

`present_files` 给成品 PNG。回复正文只写：意象一句话、会议信息核对表、需用户决策的点。
**禁止**贴 ImageGen prompt 原文、即梦参数、六段骨架。

然后**追加一行到 `references/egg-theme-history.md`**。

## 参考文件

| 文件 | 何时读 |
|---|---|
| `references/egg-profile.md` | 需要 EGG 品牌调性、愿景金句、表达规范时 |
| `references/egg-theme-history.md` | **每次出图前后**：出图前查撞车，出图后登记 |
| 通用 `references/style-library.md` | Step 2 选风格卡（**每次出图前必读**） |
| 通用 `references/prompt-skeleton.md` | Step 4 填骨架 |
| 通用 `references/theme-visual-map.md` | Step 3 匹配意象 |
| 通用 `references/guest-compositing.md` | 有嘉宾时必读 |
| 通用 `scripts/parse_activity.py` | Step 0 |
| 通用 `scripts/overlay_text.py` | Step 7（经本技能一键脚本调用） |
| 通用 `scripts/composite_guest.py` | Step 8 |

通用技能根目录：`C:/Users/Administrator/.workbuddy/skills/toastmasters-poster-prompt/`

## Python 环境

统一用隔离 venv：
`C:/Users/Administrator/.workbuddy/binaries/python/envs/default/Scripts/python.exe`
（不要用 `.../versions/3.13.12/python.exe`，其 site-packages 会被环境重置清空）

已装：Pillow、python-pptx、openpyxl、rembg、onnxruntime。
中文字体：`msyhbd.ttc`（默认）、`msyh.ttc`、`simhei.ttf`、`simkai.ttf`（楷体，风格 B 用）。
