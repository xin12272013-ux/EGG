# -*- coding: utf-8 -*-
"""
EGG「诞」头马演讲俱乐部 · 海报一键叠字

薄封装：把 EGG 的固定信息（俱乐部名 / 地址 / 标语 / 画布）自动填好，
底层调用通用技能 toastmasters-poster-prompt 的 overlay_text.py。

设计原则：不复制通用脚本。通用技能改了，这里自动受益。

用法:
    # 最少参数：只要主题 + 届次
    python make_egg_poster.py --bg bg.png --out EGG263.png \
        --theme "备稿马拉松" --session "第 263 次会议"

    # 接龙/活动信息解析后直接喂 JSON（parse_activity.py --out parsed.json）
    python make_egg_poster.py --bg bg.png --out EGG263.png --parsed parsed.json

    # 覆盖默认值
    python make_egg_poster.py --bg bg.png --out EGG263.png \
        --theme "破冰" --session "第 264 次会议" \
        --time "9月13日（周日）9:40-12:30" --slogan "让你开口，诞生无限"

    # 只打印将要执行的命令，不真的跑
    python make_egg_poster.py --bg bg.png --out x.png --theme "破冰" \
        --session "第 264 次会议" --dry-run
"""

import argparse
import json
import os
import subprocess
import sys

# ---------------------------------------------------------------- EGG 固定信息
CLUB = "上海 EGG「诞」头马演讲俱乐部"
ADDRESS = "上海市徐汇区云锦路 181 号 龙华街道水岸党群服务中心 214 室（龙华地铁站附近）"
SLOGAN = "企业家精神 · 给予帮助 · 持续成长"
CANVAS = "1080x1440"

# 通用技能路径（脚本单一数据源，不复制）
COMMON_SKILL = r"C:/Users/Administrator/.workbuddy/skills/toastmasters-poster-prompt"
OVERLAY = os.path.join(COMMON_SKILL, "scripts", "overlay_text.py")

PY = sys.executable


def _load_parsed(path):
    """读 parse_activity.py 的 JSON 输出，返回字段字典。"""
    with open(path, "r", encoding="utf-8") as f:
        d = json.load(f)
    return d


def build_cmd(args):
    """组装 overlay_text.py 的命令行。

    优先级：**显式参数 > 接龙解析值 > EGG 默认**。
    EGG 默认必须最后兜底——接龙里的地址即使写得比默认值简略，
    也是当次确认过的，优先级高于默认值。改了场地却用默认地址，比没海报更糟。
    """
    # 1) 显式参数
    vals = {k: (getattr(args, k) or "") for k in
            ("club", "theme", "session", "time", "address", "slogan")}

    # 2) 接龙解析值（只补空位，不覆盖显式参数）
    if args.parsed:
        d = _load_parsed(args.parsed)
        for key in ("club", "theme", "session", "time", "address"):
            if not vals[key] and d.get(key):
                vals[key] = d[key]
        # 嘉宾：JSON 里有嘉宾就自动开两列信息条
        if d.get("guests") and not args.no_guest:
            args.guest = True

    # 3) EGG 默认兜底，并标记哪些是默认（交付时要提示用户确认）
    used_default = {}
    for key, default in (("club", CLUB), ("address", ADDRESS), ("slogan", SLOGAN)):
        used_default[key] = not vals[key]
        if not vals[key]:
            vals[key] = default
    used_default["time"] = not vals["time"]
    used_default["session"] = not vals["session"]
    used_default["theme"] = not vals["theme"]

    cmd = [PY, OVERLAY, "--bg", args.bg, "--out", args.out,
           "--canvas", args.canvas or CANVAS]
    for key in ("club", "theme", "session", "time", "address", "slogan"):
        if vals[key]:
            cmd += ["--" + key, vals[key]]
    if args.guest:
        cmd.append("--guest")
    if args.plate:
        cmd.append("--plate")
    if args.no_logo:
        cmd.append("--no-logo")
    return cmd, vals, used_default


def main():
    ap = argparse.ArgumentParser(
        description="EGG「诞」头马演讲俱乐部海报一键叠字（自动带上 EGG 固定信息）")
    ap.add_argument("--bg", required=True, help="AI 生成的背景图（任意尺寸）")
    ap.add_argument("--out", required=True, help="输出文件")
    ap.add_argument("--theme", default="", help="会议主题大字（必填，没有主题就没有意象锚点）")
    ap.add_argument("--session", default="", help="如：第 263 次会议")
    ap.add_argument("--time", default="", help="如：9月6日（周日）9:40-12:30")
    ap.add_argument("--address", default="", help="覆盖 EGG 默认地址")
    ap.add_argument("--slogan", default="", help="覆盖默认标语")
    ap.add_argument("--club", default="", help="覆盖默认俱乐部名（一般不用）")
    ap.add_argument("--canvas", default="", help="目标画布，默认 1080x1440")
    ap.add_argument("--parsed", default="",
                    help="parse_activity.py 的 JSON 输出，自动填入届次/主题/时间/地址/嘉宾")
    ap.add_argument("--guest", action="store_true", help="有分享嘉宾，信息条改两列")
    ap.add_argument("--no-guest", action="store_true", help="强制不用两列（覆盖 --parsed 判断）")
    ap.add_argument("--plate", action="store_true", help="背景较花时给文字加柔和衬底")
    ap.add_argument("--no-logo", action="store_true", help="不贴头马 logo（默认贴左上角）")
    ap.add_argument("--dry-run", action="store_true", help="只打印命令，不执行")
    args = ap.parse_args()

    if not os.path.isfile(OVERLAY):
        sys.exit("[ERROR] 找不到通用技能脚本：%s\n"
                 "        本技能复用通用技能 toastmasters-poster-prompt 的脚本，"
                 "请确认该技能已安装。" % OVERLAY)
    if not os.path.isfile(args.bg):
        sys.exit("[ERROR] 背景图不存在：%s" % args.bg)
    if not args.theme and not args.parsed:
        sys.exit("[ERROR] 必须给 --theme（会议主题），或用 --parsed 传入解析结果。\n"
                 "        没有主题就没有意象锚点，本脚本不会替你猜。")

    cmd, vals, used_default = build_cmd(args)

    print("[EGG 海报] 俱乐部：%s" % vals["club"])
    print("[EGG 海报] 主题　：%s" % (vals["theme"] or "（未填）"))
    print("[EGG 海报] 届次　：%s" % (vals["session"] or "（未填）"))
    print("[EGG 海报] 时间　：%s" % (vals["time"] or "（未填，请务必补上）"))
    print("[EGG 海报] 地址　：%s" % vals["address"])
    print("[EGG 海报] 标语　：%s" % vals["slogan"])

    tips = []
    if used_default["address"]:
        tips.append("地址用的是 EGG 默认值，请与当次接龙核对")
    if used_default["time"]:
        tips.append("时间未填——EGG 常规是周日上午，但必须填当次接龙的具体时间")
    if used_default["slogan"]:
        tips.append("标语用的是 EGG 默认愿景三词；主题偏内省/疗愈时可改金句"
                    "「万物皆有裂痕，那是光照进来的地方」")
    if tips:
        print("[提示]")
        for t in tips:
            print("  - " + t)

    if args.dry_run:
        print("\n[DRY-RUN] 将执行：")
        print("  " + " ".join('"%s"' % c if " " in c else c for c in cmd))
        return

    print()
    subprocess.run(cmd, check=True)
    print("\n[OK] 成品：%s" % os.path.abspath(args.out))
    print("[下一步] 别忘了往 references/egg-theme-history.md 追加一行登记本期意象")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    main()
