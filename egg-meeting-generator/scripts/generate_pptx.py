import copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
import os

src = 'F:/新建文件夹(1)/EGG - 257次会议背景板.pptx'
dst = 'F:/workbuddy/2026-07-28-14-26-36/EGG - 258次会议背景板.pptx'

# Copy file
import shutil
shutil.copy2(src, dst)

prs = Presentation(dst)

def replace_text_in_shape(shape, old_text, new_text):
    """Replace text in a shape, handling both direct text and table cells"""
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            if old_text in para.text:
                # Need to handle runs for formatting preservation
                for run in para.runs:
                    if old_text in run.text:
                        run.text = run.text.replace(old_text, new_text)
                # Also check paragraph-level text
                full_text = para.text
                if old_text in full_text:
                    # Clear and rewrite
                    for run in para.runs:
                        run.text = ''
                    if para.runs:
                        para.runs[0].text = full_text.replace(old_text, new_text)
    if shape.has_table:
        for row in shape.table.rows:
            for cell in row.cells:
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)

def replace_all_text(slide, replacements):
    """Apply multiple text replacements across all shapes in a slide"""
    for shape in slide.shapes:
        for old, new in replacements:
            replace_text_in_shape(shape, old, new)

# ============ Global replacements (apply to ALL slides) ============
global_replacements = [
    ('EGG第254次中文线下会议', 'EGG第258次中文线下会议'),
    ('第254次中文线下会议【2026年7月5日】', '第258次中文线下会议【2026年8月2日】'),
    ('今日一词：创新', ''),  # Remove word of the day
    ('9:50', '9:30'),
    ('静安区西康路928号创展大厦516室', '徐汇区云锦路181号龙华街道水岸党群服务中心214室'),
    ('西康路928号创展大厦516室', '云锦路181号水岸党群服务中心214室'),
    ('龙华地铁站附近', '龙华地铁站附近'),
]

for slide in prs.slides:
    replace_all_text(slide, global_replacements)

# ============ Slide-specific replacements ============

# Slide 1: Title - 会议经理: Kim
replace_all_text(prs.slides[0], [
    ('会议经理：韩羽良@EGG', '会议经理：Kim@EGG'),
])

# Slide 2: 接待官 → SAA 待定
replace_all_text(prs.slides[1], [
    ('快乐哥@EGG', '待定'),
])

# Slide 3: 俱乐部主席欢迎辞 → still 快乐哥
# (no change needed)

# Slide 6: EGG俱乐部介绍 - update time
replace_all_text(prs.slides[5], [
    ('9:50 – 12:30', '9:30 – 12:00'),
])

# Slide 7: 主持人 → Kim (no change from template)

# Slide 8: 总评团队介绍
replace_all_text(prs.slides[7], [
    ('时间官: 岳高杰\n提问官: 小龙哥@自我探索', 
     '总评官: Ocean\n时间官: Rachel\n哼哈官: 徐朝霞\n语法官: 待定\n提问官: 吴琼'),
])

# Slide 9: 大咖分享 → 慕艺
replace_all_text(prs.slides[8], [
    ('张旻翔', '慕艺'),
    ('《用Agent亲手做出你的内容》', '《世界500强扭亏为盈的秘密》'),
])

# Slide 10: 提问官提问 → 吴琼
replace_all_text(prs.slides[9], [
    ('小龙哥@EGG', '吴琼'),
])

# Slide 11: 来宾介绍 → SAA 待定
replace_all_text(prs.slides[10], [
    ('水水@EGG', '待定@EGG'),
])

# Slide 13: 即兴主持 → Amy
replace_all_text(prs.slides[12], [
    ('Kim@EGG', 'Amy@EGG'),
])

# Slide 14: 备稿演讲1 → 程雪
replace_all_text(prs.slides[13], [
    ('陈慢慢@EGG', '程雪@EGG'),
    ('《介绍头马的导师计划》', '备稿演讲1'),
])

# Slide 15: 备稿演讲2 → 快乐哥
replace_all_text(prs.slides[14], [
    ('程雪@EGG', '快乐哥@EGG'),
    ('《接住》', '备稿演讲2'),
])

# Slide 16: 即兴评估 → 陶然
replace_all_text(prs.slides[15], [
    ('Alicia@Foodie&青蓝', '陶然'),
])

# Slide 17: 备稿个评1 → Mark
replace_all_text(prs.slides[16], [
    ('刘欣@青蓝', 'Mark'),
])

# Slide 18: 备稿个评2 → 韩羽良
replace_all_text(prs.slides[17], [
    ('陶然@见乐', '韩羽良'),
])

# Slide 19: 时间官报告 → Rachel
replace_all_text(prs.slides[18], [
    ('岳高杰', 'Rachel'),
])

# Slide 20: 投票&来宾反馈 → Ocean
replace_all_text(prs.slides[19], [
    ('快乐哥@EGG', 'Ocean'),
])

# Slide 21: 入会规则介绍 → 快乐哥 (or 程雪 since she's VPM? The template had 程雪)
# Actually 入会仪式 is separate from this. Let me keep 程雪 or change to 快乐哥.
# VPM 水泽枂 usually handles 入会, but 程雪 is secretary. Let me keep it flexible.
replace_all_text(prs.slides[20], [
    ('程雪@EGG', '快乐哥@EGG'),
])

# Slide 25: 颁奖与致谢 → 快乐哥 (no change)
# Slide 26: 感谢参与 → 会议经理 Kim
replace_all_text(prs.slides[25], [
    ('本期会议经理：韩羽良@EGG', '本期会议经理：Kim@EGG'),
])

# Slide 27: 下期预告 → 会议经理 快乐哥 or 待定
# (keep as-is or update)

# ============ Insert new slides ============
# We need: 备稿3 (Kim), 备稿个评3 (待定), 新会员入会仪式&生日会

# Function to duplicate a slide
def duplicate_slide(prs, slide_index):
    """Duplicate a slide and return the new slide"""
    template = prs.slides[slide_index]
    slide_layout = template.slide_layout
    
    # Add new slide with same layout
    new_slide = prs.slides.add_slide(slide_layout)
    
    # Copy all shapes from template
    # This is complex due to shape relationships, let's use a simpler approach
    # For now, just copy key elements
    return new_slide

# Actually duplicating slides with all formatting is very complex in python-pptx.
# Let me use a different approach: clone the XML directly.

from lxml import etree
import copy as xml_copy

def duplicate_slide_xml(prs, source_slide_idx):
    """Duplicate a slide by cloning its XML"""
    source_slide = prs.slides[source_slide_idx]
    
    # Get the slide layout used by source
    slide_layout = source_slide.slide_layout
    
    # Add a new slide
    new_slide_layout = slide_layout
    new_slide = prs.slides.add_slide(new_slide_layout)
    
    # Remove default shapes from new slide
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)
    
    # Copy all shapes from source to new slide
    for shape in source_slide.shapes:
        new_shape_elem = xml_copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(new_shape_elem)
    
    # Copy slide relationships (images, etc.)
    for rel in source_slide.part.rels.values():
        if "image" in rel.reltype or "chart" in rel.reltype:
            new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
    
    return new_slide

# After slide 15 (index 14, 备稿2): Insert 备稿3
# First, let me figure out the right place. Current structure:
# Slide 14 (idx 13): 备稿1
# Slide 15 (idx 14): 备稿2
# After that, we need 备稿3

# Let me duplicate slide 14 (备稿1 template) and modify for 备稿3
# But inserting in the middle is tricky. Let me just add at the end and reorder...

# Actually, python-pptx doesn't support slide reordering easily.
# Let me just add new slides after the existing ones.

# Add 备稿3 slide (copy from 备稿1 slide 14, index 13)
try:
    new_slide = duplicate_slide_xml(prs, 13)  # Copy 备稿1 slide
    replace_all_text(new_slide, [
        ('备稿演讲1', '备稿演讲3'),
        ('陈慢慢@EGG', 'Kim@EGG'),
        ('《介绍头马的导师计划》', '备稿演讲3'),
    ])
    print("Added 备稿3 slide")
    
    # Add 备稿个评3 (copy from 备稿个评1, slide 17 index 16)
    new_slide2 = duplicate_slide_xml(prs, 16)
    replace_all_text(new_slide2, [
        ('备稿个评1', '备稿个评3'),
        ('刘欣@青蓝', '待定'),
    ])
    print("Added 备稿个评3 slide")
    
    # Add 新会员入会仪式&生日会 (copy from 即兴评估 slide 16 index 15)
    new_slide3 = duplicate_slide_xml(prs, 15)
    replace_all_text(new_slide3, [
        ('即兴评估', '新会员入会仪式&生日会'),
        ('Alicia@Foodie&青蓝', '全体'),
    ])
    print("Added 新会员入会仪式&生日会 slide")
except Exception as e:
    print(f"Error adding slides: {e}")
    print("Continuing with base modifications only...")

# ============ Save ============
prs.save(dst)
print(f"\nPPT saved to: {dst}")
print("Done!")
