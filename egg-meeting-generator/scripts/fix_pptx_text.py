from pptx import Presentation

pptx_path = 'F:/workbuddy/2026-07-28-14-26-36/EGG - 258次会议背景板.pptx'
prs = Presentation(pptx_path)

def fix_text_in_shape(shape, old, new):
    """Replace text at paragraph level, handling runs"""
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            full = para.text
            if old in full:
                # Clear all runs
                for run in para.runs:
                    run.text = ''
                # Set new text in first run
                if para.runs:
                    para.runs[0].text = full.replace(old, new)
                else:
                    # No runs, add one
                    para.text = full.replace(old, new)

def fix_all_shapes(slide, replacements):
    for shape in slide.shapes:
        for old, new in replacements:
            fix_text_in_shape(shape, old, new)

# Fix Slide 8 (index 7): 总评团队介绍
fix_all_shapes(prs.slides[7], [
    ('时间官: 岳高杰', '时间官: Rachel'),
    ('提问官: 小龙哥@自我探索', '提问官: 吴琼'),
])

# The original text was "时间官: 岳高杰\n提问官: 小龙哥@自我探索"
# But the \n was rendered as \n vs vertical tab. Let me try more patterns
for shape in prs.slides[7].shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            text = para.text
            if '岳高杰' in text:
                for run in para.runs:
                    if '岳高杰' in run.text:
                        run.text = run.text.replace('岳高杰', 'Rachel')
            if '小龙哥' in text:
                for run in para.runs:
                    if '小龙哥' in run.text:
                        run.text = run.text.replace('小龙哥@自我探索', '吴琼')

# Fix Slide 9 (index 8): 大咖分享 - speaker name
for shape in prs.slides[8].shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            if '张旻翔' in para.text:
                for run in para.runs:
                    run.text = run.text.replace('张旻翔', '慕艺')

# Fix Slide 10 (index 9): 提问官提问
for shape in prs.slides[9].shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            if '小龙哥' in para.text:
                for run in para.runs:
                    run.text = run.text.replace('小龙哥', '吴琼')

# Fix Slide 28 (index 27): 备稿3 - should be Kim@EGG not 程雪@EGG
# Wait, I duplicated slide 14 (备稿1 slide, index 13) which had 陈慢慢@EGG
# But my replacement for 陈慢慢@EGG → 程雪@EGG was already done globally on slide 14
# Then for slide 28 I tried to replace 陈慢慢@EGG → Kim@EGG but it was already 程雪@EGG
# Let me check...
for shape in prs.slides[27].shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if '程雪' in run.text:
                    run.text = run.text.replace('程雪@EGG', 'Kim@EGG')
                if '陈慢慢' in run.text:
                    run.text = run.text.replace('陈慢慢@EGG', 'Kim@EGG')

# Fix Slide 29 (index 28): 备稿个评3 - should be 待定 not Mark
for shape in prs.slides[28].shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if 'Mark' in run.text:
                    run.text = run.text.replace('Mark', '待定')

# Also fix 总评团队介绍 to include all roles properly
for shape in prs.slides[7].shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            full = para.text
            if full and ('Rachel' in full or '吴琼' in full):
                # Check if we need to add more info
                pass

prs.save(pptx_path)
print("Fixed PPT issues, saved.")
