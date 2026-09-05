# -*- coding: utf-8 -*-
"""EGG第261次会议 Agenda Excel + PPT背景板生成 v2（修正索引+换行符处理）"""
import shutil, sys, copy as cp
from datetime import time
import openpyxl
from pptx import Presentation
from pptx.oxml.ns import qn

sys.stdout.reconfigure(encoding='utf-8')

TPL_DIR = 'F:/新建文件夹(1)'
OUT_DIR = 'F:/workbuddy/2026-08-21-07-22-48'
XLSX_OUT = OUT_DIR + '/EGG261次会议Agenda.xlsx'
PPTX_OUT = OUT_DIR + '/EGG - 261次会议背景板.pptx'

# ============================================================
# Part 1: Excel
# ============================================================
shutil.copy2(TPL_DIR + '/EGG258次会议Agenda.xlsx', XLSX_OUT)
wb = openpyxl.load_workbook(XLSX_OUT)
ws = wb['正面']

# ---- 扩表：备注区(35-40行)下移两行，行35/36并入议程表格 ----
import copy as _cp
GRID_COLS = 'BCDEFGHIJ'
old_vals, old_styles, old_heights = {}, {}, {}
for r in range(35, 41):
    for c in GRID_COLS:
        cell = ws[f'{c}{r}']
        old_vals[(r, c)] = cell.value
        old_styles[(r, c)] = _cp.copy(cell._style)
    old_heights[r] = ws.row_dimensions[r].height
for m in list(ws.merged_cells.ranges):
    if m.min_row >= 35:
        ws.unmerge_cells(str(m))
for r in range(35, 43):
    for c in GRID_COLS:
        ws[f'{c}{r}'].value = None
# 写回到37-42行（值和样式都下移）
for (r, c), v in old_vals.items():
    ws[f'{c}{r+2}'].value = v
for r in range(35, 41):
    for c in GRID_COLS:
        ws[f'{c}{r+2}']._style = _cp.copy(old_styles[(r, c)])
for r, h in old_heights.items():
    if h is not None:
        ws.row_dimensions[r+2].height = h
for m in ['B37:B41', 'C37:J38', 'C39:J39', 'C40:J40', 'C41:J41', 'C42:J42']:
    ws.merge_cells(m)
# 行35/36复制行34的表格样式
for r in (35, 36):
    for c in GRID_COLS:
        ws[f'{c}{r}']._style = _cp.copy(ws[f'{c}34']._style)
    ws.row_dimensions[r].height = ws.row_dimensions[34].height
    ws.merge_cells(f'I{r}:J{r}')
ws.print_area = 'B1:J42'  # 备注区下移后扩展打印范围

# ---- 备稿目标：本次4个备稿，标注3/4 ----
ws['C40'] = ' 3.备稿目标：'
ws['C41'] = ' 4.备稿目标：'

ws['C3'] = '第261次中文线下会议【2026年8月23日】'
ws['C4'] = '会议主题：《怎么样做一篇好的备稿演讲》'
ws['B5'] = (
    '会议时间：\n 2026年8月23日（周日）\n上午 9:40 - 中午12:30\n\n'
    '会议地点：\n上海徐汇区云锦路181号龙华街道水岸党群服务中心214室（龙华地铁站附近）\n\n'
    '费用：所有来宾和角色全免费\n\n'
    '注意事项： \n◆ 会议4大禁忌：\n 政治 * 宗教 * 性 * 低俗话题 \n◆ 会场纪律:\n'
    '- 请将手机静音；\n- 不大声喧哗和随意走动；\n- 上下舞台请与主持人握手；\n- 为发言者鼓掌。'
)
ws['C7'] = time(9, 40)

agenda = [
    (7,  '暖场游戏', (3, 4, 5), '高祥'),
    (8,  '接待官介绍', (1, 1.5, 2), 'Sherry Shi'),
    (9,  '主席致辞', (2, 3, 4), '程雪@EGG'),
    (10, '会议主持人开场', (1, 1.5, 2), '程雪@EGG'),
    (11, '总评官介绍', (1, 1.5, 2), '待定'),
    (12, '时间官介绍', (1, 1.5, 2), '茉莉晴阳'),
    (13, '语法官介绍', (1, 1.5, 2), '查莫'),
    (14, '哼哈官介绍', (1, 1.5, 2), '徐朝霞@EGG'),
    (15, '提问官介绍', (1, 1.5, 2), '盐水'),
    (16, '《怎么样做一篇好的备稿演讲》', (35, 40, 45), '毕嵘'),
    (17, '提问官提问', (1, 2, 3), '=I15'),
    (18, '来宾介绍', (5, 6, 7), 'Sherry Shi'),
    (19, '中场休息', (4, 5, 6), None),
    (20, '即兴演讲', (13, 14, 15), 'Kim@EGG'),
    (21, '备稿1', (5, 6, 7), '徐朝霞@EGG'),
    (22, '备稿2', (5, 6, 7), 'Kim@EGG'),
    (23, '备稿3', (5, 6, 7), '慢慢@EGG'),
    (24, '备稿4', (5, 6, 7), '程雪@EGG'),
    (25, '即兴评估', (4, 5, 6), 'Doris'),
    (26, '备稿评估1', (2, 2.5, 3), 'Kim@EGG'),
    (27, '备稿评估2', (2, 2.5, 3), '向莹'),
    (28, '备稿评估3', (2, 2.5, 3), '橄榄树'),
    (29, '备稿评估4', (2, 2.5, 3), '水泽枂'),
    (30, '哼哈官报告', (1, 1.5, 2), '徐朝霞@EGG'),
    (31, '语法官报告', (1, 1.5, 2), '查莫'),
    (32, '时间官报告', (1, 1.5, 2), '茉莉晴阳'),
    (33, '总评官报告', (4, 5, 6), '待定'),
    (34, '投票&来宾反馈', (4, 5, 6), 'Kim@EGG'),
    (35, '入会规则介绍', (2, 3, 4), '程雪@EGG'),
    (36, '通告&颁奖', (2, 3, 4), '程雪@EGG'),
]
for row, content, dur, role in agenda:
    ws[f'E{row}'] = content
    ws[f'F{row}'], ws[f'G{row}'], ws[f'H{row}'] = dur
    ws[f'I{row}'] = role if role else None

# 时间级联公式（行35/36为新表格行）
for r in range(8, 37):
    ws[f'C{r}'] = f'=C{r-1}+TIME(,H{r-1},60)'
    ws[f'D{r}'] = f'=C{r}+TIME(,H{r}-1,60)'

ws['B21'] = (
    '口号：EGG！破壳，诞生无限！\n\n'
    '愿景：打造中国最具企业家精神的演讲实验室。让每一次开口，都诞生无限可能。\n\n'
    '使命：提供积极互助互益的学习体验，通过赋能会员提高演讲力和领导力，获得更大自信及个人成长\n\n'
    '会议经理：Kim\n'
    'IT场控：\n'
    '接待官：Sherry Shi\n'
    '拍照官：孙博'
)

wb['背面']['C3'] = '第261次中文线下会议【2026年8月23日】'
wb.save(XLSX_OUT)
print('Excel OK:', XLSX_OUT)

# ============================================================
# Part 2: PPT
# ============================================================
shutil.copy2(TPL_DIR + '/EGG - 258次会议背景板.pptx', PPTX_OUT)
prs = Presentation(PPTX_OUT)


def set_para_text(para, new_text):
    """写入段落文本；含 \v(软换行) 时按 a:br 重建，避免 _x000B_ 字面量"""
    p = para._p
    end = p.find(qn('a:endParaRPr'))
    rPr = None
    for r in p.findall(qn('a:r')):
        el = r.find(qn('a:rPr'))
        if el is not None:
            rPr = cp.deepcopy(el)
            break
    for child in list(p):
        if child.tag in (qn('a:r'), qn('a:br'), qn('a:fld')):
            p.remove(child)

    def add(el):
        if end is not None:
            end.addprevious(el)
        else:
            p.append(el)

    segments = new_text.split('\x0b')
    for i, seg in enumerate(segments):
        if i > 0:
            br = p.makeelement(qn('a:br'), {})
            if rPr is not None:
                br.append(cp.deepcopy(rPr))
            add(br)
        r = p.makeelement(qn('a:r'), {})
        if rPr is not None:
            r.append(cp.deepcopy(rPr))
        t = r.makeelement(qn('a:t'), {})
        t.text = seg
        r.append(t)
        add(r)


def replace_in_shape(shape, old, new):
    if not shape.has_text_frame:
        return False
    hit = False
    for para in shape.text_frame.paragraphs:
        if old in para.text:
            set_para_text(para, para.text.replace(old, new))
            hit = True
    return hit


def replace_all(slide, replacements):
    for shape in slide.shapes:
        for old, new in replacements:
            replace_in_shape(shape, old, new)


# ---- 全局：会议号 ----
for slide in prs.slides:
    replace_all(slide, [('第258次', '第261次')])

S = prs.slides
replace_all(S[1], [('待定', 'Sherry Shi')])                      # 2 接待官
replace_all(S[2], [('快乐哥@EGG', '程雪@EGG')])                   # 3 主席欢迎辞
replace_all(S[5], [('9:30', '9:40')])                            # 6 会议时间
replace_all(S[6], [('Kim@EGG', '程雪@EGG')])                     # 7 主持人
replace_all(S[8], [('慕艺@国学', '毕嵘'),
                   ('《世界500强扭亏为盈的秘密》', '《怎么样做一篇好的备稿演讲》')])  # 9 大咖
replace_all(S[9], [('吴琼@EGG', '盐水')])                        # 10 提问官
replace_all(S[11], [('程雪@EGG', 'Sherry Shi')])                 # 12 来宾介绍
replace_all(S[13], [('Amy@EGG', 'Kim@EGG')])                     # 14 即兴主持
replace_all(S[14], [('程雪@EGG', '徐朝霞@EGG')])                 # 15 备稿1
replace_all(S[15], [('快乐哥@EGG', 'Kim@EGG')])                  # 16 备稿2
replace_all(S[16], [('Kim@EGG', '慢慢@EGG')])                    # 17 备稿3
replace_all(S[17], [('陶然', 'Doris')])                          # 18 即兴评估
replace_all(S[18], [('Mark', 'Kim')])                            # 19 备稿个评1
replace_all(S[19], [('韩羽良', '向莹')])                          # 20 备稿个评2
replace_all(S[20], [('待定', '橄榄树')])                          # 21 备稿个评3
replace_all(S[22], [('待定', '查莫')])                            # 23 语法官报告
replace_all(S[23], [('Rachel', '茉莉晴阳')])                     # 24 时间官报告
replace_all(S[24], [('Ocean', '待定')])                          # 25 总评官报告
replace_all(S[25], [('快乐哥@EGG', 'Kim@EGG')])                  # 26 投票&来宾反馈
replace_all(S[30], [('快乐哥@EGG', '程雪@EGG')])                 # 32 颁奖与致谢

# ---- Slide 8 总评团队：整段重建为5行 ----
TEAM_LINES = ['总评官：待定', '时间官: 茉莉晴阳', '语法官：查莫', '哼哈官：徐朝霞@EGG', '提问官: 盐水']
done8 = False
for shape in S[7].shapes:
    if shape.has_text_frame:
        for para in shape.text_frame.paragraphs:
            if '总评官' in para.text:
                set_para_text(para, '\x0b'.join(TEAM_LINES))
                done8 = True
print('Slide8 重建:', done8)


def duplicate_slide_xml(prs, source_idx):
    src = prs.slides[source_idx]
    new_slide = prs.slides.add_slide(src.slide_layout)
    for shape in list(new_slide.shapes):
        sp = shape._element
        sp.getparent().remove(sp)
    for shape in src.shapes:
        new_slide.shapes._spTree.append(cp.deepcopy(shape._element))
    for rel in src.part.rels.values():
        if 'image' in rel.reltype or 'chart' in rel.reltype:
            new_slide.part.rels.get_or_add(rel.reltype, rel.target_part)
    return new_slide


def slide_text(slide):
    return ' '.join(s.text_frame.text for s in slide.shapes if s.has_text_frame)


def move_after(prs, moving_idx, target_idx):
    lst = prs.slides._sldIdLst
    ids = list(lst)
    el = ids[moving_idx]
    lst.remove(el)
    lst.insert(target_idx + 1, el)


dup1 = duplicate_slide_xml(prs, 16)   # 备稿演讲4
replace_all(dup1, [('备稿演讲3', '备稿演讲4'), ('慢慢@EGG', '程雪@EGG')])
dup2 = duplicate_slide_xml(prs, 20)   # 备稿个评4
replace_all(dup2, [('备稿个评3', '备稿个评4'), ('橄榄树', '水泽枂')])

n = len(prs.slides)  # 35
move_after(prs, n - 2, 16)   # 备稿4 → 备稿3后
move_after(prs, n - 1, 21)   # 个评4 → 个评3后

# ---- 删除拜师仪式页（模板index 10，移动操作不影响其位置）----
sldIdLst = prs.slides._sldIdLst
target = list(sldIdLst)[10]
rId = target.get(qn('r:id'))
sldIdLst.remove(target)
prs.part.drop_rel(rId)
print('拜师仪式页已删除, 剩余页数:', len(prs.slides))

prs.save(PPTX_OUT)
print('PPT OK:', PPTX_OUT)

# ============================================================
# 验证
# ============================================================
import zipfile, re
z = zipfile.ZipFile(PPTX_OUT)
bad = [n for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml') and b'_x000B_' in z.read(n)]
print('含_x000B_字面量的页:', bad if bad else '无')

prs2 = Presentation(PPTX_OUT)
print('\n=== PPT 共', len(prs2.slides), '页 ===')
for i, s in enumerate(prs2.slides):
    t = slide_text(s).replace('\x0b', '|').replace('\n', '|')
    print(f'{i+1:2d}. {t[:80]}')
