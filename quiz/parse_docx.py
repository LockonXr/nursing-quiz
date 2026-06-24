#!/usr/bin/env python3
"""解析二季度历年考试试卷.docx → questions.json (v4 - stream approach)"""
import re, json
from pathlib import Path
from docx import Document

DOCX = Path(__file__).parent.parent / "二季度历年考试试卷.docx"
OUT = Path(__file__).parent / "questions.json"

# 读取并清理所有段落
doc = Document(str(DOCX))
raw_paras = []
for p in doc.paragraphs:
    t = p.text
    for sp in ['\xa0', ' ', ' ', '　', '\t', ' ', '\r']:
        t = t.replace(sp, ' ')
    t = t.strip()
    if t:
        raw_paras.append(t)

# ===== 合并全文档为流式文本 =====
# 段落间用换行分隔（方便后续按行号匹配）
full_text = '\n'.join(raw_paras)

def clean(s):
    return re.sub(r'\s+', ' ', s).strip()

def extract_answer(text):
    """提取答案: 返回 (剩余文本, 答案)"""
    # 1. 括号答案: ( B ) ( ABCD ) ( √ ) ( 全部 ) ( A B ) 等
    for m in re.finditer(r'[\(（]\s*([A-F√×\s]+|全部)\s*[\)）]', text):
        ans = re.sub(r'\s+', '', m.group(1))
        if not ans:
            continue  # 跳过空括号
        if '全部' in ans:
            ans = 'ABCD'
        rest = clean(text[:m.start()] + ' ' + text[m.end():])
        return rest, ans

    # 2. 找选项A位置，往前提取答案字母
    a_m = re.search(r'(?<!\w)A\s*[、．\.。]', text)
    if a_m:
        before = text[:a_m.start()]
        # before末尾的连续大写字母即为答案（可能跟？）
        m = re.search(r'([A-F]{1,6})\s*[？?]?\s*$', before)
        if m:
            ans = m.group(1)
            rest = clean(before[:m.start()] + ' ' + text[a_m.start():])
            return rest, ans

    return text, ''

def parse_options(text):
    """从文本解析选项字典"""
    opts = {}
    # 找所有选项标签位置
    pattern = r'(?<!\w)([A-F])\s*[、．\.。]\s*'
    matches = list(re.finditer(pattern, text))
    for i, m in enumerate(matches):
        label = m.group(1)
        start = m.end()
        end = matches[i+1].start() if i+1 < len(matches) else len(text)
        value = clean(text[start:end])
        # 移除尾部残留的题号/章节标题
        value = re.sub(r'\s+\d{1,2}\s*[、．\.。]?\s*$', '', value).strip()
        # 移除尾部括号答案
        value = re.sub(r'[\(（]\s*[A-F√×\s]*\s*[\)）]\s*$', '', value).strip()
        if len(value) >= 2:
            opts[label] = value
    return opts

def parse_questions_from_chunk(text, section_name):
    """从一段题型文本中提取所有题目"""
    # 按题号切分: "1、", "1.", "1．", "1 ", "1)" 等
    parts = re.split(r'(?:^|\n|\s{2,})(\d{1,2})\s*[、．\.。)]\s*', text)
    questions = []
    for i in range(1, len(parts), 2):
        qnum = parts[i]
        qtext = clean(parts[i+1]) if i+1 < len(parts) else ''
        if len(qtext) < 5:
            continue

        # 是非题特殊处理
        if section_name == '是非题':
            tf_m = re.search(r'[\(（]\s*([√×])\s*[\)）]', qtext)
            answer = tf_m.group(1) if tf_m else ''
            stem = clean(qtext[:tf_m.start()] + qtext[tf_m.end():]) if tf_m else qtext
            options = {'正确': '√', '错误': '×'}
            # 移除题干尾部句号
            stem = re.sub(r'[。．.]$', '', stem).strip()
            if len(stem) >= 3:
                questions.append({
                    'stem': stem, 'options': options,
                    'answer': answer, 'note': ''
                })
            continue

        # 提取答案
        cleaned, answer = extract_answer(qtext)
        # 解析选项
        options = parse_options(cleaned)

        # 清理题干: 移除选项文本
        stem = cleaned
        for lbl, txt in sorted(options.items(), key=lambda x: -len(x[1])):
            stem = stem.replace(txt, '')
        stem = re.sub(r'\s*[A-F]\s*[、．\.。]\s*', ' ', stem)
        stem = re.sub(r'[\(（]\s*[A-F√×\s]*\s*[\)）]\s*$', '', stem).strip()
        stem = re.sub(r'\s+[A-F]{1,6}\s*$', '', stem).strip()
        stem = clean(stem)

        if not stem or len(stem) < 2:
            continue

        note = '旧版' if '旧版' in stem else ''
        stem = stem.replace('旧版', '').strip()

        questions.append({
            'stem': stem, 'options': options,
            'answer': answer, 'note': note
        })

    return questions


# ===== 主流程: 流式解析 =====

# Step 1: 按试卷标题切分（直接在 full_text 上搜索，不包裹）
paper_titles = list(re.finditer(r'(202\d年试卷（[一二三四五六]）)', full_text))
unknown_titles = list(re.finditer(r'(不知年份)', full_text))
all_titles = sorted(paper_titles + unknown_titles, key=lambda x: x.start())

exam_papers = []
for idx, title_match in enumerate(all_titles):
    paper_name = title_match.group(1)
    # 试卷正文从标题后开始
    body_start = title_match.end()
    body_end = all_titles[idx+1].start() if idx+1 < len(all_titles) else len(full_text)
    # 跳过标题行后的换行符
    paper_text = full_text[body_start:body_end].strip()

    # Step 2: 在试卷文本内按题型标题切分
    # 题型标题模式: "一、单选题" "单选题" "二、多选题(2分/题)" 等
    sec_pattern = re.compile(
        r'(?:[一二三四五六]+\s*[、．]?\s*)?(单选题|多选题|是非题|简答题)\s*(?:\(\d+分/题\))?'
    )
    sec_matches = list(sec_pattern.finditer(paper_text))

    sections_data = []
    for si, sm in enumerate(sec_matches):
        sname = sm.group(1)
        sstart = sm.end()
        send = sec_matches[si+1].start() if si+1 < len(sec_matches) else len(paper_text)
        sec_text = paper_text[sstart:send].strip()

        if not sec_text:
            continue

        questions = parse_questions_from_chunk(sec_text, sname)
        qtype_map = {
            '单选题': 'single', '多选题': 'multi',
            '是非题': 'truefalse', '简答题': 'essay'
        }

        qlist = []
        for qi, q in enumerate(questions):
            qlist.append({
                'id': f'{paper_name}_{sname}_{qi+1}',
                'paper': paper_name,
                'type': qtype_map.get(sname, 'single'),
                'stem': q['stem'],
                'options': q['options'],
                'answer': q['answer'],
                'note': q.get('note', ''),
            })

        if qlist:
            sections_data.append({'section': sname, 'questions': qlist})

    exam_papers.append({'paper': paper_name, 'sections': sections_data})


# ===== 统计 =====
print(f'Exam papers: {len(exam_papers)}')
total = 0
no_ans = 0
for ep in exam_papers:
    counts = ', '.join(f'{s["section"]}:{len(s["questions"])}'
                       for s in ep['sections'])
    ep_total = sum(len(s['questions']) for s in ep['sections'])
    ep_no_ans = sum(1 for s in ep['sections'] for q in s['questions'] if not q['answer'])
    total += ep_total
    no_ans += ep_no_ans
    missing = f'  !! missing {ep_no_ans} answers' if ep_no_ans > 0 else ''
    print(f'  {ep["paper"]}: {counts} ({ep_total}){missing}')

print(f'\nTotal: {total}')
print(f'Missing answers: {no_ans}')

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(exam_papers, f, ensure_ascii=False, indent=2)
print(f'Output: {OUT}')
