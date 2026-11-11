# -*- coding: utf-8 -*-
"""모의고사 성적통지표 PDF 파서.

입력: 학생별 성적통지표 PDF (각 학생 2쪽: 성적표 + 해설)
출력: 학생 목록 [{학년, 반, 번호, 이름, 과목: {과목명: {원점수, 표준점수, 백분위, 등급}}}]
"""
from collections import defaultdict

import fitz


def _num(x):
    """숫자 문자열을 float/int로. 실패 시 None."""
    if x is None:
        return None
    try:
        f = float(str(x).replace(',', ''))
        return int(f) if f == int(f) else f
    except (ValueError, TypeError):
        return None


def parse_student_page(page):
    """한 학생의 성적표 페이지에서 성적 추출."""
    words = page.get_text('words')
    rows = defaultdict(list)
    for w in words:
        rows[round(w[1])].append(w)
    for y in rows:
        rows[y] = sorted(rows[y], key=lambda w: w[0])

    # 학생 정보: y~37 행 (학교명/학년/반/번호/성명/실시일)
    info = {}
    for y in sorted(rows):
        if 30 <= y <= 45:
            for w in rows[y]:
                x, txt = w[0], w[4]
                if 540 <= x <= 560:
                    info['학년'] = txt
                elif 580 <= x <= 595:
                    info['반'] = txt
                elif 615 <= x <= 630:
                    info['번호'] = txt
                elif 665 <= x <= 700:
                    info['이름'] = txt
    if '이름' not in info:
        return None

    # 과목 행 (y 70~175)
    subjects = {}
    for y in sorted(rows):
        if not (70 <= y <= 175):
            continue
        line = rows[y]
        first = None
        for w in line:
            if 55 <= w[0] <= 100:
                first = w
                break
        if first is None:
            continue
        t = first[4]
        subj = {'국': '국어', '수': '수학', '영': '영어',
                '한국사': '한국사', '사회': '사회탐구', '과학': '과학탐구'}.get(t)
        if subj is None:
            continue
        vals = {}
        for w in line:
            x, txt = w[0], w[4]
            if 130 <= x <= 150:
                vals['배점'] = txt
            elif 170 <= x <= 190:
                vals['원점수'] = txt
            elif 245 <= x <= 265:
                vals['표준점수'] = txt
            elif 355 <= x <= 380:
                vals['백분위'] = txt
            elif 405 <= x <= 420:
                vals['등급'] = txt
        subjects[subj] = vals
    info['과목'] = subjects
    return info


def parse_pdf_bytes(data: bytes):
    """PDF 바이트에서 학생 목록 추출."""
    doc = fitz.open(stream=data, filetype='pdf')
    students = []
    for i in range(len(doc)):
        if i % 2 == 1:  # 짝수 페이지는 해설 페이지
            continue
        st = parse_student_page(doc[i])
        if st:
            students.append(st)
    return students


def parse_pdf(path: str):
    with open(path, 'rb') as f:
        return parse_pdf_bytes(f.read())


def normalize(students):
    """파싱 결과를 분석용 dict 목록으로 정리."""
    out = []
    for st in students:
        rec = {
            '학년': _num(st.get('학년')),
            '반': _num(st.get('반')),
            '번호': _num(st.get('번호')),
            '이름': st.get('이름'),
        }
        for subj in ['국어', '수학', '영어', '한국사', '사회탐구', '과학탐구']:
            v = st.get('과목', {}).get(subj, {})
            rec[subj] = {
                '원점수': _num(v.get('원점수')),
                '표준점수': _num(v.get('표준점수')),
                '백분위': _num(v.get('백분위')),
                '등급': _num(v.get('등급')),
            }
        # 표준점수 합 = 국어+수학+사회탐구+과학탐구 (있는 것만)
        tot = 0
        for s in ['국어', '수학', '사회탐구', '과학탐구']:
            p = rec[s]['표준점수']
            if p is not None:
                tot += p
        rec['표준점수합'] = tot if tot else None
        out.append(rec)
    return out


def add_ranks(records):
    """표준점수 합 기준 전교 석차 부여 (동점자 동일 석차, competition ranking)."""
    scored = sorted(
        [r for r in records if r['표준점수합'] is not None],
        key=lambda r: -r['표준점수합'])
    rank_of = {}
    prev_score, prev_rank = None, 0
    for i, r in enumerate(scored):
        if r['표준점수합'] == prev_score:
            rank = prev_rank
        else:
            rank = i + 1
        rank_of[id(r)] = rank
        prev_score, prev_rank = r['표준점수합'], rank
    for r in records:
        r['전교석차'] = rank_of.get(id(r))
    return records
