# -*- coding: utf-8 -*-
"""모의고사 성적 분석 엑셀 생성 (원점수 포함).

시트: 반별 / 석차별 / 등급별 비율
원본 2026_10__2.xlsx 서식을 그대로 재현하고 원점수 열을 추가함.
"""
import io

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

FONT = Font(name='맑은 고딕', size=10)
FONT_B = Font(name='맑은 고딕', size=10, bold=True)
FONT11 = Font(name='맑은 고딕', size=11)
FONT11_B = Font(name='맑은 고딕', size=11, bold=True)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
THIN = Side(style='thin', color='A6A6A6')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_GRAY = PatternFill('solid', fgColor='F2F2F2')
FILL_KOR = PatternFill('solid', fgColor='E2EFDA')   # 국어 초록
FILL_MATH = PatternFill('solid', fgColor='FFF2CC')  # 수학 노랑
FILL_ENG = PatternFill('solid', fgColor='D9E1F2')   # 영어 파랑
FILL_HIST = PatternFill('solid', fgColor='F2F2F2')  # 한국사 회색
FILL_TAM = PatternFill('solid', fgColor='FCE4D6')   # 탐구 주황
FILL_GRADE = PatternFill('solid', fgColor='D9D9D9')  # 등급별 비율 헤더

# (과목명, 하위 열, fill)
SUBJECTS = [
    ('국어', ['원점수', '표준점수', '백분위', '등급'], FILL_KOR),
    ('수학', ['원점수', '표준점수', '백분위', '등급'], FILL_MATH),
    ('영어', ['원점수', '등급'], FILL_ENG),
    ('한국사', ['원점수', '등급'], FILL_HIST),
    ('사회탐구', ['원점수', '표준점수', '백분위', '등급'], FILL_TAM),
    ('과학탐구', ['원점수', '표준점수', '백분위', '등급'], FILL_TAM),
]


def _header(ws):
    """반별/석차별 시트 1~2행 헤더."""
    ws.row_dimensions[1].height = 24.95
    ws.row_dimensions[2].height = 24.95
    # 반/번호/이름 (세로 병합)
    for col, name in [(1, '반'), (2, '번호'), (3, '이름')]:
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        c = ws.cell(row=1, column=col, value=name)
        c.font = FONT_B
        c.fill = FILL_GRAY
        c.alignment = CENTER
        c.border = BORDER
    col = 4
    for subj, subs, fill in SUBJECTS:
        n = len(subs)
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + n - 1)
        c = ws.cell(row=1, column=col, value=subj)
        c.font = FONT_B
        c.fill = fill
        c.alignment = CENTER
        c.border = BORDER
        for j, sub in enumerate(subs):
            c2 = ws.cell(row=2, column=col + j, value=sub)
            c2.font = FONT_B
            c2.fill = fill
            c2.alignment = CENTER
            c2.border = BORDER
        col += n
    # 표준점수 합 / 전교 석차 (세로 병합)
    for name in ['표준점수 합', '전교 석차']:
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)
        c = ws.cell(row=1, column=col, value=name)
        c.font = FONT_B
        c.fill = FILL_GRAY
        c.alignment = CENTER
        c.border = BORDER
        col += 1
    # 열 너비
    widths = {1: 6, 2: 6, 3: 10}
    col = 4
    for subj, subs, _ in SUBJECTS:
        for sub in subs:
            widths[col] = 9 if sub in ('표준점수', '백분위', '원점수') else 7
            col += 1
    widths[col] = 12
    widths[col + 1] = 10
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w


def _student_row(rec):
    """한 학생의 행 값 목록."""
    row = [rec['반'], rec['번호'], rec['이름']]
    for subj, subs, _ in SUBJECTS:
        d = rec[subj]
        for sub in subs:
            row.append(d.get(sub))
    row.append(rec['표준점수합'])
    row.append(rec['전교석차'])
    return row


def _fill_rows(ws, records):
    for i, rec in enumerate(records):
        r = i + 3
        for j, v in enumerate(_student_row(rec), start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.font = FONT
            c.alignment = CENTER
            c.border = BORDER


def _grade_sheet(wb, records):
    """등급별 비율 시트."""
    ws = wb.create_sheet('등급별 비율')
    ws.merge_cells('A1:B1')
    for col, name in [(1, '구   분'), (3, '1등급'), (4, '2등급'), (5, '3등급'),
                      (6, '4등급'), (7, '5등급'), (8, '6등급'), (9, '7등급'),
                      (10, '8등급'), (11, '9등급'), (12, '총인원')]:
        c = ws.cell(row=1, column=col, value=name)
        c.font = FONT11_B
        c.fill = FILL_GRADE
        c.alignment = CENTER
        c.border = BORDER
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 9
    for col in range(3, 13):
        ws.column_dimensions[get_column_letter(col)].width = 9
    row = 2
    for subj, _, _ in SUBJECTS:
        counts = [0] * 9
        total = 0
        for rec in records:
            g = rec[subj]['등급']
            if g is not None and 1 <= g <= 9:
                counts[g - 1] += 1
                total += 1
        ws.merge_cells(start_row=row, start_column=1, end_row=row + 1, end_column=1)
        c = ws.cell(row=row, column=1, value=subj)
        c.font = FONT11
        c.fill = FILL_GRADE
        c.alignment = CENTER
        c.border = BORDER
        for k, label in enumerate(['인원', '비율(%)']):
            c = ws.cell(row=row + k, column=2, value=label)
            c.font = FONT11
            c.fill = FILL_GRADE
            c.alignment = CENTER
            c.border = BORDER
            for g in range(9):
                if k == 0:
                    v = counts[g]
                else:
                    v = round(counts[g] / total * 100, 2) if total else None
                c2 = ws.cell(row=row + k, column=3 + g, value=v)
                c2.font = FONT11
                c2.alignment = CENTER
                c2.border = BORDER
            c3 = ws.cell(row=row, column=12, value=total if k == 0 else None)
            c3.font = FONT11
            c3.alignment = CENTER
            c3.border = BORDER
        row += 2
    return ws


def build_workbook(records):
    """분석 엑셀 생성 → bytes."""
    wb = Workbook()
    ws1 = wb.active
    ws1.title = '반별'
    _header(ws1)
    by_class = sorted(records, key=lambda r: (
        r['반'] if r['반'] is not None else 99,
        r['번호'] if r['번호'] is not None else 99))
    _fill_rows(ws1, by_class)

    ws2 = wb.create_sheet('석차별')
    _header(ws2)
    by_rank = sorted(records, key=lambda r: (
        r['전교석차'] if r['전교석차'] is not None else 9999,
        r['반'] if r['반'] is not None else 99,
        r['번호'] if r['번호'] is not None else 99))
    _fill_rows(ws2, by_rank)

    _grade_sheet(wb, records)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
