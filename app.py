# -*- coding: utf-8 -*-
"""모의고사 성적 분석 — 성적통지표 PDF → 원점수 포함 분석 엑셀

사용법: 반별 성적통지표 PDF(예: 2-1.pdf, 2-2.pdf, ...)를 모두 업로드하면
반별 / 석차별 / 등급별 비율 3개 시트가 있는 엑셀을 만들어 줍니다.
원점수·표준점수·백분위·등급이 모두 들어갑니다.
"""
import streamlit as st

from mock_excel import build_workbook
from mock_parser import add_ranks, normalize, parse_pdf_bytes

st.set_page_config(page_title='모의고사 성적 분석', layout='centered')
st.title('📊 모의고사 성적 분석')
st.write(
    '학생별 성적통지표 PDF를 올리면 원점수가 포함된 분석 엑셀을 만들어 줍니다. '
    '반별 PDF를 **모두** 올리면 전교 석차가 맞게 계산됩니다.'
)

files = st.file_uploader(
    '성적통지표 PDF 업로드 (여러 개 가능)',
    type=['pdf'], accept_multiple_files=True)

if files:
    all_students = []
    problems = []
    with st.spinner('PDF를 읽는 중…'):
        for f in files:
            try:
                sts = parse_pdf_bytes(f.read())
                if not sts:
                    problems.append(f'{f.name}: 학생을 찾지 못했습니다.')
                all_students.extend(sts)
            except Exception as e:  # noqa: BLE001
                problems.append(f'{f.name}: {e}')
    if problems:
        for p in problems:
            st.warning(p)
    if all_students:
        recs = add_ranks(normalize(all_students))
        classes = sorted({r['반'] for r in recs if r['반'] is not None})
        st.success(f'총 {len(recs)}명 ({", ".join(f"{c}반" for c in classes)})을 읽었습니다.')
        # 미리보기: 반별 상위 5명
        st.subheader('미리보기 (표준점수 합 상위 5명)')
        top5 = sorted(
            [r for r in recs if r['전교석차'] is not None],
            key=lambda r: r['전교석차'])[:5]
        st.table([{
            '반': r['반'], '번호': r['번호'], '이름': r['이름'],
            '국어 원점수': r['국어']['원점수'],
            '수학 원점수': r['수학']['원점수'],
            '표준점수 합': r['표준점수합'], '전교 석차': r['전교석차'],
        } for r in top5])
        # 미응시/결측 안내
        missing = [r for r in recs if any(
            r[s]['표준점수'] is None for s in ['국어', '수학', '사회탐구', '과학탐구'])]
        if missing:
            st.info(f'표준점수가 비어 있는 학생 {len(missing)}명: '
                    + ', '.join(f"{r['반']}반 {r['번호']}번 {r['이름']}" for r in missing[:10])
                    + (' …' if len(missing) > 10 else ''))
        if st.button('📥 분석 엑셀 만들기'):
            with st.spinner('엑셀을 만드는 중…'):
                data = build_workbook(recs)
            st.download_button(
                '⬇️ 분석 엑셀 다운로드',
                data=data,
                file_name='모의고사_성적분석.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            st.caption('시트: 반별 / 석차별 / 등급별 비율 — 원점수 포함')
else:
    st.info('PDF 파일을 올려 주세요. (예: 2-1.pdf, 2-2.pdf, …)')
    st.caption('성적통지표 PDF는 학생 1명당 2쪽(성적표+해설)으로 되어 있습니다.')
