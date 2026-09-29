"""두 엑셀 파일을 사업자등록번호 기준으로 비교 (2026-09-29)
A: 수임업체 세부내역 (시트 '세부내역', 5행 헤더)
B: 기장의무 구분 목록 (Sheet1)
사용: python compare.py A.xlsx B.xlsx 결과.xlsx
"""
import re, sys
import pandas as pd

n = lambda s: re.sub(r'\D', '', str(s)) if pd.notna(s) else ''
t = lambda s: str(s).strip() if pd.notna(s) else ''

fa, fb, out = sys.argv[1:4]
a = pd.read_excel(fa, header=4, dtype=str)
b = pd.read_excel(fb, header=0, dtype=str).iloc[1:]
b.columns = ['구분', '기장의무구분', '대표자명', '사업자등록번호', '주민(법인)번호', '상호']
a['k'] = a['사업자등록번호'].map(n)
b['k'] = b['사업자등록번호'].map(n)
b_nobrn = b[b.k == '']
b = b[b.k != '']

acols = ['구분', '(신)업체코드', '업체명', '대표자명', '주민/법인 등록번호', '사업자등록번호', '부가세대상', '비고1']
a_only = a[~a.k.isin(b.k)][acols]
b_only = b[~b.k.isin(a.k)].drop(columns='k')

m = a.merge(b, on='k', suffixes=('_A', '_B'))
rows = []
for _, r in m.iterrows():
    def add(item, va, vb):
        rows.append({'사업자등록번호': r['사업자등록번호_A'], '업체명(A)': r['업체명'],
                     '항목': item, 'A 파일': va, 'B 파일': vb})
    if t(r['대표자명_A']) != t(r['대표자명_B']):
        add('대표자명', r['대표자명_A'], r['대표자명_B'])
    if n(r['주민/법인 등록번호']) != n(r['주민(법인)번호']):
        add('주민/법인번호', r['주민/법인 등록번호'], r['주민(법인)번호'])
    va, vb = t(r['부가세대상']), t(r['기장의무구분'])
    if (va == '간이과세') != (vb == '간이'):
        add('과세유형', va, vb)
diff = pd.DataFrame(rows)

with pd.ExcelWriter(out) as w:
    diff.to_excel(w, sheet_name='항목불일치', index=False)
    a_only.to_excel(w, sheet_name='A에만있음(B누락)', index=False)
    b_only.to_excel(w, sheet_name='B에만있음(A누락)', index=False)
    b_nobrn.drop(columns='k').to_excel(w, sheet_name='B_사업자번호없음', index=False)
print(f'공통 {len(m)} / 불일치 {len(diff)} / A만 {len(a_only)} / B만 {len(b_only)}')
