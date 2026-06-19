"""대입통합검색기_데이터정리본.xlsx → main.csv / min.csv / jong.csv

데이터 업데이트 절차:
  1. 대입통합검색기_데이터정리본.xlsx 에서 데이터 수정
  2. python build_csv.py 실행
  3. git add data/ && git commit && git push
"""
import openpyxl, csv, os, sys, re
sys.stdout.reconfigure(encoding='utf-8')

SRC = r'C:\Users\user\search\대입통합검색기_데이터템플릿.xlsx'
DATA_DIR = r'C:\Users\user\search\data'
os.makedirs(DATA_DIR, exist_ok=True)

SHEETS = [
    ('메인',     'main.csv'),
    ('최저기준', 'min.csv'),
    ('전형방법', 'jong.csv'),  # 시트명 '종합전형' → '전형방법'으로 변경됨
]

wb = openpyxl.load_workbook(SRC, data_only=True, read_only=True)

for sheet_name, out_name in SHEETS:
    if sheet_name not in wb.sheetnames:
        print(f'⚠ {sheet_name} 시트 없음')
        continue
    ws = wb[sheet_name]
    out_path = os.path.join(DATA_DIR, out_name)
    rows = []
    for row in ws.iter_rows(values_only=True):
        # 모든 값 None인 행은 건너뜀
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue
        rows.append([('' if v is None else str(v).strip()) for v in row])
    # 메인 시트: 중복된 '충원' 헤더를 직전 전형 기준으로 고유화 (교과1/충원 등)
    if sheet_name == '메인' and rows:
        hdr = rows[0]
        prefix = None
        for i, h in enumerate(hdr):
            if '/전형' in h:
                prefix = h.split('/')[0]      # 교과1, 교과2, 종합1 ...
            elif h == '정시백분위':
                prefix = '정시'
            elif h == '충원' and prefix:
                hdr[i] = prefix + '/충원'
        rows[0] = hdr
        # 권역 정규화 — 데이터를 검색기 6권역 필터에 맞춤
        KWON_MAP = {'서울권':'서울', '부울경권':'경상권', '대경권':'경상권', '호남권':'전라권'}
        kwon_idx = hdr.index('권역') if '권역' in hdr else 0
        for r in rows[1:]:
            if kwon_idx < len(r) and r[kwon_idx] in KWON_MAP:
                r[kwon_idx] = KWON_MAP[r[kwon_idx]]
        # 중복 모집단위 정리 — '학부' 접두사만 다른 중복은 완성도 높은 행만 남김
        # (단과대학 '○○대학' 접두는 보존 → 단과대별 별도 모집 유지)
        univ_i = hdr.index('대학명') if '대학명' in hdr else None
        dept_i = hdr.index('2026 수시 모집단위') if '2026 수시 모집단위' in hdr else None
        if univ_i is not None and dept_i is not None:
            cut_cols = [hdr.index(c) for c in
                        ['교과1/70%','교과2/70%','교과3/70%','종합1/70%','종합2/70%','정시백분위'] if c in hdr]
            def norm_dept(d):
                d = d.strip()
                m = re.match(r'^\S*학부\s+(.+)$', d)   # '○○학부 △△전공' → '△△전공'
                if m: d = m.group(1)
                return d.replace(' ', '')
            def score(r):
                return sum(1 for c in cut_cols if c < len(r) and r[c].strip())
            groups = {}
            for r in rows[1:]:
                groups.setdefault((r[univ_i].strip(), norm_dept(r[dept_i])), []).append(r)
            chosen = set()
            newrows = [hdr]
            for r in rows[1:]:
                key = (r[univ_i].strip(), norm_dept(r[dept_i]))
                g = groups[key]
                if len(g) == 1:
                    newrows.append(r); continue
                if key in chosen:
                    continue                         # 그룹 대표 이미 추가됨 → 중복 제거
                newrows.append(max(g, key=score))    # 완성도 최고 행만 (동점 시 첫째)
                chosen.add(key)
            removed = len(rows) - len(newrows)
            rows = newrows
            if removed:
                print(f'    └ 중복 모집단위 {removed}건 정리됨')
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f)
        w.writerows(rows)
    print(f'  {out_name}: {len(rows)-1:,} 데이터 행 ({os.path.getsize(out_path)/1024:.1f} KB)')

print('\n완료. 다음 명령으로 배포:')
print('  git add data/ && git commit -m "데이터 업데이트" && git push')
