"""대입통합검색기_데이터정리본.xlsx → main.csv / min.csv / jong.csv

데이터 업데이트 절차:
  1. 대입통합검색기_데이터정리본.xlsx 에서 데이터 수정
  2. python build_csv.py 실행
  3. git add data/ && git commit && git push
"""
import openpyxl, csv, os, sys
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
    with open(out_path, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f)
        w.writerows(rows)
    print(f'  {out_name}: {len(rows)-1:,} 데이터 행 ({os.path.getsize(out_path)/1024:.1f} KB)')

print('\n완료. 다음 명령으로 배포:')
print('  git add data/ && git commit -m "데이터 업데이트" && git push')
