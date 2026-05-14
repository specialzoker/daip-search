"""수시사례 376k행 → 학과별 합격 통계 JSON
- 교과: 계열별 환산등급 사용 (인문/예체능 = 국영수사 100, 자연 = 국영수과 100)
- 종합: 전교과 100 등급
- 정시: 등급 통계 생략 (지원/합격 카운트만)
"""
import pandas as pd, json, numpy as np

SRC = r'C:\클라우드파일\다운 받은 파일\수시사례데이터통합_260423.xlsx'
OUT = r'C:\Users\user\search\cases.json'

# 원본 파일은 2줄 헤더 (row0=대분류, row1=가중치) — skiprows로 가중치 행 건너뜀
print('로딩...')
df = pd.read_excel(SRC, sheet_name='DB', header=0, skiprows=[1])
print(f'전체 행: {len(df):,}, 컬럼: {len(df.columns)}')

# 컬럼 위치 (원본 파일 기준)
COL = {
    'series': 4, 'univ': 5, 'dept': 6, 'time': 7, 'type': 8, 'subtype': 9,
    'n_admit': 11, 'step1': 12, 'final': 13, 'rank': 15,
    'grade_all':   23,  # 전교과 (100% 가중)
    'g_kuksuyoungsa': 32,  # 국영수사 (100%)
    'g_kuksuyounggwa':39,  # 국영수과 (100%)
}

# 시리즈 추출 (안전한 iloc 접근)
def col(i): return df.iloc[:, i]

univ_s   = col(COL['univ']).astype(str).str.strip()
dept_s   = col(COL['dept']).astype(str).str.strip()
series_s = col(COL['series']).astype(str).str.strip()
time_s   = col(COL['time']).astype(str).str.strip()
type_s   = col(COL['type']).astype(str).str.strip()
final_s  = col(COL['final']).astype(str).str.strip()
n_admit  = pd.to_numeric(col(COL['n_admit']), errors='coerce')
rank_v   = pd.to_numeric(col(COL['rank']), errors='coerce')
g_all    = pd.to_numeric(col(COL['grade_all']),       errors='coerce')
g_kss    = pd.to_numeric(col(COL['g_kuksuyoungsa']),  errors='coerce')
g_ksg    = pd.to_numeric(col(COL['g_kuksuyounggwa']), errors='coerce')

# 카테고리화
def cat_of(t, tm):
    t = str(t) if t is not None else ''
    tm = str(tm) if tm is not None else ''
    if '정시' in tm or '수능위주' in t: return '정시'
    if '교과' in t: return '교과'
    if '종합' in t: return '종합'
    if '논술' in t: return '논술'
    return None

cats = [cat_of(t, tm) for t, tm in zip(type_s, time_s)]

# 유효 행만 (univ/dept 있고, 카테고리 매칭됨)
work = pd.DataFrame({
    'univ': univ_s, 'dept': dept_s, 'series': series_s, 'cat': cats,
    'final': final_s, 'n_admit': n_admit, 'rank': rank_v,
    'g_all': g_all, 'g_kss': g_kss, 'g_ksg': g_ksg,
})
work = work[work['univ'].notna() & work['dept'].notna() & (work['univ']!='nan') & (work['dept']!='nan')]
work = work[work['cat'].isin(['교과','종합','정시'])]
work['pass'] = work['final'].isin(['합','추합'])
work['chuhab'] = work['final'] == '추합'  # 추합 자체 카운트용
print(f'유효 행: {len(work):,}')

# 계열별 비교등급 결정
# - 교과 + 인문/예체능 → 국영수사 (없으면 전교과)
# - 교과 + 자연 → 국영수과 (없으면 전교과)
# - 종합 → 전교과
# - 정시 → 없음
def compare_grade(row):
    if row['cat']=='종합':
        return row['g_all'] if pd.notna(row['g_all']) and row['g_all']>0 else None
    if row['cat']=='교과':
        s = row['series']
        if s in ('자연',):
            v = row['g_ksg']
            if pd.notna(v) and v>0: return v
        else:  # 인문/예체능/공통 등
            v = row['g_kss']
            if pd.notna(v) and v>0: return v
        # 폴백
        v = row['g_all']
        if pd.notna(v) and v>0: return v
    return None

work['cmp'] = work.apply(compare_grade, axis=1)

print('집계 시작...')
result = {}
for (univ, dept, cat), g in work.groupby(['univ','dept','cat']):
    n_total = len(g)
    passed = g[g['pass']]
    n_pass = len(passed)
    if n_pass < 1: continue
    key = f'{univ}|{dept}'
    if key not in result: result[key] = {}
    entry = {'지원': int(n_total), '합격': int(n_pass),
             '합격률': round(n_pass / n_total * 100, 1)}
    if cat in ('교과','종합'):
        gp = passed['cmp'].dropna()
        if len(gp) >= 1:
            entry['평균'] = round(float(gp.mean()), 2)
            entry['최소'] = round(float(gp.min()),  2)
            entry['최대'] = round(float(gp.max()),  2)
            entry['p25'] = round(float(gp.quantile(0.25)), 2)
            entry['p50'] = round(float(gp.quantile(0.50)), 2)
            entry['p75'] = round(float(gp.quantile(0.75)), 2)
            entry['n_등급'] = int(len(gp))
            if cat == '교과':
                entry['기준'] = '계열별 환산(국영수사/국영수과 100%)'
            else:
                entry['기준'] = '전교과 100%'
    # 추합 — 최종단계 == '추합'
    n_chuhab = int(passed['chuhab'].sum())
    if n_chuhab: entry['추합'] = n_chuhab
    entry['최초합격'] = int(n_pass - n_chuhab)  # 추합 제외 최초합격수
    result[key][cat] = entry

print(f'학과 수: {len(result):,}')

with open(OUT, 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, separators=(',',':'))

import os
print(f'저장 완료: {os.path.getsize(OUT)/1024:.1f} KB')
