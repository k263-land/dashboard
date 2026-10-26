import sys, json
import pandas as pd, numpy as np
srcs, out = sys.argv[1].split(','), sys.argv[2]   # 결산DB 파일 여러 개(쉼표 구분) 통합 가능
_p=[]
for src in srcs:
    _d = pd.read_excel(src, dtype=object)
    _d.columns = [str(c).replace('\n','') for c in _d.columns]
    _p.append(_d)
df = pd.concat(_p, ignore_index=True)
df = df[df['직영구분'].notna() & (df['년월'].astype(str)!='년월')].copy()   # 빈 행·중복 헤더 행 제거
df['년'] = pd.to_numeric(df['년월'], errors='coerce')//100; df['월'] = pd.to_numeric(df['년월'], errors='coerce')%100
df = df[df['년'].notna()].copy()
print('결산 행',len(df))
num = ['수량','매출','환산매출','매출이익','영업이익','경상이익','재고계','재고수량계']
for c in num: df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)
df['y'] = df['년'].astype(int); df['m'] = df['월'].astype(int)
chmap = {'직영':'직영','직영외':'온라인','직영외특수':'온라인','직영특수':'B2B','특판':'특수'}
df['ch'] = df['직영구분'].map(chmap)
assert df['ch'].notna().all()
# 거래선명 끝이 'B2B'이면 B2B 채널 (예외: LG유플러스-모바일_B2B 14597771 = 직영 유지)
_vn=df['매입처'].astype(str).str.strip().str.upper(); _vc=pd.to_numeric(df['매입처코드'],errors='coerce')
_nb=_vn.str.endswith('B2B') & (_vc!=14597771) & (df['ch']!='B2B')
print('거래선명 B2B 재분류 행',int(_nb.sum()),'금액(백만)',round(pd.to_numeric(df.loc[_nb,'매출'],errors='coerce').sum()))
df.loc[_nb,'ch']='B2B'
core = ['PL','PS','PI','PM']
df['grp'] = df['신그룹'].where(df['신그룹'].isin(core), '기타')
df['item'] = df['품목.1'].astype(str).str.strip()
df['md'] = df['PM'].fillna('미지정').astype(str).str.strip()
# ---- 품목 통합(변경코드) + 수량 제외(제외=X): 기초DB ----
_b=pd.read_excel('/mnt/user-data/attachments/3b9d6fcf-efa4-5a74-9b78-592e4d13f75c/e7930bd1-__DB.xlsx',header=None,skiprows=3)
_b=_b.iloc[:,[4,8,9,10,11,12,15]]; _b.columns=['gc','org','chg','nm','md','mdb','ex']
_b['org']=pd.to_numeric(_b['org'],errors='coerce'); _b['chg']=pd.to_numeric(_b['chg'],errors='coerce'); _b=_b.dropna(subset=['org']); _b['org']=_b['org'].astype(int)
_orgs=set(_b['org']); CHG={int(o):int(k) for o,k in zip(_b['org'],_b['chg']) if pd.notna(k) and int(k)!=0 and int(k)!=o and int(k) in _orgs}
EXCL=set(_b.loc[_b['ex'].astype(str).str.strip().str.upper()=='X','org'].astype(int))
df['code']=pd.to_numeric(df['품목'],errors='coerce')
_map=df['code'].map(CHG); _m=_map.notna()
_base=df[~_m]
def _mode(col): return _base.groupby('code')[col].agg(lambda x: x.mode().iat[0] if len(x.mode()) else None)
_nm,_gr,_md=_mode('item'),_mode('grp'),_mode('md')
_bn=_b.drop_duplicates('org').set_index('org')
for i in df.index[_m]:
    t=int(_map[i]); df.at[i,'item']=_nm.get(t) or str(_bn['nm'].get(t,t)).strip(); df.at[i,'grp']=_gr.get(t) or df.at[i,'grp']; 
    if df.at[i,'ch']!='B2B': df.at[i,'md']=_md.get(t) or str(_bn['md'].get(t,df.at[i,'md'])).strip()   # B2B 행은 원본 MD 유지
print('결산 통합 적용 행',int(_m.sum()),'| 수량제외 코드',len(EXCL))

# B2B MD는 한재호·손국현만: 다른 MD면 기초DB B2B 담당(M열), 없으면 손국현
_B2M={'한재호','손국현'}
_bm=_b.drop_duplicates('org').set_index('org')['mdb'].astype(str).str.strip()
_bmod=lambda c:(lambda v:v if v in _B2M else '손국현')(_bm.get(int(CHG.get(int(c),c)),'') if pd.notna(c) else '')
_fix=(df['ch']=='B2B')&(~df['md'].isin(_B2M))
df.loc[_fix,'md']=df.loc[_fix,'code'].map(_bmod)
print('B2B MD 보정 행',int(_fix.sum()),'| B2B MD',sorted(df.loc[df['ch']=='B2B','md'].unique()))
df['grp'] = df['grp'].where(df['ch']!='B2B','B2B')   # 채널이 B2B이면 그룹=B2B (채널 우선)
df.loc[df['code'].isin(EXCL),['수량','재고수량계']]=0   # 수량제외(X): 수량·재고수량 0 (금액은 유지)
# 결산기준 매출: 2025년 이후 환산매출(없으면 매출로 대체), 이전은 매출
df['s_h'] = np.where((df['y']>=2025) & (df['환산매출']!=0), df['환산매출'], df['매출'])   # 환산: 구독=고객판매가 (2025~), 이전은 매출
df['s_k'] = df['매출']                                                              # 결산: 구독=수수료 수취분
df['sub_k'] = df['구독']                                  # 결산 기준 구독: 구독 열(수수료 수취분)
df['sub_h'] = (df['s_h']-df['매출']) + df['구독']             # 환산 기준 구독: (환산-결산 차이) + 구독 열
df['vend'] = df['매입처'].fillna('미지정').astype(str).str.strip()
topv = df.groupby('vend')['s_h'].sum().abs().sort_values(ascending=False).head(40).index.tolist()
df['vend'] = df['vend'].where(df['vend'].isin(topv), '기타 거래선')
g = df.groupby(['y','m','ch','grp','item','md','vend'], as_index=False).agg(
    s_h=('s_h','sum'), s_k=('s_k','sum'), gp=('매출이익','sum'), op=('영업이익','sum'), rc=('경상이익','sum'), sb_h=('sub_h','sum'), sb_k=('sub_k','sum'),
    inv=('재고계','sum'), invq=('재고수량계','sum'), q=('수량','sum'))
dims = {k: sorted(g[k].unique().tolist()) for k in ['ch','grp','item','md','vend']}
dims['vend'] = topv + ['기타 거래선']
dims['grp'] = [x for x in ['PL','PS','PI','PM','B2B','기타'] if x in dims['grp']]
dims['ch'] = [x for x in ['직영','온라인','B2B','특수'] if x in dims['ch']]
idx = {k:{v:i for i,v in enumerate(dims[k])} for k in dims}
cols = ['y','m','ch','grp','item','md','vend','s_h','s_k','gp','op','rc','sb_h','sb_k','inv','invq','q']
rows = {c: [] for c in cols}
for r in g.itertuples(index=False):
    rows['y'].append(int(r.y)); rows['m'].append(int(r.m))
    for k in ['ch','grp','item','md','vend']: rows[k].append(idx[k][getattr(r,k)])
    for k in ['s_h','s_k','gp','op','rc','sb_h','sb_k','inv']: rows[k].append(int(round(getattr(r,k)*1000)))  # 백만원 -> 천원
    rows['invq'].append(int(round(r.invq))); rows['q'].append(int(round(r.q)))
payload = {'dims':dims,'rows':rows,'n':len(g),'unit':'천원(재고수량·수량은 대)',
           'years':sorted(g['y'].unique().tolist()),'lastY':int(g['y'].max()),
           'lastM':int(g[g['y']==g['y'].max()]['m'].max())}
json.dump(payload, open(out,'w'), ensure_ascii=False, allow_nan=False, separators=(',',':'))
print('rows', len(g), 'dims', {k:len(v) for k,v in dims.items()}, 'last', payload['lastY'], payload['lastM'])
# reference totals (백만원) for verification
chk = g.groupby('y')[['s_h','s_k','gp','op','rc','sb_h','sb_k']].sum().round(0)
print(chk)
mm = g[(g.y==payload['lastY'])&(g.m==payload['lastM'])][['inv','invq']].sum()
print('last-month stock', mm.round(0).to_dict())
