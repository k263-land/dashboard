import sys, json, glob, os
import numpy as np, pandas as pd, openpyxl
A='/mnt/user-data/attachments/3b9d6fcf-efa4-5a74-9b78-592e4d13f75c/'
out=sys.argv[1]; TOPV=int(sys.argv[2]) if len(sys.argv)>2 else 40; OUT_S=sys.argv[3] if len(sys.argv)>3 else None; MAPF=sys.argv[4] if len(sys.argv)>4 else None
files={2019:'1a046da1-__DB_2019____.csv',2020:'97147acd-__DB_2020____.csv',2021:'ce3efde4-__DB_2021____.csv',2022:'31ef453a-__DB_2022____.csv',2023:'d9b230d2-__DB_2023____.csv',2024:'9a2c1f1d-__DB_2024____.csv',2025:'5e4775ab-__DB_2025____.csv',2026:'/절대경로/일보DB_2026년_9월.csv'}
# ---- masters (기초DB) : 프로젝트 build_data.py 규칙과 동일 ----
ws=openpyxl.load_workbook(A+'e7930bd1-__DB.xlsx',read_only=True,data_only=True)['기초']
rows=list(ws.iter_rows(min_row=1,values_only=True))
def cell(r,i): return r[i] if i<len(r) else None
GBC={100:'PL',200:'PS',300:'PI',400:'PM'}
items={}; b2b=set(); corps={}
for r in rows[3:]:
    c=cell(r,8)
    if c is not None:
        try:
            c=int(c); gc=cell(r,4)
            items[c]={'name':str(cell(r,10) or c).strip(),'grp':GBC.get(int(gc)) if isinstance(gc,(int,float)) else None,
                      'md':cell(r,11),'mdb':cell(r,12)}
        except (TypeError,ValueError): pass
    ak,al=cell(r,36),cell(r,37)
    if al=='B2B' and ak is not None and not str(ak).startswith('_'):
        try: b2b.add(int(ak))
        except (TypeError,ValueError): pass
    cc=cell(r,41)
    if cc is not None:
        try: corps[int(str(cc).lstrip("'"))]=str(cell(r,42) or '').strip()
        except (TypeError,ValueError): pass
# 지점 마스터(기초DB 지점코드/지점명/형태)
hdr=rows[2]; ci=[i for i,v in enumerate(hdr) if v=='지점코드'][0]
SMAST={}
for r in rows[3:]:
    c=cell(r,ci)
    if c is not None:
        try: SMAST[int(str(c).lstrip("'"))]=(str(cell(r,ci+1) or '').strip(),str(cell(r,ci+4) or '').strip(),str(cell(r,ci+3) or '').strip())
        except (TypeError,ValueError): pass
print('지점 마스터',len(SMAST))
# 품목 통합: 기초DB 원코드(8열) -> 변경코드(9열)
CHG={}
for r in rows[3:]:
    c,k=cell(r,8),cell(r,9)
    try:
        c=int(c); k=int(k)
        if k and k!=c and k in items: CHG[c]=k
    except (TypeError,ValueError): pass
print('품목 통합 대상',len(CHG),'->',len(set(CHG.values())))
print('items',len(items),'b2b corps',len(b2b),'corps',len(corps))
S_ON={400040}; S_SP={302420,400041}
B2B_IT={152,160,510,114,255,256,257,258}
parts=[]
for y,f in files.items():
    df=pd.read_csv(f if f.startswith('/') else A+f,encoding='cp949',dtype=str)
    df.columns=[c.strip() for c in df.columns]
    for c in ['년월','지점변경','판매구분','지점코드','품목코드','거래선코드','수량','금액']:
        df[c]=pd.to_numeric(df[c].astype(str).str.replace(',','').str.strip(),errors='coerce')
    n0=len(df); df=df.dropna(subset=['년월','지점코드','품목코드','거래선코드','판매구분'])
    print(y,'rows',n0,'dropped',n0-len(df))
    st=df['지점변경'].fillna(df['지점코드']).astype('int64')
    st0=st.copy()
    st=st.where(~((st>=300010)&(st<=300019)|(st==300030)),300010)
    corp=df['거래선코드'].astype('int64'); it=df['품목코드'].astype('int64')
    is_b2b=corp.isin(b2b)|((corp==1000001)&it.isin(B2B_IT))|(it==268)
    ch=pd.Series('직영',index=df.index,dtype=object); ch[st.isin(S_ON)]='온라인'; ch[st.isin(S_SP)]='특수'; ch[(ch=='직영')&is_b2b]='B2B'
    itm=it.map(CHG).fillna(it).astype('int64')   # 통합코드
    g=itm.map({c:v['grp'] for c,v in items.items()}); g=g.where(g.isin(['PL','PS','PI','PM']),'기타'); g=g.where(ch!='B2B','B2B')   # 채널 B2B이면 그룹=B2B
    md=itm.map({c:v['md'] for c,v in items.items()}); mdb=itm.map({c:v['mdb'] for c,v in items.items()})
    md=md.where(~is_b2b,mdb).fillna('미지정').astype(str).str.strip().replace({'nan':'미지정','':'미지정'})
    excl=None
    d=pd.DataFrame({'y':(df['년월']//100).astype(int),'m':(df['년월']%100).astype(int),'ch':ch,'grp':g,'it':it,'md':md,'st':st,'vd':corp,'sd':df['판매구분'].astype(int),'amt':df['금액'].fillna(0),'qty':df['수량'].fillna(0),'E':df['매입구분'].astype(str).str.strip().eq('E')})
    d=d[d.y==y]; parts.append(d)
d=pd.concat(parts,ignore_index=True)
# 수량제외 품목(P열=X) 수량 0 처리 -> 기초DB P열
qx={}
for r in rows[3:]:
    c=cell(r,8)
    if c is not None:
        try: qx[int(c)]=str(cell(r,15) or '').strip().upper()=='X'
        except (TypeError,ValueError): pass
d['qty']=d['qty'].where(~d['it'].map(qx).fillna(False).astype(bool),0)
d['it']=d['it'].map(CHG).fillna(d['it']).astype('int64')   # 수량제외 반영 후 통합코드로 합산
# 거래선: 상위 TOPV 외 '기타 거래선'
vs=d[(d.sd==2)&(d.y>=2022)].groupby('vd')['amt'].sum().abs().sort_values(ascending=False)
top=set(vs.head(TOPV).index.tolist()); d['vd']=d['vd'].where(d['vd'].isin(top),-1)
vname={v:(corps.get(v) or str(v)) for v in top}; vname[-1]='기타 거래선'
d['a1']=np.where(d.sd==1,d.amt,0); d['a2']=np.where(d.sd==2,d.amt,0); d['b1']=np.where(d.E,d.a1,0); d['b2']=np.where(d.E,d.a2,0); d['q1']=np.where(d.sd==1,d.qty,0); d['q2']=np.where(d.sd==2,d.qty,0)
g=d.groupby(['y','m','ch','grp','it','md','vd'],as_index=False)[['a1','a2','b1','b2','q1','q2']].sum()
print('cube rows',len(g))
if OUT_S:
    ds=d[d.y>=2022].copy(); ds['a1']=np.where(ds.sd==1,ds.amt,0); ds['a2']=np.where(ds.sd==2,ds.amt,0); ds['b1']=np.where(ds.E,ds.a1,0); ds['b2']=np.where(ds.E,ds.a2,0)
    gs=ds.groupby(['y','m','st','grp'],as_index=False)[['a1','a2','b1','b2']].sum()
    dm0=d[(d.y>=2019)&(d.sd==1)&(d.ch.isin(['직영','B2B']))]
    codes=sorted(set(gs['st'].unique().tolist())|set(dm0['st'].unique().tolist())); six={c:i for i,c in enumerate(codes)}
    GRL=['PL','PS','PI','PM','B2B','기타']
    nm=lambda c:(SMAST.get(c,('','',''))[0] or {400040:'온라인(채널)',400041:'특수(채널)',302420:'특수(채널)'}.get(c,''))
    tp=lambda c:SMAST.get(c,('','',''))[1]
    # 점포 구분: 0직영(창고형·L500 포함) 1현대 2인샵 3모바일(M) 4특수 -1없음. 폐점은 비고열(IN=인샵, 직영/L500=직영)로 폐점 전 구분 적용
    # 지점 매칭표(미매칭 지점의 폐점 전 형태): 직영→0(로드샵) IN→2(인샵) M→3(모바일) 제외→-2(점포수 제외)
    XM={}
    if MAPF:
        _m=pd.read_excel(MAPF,sheet_name='미매칭'); _m=_m[_m['지점코드'].notna()]
        _tm={'직영':0,'현대':1,'IN':2,'인샵':2,'M':3,'모바일':3,'특수':4,'제외':-2}
        XM={int(c):_tm[str(t).strip()] for c,t in zip(_m['지점코드'],_m['기초DB 형태']) if str(t).strip() in _tm}
        print('지점 매칭표',len(XM),{k:list(XM.values()).count(k) for k in set(XM.values())})
    def sty(c):
        if c in XM: return XM[c]
        n,f,bk=SMAST.get(c,('','',''))
        if f in('직영','창고형','L500'): return 0
        if f=='현대': return 1
        if f=='M': return 3
        if f=='특수': return 4
        if f=='폐점': return 2 if bk=='IN' else (0 if bk in('직영','L500') else -1)
        return -1
    dm=dm0.groupby(['y','m','st'],as_index=False)['amt'].sum()
    SP={'sty':[sty(c) for c in codes],'sm':{'y':[int(x) for x in dm.y],'m':[int(x) for x in dm.m],'st':[six[c] for c in dm.st if c in six],'a':[int(round(x/1000)) for x in dm.amt]},'sc':codes,'sn':[nm(c) for c in codes],'st':[tp(c) for c in codes],'grp':GRL,
        'rows':{'y':[int(x) for x in gs.y],'m':[int(x) for x in gs.m],'st':[six[c] for c in gs.st],'grp':[GRL.index(x) for x in gs.grp],
                'a1':[int(round(x/1000)) for x in gs.a1],'a2':[int(round(x/1000)) for x in gs.a2],'b1':[int(round(x/1000)) for x in gs.b1],'b2':[int(round(x/1000)) for x in gs.b2]}}
    assert len(SP['sm']['st'])==len(SP['sm']['y'])
    print('점포구분',{k:SP['sty'].count(k) for k in (0,1,2,3,4,-1,-2)})
    json.dump(SP,open(OUT_S,'w',encoding='utf-8'),ensure_ascii=False,allow_nan=False,separators=(',',':'))
    print('지점 큐브 rows',len(gs),'stores',len(codes),'이름없음',sum(1 for c in codes if not nm(c)),'MB',round(os.path.getsize(OUT_S)/1e6,2))
    t=gs[(gs.y==2026)&(gs.m<=7)]; print('2026 1~7 지점합 매출(백만)',round(t.a2.sum()/1e6))
dims={'ch':['직영','온라인','B2B','특수'],'grp':['PL','PS','PI','PM','B2B','기타']}
dims['item']=sorted({items.get(int(c),{'name':str(int(c))})['name'] for c in g['it'].unique()})
dims['md']=sorted(g['md'].unique().tolist())
vorder=[v for v in vs.head(TOPV).index.tolist()]+[-1]; dims['vend']=[vname[v] for v in vorder]
ix={k:{v:i for i,v in enumerate(dims[k])} for k in dims}
vix={v:i for i,v in enumerate(vorder)}
iname=lambda c: items.get(int(c),{'name':str(int(c))})['name']
R={k:[] for k in ['y','m','ch','grp','item','md','vend','a1','a2','b1','b2','q1','q2']}
for r in g.itertuples(index=False):
    R['y'].append(int(r.y));R['m'].append(int(r.m));R['ch'].append(ix['ch'][r.ch]);R['grp'].append(ix['grp'][r.grp])
    R['item'].append(ix['item'][iname(r.it)]);R['md'].append(ix['md'][r.md]);R['vend'].append(vix[r.vd])
    R['a1'].append(int(round(r.a1/1000)));R['a2'].append(int(round(r.a2/1000)));R['b1'].append(int(round(r.b1/1000)));R['b2'].append(int(round(r.b2/1000)));R['q1'].append(int(round(r.q1)));R['q2'].append(int(round(r.q2)))
pay={'dims':dims,'rows':R,'years':sorted(g['y'].unique().tolist()),'lastY':int(g.y.max()),'lastM':int(g[g.y==g.y.max()].m.max())}
json.dump(pay,open(out,'w',encoding='utf-8'),ensure_ascii=False,allow_nan=False,separators=(',',':'))
print('size MB',round(os.path.getsize(out)/1e6,2),'last',pay['lastY'],pay['lastM'],'vendors',len(dims['vend']),'items',len(dims['item']))
t=g.groupby('y')[['a1','a2']].sum()/1e6; print(t.round(0))
t=g[(g.m<=7)].groupby('y')[['a2']].sum()/1e6; print('1~7월 매출', t.round(0).to_dict())
print(g[g.y==2026].groupby('ch')[['a2']].sum().div(1e6).round(0).to_dict())
print(g[g.y==2026].groupby('grp')[['a2']].sum().div(1e6).round(0).to_dict())

print('구독(E) 매출구분 by y(백만):',(g.groupby('y')['b2'].sum()/1e6).round(0).to_dict(),' 판매구분:',(g.groupby('y')['b1'].sum()/1e6).round(0).to_dict())
