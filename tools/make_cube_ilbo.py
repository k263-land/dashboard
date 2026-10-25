import sys, json, glob, os
import numpy as np, pandas as pd, openpyxl
A='/mnt/user-data/attachments/3b9d6fcf-efa4-5a74-9b78-592e4d13f75c/'
out=sys.argv[1]; TOPV=int(sys.argv[2]) if len(sys.argv)>2 else 40
files={2019:'1a046da1-__DB_2019____.csv',2020:'97147acd-__DB_2020____.csv',2021:'ce3efde4-__DB_2021____.csv',2022:'31ef453a-__DB_2022____.csv',2023:'d9b230d2-__DB_2023____.csv',2024:'9a2c1f1d-__DB_2024____.csv',2025:'5e4775ab-__DB_2025____.csv',2026:'27758c4b-__DB_2026__7_.csv'}
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
    df=pd.read_csv(A+f,encoding='cp949',dtype=str)
    df.columns=[c.strip() for c in df.columns]
    for c in ['년월','지점변경','판매구분','지점코드','품목코드','거래선코드','수량','금액']:
        df[c]=pd.to_numeric(df[c].astype(str).str.replace(',','').str.strip(),errors='coerce')
    n0=len(df); df=df.dropna(subset=['년월','지점코드','품목코드','거래선코드','판매구분'])
    print(y,'rows',n0,'dropped',n0-len(df))
    st=df['지점변경'].fillna(df['지점코드']).astype('int64')
    st=st.where(~((st>=300010)&(st<=300019)|(st==300030)),300010)
    corp=df['거래선코드'].astype('int64'); it=df['품목코드'].astype('int64')
    is_b2b=corp.isin(b2b)|((corp==1000001)&it.isin(B2B_IT))|(it==268)
    ch=pd.Series('직영',index=df.index,dtype=object); ch[st.isin(S_ON)]='온라인'; ch[st.isin(S_SP)]='특수'; ch[(ch=='직영')&is_b2b]='B2B'
    itm=it.map(CHG).fillna(it).astype('int64')   # 통합코드
    g=itm.map({c:v['grp'] for c,v in items.items()}); g=g.where(g.isin(['PL','PS','PI','PM']),'기타'); g=g.where(ch!='B2B','B2B')   # 채널 B2B이면 그룹=B2B
    md=itm.map({c:v['md'] for c,v in items.items()}); mdb=itm.map({c:v['mdb'] for c,v in items.items()})
    md=md.where(~is_b2b,mdb).fillna('미지정').astype(str).str.strip().replace({'nan':'미지정','':'미지정'})
    excl=None
    d=pd.DataFrame({'y':(df['년월']//100).astype(int),'m':(df['년월']%100).astype(int),'ch':ch,'grp':g,'it':it,'md':md,'vd':corp,'sd':df['판매구분'].astype(int),'amt':df['금액'].fillna(0),'qty':df['수량'].fillna(0),'E':df['매입구분'].astype(str).str.strip().eq('E')})
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
