import pandas as pd, json, sys
from datetime import datetime
src, mes, pagado, cliente = sys.argv[1], sys.argv[2], float(sys.argv[3]), sys.argv[4]
OUTJSON = sys.argv[5] if len(sys.argv) > 5 else 'metrics.json'   # ej: export.xlsx 2026-10 10995 "Natisa"
d=pd.read_excel(src); d['t']=pd.to_datetime(d['Created at']); d['m']=d.t.dt.to_period('M')
d['Operation']=d.Operation.str.strip(); d['cli']=d.Phone.astype(str)
d['nombre']=(d.Name.fillna('')+' '+d['Last name'].fillna('')).str.replace(r'\s+',' ',regex=True).str.strip()
M=pd.Period(mes); P=M-1
cut=d[d.m<=M]
v=cut[cut.Operation=='Visit logged'].sort_values('t').copy(); v['k']=v.groupby('cli').cumcount()
def vm(m): return v[v.m==m]
def ing(m): return float(vm(m)['Purchase sum'].sum())
def tick(m): x=vm(m); return float(x['Purchase sum'].mean()) if len(x) else 0
first=cut.groupby('cli').t.min().dt.to_period('M')
def nuevos(m): return int((first==m).sum())
import re
cashmask=d.Operation.str.contains('cashback',case=False)&d.Operation.str.contains('redeem|spent|used|withdraw',case=False)
def pts(m): x=cut[(cut.m==m)&(cut.Operation=='Points redeemed')]; return float(x.Amount.sum())
def cash(m): x=cut[(cut.m==m)&cashmask.reindex(cut.index)]; return float(x.Amount.sum())
HAS_PTS=bool((d.Operation=='Points redeemed').any()); HAS_CASH=bool(cashmask.any())
MODE='both' if HAS_PTS and HAS_CASH else 'puntos' if HAS_PTS else 'cashback' if HAS_CASH else 'recompensas'
def cnt(m):
    x=cut[cut.m==m]; return int(x.Operation.isin(['Coupon redeemed','Rewards redeemed','Stamps redeemed']).sum())
def prim(m): return pts(m) if MODE in ('both','puntos') else cash(m) if MODE=='cashback' else cnt(m)
def canjes(m):
    x=cut[cut.m==m]; return int(x.Operation.isin(['Coupon redeemed','Points redeemed','Rewards redeemed','Stamps redeemed']).sum())
prev3=[M-i for i in (1,2,3)]
avg=lambda f: sum(f(m) for m in prev3)/3
o=vm(M)
g=o.groupby('cli').agg(n=('t','size'),gasto=('Purchase sum','sum'),nom=('nombre','first'))
bg=g.sort_values('gasto',ascending=False).iloc[0]; bv=g.sort_values(['n','gasto'],ascending=False).iloc[0]
rep_acum=float(v[v.k>0]['Purchase sum'].sum()); tot_acum=float(v['Purchase sum'].sum())
rep_mes={str(m):float(vm(m)[vm(m).k>0]['Purchase sum'].sum()) for m in [M-i for i in range(5,-1,-1)]}
med=float(o['Purchase sum'].median())
big=o[o['Purchase sum']>5*med].sort_values('Purchase sum',ascending=False)
oo=o.sort_values(['cli','t']).copy(); oo['dt']=oo.groupby('cli').t.diff().dt.total_seconds()/60
dob=oo[oo.dt<=10].sort_values('dt')
def row(r): return dict(fecha=r.t.strftime('%d %b · %H:%M'),monto=float(r['Purchase sum']),nombre=r.nombre,dt=float(r.get('dt',0) or 0))
MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre']
out=dict(cliente=cliente,mes=str(M),mes_label=MESES[M.month-1].upper()+' '+str(M.year),mes_prev=MESES[P.month-1],pagado=pagado,
 ing=ing(M),ing_prev=ing(P),ing_avg=avg(ing),acum=tot_acum,rep_acum=rep_acum,
 roi_rep=rep_acum/pagado,roi_tot=tot_acum/pagado,rep_mes=rep_mes,
 nuevos=nuevos(M),nuevos_prev=nuevos(P),nuevos_avg=avg(nuevos),
 clientes=int(v.cli.nunique()),activos=int(o.cli.nunique()),
 ticket=tick(M),ticket_prev=tick(P),ticket_avg=avg(tick),
 canje_mode=MODE,canje=prim(M),canje_prev=prim(P),canje_avg=avg(prim),canje_pts=pts(M),canje_cash=cash(M),canje_n=cnt(M),canjes=canjes(M),canjes_prev=canjes(P),canjes_avg=avg(canjes),
 mejor_gasto=dict(nombre=bg.nom,gasto=float(bg.gasto),n=int(bg.n)),
 mejor_visitas=dict(nombre=bv.nom,gasto=float(bv.gasto),n=int(bv.n)),
 visitas=int(len(o)),mediana=med,n_big=int(len(big)),n_dob=int(len(dob)),
 top_big=row(big.iloc[0]) if len(big) else None,
 top_dob=row(dob.iloc[0]) if len(dob) else None)
warn=[]
if len(cliente)>26: warn.append('nombre de negocio largo')
if max(len(bg.nom),len(bv.nom))>24: warn.append('nombre de mejor cliente largo')
out['warnings']=warn
json.dump(out,open(OUTJSON,'w'),ensure_ascii=False,indent=1,default=float)
print(json.dumps(out,ensure_ascii=False,indent=1,default=float))
