import os, sys, shutil, tempfile, formulas, pandas as pd, openpyxl
f=os.path.join(tempfile.mkdtemp(),'vv.xlsx'); shutil.copy(sys.argv[1],f)
sol=formulas.ExcelModel().loads(f).finish().calculate()
vals={}
for k,v in sol.items():
    if ':' in k.split('!')[-1]: continue
    try: vals[k.split(']')[1].upper()]=v.value[0,0]
    except Exception: pass
g=lambda sh,c: vals.get((sh+"'!"+c).upper())
wb=openpyxl.load_workbook(f)
fe=[k for k,v in vals.items() if isinstance(v,str) and v.startswith('#') and v!='#']
print('formula errors:',fe[:5])
df=pd.read_excel(f,sheet_name='XTab DS_1')
df['kbn']=df.groupby('品目コード')['保管場所'].transform('nunique').map(lambda n:'重複' if n>=2 else '単独')
bad=0
for (loc,n),d in df.groupby(['保管場所','保管場所名'],sort=False):
    sh='倉庫別_'+n; ws=wb[sh]; nrow=len(d)
    for i,(_,row) in enumerate(d.iterrows(),2):
        if (g(sh,f'B{i}'),g(sh,f'F{i}'),g(sh,f'G{i}'),g(sh,f'H{i}'))!=(row.品目コード,row.kbn,row.在庫数量,row.在庫金額): bad+=1
    t=nrow+2
    print(sh,nrow,'rows | total',g(sh,f'G{t}'),g(sh,f'H{t}'),'exp',d.在庫数量.sum(),d.在庫金額.sum(),
          '| dup',g(sh,f'H{t+1}'),'exp',d[d.kbn=='重複'].在庫金額.sum(),'| single',g(sh,f'H{t+2}'),'exp',d[d.kbn=='単独'].在庫金額.sum())
sh='商品別_合算'; ws=wb[sh]; nc=ws.max_column; Q=ws.cell(1,nc-1).column_letter; A=ws.cell(1,nc).column_letter
agg=df.groupby('品目コード',sort=False).agg(q=('在庫数量','sum'),a=('在庫金額','sum'),n=('保管場所','nunique'),k=('kbn','first'))
for i,(code,row) in enumerate(agg.iterrows(),2):
    if (g(sh,f'B{i}'),g(sh,f'D{i}'),g(sh,f'E{i}'),g(sh,f'{Q}{i}'),g(sh,f'{A}{i}'))!=(code,row.n,row.k,row.q,row.a): bad+=1; print('bad',code)
t=len(agg)+2
print(sh,len(agg),'rows | total',g(sh,f'{Q}{t}'),g(sh,f'{A}{t}'),'exp',agg.q.sum(),agg.a.sum(),'| dup',g(sh,f'{A}{t+1}'),'| single',g(sh,f'{A}{t+2}'))
print('row mismatches:',bad)
sys.exit(1 if bad or fe else 0)
