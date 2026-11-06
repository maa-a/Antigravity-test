"""払出後の在庫一覧Excelを作る（全品目 × 倉庫別に、払出前・払出数・払出後を並べる）。

使い方:
  python3 after_stock.py <SAP在庫.xlsx> <払出リスト.tsv> <出力.xlsx> [既定の払出元保管場所コード=CX01]

払出リスト.tsv は1行1件、ヘッダーなし、タブ区切り:
  品目コード<TAB>数量                  … 既定の払出元から払い出す
  品目コード<TAB>数量<TAB>保管場所コード … 指定した倉庫から払い出す（同じ品目を複数行に分けてよい）
数量のカンマ可。在庫・払出後・合算・判定はすべて数式で、元データ／払出リストのシートを参照する。
"""
import sys
from collections import OrderedDict
import openpyxl
from openpyxl.utils import get_column_letter as CL
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.properties import CalcProperties

sap_path, req_path, out = sys.argv[1], sys.argv[2], sys.argv[3]
default_loc = sys.argv[4] if len(sys.argv) > 4 else 'CX01'

F = 'Calibri'
HFILL = PatternFill('solid', fgColor='305496'); GFILL = PatternFill('solid', fgColor='8EA9DB')
TFILL = PatternFill('solid', fgColor='D9E1F2'); AFILL = PatternFill('solid', fgColor='E2EFDA')
THIN = Side(style='thin', color='BFBFBF'); BD = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NUM = '#,##0;[Red]-#,##0'

src = openpyxl.load_workbook(sap_path)['XTab DS_1']
rows = [[c.value for c in r] for r in src.iter_rows(min_row=2, max_row=src.max_row) if r[1].value]
items = OrderedDict(); whs = OrderedDict()
for jan, code, txt, l, ln, q, a in rows:
    items.setdefault(code, (jan, txt)); whs.setdefault(l, ln)

req = []  # (code, loc, qty)
for line in open(req_path, encoding='utf-8'):
    p = [x.strip() for x in line.rstrip('\n').split('\t')]
    if len(p) < 2 or not p[0]: continue
    loc = p[2] if len(p) > 2 and p[2] else default_loc
    req.append((p[0], loc, int(p[1].replace(',', ''))))
bad_loc = sorted({l for _, l, _ in req if l not in whs})
if bad_loc:
    sys.exit(f'元データにない保管場所: {bad_loc}')
missing = sorted({c for c, _, _ in req if c not in items})

wb = openpyxl.Workbook()
ws = wb.active; ws.title = '払出後在庫一覧'
sd = wb.create_sheet('XTab DS_1'); rq = wb.create_sheet('払出リスト')


def head(sh, row, col, text, fill=HFILL, width=None):
    c = sh.cell(row, col, text); c.font = Font(name=F, bold=True, color='FFFFFF'); c.fill = fill; c.border = BD
    c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    if width: sh.column_dimensions[c.column_letter].width = width
    return c


# 元データ（そのまま）
for i, w in enumerate([18, 24, 52, 10, 12, 12, 14], 1):
    head(sd, 1, i, src.cell(1, i).value, width=w)
for r, row in enumerate(rows, 2):
    for c, v in enumerate(row, 1):
        cell = sd.cell(r, c, v); cell.font = Font(name=F); cell.number_format = NUM if c >= 6 else '@'
last = len(rows) + 1
sd.freeze_panes = 'A2'
S = lambda col: f"'XTab DS_1'!${col}$2:${col}${last}"

# 払出リスト（入力値は青字）
for i, (h, w) in enumerate([('品目コード', 24), ('品目名（SAP）', 60), ('払出元 保管場所', 14), ('払出元 保管場所名', 14), ('払出数', 12)], 1):
    head(rq, 1, i, h, width=w)
for r, (code, loc, q) in enumerate(req, 2):
    rq.cell(r, 1, code).number_format = '@'
    rq.cell(r, 2, items[code][1] if code in items else '（元データに該当なし）')
    rq.cell(r, 3, loc).font = Font(name=F, color='0000FF')
    rq.cell(r, 4, whs[loc])
    c = rq.cell(r, 5, q); c.number_format = NUM; c.font = Font(name=F, color='0000FF')
    if loc != default_loc:
        for i in range(1, 6): rq.cell(r, i).fill = PatternFill('solid', fgColor='FCE4D6')
rlast = len(req) + 1
rq.cell(rlast + 1, 2, '合計').font = Font(name=F, bold=True)
rq.cell(rlast + 1, 5, f'=SUM(E2:E{rlast})').number_format = NUM
rq.freeze_panes = 'A2'
RQ = lambda col: f"払出リスト!${col}$2:${col}${rlast}"

# 一覧
locs = list(whs.items())
ws['A1'] = '払出後在庫一覧'; ws['A1'].font = Font(name=F, bold=True, size=14)
ws['A2'] = f'元データ：{sap_path.rsplit("/", 1)[-1]}　払出リスト：{len({c for c, _, _ in req})}品目・{len(req)}行・合計 {sum(q for _, _, q in req):,}'
other = [(c, whs[l], q) for c, l, q in req if l != default_loc]
ws['A3'] = (f'払出元は{whs[default_loc]}（{default_loc}）。' +
            ('他倉庫から払い出す分：' + '、'.join(f'{items.get(c, ("", c))[1].split("　")[-1]} {q:,}（{n}）' for c, n, q in other) + '。' if other else '') +
            '払出後の残数が0は問題なし、マイナスのみ「不足」。')
for a in ('A2', 'A3'): ws[a].font = Font(name=F, color='595959')

H1, H2 = 5, 6
for i, (h, w) in enumerate([('JANコード(EAN/UPC)', 16), ('品目コード', 22), ('品目テキスト', 48)], 1):
    head(ws, H1, i, h, width=w); head(ws, H2, i, None); ws.merge_cells(start_row=H1, start_column=i, end_row=H2, end_column=i)
col = 4; blocks = []
for l, n in locs + [(None, '合算')]:
    head(ws, H1, col, n if l else '合算（全倉庫）', fill=GFILL if l else HFILL)
    ws.merge_cells(start_row=H1, start_column=col, end_row=H1, end_column=col + 2)
    for j, h in enumerate(['払出前', '払出数', '払出後']):
        head(ws, H2, col + j, h, width=11)
    blocks.append((l, col)); col += 3
jcol = col
head(ws, H1, jcol, '判定', width=8); head(ws, H2, jcol, None); ws.merge_cells(start_row=H1, start_column=jcol, end_row=H2, end_column=jcol)
ncol = jcol

r = H2 + 1
for code, (jan, txt) in items.items():
    for c, v in enumerate([jan, code, txt], 1):
        ws.cell(r, c, v).number_format = '@'
    after_cells = []
    for l, c0 in blocks:
        L = lambda k: ws.cell(r, c0 + k).column_letter
        if l:
            ws.cell(r, c0, f'=SUMIFS({S("F")},{S("B")},$B{r},{S("D")},"{l}")')
            ws.cell(r, c0 + 1, f'=SUMIFS({RQ("E")},{RQ("A")},$B{r},{RQ("C")},"{l}")')
            ws.cell(r, c0 + 2, f'={L(0)}{r}-{L(1)}{r}')
            after_cells.append(f'{L(2)}{r}')
        else:
            sub = [bc for bl, bc in blocks if bl]
            for k in range(3):
                ws.cell(r, c0 + k, '=' + '+'.join(f'{ws.cell(r, bc + k).column_letter}{r}' for bc in sub))
    tot = blocks[-1][1]
    TQ = ws.cell(r, tot + 1).column_letter
    ws.cell(r, jcol, f'=IF({TQ}{r}=0,"-",IF(MIN({",".join(after_cells)})<0,"不足","OK"))')
    for c in range(1, ncol + 1):
        cell = ws.cell(r, c); cell.border = BD; cell.font = Font(name=F)
        if 4 <= c < jcol: cell.number_format = NUM
    for _, c0 in blocks:
        ws.cell(r, c0 + 2).fill = AFILL; ws.cell(r, c0 + 2).font = Font(name=F, bold=True)
    ws.cell(r, jcol).alignment = Alignment(horizontal='center')
    r += 1
end = r - 1
ws.cell(r, 3, '合計').font = Font(name=F, bold=True)
for c in range(4, jcol):
    Lc = CL(c)
    ws.cell(r, c, f'=SUM({Lc}{H2+1}:{Lc}{end})').number_format = NUM
for c in range(1, ncol + 1):
    ws.cell(r, c).fill = TFILL; ws.cell(r, c).border = BD; ws.cell(r, c).font = Font(name=F, bold=True)
J = CL(jcol)
TQ = CL(blocks[-1][1] + 1)
s = r + 2
summary = [('払出対象の品目数', f'=COUNTIF({J}{H2+1}:{J}{end},"<>-")'),
           ('不足（マイナス）品目数', f'=COUNTIF({J}{H2+1}:{J}{end},"不足")'),
           ('払出リストの合計（照合用）', f'=払出リスト!E{rlast+1}'),
           ('一覧の払出数合計', f'=SUM({TQ}{H2+1}:{TQ}{end})')]
for i, (k, f) in enumerate(summary):
    ws.cell(s + i, 2, k).font = Font(name=F, bold=True)
    c = ws.cell(s + i, 3, f); c.font = Font(name=F, bold=True); c.alignment = Alignment(horizontal='left'); c.number_format = NUM
ws.cell(s + len(summary) + 1, 2, '※払出数は「払出リスト」シート（青字）から倉庫別に集計。払出元を変えるときはリストの保管場所コードを書き換える。').font = Font(name=F, italic=True, color='595959')

area = f'A{H2+1}:{J}{end}'
ws.conditional_formatting.add(area, FormulaRule(formula=[f'${J}{H2+1}="不足"'], fill=PatternFill('solid', fgColor='F8CBAD'), font=Font(color='9C0006', bold=True)))
ws.conditional_formatting.add(f'A{H2+1}:C{end}', FormulaRule(formula=[f'${TQ}{H2+1}>0'], fill=PatternFill('solid', fgColor='FFF2CC')))
for l, c0 in blocks:
    if l and l != default_loc:
        Lq = CL(c0 + 1)
        ws.conditional_formatting.add(f'{Lq}{H2+1}:{Lq}{end}', FormulaRule(formula=[f'{Lq}{H2+1}>0'], fill=PatternFill('solid', fgColor='FCE4D6'), font=Font(bold=True)))
ws.freeze_panes = f'D{H2+1}'; ws.auto_filter.ref = f'A{H2}:{J}{end}'

wb.calculation = CalcProperties(fullCalcOnLoad=True)
wb.save(out)
print(out, f'品目 {len(items)} / 払出 {len(req)} 行 合計 {sum(q for _, _, q in req)} / 元データにない品目 {missing}')
