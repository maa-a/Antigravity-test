"""SAPUI5 エクスポート形式の在庫Excelを、build.py が読める XTab DS_1 形式に変換する。

使い方: python3 convert_sapui5.py <SAPUI5エクスポート.xlsx> <出力.xlsx> [--prev 前回のSAP在庫.xlsx]

--prev を付けると、前回データ（XTab DS_1 形式）から次を補う:
  - JANコード
  - 前回は登録があり、今回のエクスポートに無い「品目×保管場所」を在庫0の行として追加
    （エクスポートは在庫0の行を出さないため。0の行も省略せず出力するルール）
  補った行は灰色の斜体にし、「補完メモ」シートに一覧を残す。

入力列: 品目 / 品目テキスト / 利用可能在庫 / 利用可能在庫の金額 / 保管場所 / 原価
出力:   XTab DS_1（JANコード(EAN/UPC), 品目コード, 品目テキスト, 保管場所, 保管場所名, 在庫数量, 在庫金額）
        と元のエクスポートシート（そのまま）。
JANコードはエクスポートに無いため空欄。保管場所名は既知のコードから補う。
"""
import sys
import openpyxl
from openpyxl.styles import Font

LOC_NAMES = {'CX01': 'CX催事', 'DB00': 'DBS倉庫', 'ZZ01': 'その他'}
QF = '#,##0 "PC";-#,##0 "PC"'
AF = '#,##0 "JPY";-#,##0 "JPY"'

args = sys.argv[1:]
prev = None
if '--prev' in args:
    i = args.index('--prev'); prev = args[i + 1]; del args[i:i + 2]
src_path, out = args[0], args[1]
src_wb = openpyxl.load_workbook(src_path)
src = src_wb.worksheets[0]
hdr = {c.value: i for i, c in enumerate(src[1])}
need = ['品目', '品目テキスト', '利用可能在庫', '利用可能在庫の金額', '保管場所']
miss = [h for h in need if h not in hdr]
if miss:
    sys.exit(f'想定の列がありません: {miss}')

rows = []
for r in src.iter_rows(min_row=2, values_only=True):
    if not r[hdr['品目']]:
        continue
    loc = r[hdr['保管場所']]
    rows.append([None, r[hdr['品目']], r[hdr['品目テキスト']], loc, LOC_NAMES.get(loc, loc),
                 r[hdr['利用可能在庫']] or 0, r[hdr['利用可能在庫の金額']] or 0])
filled = []
if prev:
    pws = openpyxl.load_workbook(prev, data_only=True)['XTab DS_1']
    jan, known = {}, {}
    for r in pws.iter_rows(min_row=2, values_only=True):
        if not r[1]: continue
        if r[0]: jan.setdefault(r[1], r[0])
        known.setdefault((r[1], r[3]), r)
    have = {(r[1], r[3]) for r in rows}
    for r in rows:
        r[0] = jan.get(r[1])
    for (code, loc), r in known.items():
        if (code, loc) not in have:
            new = [r[0], code, r[2], loc, r[4] or LOC_NAMES.get(loc, loc), 0, 0]
            rows.append(new); filled.append(new)
# 品目コード順・保管場所順に揃える
rows.sort(key=lambda x: (x[1], x[3]))

wb = openpyxl.Workbook()
ws = wb.active; ws.title = 'XTab DS_1'
ws.append(['JANコード(EAN/UPC)', '品目コード', '品目テキスト', '保管場所', '保管場所名', '在庫数量', '在庫金額'])
for row in rows:
    ws.append(row)
for row in ws.iter_rows():
    for c in row:
        c.font = Font(name='Calibri')
for r in range(2, ws.max_row + 1):
    ws.cell(r, 6).number_format = QF; ws.cell(r, 7).number_format = AF
for col in 'ABCDEFG':
    ws.column_dimensions[col].width = 13
ws.column_dimensions['B'].width = 22; ws.column_dimensions['C'].width = 52

fill_keys = {(r[1], r[3]) for r in filled}
for r in range(2, ws.max_row + 1):
    if (ws.cell(r, 2).value, ws.cell(r, 4).value) in fill_keys:
        for c in range(1, 8):
            ws.cell(r, c).font = Font(name='Calibri', italic=True, color='808080')
if filled:
    memo = wb.create_sheet('補完メモ')
    memo.append(['前回データから在庫0として補った行（今回のエクスポートに無い品目×保管場所）'])
    memo.append([f'前回データ: {prev.rsplit("/", 1)[-1]}'])
    memo.append(['品目コード', '品目テキスト', '保管場所', '保管場所名', '在庫数量', '在庫金額'])
    for r in sorted(filled, key=lambda x: (x[1], x[3])):
        memo.append([r[1], r[2], r[3], r[4], 0, 0])
    memo.column_dimensions['A'].width = 22; memo.column_dimensions['B'].width = 52
    for row in memo.iter_rows():
        for c in row: c.font = Font(name='Calibri', bold=c.row in (1, 3))

# 元のエクスポートもそのまま残す
cp = wb.create_sheet(src.title[:31])
for row in src.iter_rows(values_only=True):
    cp.append(list(row))
wb.save(out)
print(out, len(rows), '行', sorted({r[3] for r in rows}), '補完', len(filled), '行')
