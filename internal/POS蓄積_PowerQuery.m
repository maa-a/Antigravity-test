// POS蓄積：フォルダ内のPOS CSVをすべて読み込み、標準の列にそろえて1つの表にする
let
    Settings = Excel.CurrentWorkbook(){[Name = "tblSettings"]}[Content],
    GetSetting = (key as text) as text =>
        let v = Settings{[項目 = key]}[値] in if v = null then "" else Text.From(v),
    FolderPath = GetSetting("CSVフォルダ"),
    Encoding = Number.From(GetSetting("文字コード")),
    Delimiter = GetSetting("区切り文字"),
    EventKey = GetSetting("催事ファイル名キーワード"),
    ColumnMap = Table.SelectRows(
        Excel.CurrentWorkbook(){[Name = "tblColumns"]}[Content],
        each [POSのCSVの列名] <> null and Text.Trim(Text.From([POSのCSVの列名])) <> ""),
    CategoryMap = Table.SelectRows(
        Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "tblCategory"]}[Content], {"POSの部門名", "区分"}),
        each [POSの部門名] <> null),
    Files = Table.SelectRows(
        Folder.Files(FolderPath),
        each Text.Lower([Extension]) = ".csv" and not Text.StartsWith([Name], "~")),
    Loaded = Table.AddColumn(Files, "Data", (f) =>
        Table.AddColumn(
            Table.PromoteHeaders(
                Csv.Document(f[Content], [Delimiter = Delimiter, Encoding = Encoding, QuoteStyle = QuoteStyle.Csv]),
                [PromoteAllScalars = true]),
            "ファイル名", each f[Name])),
    Combined = Table.Combine(Loaded[Data]),
    Renamed = Table.RenameColumns(
        Combined,
        List.Zip({List.Transform(ColumnMap[POSのCSVの列名], Text.From), ColumnMap[標準列名]}),
        MissingField.Ignore),
    Standard = {"売上日", "部門", "商品名", "数量", "税抜金額", "税込金額", "税率"},
    Filled = List.Accumulate(Standard, Renamed,
        (t, c) => if Table.HasColumns(t, c) then t else Table.AddColumn(t, c, each null)),
    Typed = Table.TransformColumnTypes(Filled,
        {{"売上日", type date}, {"部門", type text}, {"商品名", type text},
         {"数量", type number}, {"税抜金額", type number}, {"税込金額", type number}}, "ja-JP"),
    NoBlank = Table.SelectRows(Typed, each [売上日] <> null),
    AddTax = Table.AddColumn(NoBlank, "消費税",
        each if [税込金額] <> null and [税抜金額] <> null then [税込金額] - [税抜金額] else null, type number),
    AddMonth = Table.AddColumn(AddTax, "計上月", each Date.StartOfMonth([売上日]), type date),
    AddRate = Table.AddColumn(AddMonth, "税率区分", each
        let
            raw = if [税率] = null then null else Text.Trim(Text.Remove(Text.From([税率]), {"%", "％"})),
            n = if raw = null or raw = "" then null else try Number.From(raw, "en-US") otherwise null,
            v = if n = null then null else if n > 1 then n / 100 else n
        in
            if v = null then "不明"
            else if Number.Round(v, 2) = 0.1 then "10%"
            else if Number.Round(v, 2) = 0.08 then "8%"
            else if v = 0 then "非課税"
            else "不明", type text),
    Joined = Table.NestedJoin(AddRate, {"部門"}, CategoryMap, {"POSの部門名"}, "Map", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Map", {"区分"}, {"区分0"}),
    AddCategory = Table.AddColumn(Expanded, "区分", each
        let k = if [区分0] = null or [区分0] = "" then "未設定" else [区分0]
        in if k = "物販" and EventKey <> "" and Text.Contains([ファイル名], EventKey) then "POPUP催事物販" else k,
        type text),
    Output = Table.SelectColumns(AddCategory,
        {"計上月", "売上日", "区分", "部門", "商品名", "数量", "税抜金額", "消費税", "税込金額", "税率区分", "ファイル名"}),
    Sorted = Table.Sort(Output, {{"売上日", Order.Ascending}, {"ファイル名", Order.Ascending}})
in
    Sorted
