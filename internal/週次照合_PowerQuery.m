// 週次照合：複数日の期間のファイル（週次など）は取り込まず、その合計を
// 取り込んだ日別ファイル（クエリ「POS蓄積」）の同じ期間の合計と照合する
let
    Settings = Excel.CurrentWorkbook(){[Name = "tblSettings"]}[Content],
    GetSetting = (key as text) as text =>
        let v = try Settings{[項目 = key]}[値] otherwise null in if v = null then "" else Text.From(v),
    FolderPath = GetSetting("CSVフォルダ"),
    Encoding = Number.From(GetSetting("文字コード")),
    Delimiter = GetSetting("区切り文字"),
    EventKey = GetSetting("催事ファイル名キーワード"),
    T = (x) as text => if x = null then "" else Text.Trim(Text.From(x)),
    Num = (x) as number => let v = try Number.From(Text.Remove(T(x), {","}), "ja-JP") otherwise null in if v = null then 0 else v,
    Sum0 = (l as list) as number => let v = List.Sum(l) in if v = null then 0 else v,
    ColumnMap = Excel.CurrentWorkbook(){[Name = "tblColumns"]}[Content],
    Col = (std as text) as text =>
        let r = Table.SelectRows(ColumnMap, each [標準列名] = std) in if Table.RowCount(r) = 0 then "" else T(r{0}[POSのCSVの列名]),
    DigitRuns = (s as text) as list =>
        let
            st = List.Accumulate(Text.ToList(s), [runs = {}, cur = ""], (a, ch) =>
                if Text.Contains("0123456789", ch) then [runs = a[runs], cur = a[cur] & ch]
                else [runs = if a[cur] <> "" then a[runs] & {a[cur]} else a[runs], cur = ""]),
            all = if st[cur] <> "" then st[runs] & {st[cur]} else st[runs]
        in
            List.Select(all, each Text.Length(_) = 8),
    ToDate = (t as text) as nullable date =>
        try #date(Number.From(Text.Start(t, 4)), Number.From(Text.Middle(t, 4, 2)), Number.From(Text.End(t, 2))) otherwise null,
    IsEvent = (n as text) as logical => EventKey <> "" and Text.Contains(n, EventKey),
    Files0 = Table.SelectRows(Folder.Files(FolderPath),
        each Text.Lower([Extension]) = ".csv" and not Text.StartsWith([Name], "~")),
    Files1 = Table.AddColumn(Files0, "期間開始",
        each let d = DigitRuns([Name]) in if List.Count(d) >= 1 then ToDate(d{0}) else null, type nullable date),
    Files2 = Table.AddColumn(Files1, "期間終了",
        each let d = DigitRuns([Name]) in if List.Count(d) >= 2 then ToDate(d{1}) else [期間開始], type nullable date),
    Multi = Table.SelectRows(Files2, each [期間開始] <> null and [期間終了] <> null and [期間終了] > [期間開始]),
    // 照合相手は、取り込んだ1日分のファイルだけ
    Daily = Table.Buffer(Table.SelectRows(
        Table.SelectColumns(#"POS蓄積", {"期間開始", "期間終了", "販売点数", "税抜金額", "税込金額", "ファイル名"}),
        each [期間開始] <> null and [期間開始] = [期間終了])),
    cCode = Col("商品コード"), cName = Col("商品名"), cQty = Col("数量"), cEx = Col("税抜金額"), cInc = Col("税込金額"),
    Checked = Table.AddColumn(Multi, "R", (f) =>
        let
            t0 = Table.PromoteHeaders(
                Csv.Document(f[Content], [Delimiter = Delimiter, Encoding = Encoding, QuoteStyle = QuoteStyle.Csv]),
                [PromoteAllScalars = true]),
            items = List.Select(Table.ToRecords(t0), (r) =>
                let c = T(Record.FieldOrDefault(r, cCode, "")), n = T(Record.FieldOrDefault(r, cName, ""))
                in c <> "合計" and n <> "合計" and (c <> "" or n <> "")),
            ev = IsEvent(f[Name]),
            d = Table.SelectRows(Daily, each [期間開始] <> null and [期間開始] >= f[期間開始]
                and [期間終了] <= f[期間終了] and IsEvent([ファイル名]) = ev),
            days = List.Dates(f[期間開始], Duration.Days(f[期間終了] - f[期間開始]) + 1, #duration(1, 0, 0, 0)),
            missing = List.Difference(days, List.Distinct(d[期間開始])),
            wQty = Sum0(List.Transform(items, each Num(Record.FieldOrDefault(_, cQty, null)))),
            wEx = Sum0(List.Transform(items, each Num(Record.FieldOrDefault(_, cEx, null)))),
            wInc = Sum0(List.Transform(items, each Num(Record.FieldOrDefault(_, cInc, null)))),
            dQty = Sum0(d[販売点数]), dEx = Sum0(d[税抜金額]), dInc = Sum0(d[税込金額])
        in
            [区別 = if ev then "催事" else "店舗", 日数 = List.Count(days),
             日別ファイル数 = List.Count(List.Distinct(d[ファイル名])),
             日別ファイルがない日 = Text.Combine(List.Transform(missing, each Date.ToText(_, "M/d")), "、"),
             週次_販売点数 = wQty, 日別_販売点数 = dQty, 週次_税抜 = wEx, 日別_税抜 = dEx, 差額_税抜 = wEx - dEx,
             週次_税込 = wInc, 日別_税込 = dInc, 差額_税込 = wInc - dInc,
             判定 = if wQty = dQty and wEx = dEx and wInc = dInc then "OK" else "NG：日別の合計と合わない"]),
    Expanded = Table.ExpandRecordColumn(Checked, "R",
        {"区別", "日数", "日別ファイル数", "日別ファイルがない日", "週次_販売点数", "日別_販売点数", "週次_税抜", "日別_税抜",
         "差額_税抜", "週次_税込", "日別_税込", "差額_税込", "判定"}),
    Output = Table.SelectColumns(Expanded,
        {"Name", "期間開始", "期間終了", "区別", "日数", "日別ファイル数", "日別ファイルがない日", "週次_販売点数", "日別_販売点数",
         "週次_税抜", "日別_税抜", "差額_税抜", "週次_税込", "日別_税込", "差額_税込", "判定"}),
    Renamed = Table.RenameColumns(Output, {{"Name", "ファイル名"}}),
    Typed = Table.TransformColumnTypes(Renamed, {{"期間開始", type date}, {"期間終了", type date}}),
    Sorted = Table.Sort(Typed, {{"期間開始", Order.Ascending}, {"区別", Order.Ascending}})
in
    Sorted
