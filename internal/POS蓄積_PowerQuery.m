// POS蓄積：フォルダ内のPOS CSV（商品別売上）を読み込み、標準の列にそろえて1つの表にする
// 取込対象が「日別ファイルのみ」のときは、ファイル名の2つの日付が同じ（1日分）ファイルだけを取り込む
// 計上月：CSVに日付列があればその月、なければファイル名の8桁の日付（yyyymmdd）の1つ目の月
let
    Settings = Excel.CurrentWorkbook(){[Name = "tblSettings"]}[Content],
    GetSetting = (key as text) as text =>
        let v = try Settings{[項目 = key]}[値] otherwise null in if v = null then "" else Text.From(v),
    FolderPath = GetSetting("CSVフォルダ"),
    Encoding = Number.From(GetSetting("文字コード")),
    Delimiter = GetSetting("区切り文字"),
    EventKey = GetSetting("催事ファイル名キーワード"),
    ImportAll = GetSetting("取込対象") = "すべて",
    T = (x) as text => if x = null then "" else Text.Trim(Text.From(x)),
    NormRate = (x) as nullable text =>
        let
            t = Text.Remove(T(x), {"%", "％"}),
            n = try Number.From(t, "en-US") otherwise null,
            v = if n = null then null else if n > 1 then n / 100 else n
        in
            if t = "" then null
            else if t = "非課税" then "非課税"
            else if v = null then null
            else if v >= 0.09 then "10%" else if v >= 0.065 then "8%" else "非課税",
    ColumnMap = Table.SelectRows(Excel.CurrentWorkbook(){[Name = "tblColumns"]}[Content],
        each T([POSのCSVの列名]) <> ""),
    CategoryMap = Table.SelectRows(
        Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "tblCategory"]}[Content], {"POSの部門名", "区分", "税率"}),
        each T([POSの部門名]) <> ""),
    ItemMap = Table.SelectRows(
        Table.TransformColumns(
            Table.SelectColumns(Excel.CurrentWorkbook(){[Name = "tblItem"]}[Content], {"商品コード", "区分"}),
            {{"商品コード", each T(_), type text}}),
        each [商品コード] <> "" and T([区分]) <> ""),
    // ファイル名から8桁の数字（yyyymmdd）を順に取り出す
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
    Files0 = Table.SelectRows(Folder.Files(FolderPath),
        each Text.Lower([Extension]) = ".csv" and not Text.StartsWith([Name], "~")),
    Files1 = Table.AddColumn(Files0, "期間開始",
        each let d = DigitRuns([Name]) in if List.Count(d) >= 1 then ToDate(d{0}) else null, type nullable date),
    Files2 = Table.AddColumn(Files1, "期間終了",
        each let d = DigitRuns([Name]) in if List.Count(d) >= 2 then ToDate(d{1}) else [期間開始], type nullable date),
    // 複数日の期間のファイル（週次など）は照合用。取り込まない（「週次照合」クエリで使う）
    Files = if ImportAll then Files2 else Table.SelectRows(Files2, each
        not ([期間開始] <> null and [期間終了] <> null and [期間終了] > [期間開始])),
    // 期間が重なるファイル（同じ日の二重取込）を検出。店舗と催事は別々に判定
    Periods = Table.Buffer(Table.SelectColumns(Files, {"Name", "期間開始", "期間終了"})),
    IsEvent = (n as text) as logical => EventKey <> "" and Text.Contains(n, EventKey),
    Overlap = (n as text, s as nullable date, e as nullable date) as logical =>
        if s = null or e = null then false
        else List.AnyTrue(List.Transform(Table.ToRecords(Periods), (p) =>
            p[Name] <> n and p[期間開始] <> null and p[期間終了] <> null
            and p[期間開始] <= e and p[期間終了] >= s and IsEvent(p[Name]) = IsEvent(n))),
    Loaded = Table.AddColumn(Files, "Data", (f) =>
        let
            t0 = Table.PromoteHeaders(
                Csv.Document(f[Content], [Delimiter = Delimiter, Encoding = Encoding, QuoteStyle = QuoteStyle.Csv]),
                [PromoteAllScalars = true]),
            note =
                (if f[期間開始] = null then "ファイル名に日付なし " else "")
                & (if f[期間開始] <> null and f[期間終了] <> null
                      and Date.StartOfMonth(f[期間開始]) <> Date.StartOfMonth(f[期間終了]) then "期間が月をまたぐ " else "")
                & (if Overlap(f[Name], f[期間開始], f[期間終了]) then "期間が他のファイルと重複 " else ""),
            t1 = Table.AddColumn(t0, "ファイル名", each f[Name]),
            t2 = Table.AddColumn(t1, "期間開始", each f[期間開始]),
            t3 = Table.AddColumn(t2, "期間終了", each f[期間終了]),
            t4 = Table.AddColumn(t3, "注意0", each note)
        in
            t4),
    Combined = Table.Combine(Loaded[Data]),
    Renamed = Table.RenameColumns(Combined,
        List.Zip({List.Transform(ColumnMap[POSのCSVの列名], T), ColumnMap[標準列名]}), MissingField.Ignore),
    Standard = {"売上日", "商品コード", "商品名", "部門", "数量", "返品数", "税抜金額", "消費税", "税込金額", "原価"},
    Filled = List.Accumulate(Standard, Renamed,
        (t, c) => if Table.HasColumns(t, c) then t else Table.AddColumn(t, c, each null)),
    // 「合計」行と空行を除く
    NoTotal = Table.SelectRows(Filled, each
        T([商品コード]) <> "合計" and T([商品名]) <> "合計" and (T([商品コード]) <> "" or T([商品名]) <> "")),
    Typed = Table.TransformColumnTypes(NoTotal, {
        {"売上日", type date}, {"商品コード", type text}, {"商品名", type text}, {"部門", type text},
        {"数量", type number}, {"返品数", type number}, {"税抜金額", type number},
        {"消費税", type number}, {"税込金額", type number}, {"原価", type number}}, "ja-JP"),
    AddTax = Table.AddColumn(Typed, "消費税2", each
        if [消費税] <> null then [消費税]
        else if [税込金額] <> null and [税抜金額] <> null then [税込金額] - [税抜金額] else null, type nullable number),
    AddMonth = Table.AddColumn(AddTax, "計上月", each
        if [売上日] <> null then Date.StartOfMonth([売上日])
        else if [期間開始] <> null then Date.StartOfMonth([期間開始]) else null, type nullable date),
    Joined = Table.NestedJoin(AddMonth, {"部門"}, CategoryMap, {"POSの部門名"}, "Map", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Joined, "Map", {"区分", "税率"}, {"区分_部門", "税率_部門"}),
    JoinedItem = Table.NestedJoin(Expanded, {"商品コード"}, ItemMap, {"商品コード"}, "Item", JoinKind.LeftOuter),
    ExpandedItem = Table.ExpandTableColumn(JoinedItem, "Item", {"区分"}, {"区分_商品"}),
    AddCategory = Table.AddColumn(ExpandedItem, "区分", each
        let
            k0 = if T([区分_商品]) <> "" then T([区分_商品])
                 else if T([区分_部門]) <> "" then T([区分_部門]) else "未設定"
        in
            if k0 = "物販" and IsEvent([ファイル名]) then "POPUP催事物販" else k0, type text),
    AddRate = Table.AddColumn(AddCategory, "税率区分", each
        let r = NormRate([税率_部門])
        in if r <> null then r
           else if [税抜金額] = null or [税抜金額] = 0 or [消費税2] = null then "0円"
           else if [消費税2] / [税抜金額] >= 0.09 then "10%"
           else if [消費税2] / [税抜金額] >= 0.065 then "8%" else "不明", type text),
    AddNote = Table.AddColumn(AddRate, "注意", each
        Text.Trim([注意0] & (if [計上月] = null then "計上月不明" else "")), type text),
    Output = Table.SelectColumns(AddNote,
        {"計上月", "期間開始", "期間終了", "区分", "部門", "商品コード", "商品名", "数量", "返品数",
         "税抜金額", "消費税2", "税込金額", "税率区分", "原価", "ファイル名", "注意"}),
    Renamed2 = Table.RenameColumns(Output, {{"数量", "販売点数"}, {"消費税2", "消費税"}}),
    Sorted = Table.Sort(Renamed2, {{"計上月", Order.Ascending}, {"ファイル名", Order.Ascending}, {"区分", Order.Ascending}, {"税抜金額", Order.Descending}})
in
    Sorted
