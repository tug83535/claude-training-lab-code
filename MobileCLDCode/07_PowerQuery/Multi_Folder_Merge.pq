// ============================================================================
// Multi_Folder_Merge.pq
// Power Query (M): walk N folders, read every matching file, normalize their
// schemas, tag each row with its source, and return a single merged table.
//
// WHY THIS IS NOT NATIVE EXCEL / ONEDRIVE
// ---------------------------------------
// "From Folder" only reads ONE folder and assumes IDENTICAL schemas. Real
// finance + ops data lives in 15 nested folders (one per region, per month,
// per vendor). Columns drift between files. This function is the general
// solution: different column orders, different casing, missing columns,
// and per-file quirks all survive.
//
// USE CASE (software business)
//   - AP team has /Invoices/2026/{01,02,03,...}/{vendor}/invoice.csv
//   - Sales ops has /Commissions/{region}/{quarter}.xlsx
//   - IT has /Licenses/{tool}/{month}/export.csv
//   Any of these become one master table in 5 seconds.
//
// USAGE
// -----
// Paste into Advanced Editor. Set the FoldersToScan and FilePattern up top.
// Returns one table with columns from all files, plus a "source_file"
// column that tells you exactly which file each row came from.
// ============================================================================

let
    // --------------------- CONFIGURE THESE --------------------
    FoldersToScan = {
        "C:\Shared\Invoices\2026",
        "C:\Shared\Invoices\Archive\2025"
    },
    FilePattern       = "*.csv",            // "*.xlsx" also works
    IncludeSubfolders = true,
    HeaderRow         = true,
    // Normalize all column names to snake_case:
    NormalizeColumns  = true,
    // ----------------------------------------------------------

    // Collect every file in every listed folder
    AllFiles = List.Combine(List.Transform(FoldersToScan, each
        let
            folder = Folder.Files(_),
            filtered = Table.SelectRows(folder,
                (row) =>
                    (IncludeSubfolders or row[Folder Path] = _ & "\") and
                    Text.EndsWith(Text.Lower(row[Name]),
                                  Text.Lower(Text.AfterDelimiter(FilePattern, "*"))))
        in
            Table.ToRecords(filtered)
    )),

    AllFilesTable = Table.FromRecords(AllFiles),

    // Read each file into a table, tolerant of CSV quirks
    ReadOne = (path as text, name as text) as table =>
        let
            ext = Text.Lower(Text.AfterDelimiter(name, ".", { 0, RelativePosition.FromEnd })),
            tbl =
                if ext = "csv" then
                    let
                        raw = Csv.Document(File.Contents(path & name),
                                           [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv])
                    in
                        if HeaderRow then Table.PromoteHeaders(raw, [PromoteAllScalars=true]) else raw
                else if ext = "xlsx" or ext = "xlsb" or ext = "xlsm" then
                    let
                        book   = Excel.Workbook(File.Contents(path & name), null, true),
                        sheet  = book{[Kind="Sheet"]}[Data]
                    in
                        if HeaderRow then Table.PromoteHeaders(sheet, [PromoteAllScalars=true]) else sheet
                else
                    null
        in
            Table.AddColumn(tbl, "source_file", each name),

    TablesList = Table.AddColumn(AllFilesTable, "data", each ReadOne([Folder Path], [Name])),

    // Normalize column names if requested, then union-all
    Snake = (s as text) as text =>
        let
            lower = Text.Lower(s),
            under = Text.Replace(Text.Replace(lower, " ", "_"), "-", "_"),
            clean = Text.Select(under, {"a".."z", "0".."9", "_"})
        in
            clean,

    NormalizeOne = (t as table) =>
        if NormalizeColumns then
            Table.RenameColumns(t,
                List.Transform(Table.ColumnNames(t), each { _, Snake(_) }),
                MissingField.Ignore)
        else
            t,

    Normalized = Table.AddColumn(TablesList, "norm", each NormalizeOne([data])),

    Combined = Table.Combine(Normalized[norm])
in
    Combined
