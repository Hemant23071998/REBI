Attribute VB_Name = "ExcelPipeline"
'==============================================================================
' Excel pipeline using Power Query  (requires Excel 2016 / 2019 / 2021 / 365)
'
'  1-2. Power Query reads every Excel file in a folder and combines them
'  3.   The combined query is loaded into a worksheet table
'  4.   Rows are filtered to the chosen dates (single days and/or ranges)
'  5.   Saved as  <folder>\Output\Filtered_<first>_to_<last>.xlsx
'  6.   Sorted: Phone Number A-Z, Website A-Z, Brand Name A-Z
'  7.   Duplicates removed on Brand Name only
'  8.   Header row deleted
'  9.   Saved as  <folder>\Output\Final_<first>_to_<last>.txt  (tab-delimited)
'
' HOW TO USE
'   Alt+F11 > File > Import File... > pick this .bas file
'   (import into PERSONAL.XLSB to have it in every workbook, or into any .xlsm)
'   Run macro  RunExcelPipeline  (Alt+F8). It asks for the folder and the dates.
'
' DATES - MM-DD-YYYY, single days and/or ranges, separated by commas:
'   10-05-2026, 10-06-2026, 10-07-2026, 10-15-2026, 10-16-2026
'   10-05-2026 to 10-07-2026, 10-15-2026 to 10-16-2026
'==============================================================================
Option Explicit

' ---- settings ----------------------------------------------------------------
Private Const DATE_COLUMN As String = ""          ' "" = first column with "date" in its name
Private Const DAY_FIRST As Boolean = False        ' False = MM-DD-YYYY (10-05-2026 = 5 Oct); True = DD-MM-YYYY
Private Const TXT_DATE_FORMAT As String = "mm-dd-yyyy"
Private Const UNICODE_TXT As Boolean = False      ' True = UTF-16 .txt (keeps non-English characters)
Private Const QUERY_NAME As String = "CombinedData"

Private SORT_COLUMNS As Variant                   ' set in RunExcelPipeline
Private Const DEDUPE_COLUMN As String = "Brand Name"
' ------------------------------------------------------------------------------

Public Sub RunExcelPipeline()
    SORT_COLUMNS = Array("Phone Number", "Website", "Brand Name")

    Dim folder As String, spec As String, days As Object
    folder = PickFolder()
    If folder = "" Then Exit Sub

    spec = InputBox("Dates to keep (MM-DD-YYYY), comma-separated." & vbLf & _
                    "Ranges allowed as  A to B" & vbLf & vbLf & _
                    "e.g.  10-05-2026, 10-06-2026, 10-15-2026 to 10-16-2026", "Dates")
    If Trim$(spec) = "" Then Exit Sub

    On Error GoTo Fail
    Set days = ParseDateSpec(spec)            ' Dictionary: key = CLng(date)

    Dim minD As Long, maxD As Long, k As Variant
    minD = 2147483647: maxD = 0
    For Each k In days.Keys
        If k < minD Then minD = k
        If k > maxD Then maxD = k
    Next k

    Dim outDir As String, tag As String
    outDir = folder & "\Output"
    If Dir(outDir, vbDirectory) = "" Then MkDir outDir
    tag = Format$(CDate(minD), "yyyymmdd") & "_to_" & Format$(CDate(maxD), "yyyymmdd")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' ---- Steps 1-4: Power Query --------------------------------------------
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Filtered"
    wb.Queries.Add Name:=QUERY_NAME, Formula:=BuildQuery(folder, days)

    Set lo = ws.ListObjects.Add(SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & QUERY_NAME & ";Extended Properties=""""", _
        Destination:=ws.Range("$A$1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & QUERY_NAME & "]")
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = True
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    lo.Name = QUERY_NAME

    If lo.ListRows.Count = 0 Then
        Application.DisplayAlerts = True: Application.ScreenUpdating = True
        MsgBox "No rows matched the chosen dates.", vbExclamation
        Exit Sub
    End If

    Dim dateIdx As Long, r As Long, found As Object, missing As String
    dateIdx = FindDateColumn(lo.HeaderRowRange)
    lo.ListColumns(dateIdx).DataBodyRange.NumberFormat = TXT_DATE_FORMAT

    Set found = CreateObject("Scripting.Dictionary")
    For r = 1 To lo.ListRows.Count
        found(CLng(lo.DataBodyRange(r, dateIdx).Value)) = True
    Next r
    For Each k In days.Keys
        If Not found.Exists(k) Then missing = missing & Format$(CDate(k), TXT_DATE_FORMAT) & "  "
    Next k

    ' ---- Step 5: save filtered data ----------------------------------------
    Dim xlsxPath As String
    xlsxPath = outDir & "\Filtered_" & tag & ".xlsx"
    wb.SaveAs Filename:=xlsxPath, FileFormat:=xlOpenXMLWorkbook

    ' ---- Steps 6-9 on a values-only copy -----------------------------------
    Dim wbF As Workbook, wsF As Worksheet, rng As Range, i As Long, idx As Long
    Dim filteredRows As Long, finalRows As Long
    filteredRows = lo.ListRows.Count

    Set wbF = Workbooks.Add(xlWBATWorksheet)
    Set wsF = wbF.Worksheets(1)
    lo.Range.Copy
    wsF.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    Set rng = wsF.Range("A1").CurrentRegion
    wsF.Columns(dateIdx).NumberFormat = TXT_DATE_FORMAT

    ' Step 6: sort (Excel's own sort: case-insensitive, blanks last)
    With wsF.Sort
        .SortFields.Clear
        For i = LBound(SORT_COLUMNS) To UBound(SORT_COLUMNS)
            idx = FindHeader(rng.Rows(1), CStr(SORT_COLUMNS(i)))
            If idx = 0 Then Err.Raise vbObjectError + 1, , "Column '" & SORT_COLUMNS(i) & "' not found."
            .SortFields.Add Key:=rng.Columns(idx).Offset(1).Resize(rng.Rows.Count - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next i
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With

    ' Step 7: remove duplicates on Brand Name only (keeps the first row after sorting)
    idx = FindHeader(rng.Rows(1), DEDUPE_COLUMN)
    rng.RemoveDuplicates Columns:=idx, Header:=xlYes
    finalRows = wsF.Range("A1").CurrentRegion.Rows.Count - 1

    ' Step 8: delete header row
    wsF.Rows(1).Delete

    ' Step 9: save as tab-delimited text
    Dim txtPath As String
    txtPath = outDir & "\Final_" & tag & ".txt"
    wbF.SaveAs Filename:=txtPath, FileFormat:=IIf(UNICODE_TXT, xlUnicodeText, xlText)
    wbF.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Done." & vbLf & vbLf & _
           "Filtered rows: " & filteredRows & vbLf & xlsxPath & vbLf & vbLf & _
           "Final rows (after removing duplicate brands): " & finalRows & vbLf & txtPath & _
           IIf(missing <> "", vbLf & vbLf & "No rows found for: " & missing, ""), vbInformation
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Power Query (M) - backticks are swapped for double quotes at the end
'------------------------------------------------------------------------------
Private Function BuildQuery(ByVal folder As String, ByVal days As Object) As String
    Dim m As String, culture As String, dateMatch As String, dayList As String, k As Variant

    culture = IIf(DAY_FIRST, "en-GB", "en-US")
    If DATE_COLUMN = "" Then
        dateMatch = "Text.Contains(Text.Lower(_), `date`)"
    Else
        dateMatch = "Text.Lower(_) = `" & LCase$(Application.WorksheetFunction.Trim(DATE_COLUMN)) & "`"
    End If
    For Each k In days.Keys
        dayList = dayList & IIf(dayList = "", "", ", ") & _
            "#date(" & Year(CDate(k)) & "," & Month(CDate(k)) & "," & Day(CDate(k)) & ")"
    Next k

    m = "let" & vbLf
    m = m & "    Source = Folder.Files(`" & folder & "`)," & vbLf
    m = m & "    ExcelFiles = Table.SelectRows(Source, each List.Contains({`.xlsx`, `.xlsm`, `.xls`}, Text.Lower([Extension])) and not Text.StartsWith([Name], `~$`))," & vbLf
    ' tidy header spacing so "Brand  Name" matches "Brand Name"
    m = m & "    Tidy = (t as table) as table => Table.TransformColumnNames(t, each Text.Combine(List.Select(Text.Split(Text.Trim(Text.From(_)), ` `), each _ <> ``), ` `))," & vbLf
    m = m & "    FirstSheet = (bin as binary) as table => let wb = Excel.Workbook(bin, true), sh = Table.SelectRows(wb, each [Kind] = `Sheet`) in Tidy(sh{0}[Data])," & vbLf
    m = m & "    Combined = Table.Combine(List.Transform(ExcelFiles[Content], FirstSheet))," & vbLf
    m = m & "    NonBlank = Table.SelectRows(Combined, each List.NonNullCount(List.ReplaceValue(Record.FieldValues(_), ``, null, Replacer.ReplaceValue)) > 0)," & vbLf
    m = m & "    DateCol = let c = List.First(List.Select(Table.ColumnNames(NonBlank), each " & dateMatch & ")) in if c = null then error `No date column found` else c," & vbLf
    ' everything except the date column becomes text (keeps phone leading zeros, no 9.87E+09)
    m = m & "    ToText = (v) => if v = null then null else if Value.Is(v, type number) and Number.Round(v) = v then Number.ToText(v, `0`) else Text.From(v)," & vbLf
    m = m & "    ToDate = (v) => if v = null then null else if Value.Is(v, type datetime) then DateTime.Date(v) else if Value.Is(v, type date) then v" & _
            " else if Value.Is(v, type number) then Date.From(v)" & _
            " else let t = Text.Trim(Text.From(v)) in try Date.FromText(t, `" & culture & "`) otherwise try DateTime.Date(DateTime.FromText(t, `" & culture & "`)) otherwise null," & vbLf
    m = m & "    AsText = Table.TransformColumns(NonBlank, List.Transform(List.RemoveItems(Table.ColumnNames(NonBlank), {DateCol}), (c) => {c, ToText, type text}))," & vbLf
    m = m & "    AsDate = Table.TransformColumns(AsText, {{DateCol, ToDate, type date}})," & vbLf
    m = m & "    Days = {" & dayList & "}," & vbLf
    m = m & "    Filtered = Table.SelectRows(AsDate, each List.Contains(Days, Record.Field(_, DateCol)))" & vbLf
    m = m & "in" & vbLf & "    Filtered"

    BuildQuery = Replace(m, "`", Chr(34))
End Function

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function PickFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing the Excel files"
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

' "2026-10-05, 2026-10-07 to 2026-10-09" -> Dictionary of CLng(date)
Private Function ParseDateSpec(ByVal spec As String) As Object
    Dim d As Object, items() As String, item As Variant, parts() As String
    Dim a As Date, b As Date, x As Long
    Set d = CreateObject("Scripting.Dictionary")

    spec = LCase$(Application.WorksheetFunction.Trim(spec))
    spec = Replace(spec, " to ", "..")
    spec = Replace(Replace(spec, ";", ","), " ", ",")
    items = Split(spec, ",")

    For Each item In items
        If Trim$(item) <> "" Then
            If InStr(item, "..") > 0 Then
                parts = Split(item, "..")
                a = ParseOneDate(parts(0)): b = ParseOneDate(parts(1))
                If b < a Then Err.Raise vbObjectError + 2, , "Range '" & item & "' ends before it starts."
                For x = CLng(a) To CLng(b): d(x) = True: Next x
            Else
                d(CLng(ParseOneDate(item))) = True
            End If
        End If
    Next item
    If d.Count = 0 Then Err.Raise vbObjectError + 3, , "No dates given."
    Set ParseDateSpec = d
End Function

' Accepts MM-DD-YYYY (DD-MM-YYYY when DAY_FIRST = True); YYYY-MM-DD also works
Private Function ParseOneDate(ByVal t As String) As Date
    Dim p() As String, y As Long, mo As Long, dy As Long
    On Error GoTo Bad
    t = Trim$(Replace(Replace(t, "/", "-"), ".", "-"))
    p = Split(t, "-")
    If UBound(p) <> 2 Then GoTo Bad
    If Not (IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2))) Then GoTo Bad
    If Len(p(0)) = 4 Then
        y = p(0): mo = p(1): dy = p(2)
    ElseIf DAY_FIRST Then
        dy = p(0): mo = p(1): y = p(2)
    Else
        mo = p(0): dy = p(1): y = p(2)
    End If
    Dim dt As Date
    dt = DateSerial(y, mo, dy)
    If Year(dt) <> y Or Month(dt) <> mo Or Day(dt) <> dy Then GoTo Bad   ' rejects 2026-02-30 etc.
    ParseOneDate = dt
    Exit Function
Bad:
    On Error GoTo 0
    Err.Raise vbObjectError + 4, , "Could not understand date '" & t & "'. Use MM-DD-YYYY, e.g. 10-05-2026."
End Function

Private Function NormHeader(ByVal s As String) As String
    NormHeader = LCase$(Application.WorksheetFunction.Trim(s))
End Function

Private Function FindHeader(ByVal headerRow As Range, ByVal wanted As String) As Long
    Dim c As Range
    For Each c In headerRow.Cells
        If NormHeader(CStr(c.Value)) = NormHeader(wanted) Then FindHeader = c.Column - headerRow.Column + 1: Exit Function
    Next c
End Function

Private Function FindDateColumn(ByVal headerRow As Range) As Long
    Dim c As Range
    If DATE_COLUMN <> "" Then FindDateColumn = FindHeader(headerRow, DATE_COLUMN): Exit Function
    For Each c In headerRow.Cells
        If InStr(1, CStr(c.Value), "date", vbTextCompare) > 0 Then FindDateColumn = c.Column - headerRow.Column + 1: Exit Function
    Next c
End Function
