Attribute VB_Name = "ExcelPipeline"
'==============================================================================
' Excel pipeline using Power Query  (requires Excel 2016 / 2019 / 2021 / 365)
'
' Saves two Excel files:
'  1. COMBINED_FOLDER\Combined_<first>_to_<last>.xlsx
'       Power Query joins every Excel file in the input folder and keeps only
'       the chosen dates; everything else exactly as-is
'  2. OUTPUT_FOLDER\Final_<first>_to_<last>.xlsx
'       Power Query keeps only the chosen dates, then: sort (Phone, Website,
'       Brand), remove duplicate brands, delete the DROP_COLUMNS, remove "-"
'       from phone numbers, remove "^" from every column
'
' HOW TO USE
'   Alt+F11 > File > Import File... > pick this .bas file
'   (import into PERSONAL.XLSB to have it in every workbook, or into any .xlsm)
'   Run macro  RunExcelPipeline  (Alt+F8).
'
' DATES - MM-DD-YYYY, single days and/or ranges, separated by commas:
'   10-05-2026, 10-06-2026, 10-07-2026, 10-15-2026, 10-16-2026
'   10-05-2026 to 10-07-2026, 10-15-2026 to 10-16-2026
'
' COLUMNS are found by their HEADER TEXT (any position), ignoring upper/lower
' case and extra spaces.
'==============================================================================
Option Explicit

'==============================================================================
'  SETTINGS - EDIT HERE
'==============================================================================

' ---- Folders ----
Private Const INPUT_FOLDER As String = ""        ' e.g. "C:\Data\Input";  "" = pick folder each time
Private Const COMBINED_FOLDER As String = ""     ' where Combined_*.xlsx is saved;  "" = same as OUTPUT_FOLDER
Private Const OUTPUT_FOLDER As String = ""       ' where Final_*.xlsx is saved;  "" = <input folder>\Output

' ---- Column headers (type them exactly as they appear in row 1 of your files) ----
Private Const PHONE_COLUMN As String = "Phone Number"
Private Const WEBSITE_COLUMN As String = "Website"
Private Const BRAND_COLUMN As String = "Brand Name"   ' also the column used to remove duplicates
Private Const DATE_COLUMN As String = ""              ' "" = first column with "date" in its name

' ---- Columns to DELETE from the Final file (exact header names, comma-separated) ----
Private Const DROP_COLUMNS As String = ""             ' e.g. "Email, Address, Notes"

' ---- Characters to remove from the Final file (every character listed is removed) ----
Private Const PHONE_REMOVE_CHARS As String = "-"      ' from the phone column only
Private Const REMOVE_CHARS_EVERYWHERE As String = "^" ' from every column

' ---- Dates ----
Private Const DAY_FIRST As Boolean = False            ' False = MM-DD-YYYY, True = DD-MM-YYYY
Private Const DATE_DISPLAY_FORMAT As String = "mm-dd-yyyy"

'==============================================================================

Public Sub RunExcelPipeline()
    Dim sortCols As Variant
    sortCols = Array(PHONE_COLUMN, WEBSITE_COLUMN, BRAND_COLUMN)

    Dim folder As String, spec As String, days As Object
    folder = INPUT_FOLDER
    If folder = "" Then folder = PickFolder()
    If folder = "" Then Exit Sub
    If Right$(folder, 1) = "\" Then folder = Left$(folder, Len(folder) - 1)
    If Dir(folder, vbDirectory) = "" Then MsgBox "Folder not found: " & folder, vbCritical: Exit Sub

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

    Dim outDir As String, combDir As String, tag As String
    outDir = IIf(OUTPUT_FOLDER = "", folder & "\Output", OUTPUT_FOLDER)
    combDir = IIf(COMBINED_FOLDER = "", outDir, COMBINED_FOLDER)
    EnsureFolder outDir
    EnsureFolder combDir
    tag = Format$(CDate(minD), "yyyymmdd") & "_to_" & Format$(CDate(maxD), "yyyymmdd")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim i As Long, idx As Long, warnings As String

    ' ---- Copy 1: combine all files as-is, chosen dates only, and save --------
    Dim wbC As Workbook, loC As ListObject, combinedPath As String, combinedRows As Long
    Set wbC = Workbooks.Add(xlWBATWorksheet)
    wbC.Worksheets(1).Name = "Combined"
    Set loC = LoadQuery(wbC, "Combined", BuildRawQuery(folder, days))
    For i = LBound(sortCols) To UBound(sortCols)
        If FindHeader(loC.HeaderRowRange, CStr(sortCols(i))) = 0 Then _
            Err.Raise vbObjectError + 1, , "Column '" & sortCols(i) & "' not found." & vbLf & _
                "Headers found: " & HeaderList(loC.HeaderRowRange) & vbLf & _
                "Change PHONE_COLUMN / WEBSITE_COLUMN / BRAND_COLUMN in the SETTINGS."
    Next i
    combinedRows = loC.ListRows.Count
    If combinedRows > 0 Then   ' display real dates as mm-dd-yyyy (values unchanged)
        idx = FindDateColumn(loC.HeaderRowRange)
        If idx > 0 Then loC.ListColumns(idx).DataBodyRange.NumberFormat = DATE_DISPLAY_FORMAT
    End If
    combinedPath = combDir & "\Combined_" & tag & ".xlsx"
    wbC.SaveAs Filename:=combinedPath, FileFormat:=xlOpenXMLWorkbook
    wbC.Close SaveChanges:=False

    ' ---- Copy 2: Power Query date filter (temporary workbook) ---------------
    Dim wb As Workbook, lo As ListObject
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets(1).Name = "Filtered"
    Set lo = LoadQuery(wb, "Filtered", BuildQuery(folder, days))

    If lo.ListRows.Count = 0 Then
        wb.Close SaveChanges:=False
        Application.DisplayAlerts = True: Application.ScreenUpdating = True
        MsgBox "Combined file saved:" & vbLf & combinedPath & vbLf & vbLf & _
               "No rows matched the chosen dates.", vbExclamation
        Exit Sub
    End If

    Dim dateIdx As Long, dateHeader As String, r As Long, found As Object, missing As String
    dateIdx = FindDateColumn(lo.HeaderRowRange)
    dateHeader = CStr(lo.HeaderRowRange.Cells(1, dateIdx).Value)
    lo.ListColumns(dateIdx).DataBodyRange.NumberFormat = DATE_DISPLAY_FORMAT

    Set found = CreateObject("Scripting.Dictionary")
    For r = 1 To lo.ListRows.Count
        found(CLng(lo.DataBodyRange(r, dateIdx).Value)) = True
    Next r
    For Each k In days.Keys
        If Not found.Exists(k) Then missing = missing & Format$(CDate(k), DATE_DISPLAY_FORMAT) & "  "
    Next k

    ' ---- Values-only copy for the Final file -------------------------------
    Dim wbF As Workbook, wsF As Worksheet, rng As Range
    Dim filteredRows As Long, finalRows As Long
    filteredRows = lo.ListRows.Count
    Set wbF = Workbooks.Add(xlWBATWorksheet)
    Set wsF = wbF.Worksheets(1)
    wsF.Name = "Final"
    lo.Range.Copy
    wsF.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    wb.Close SaveChanges:=False
    Set rng = wsF.Range("A1").CurrentRegion
    wsF.Columns(dateIdx).NumberFormat = DATE_DISPLAY_FORMAT

    ' Sort (Excel's own sort: case-insensitive, blanks last)
    With wsF.Sort
        .SortFields.Clear
        For i = LBound(sortCols) To UBound(sortCols)
            idx = FindHeader(rng.Rows(1), CStr(sortCols(i)))
            .SortFields.Add Key:=rng.Columns(idx).Offset(1).Resize(rng.Rows.Count - 1), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        Next i
        .SetRange rng
        .Header = xlYes
        .MatchCase = False
        .Apply
    End With

    ' Remove duplicates on Brand Name only (keeps the first row after sorting)
    idx = FindHeader(rng.Rows(1), BRAND_COLUMN)
    rng.RemoveDuplicates Columns:=idx, Header:=xlYes
    finalRows = wsF.Range("A1").CurrentRegion.Rows.Count - 1

    ' Delete the DROP_COLUMNS
    Dim dropName As Variant, deleted As String
    If Trim$(DROP_COLUMNS) <> "" Then
        For Each dropName In Split(DROP_COLUMNS, ",")
            If Trim$(dropName) <> "" Then
                idx = FindHeader(wsF.Range("A1").CurrentRegion.Rows(1), Trim$(dropName))
                If idx = 0 Then
                    warnings = warnings & "Column to delete not found: " & Trim$(dropName) & vbLf
                Else
                    wsF.Columns(idx).Delete
                    deleted = deleted & IIf(deleted = "", "", ", ") & Trim$(dropName)
                End If
            End If
        Next dropName
    End If

    ' Remove "-" from phone numbers and "^" from every cell
    CleanCharacters wsF.Range("A1").CurrentRegion, _
                    FindHeader(wsF.Range("A1").CurrentRegion.Rows(1), PHONE_COLUMN), _
                    FindHeader(wsF.Range("A1").CurrentRegion.Rows(1), dateHeader)

    ' Save the Final file (left open so you can check it)
    Dim finalPath As String
    wsF.Columns.AutoFit
    finalPath = outDir & "\Final_" & tag & ".xlsx"
    wbF.SaveAs Filename:=finalPath, FileFormat:=xlOpenXMLWorkbook

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Done." & vbLf & vbLf & _
           "Combined (chosen dates, as-is): " & combinedRows & " rows" & vbLf & combinedPath & vbLf & vbLf & _
           "Rows on chosen dates: " & filteredRows & vbLf & _
           "Final (after removing duplicate brands): " & finalRows & " rows" & vbLf & finalPath & _
           IIf(deleted <> "", vbLf & vbLf & "Deleted columns: " & deleted, "") & _
           IIf(missing <> "", vbLf & vbLf & "No rows found for: " & missing, "") & _
           IIf(warnings <> "", vbLf & vbLf & warnings, ""), vbInformation
    Exit Sub

Fail:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Add a Power Query to a workbook and load it into a table on its first sheet
'------------------------------------------------------------------------------
Private Function LoadQuery(ByVal wb As Workbook, ByVal qName As String, ByVal formula As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    Set ws = wb.Worksheets(1)
    wb.Queries.Add Name:=qName, Formula:=formula
    Set lo = ws.ListObjects.Add(SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties=""""", _
        Destination:=ws.Range("$A$1"))
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & qName & "]")
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = True
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    lo.Name = qName
    Set LoadQuery = lo
End Function

'------------------------------------------------------------------------------
' Power Query (M) - backticks are swapped for double quotes at the end
'------------------------------------------------------------------------------
' Copy 1: first sheet of every file appended as-is, only rows on the chosen dates
Private Function BuildRawQuery(ByVal folder As String, ByVal days As Object) As String
    Dim m As String
    m = "let" & vbLf
    m = m & "    Source = Folder.Files(`" & folder & "`)," & vbLf
    m = m & "    ExcelFiles = Table.SelectRows(Source, each List.Contains({`.xlsx`, `.xlsm`, `.xls`}, Text.Lower([Extension])) and not Text.StartsWith([Name], `~$`))," & vbLf
    m = m & "    FirstSheet = (bin as binary) as table => let wb = Excel.Workbook(bin, true), sh = Table.SelectRows(wb, each [Kind] = `Sheet`) in sh{0}[Data]," & vbLf
    m = m & "    Combined = Table.Combine(List.Transform(ExcelFiles[Content], FirstSheet))," & vbLf
    m = m & MDateSteps("Combined", days)
    ' the date is only read to decide which rows to keep; no value is changed
    m = m & "    Filtered = Table.SelectRows(Combined, each List.Contains(Days, ToDate(Record.Field(_, DateCol))))" & vbLf
    m = m & "in" & vbLf & "    Filtered"
    BuildRawQuery = Replace(m, "`", Chr(34))
End Function

' Shared M steps: find the date column in <tableStep>, read any date value, list of chosen days
Private Function MDateSteps(ByVal tableStep As String, ByVal days As Object) As String
    Dim m As String, culture As String, dateMatch As String, dayList As String, k As Variant

    culture = IIf(DAY_FIRST, "en-GB", "en-US")
    If DATE_COLUMN = "" Then
        dateMatch = "Text.Contains(Text.Lower(_), `date`)"
    Else
        dateMatch = "Norm(_) = `" & NormHeader(DATE_COLUMN) & "`"
    End If
    For Each k In days.Keys
        dayList = dayList & IIf(dayList = "", "", ", ") & _
            "#date(" & Year(CDate(k)) & "," & Month(CDate(k)) & "," & Day(CDate(k)) & ")"
    Next k

    m = "    Norm = (s) => Text.Lower(Text.Combine(List.Select(Text.Split(Text.Trim(Text.From(s)), ` `), each _ <> ``), ` `))," & vbLf
    m = m & "    DateCol = let c = List.First(List.Select(Table.ColumnNames(" & tableStep & "), each " & dateMatch & ")) in if c = null then error `No date column found` else c," & vbLf
    m = m & "    ToDate = (v) => if v = null then null else if Value.Is(v, type datetime) then DateTime.Date(v) else if Value.Is(v, type date) then v" & _
            " else if Value.Is(v, type number) then Date.From(v)" & _
            " else let t = Text.Trim(Text.From(v)) in try Date.FromText(t, `" & culture & "`) otherwise try DateTime.Date(DateTime.FromText(t, `" & culture & "`)) otherwise null," & vbLf
    m = m & "    Days = {" & dayList & "}," & vbLf
    MDateSteps = m
End Function

' Copy 2: tidy headers, keep text as text, keep only the chosen dates
Private Function BuildQuery(ByVal folder As String, ByVal days As Object) As String
    Dim m As String

    m = "let" & vbLf
    m = m & "    Source = Folder.Files(`" & folder & "`)," & vbLf
    m = m & "    ExcelFiles = Table.SelectRows(Source, each List.Contains({`.xlsx`, `.xlsm`, `.xls`}, Text.Lower([Extension])) and not Text.StartsWith([Name], `~$`))," & vbLf
    ' tidy header spacing so "Brand  Name" matches "Brand Name"
    m = m & "    Tidy = (t as table) as table => Table.TransformColumnNames(t, each Text.Combine(List.Select(Text.Split(Text.Trim(Text.From(_)), ` `), each _ <> ``), ` `))," & vbLf
    m = m & "    FirstSheet = (bin as binary) as table => let wb = Excel.Workbook(bin, true), sh = Table.SelectRows(wb, each [Kind] = `Sheet`) in Tidy(sh{0}[Data])," & vbLf
    m = m & "    Combined = Table.Combine(List.Transform(ExcelFiles[Content], FirstSheet))," & vbLf
    m = m & "    NonBlank = Table.SelectRows(Combined, each List.NonNullCount(List.ReplaceValue(Record.FieldValues(_), ``, null, Replacer.ReplaceValue)) > 0)," & vbLf
    m = m & MDateSteps("NonBlank", days)
    ' everything except the date column becomes text (keeps phone leading zeros, no 9.87E+09)
    m = m & "    ToText = (v) => if v = null then null else if Value.Is(v, type number) and Number.Round(v) = v then Number.ToText(v, `0`) else Text.From(v)," & vbLf
    m = m & "    AsText = Table.TransformColumns(NonBlank, List.Transform(List.RemoveItems(Table.ColumnNames(NonBlank), {DateCol}), (c) => {c, ToText, type text}))," & vbLf
    m = m & "    AsDate = Table.TransformColumns(AsText, {{DateCol, ToDate, type date}})," & vbLf
    m = m & "    Filtered = Table.SelectRows(AsDate, each List.Contains(Days, Record.Field(_, DateCol)))" & vbLf
    m = m & "in" & vbLf & "    Filtered"

    BuildQuery = Replace(m, "`", Chr(34))
End Function

'------------------------------------------------------------------------------
' Remove PHONE_REMOVE_CHARS from the phone column and REMOVE_CHARS_EVERYWHERE
' from every cell. Works on an array and writes back as text, so phone numbers
' like 0123456789 keep their leading zero.
'------------------------------------------------------------------------------
Private Sub CleanCharacters(ByVal rng As Range, ByVal phoneIdx As Long, ByVal dateIdx As Long)
    Dim arr As Variant, r As Long, c As Long, i As Long, s As String
    arr = rng.Value
    For r = 2 To UBound(arr, 1)
        For c = 1 To UBound(arr, 2)
            If VarType(arr(r, c)) = vbString Then
                s = arr(r, c)
                If c = phoneIdx Then
                    For i = 1 To Len(PHONE_REMOVE_CHARS): s = Replace(s, Mid$(PHONE_REMOVE_CHARS, i, 1), ""): Next i
                End If
                For i = 1 To Len(REMOVE_CHARS_EVERYWHERE): s = Replace(s, Mid$(REMOVE_CHARS_EVERYWHERE, i, 1), ""): Next i
                arr(r, c) = s
            End If
        Next c
    Next r
    For c = 1 To UBound(arr, 2)
        If c <> dateIdx Then rng.Columns(c).NumberFormat = "@"
    Next c
    rng.Value = arr
End Sub

'------------------------------------------------------------------------------
' Helpers
'------------------------------------------------------------------------------
Private Function PickFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select the folder containing the Excel files"
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

Private Sub EnsureFolder(ByVal path As String)
    Dim parent As String
    If Right$(path, 1) = "\" Then path = Left$(path, Len(path) - 1)
    If path = "" Or Dir(path, vbDirectory) <> "" Then Exit Sub
    parent = Left$(path, InStrRev(path, "\") - 1)
    If parent <> "" And Right$(parent, 1) <> ":" Then EnsureFolder parent
    MkDir path
End Sub

' "10-05-2026, 10-07-2026 to 10-09-2026" -> Dictionary of CLng(date)
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
    Dim p() As String, y As Long, mo As Long, dy As Long, dt As Date
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
    dt = DateSerial(y, mo, dy)
    If Year(dt) <> y Or Month(dt) <> mo Or Day(dt) <> dy Then GoTo Bad   ' rejects 02-30-2026 etc.
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

Private Function HeaderList(ByVal headerRow As Range) As String
    Dim c As Range
    For Each c In headerRow.Cells
        HeaderList = HeaderList & IIf(HeaderList = "", "", ", ") & c.Value
    Next c
End Function

Private Function FindDateColumn(ByVal headerRow As Range) As Long
    Dim c As Range
    If DATE_COLUMN <> "" Then FindDateColumn = FindHeader(headerRow, DATE_COLUMN): Exit Function
    For Each c In headerRow.Cells
        If InStr(1, CStr(c.Value), "date", vbTextCompare) > 0 Then FindDateColumn = c.Column - headerRow.Column + 1: Exit Function
    Next c
End Function
