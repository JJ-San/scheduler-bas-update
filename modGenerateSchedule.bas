Attribute VB_Name = "modGenerateSchedule"
'=============================================================================
' INVOICE SCHEDULE GENERATOR v2
' Smart page breaks | Multi-contractor | Unpaid highlighting | Dynamic summary
'=============================================================================

Option Explicit

' Read by modUpdater.CheckForUpdates to decide whether a different release
' exists. EQUALITY compare (any mismatch prompts); tmp/release.py separately
' enforces forward-only numbering. Format "MAJOR.MINOR": each client update is a
' whole number (Update 4 = 4.0), fixes after a release take the decimal (4.1).
' Trial and test builds carry a label ("4.0-trial") that release.py refuses to
' publish. Never reuse a burned number (see CLAUDE.md) - equality would skip it.
Public Const MODULE_VERSION As String = "4.0"

' Input-sheet layout generation this module expects. EnsureSchema upgrades
' older workbooks in place (idempotent, additive-only) on the first Generate
' after a module update; the workbook's current generation is stamped in the
' hidden defined name "SchemaVersion". Bump ONLY when the entry-sheet schema
' changes, and add a matching migration step to EnsureSchema.
' 6 (Update 4): RETENTION CAP ($) on COMMITTED col H (migration 6).
Public Const SCHEMA_VERSION As Long = 6

' Cell on REPORT_SETTINGS where the human-visible version label lives.
' WriteVersionLabel writes "Workbook Version: X.Y" here every time GenerateSchedule
' runs, so the label always reflects the loaded macro. Move the cell? Update this constant.
Public Const VERSION_CELL As String = "D10"

Private Const FONT_NAME As String = "Aptos"
Private Const FONT_SIZE As Double = 9
Private Const GST_RATE As Double = 1.15
Private Const MONEY_FMT As String = "$#,##0.00;($#,##0.00)"
Private Const MONEY_FMT_WHOLE As String = "$#,##0;($#,##0)"
Private Const DATE_FMT As String = "d-mmm-yy"
Private Const PCT_FMT As String = "0%"
Private Const ROWS_PER_PAGE As Long = 33
Private Const FEE_FILL As Long = 15263976    ' #E8E8E8 brand light grey (unified)
Private Const UNPAID_FILL As Long = 15983807  ' #BFE4F3 brand aqua tint (attention / unpaid)
Private Const HEADER_FILL As Long = 15263976  ' #E8E8E8 brand light grey
Private Const CASHFLOW_FILL As Long = 9359529   ' #A9D08E Green Accent 6 Lighter 40% (semantic: money in)
Private Const BRAND_NAVY As Long = 5850151       ' #274459 primary brand (text, borders)
Private Const BRAND_AQUA As Long = 14854951      ' #27ABE2 brand flourish (borders only, never text)

' ---------------------------------------------------------------------------
' Module-level row trackers (used by Invoice Summary / YTD)
' ---------------------------------------------------------------------------
Private mFirstFeeDataRow As Long   ' first invoice OR committed-fee row across Consultant+Council+MilnePM
Private mLastFeeDataRow As Long    ' last  invoice OR committed-fee row across Consultant+Council+MilnePM
Private mTotalFeeCommittedRow As Long ' TOTAL INVOICES PROCESSED row
Private mMiscStartRow As Long
Private mMiscEndRow As Long
Private mMiscTotalFeeRow As Long
Private mCsStartRow As Long
Private mCsEndRow As Long
Private mCsTotalFeeRow As Long
Private mCfStartRow As Long        ' Cash Flow In first data row
Private mCfEndRow As Long          ' Cash Flow In last data row
Private mCfTotalFeeRow As Long

' Contractor tracking (arrays for multiple contractors)
Private mContrStartRows() As Long
Private mContrEndRows() As Long
Private mContrTotalFeeRows() As Long
Private mContrCount As Long

' Company -> discipline prefix (COMMITTED col B), rebuilt on every run.
' Companies without an entry render without a prefix.
Private mCompDisc As Object

' ============================= PUBLIC MACROS ================================

' Writes "Workbook Version: X.Y" to the version label cell on REPORT_SETTINGS.
' Called at the top of GenerateSchedule (so every regenerate refreshes the label)
' and by modUpdater after a successful update (so the label refreshes immediately).
Public Sub WriteVersionLabel()
    On Error Resume Next
    ThisWorkbook.Worksheets("REPORT_SETTINGS").Range(VERSION_CELL).Value = _
        "Workbook Version: " & MODULE_VERSION
End Sub

'=============================================================================
' SCHEMA MIGRATIONS
' The .bas swap is the only remote delivery channel (Check for Updates), so
' workbook-side changes ride inside the module: EnsureSchema compares the
' stamped workbook generation against SCHEMA_VERSION and, once per update,
' applies additive migrations (new columns, validations, sheet-module guard).
' Every step is check-before-do, so a partial or repeated run converges.
' NEVER destructive: migrations may add, move or restyle — never delete data.
'=============================================================================

Private Sub EnsureSchema()
    On Error GoTo Bail
    If ReadSchemaVersion() >= SCHEMA_VERSION Then Exit Sub

    Dim wsCE As Worksheet, wsCF As Worksheet
    Set wsCE = ThisWorkbook.Sheets("CLAIM_ENTRY")
    Set wsCF = ThisWorkbook.Sheets("COMMITTED")

    ' --- Migration 3 (Update 3): release column + discipline column ---
    ' Header compares are whitespace-tolerant (NormHdr): users legitimately
    ' Alt+Enter headers to wrap them, and that must not read as "missing".
    Dim ceHdr As String: ceHdr = NormHdr(wsCE.Range("G1").Value)
    Dim needCE As Boolean, needCF As Boolean
    needCE = (ceHdr <> "RETENTION RELEASE ($ OR %)" And ceHdr <> "RETENTION RELEASE $")
    needCF = (NormHdr(wsCF.Range("B1").Value) <> "DISCIPLINE")
    ' Migration 6: a schema-5 workbook has needCE = needCF = False, so the cap
    ' column must take part in the backup decision itself
    Dim needCap As Boolean
    needCap = (NormHdr(wsCF.Range("H1").Value) <> "RETENTION CAP ($)")

    ' Safety copy before the first structural touch (freshly built workbooks
    ' skip straight to the stamp). Backup failure never blocks the migration.
    If needCE Or needCF Or needCap Then
        On Error Resume Next
        ThisWorkbook.SaveCopyAs Environ$("TEMP") & "\" & _
            Replace(ThisWorkbook.Name, ".xlsm", "") & "_preSchema" & _
            SCHEMA_VERSION & "_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
        On Error GoTo Bail
    End If

    If needCE Then
        ' New col G "RETENTION RELEASE $" on CLAIM_ENTRY (appended — nothing shifts)
        wsCE.Range("F1").Copy
        wsCE.Range("G1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsCE.Range("G1").Value = "RETENTION RELEASE ($ or %)"
        wsCE.Columns("G").ColumnWidth = 16
        ' Fold the new column into the sheet's entry Table if one exists
        On Error Resume Next
        If wsCE.ListObjects.Count > 0 Then
            Dim loCE As Object: Set loCE = wsCE.ListObjects(1)
            If loCE.Range.Column + loCE.Range.Columns.Count - 1 < 7 Then
                loCE.Resize wsCE.Range(loCE.Range.Cells(1, 1), _
                    wsCE.Cells(loCE.Range.Row + loCE.Range.Rows.Count - 1, 7))
            End If
        End If
        On Error GoTo Bail
    End If

    If needCF Then
        ' New col B "DISCIPLINE" on COMMITTED. A whole-column insert shifts
        ' data, formats, validations, the entry Table AND the below-table
        ' plumbing together — RETENTIONS (with its dropdown) lands on G intact.
        wsCF.Columns("B").Insert
        wsCF.Range("A1").Copy
        wsCF.Range("B1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsCF.Range("B1").Value = "DISCIPLINE"
        wsCF.Columns("B").ColumnWidth = 18
    End If

    ' --- Migration 4 (3.1 repair): a column insert INHERITS the left
    '     neighbour's data validation, so 3.0's migration gave DISCIPLINE the
    '     TYPE dropdown. Strip any validation from col B — discipline is free
    '     text. Runs for fresh migrations (right after the insert above) AND
    '     heals workbooks already migrated by 3.0. Idempotent. ---
    If UCase(Trim(CStr(wsCF.Range("B1").Value))) = "DISCIPLINE" Then
        On Error Resume Next
        wsCF.Columns("B").Validation.Delete
        On Error GoTo Bail
    End If

    ' --- Migration 5 (3.2): the release column now accepts $ or % — rename
    '     the header on workbooks migrated before the rename. ---
    If NormHdr(wsCE.Range("G1").Value) = "RETENTION RELEASE $" Then
        wsCE.Range("G1").Value = "RETENTION RELEASE ($ or %)"
    End If

    ' --- Migration 6 (Update 4): per-contractor retention cap, COMMITTED col H ---
    If needCap Then
        ' The below-table note ("Drag blue corner down" / "to add more rows to
        ' the table") lives in col H, sometimes on rows inside the data range.
        ' Move it to col I on the same row first (only when I is empty) so it
        ' never sits in the cap column.
        Dim nRow As Long, nTxt As String
        For nRow = 2 To 600
            nTxt = ""
            If Not IsError(wsCF.Cells(nRow, 8).Value) Then nTxt = LCase(Trim(CStr(wsCF.Cells(nRow, 8).Value)))
            If Left$(nTxt, 16) = "drag blue corner" Or Left$(nTxt, 16) = "to add more rows" Then
                If IsEmpty(wsCF.Cells(nRow, 9).Value) Then
                    wsCF.Cells(nRow, 8).Cut Destination:=wsCF.Cells(nRow, 9)
                End If
            End If
        Next nRow

        ' Header styled like RETENTIONS. The value goes in BEFORE the Table
        ' resize, otherwise Excel names the new column "Column1". If the Table
        ' already grew into H (typing in H auto-extends it), this renames it.
        wsCF.Range("G1").Copy
        wsCF.Range("H1").PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wsCF.Range("H1").Value = "RETENTION CAP ($)"
        wsCF.Columns("H").ColumnWidth = 16

        ' Fold the new column into the entry Table if it stops short of H
        On Error Resume Next
        If wsCF.ListObjects.Count > 0 Then
            Dim loCF As Object: Set loCF = wsCF.ListObjects(1)
            If loCF.Range.Column + loCF.Range.Columns.Count - 1 < 8 Then
                loCF.Resize wsCF.Range(loCF.Range.Cells(1, 1), _
                    wsCF.Cells(loCF.Range.Row + loCF.Range.Rows.Count - 1, 8))
            End If
        End If
        On Error GoTo Bail

        ' Dollars only: Delete first (Add raises 1004 over existing
        ' validation). A typed 5% arrives as 0.05 and is rejected.
        With wsCF.Range("H2:H502").Validation
            .Delete
            .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlGreaterEqual, Formula1:="1"
            .IgnoreBlank = True
            .ErrorTitle = "Retention cap"
            .ErrorMessage = "Enter the cap in dollars (for example 30000), or leave blank for no cap."
        End With
        wsCF.Range("H2:H502").NumberFormatLocal = MONEY_FMT_WHOLE
    End If

    ' CommittedCompanies (company dropdowns on CLAIM_ENTRY col A and
    ' FEES_ENTRY col B) must point at COMMITTED col C. Outside the needCap
    ' guard so it also heals already-migrated workbooks. Check-before-do:
    ' any end row on col C is left alone (row inserts legitimately grow it
    ' past 490); only a reference off col C (the stale $B$23:$B$490 dev
    ' artefact) is repointed.
    On Error Resume Next
    Dim ccRef As String: ccRef = ""
    ccRef = Replace(ThisWorkbook.Names("CommittedCompanies").RefersTo, "'", "")
    If Len(ccRef) > 0 Then
        If Not (ccRef Like "=COMMITTED!$C$2:$C$*") Then
            ThisWorkbook.Names("CommittedCompanies").RefersTo = "=COMMITTED!$C$2:$C$490"
        End If
    End If
    On Error GoTo Bail

    ' The retention guard (COMMITTED sheet module) hardcodes its column, so a
    ' column shift strands it, so refresh it to cols G:H (retention type and
    ' cap). Needs VBA project access, which Check for Updates users already have. If unavailable, leave the
    ' schema unstamped so the whole (idempotent) migration retries next run.
    If Not InstallRetentionGuard() Then
        LogMacroNote "EnsureSchema: retention guard deferred - VBA project access unavailable"
        Exit Sub
    End If

    WriteSchemaVersion SCHEMA_VERSION
    Exit Sub
Bail:
    ' Never let a migration hiccup block schedule generation — the read-side
    ' legacy shim keeps unmigrated workbooks fully functional.
    LogMacroNote "EnsureSchema: bailed - Err " & Err.Number & " " & Err.Description
End Sub

' Header text normalised for comparison: uppercased, line breaks flattened,
' runs of whitespace collapsed. Users Alt+Enter headers for display; that
' formatting must never be mistaken for a schema difference.
Private Function NormHdr(v As Variant) As String
    Dim s As String
    s = UCase(Trim(CStr(v)))
    s = Replace(Replace(s, vbCr, " "), vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    NormHdr = s
End Function

' COMMITTED TYPE -> canonical spelling of one of the six known types (Trim,
' whitespace runs collapsed, case-insensitive). Anything else, blank included,
' returns "" so the auto-fill treats it as undetected instead of an explicit
' TYPE that no section would ever render (item 15: explicit TYPE wins).
Private Function CanonType(v As Variant) As String
    Dim s As String, t As Variant
    s = NormHdr(v)
    For Each t In Array("Consultant", "Council", "Misc Fees", "Client Supply", "Cash Flow In", "Contractor")
        If s = UCase(CStr(t)) Then CanonType = CStr(t): Exit Function
    Next t
    CanonType = ""
End Function

Private Function ReadSchemaVersion() As Long
    On Error Resume Next
    Dim s As String
    s = ThisWorkbook.Names("SchemaVersion").RefersTo
    ReadSchemaVersion = CLng(Val(Replace(Replace(s, "=", ""), """", "")))
End Function

Private Sub WriteSchemaVersion(v As Long)
    On Error Resume Next
    ThisWorkbook.Names("SchemaVersion").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="SchemaVersion", RefersTo:="=" & v, Visible:=False
End Sub

' Diagnostic breadcrumb, same channel the PageSetup guard uses.
Private Sub LogMacroNote(msg As String)
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="LastMacroError", _
        RefersTo:="=""" & Replace(msg, """", """""") & """", Visible:=False
End Sub

' (Re)installs the COMMITTED Worksheet_Change guard watching cols G:H.
' Mirrors tmp/install_retention_guard.py — keep the two in sync.
Private Function InstallRetentionGuard() As Boolean
    On Error GoTo Fail
    Dim cm As Object
    Set cm = ThisWorkbook.VBProject.VBComponents( _
        ThisWorkbook.Sheets("COMMITTED").CodeName).CodeModule
    Dim startL As Long, cnt As Long
    On Error Resume Next
    startL = cm.ProcStartLine("Worksheet_Change", 0)
    cnt = cm.ProcCountLines("Worksheet_Change", 0)
    On Error GoTo Fail
    If startL > 0 Then cm.DeleteLines startL, cnt
    cm.AddFromString GuardCode()
    InstallRetentionGuard = True
    Exit Function
Fail:
    InstallRetentionGuard = False
End Function

Private Function GuardCode() As String
    Dim c As String
    c = "Private Sub Worksheet_Change(ByVal Target As Range)" & vbCrLf
    c = c & "    ' Enforce: cols G:H (RETENTIONS, RETENTION CAP) only accept values on Contractor rows." & vbCrLf
    c = c & "    ' Installed by modGenerateSchedule.EnsureSchema - do not edit here." & vbCrLf
    c = c & "    Dim isect As Range" & vbCrLf
    c = c & "    Set isect = Application.Intersect(Target, Me.Range(""G2:H502""))" & vbCrLf
    c = c & "    If isect Is Nothing Then Exit Sub" & vbCrLf
    c = c & "    Dim cell As Range, bad As Range" & vbCrLf
    c = c & "    For Each cell In isect" & vbCrLf
    c = c & "        If Len(Trim(CStr(cell.Value))) > 0 Then" & vbCrLf
    c = c & "            If Trim(CStr(Me.Cells(cell.Row, 1).Value)) <> ""Contractor"" Then" & vbCrLf
    c = c & "                If bad Is Nothing Then Set bad = cell Else Set bad = Application.Union(bad, cell)" & vbCrLf
    c = c & "            End If" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    Next cell" & vbCrLf
    c = c & "    If Not bad Is Nothing Then" & vbCrLf
    c = c & "        Application.EnableEvents = False" & vbCrLf
    c = c & "        bad.ClearContents" & vbCrLf
    c = c & "        Application.EnableEvents = True" & vbCrLf
    c = c & "        If Application.Interactive Then" & vbCrLf
    c = c & "            MsgBox ""Retention settings only apply to Contractor rows. Value cleared."", _" & vbCrLf
    c = c & "                   vbExclamation, ""Retention not allowed""" & vbCrLf
    c = c & "        End If" & vbCrLf
    c = c & "    End If" & vbCrLf
    c = c & "End Sub" & vbCrLf
    GuardCode = c
End Function

Public Sub GenerateSchedule()

    WriteVersionLabel
    EnsureSchema
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    ' Section + row trackers are declared early so the error handler always
    ' sees a meaningful value even if a failure fires before any section
    ' updates run (e.g. inside the sheet-clear or PAYMENTS-protect path).
    Dim curRow As Long: curRow = 0
    Dim curSection As String: curSection = "Init"
    On Error GoTo ErrHandler

    ' --- Sheet references ---
    Dim wsDE As Worksheet, wsCE As Worksheet, wsPay As Worksheet, wsCF As Worksheet, wsRS As Worksheet
    Set wsDE = ThisWorkbook.Sheets("FEES_ENTRY")
    Set wsCE = ThisWorkbook.Sheets("CLAIM_ENTRY")
    Set wsCF = ThisWorkbook.Sheets("COMMITTED")
    Set wsPay = ThisWorkbook.Sheets("PAYMENTS")
    Set wsRS = ThisWorkbook.Sheets("REPORT_SETTINGS")

    ' --- Clear PAYMENTS (preserve header/footer images) ---
    wsPay.Cells.Clear
    wsPay.Cells.UnMerge
    wsPay.ResetAllPageBreaks
    ' Defensive: unhide any rows on PAYMENTS. The sheet is fully regenerated,
    ' so a hidden row is always stale state from a prior run / manual edit.
    wsPay.Rows.Hidden = False
    ' Clear any stale PrintArea -- if it's set to a small range (e.g. left over
    ' from an earlier failed run or a manual user action), HPageBreaks.Add later
    ' can throw Err 1004 when the target row falls on the PrintArea boundary.
    wsPay.PageSetup.PrintArea = ""

    With wsPay.Cells.Font
        .Name = FONT_NAME
        .Size = FONT_SIZE
    End With
    wsPay.Cells.RowHeight = 12.75

    ' Column widths
    wsPay.Columns("A").ColumnWidth = 43.0
    wsPay.Columns("B").ColumnWidth = 11.73
    wsPay.Columns("C").ColumnWidth = 11.73
    wsPay.Columns("D").ColumnWidth = 11.73
    wsPay.Columns("E").ColumnWidth = 11.73
    wsPay.Columns("F").ColumnWidth = 15    ' "LESS PREVIOUS" stays on one header line at 80-90% zoom
    wsPay.Columns("G").ColumnWidth = 11.73
    wsPay.Columns("H").ColumnWidth = 3.34
    wsPay.Columns("I").ColumnWidth = 11    ' "COMMITTED" stays on one header line at 80-90% zoom
    wsPay.Columns("J").ColumnWidth = 18
    wsPay.Columns("K").ColumnWidth = 8

    ' --- Read DATA_ENTRY ---
    Dim lastRowDE As Long
    lastRowDE = wsDE.Cells(wsDE.Rows.Count, "A").End(xlUp).Row
    Dim dataArr() As Variant
    Dim hasData As Boolean: hasData = False
    If lastRowDE >= 2 Then
        dataArr = wsDE.Range("A2:G" & lastRowDE).Value
        hasData = True
    Else
        ' Empty sheet: create a 1-row dummy array (no type will match empty strings)
        ReDim dataArr(1 To 1, 1 To 7)
        Dim dc As Long
        For dc = 1 To 7: dataArr(1, dc) = "": Next dc
    End If

    ' --- Read CLAIM_ENTRY (7 columns: Contractor, Date, Payment No, Value, Paid, Comments,
    '     Retention Release $). Col G is new in Update 3 — on older workbooks without the
    '     header the range read returns Empty for it, so no release rows are detected and
    '     behaviour is identical to Update 2. ---
    Dim lastRowCE As Long
    lastRowCE = wsCE.Cells(wsCE.Rows.Count, "A").End(xlUp).Row
    Dim claimArr() As Variant
    Dim hasClaims As Boolean: hasClaims = False
    If lastRowCE >= 2 Then
        claimArr = wsCE.Range("A2:G" & lastRowCE).Value
        hasClaims = True
    End If

    ' --- Read COMMITTED (8 columns: Type, Discipline, Company, FeeItem, Amount,
    '     DateAppointed, Retention, RetentionCap). Retention Cap (col H) is new
    '     in Update 4 (schema 6). Discipline (col B) is new in Update 3;
    '     pre-Update-3 workbooks have 6 columns in the old order — see the
    '     legacy-layout shim right below the read. ---
    Dim lastRowCF As Long
    lastRowCF = wsCF.Cells(wsCF.Rows.Count, "A").End(xlUp).Row
    Dim feesArr() As Variant
    Dim hasFees As Boolean: hasFees = False
    If lastRowCF >= 2 Then
        feesArr = wsCF.Range("A2:H" & lastRowCF).Value
        hasFees = True
    End If

    ' Legacy-layout shim: pre-Update-3 workbooks have no DISCIPLINE column
    ' (B1 = "COMPANY"; 6 columns Type|Company|FeeItem|Amount|Date|Retention).
    ' Rebuild the array into the Update-3 shape (Discipline at col 2) so the
    ' rest of the macro only ever deals with one layout.
    If hasFees Then
        If UCase(Trim(CStr(wsCF.Range("B1").Value))) <> "DISCIPLINE" Then
            Dim legacyArr() As Variant: legacyArr = feesArr
            ReDim feesArr(1 To UBound(legacyArr, 1), 1 To 8)
            Dim lgR As Long
            For lgR = 1 To UBound(legacyArr, 1)
                feesArr(lgR, 1) = legacyArr(lgR, 1)   ' Type
                feesArr(lgR, 2) = Empty               ' Discipline (absent on legacy)
                feesArr(lgR, 3) = legacyArr(lgR, 2)   ' Company
                feesArr(lgR, 4) = legacyArr(lgR, 3)   ' Fee Item
                feesArr(lgR, 5) = legacyArr(lgR, 4)   ' Amount
                feesArr(lgR, 6) = legacyArr(lgR, 5)   ' Date Appointed
                feesArr(lgR, 7) = legacyArr(lgR, 6)   ' Retention
                feesArr(lgR, 8) = Empty               ' Retention Cap (absent on legacy)
            Next lgR
        End If
    End If

    ' Canonicalise every COMMITTED TYPE ("misc fees " -> "Misc Fees"); an
    ' unrecognised value becomes "" and is auto-detected below
    Dim ftIdx As Long
    If hasFees Then
        For ftIdx = 1 To UBound(feesArr, 1)
            feesArr(ftIdx, 1) = CanonType(feesArr(ftIdx, 1))
        Next ftIdx
    End If

    ' --- Auto-fill BLANK fee types from a company->type lookup ---
    ' Item 15 (Update 4): this used to OVERWRITE every COMMITTED row's TYPE
    ' with the company's first FEES_ENTRY type, so a Misc Fees row for a
    ' consultant rolled into its committed block, and a contractor with a
    ' Misc Fees invoice lost every row to Misc Fees. Now an explicit TYPE
    ' always wins; only blank rows are filled, CLAIM_ENTRY (Contractor) first,
    ' then FEES_ENTRY, first seen wins.
    If (hasData Or hasClaims) And hasFees Then
        Dim compTypeMap As Object: Set compTypeMap = CreateObject("Scripting.Dictionary")
        Dim ctmIdx As Long, cName As String
        If hasClaims Then
            For ctmIdx = 1 To UBound(claimArr, 1)
                cName = CStr(claimArr(ctmIdx, 1))
                If Len(cName) > 0 And Not compTypeMap.Exists(cName) Then
                    compTypeMap.Add cName, "Contractor"
                End If
            Next ctmIdx
        End If
        If hasData Then
            For ctmIdx = 1 To UBound(dataArr, 1)
                cName = CStr(dataArr(ctmIdx, 2))
                If Len(cName) > 0 And Not compTypeMap.Exists(cName) Then
                    compTypeMap.Add cName, CStr(dataArr(ctmIdx, 1))
                End If
            Next ctmIdx
        End If
        For ftIdx = 1 To UBound(feesArr, 1)
            Dim fComp As String: fComp = CStr(feesArr(ftIdx, 3))
            If Len(feesArr(ftIdx, 1)) = 0 And Len(fComp) > 0 And compTypeMap.Exists(fComp) Then
                feesArr(ftIdx, 1) = compTypeMap(fComp)
            End If
        Next ftIdx
    End If

    ' --- Sort data by date within each type+company ---
    SortDataByDate dataArr
    If hasClaims Then SortDataByDate claimArr
    If hasFees Then SortFeesByDate feesArr

    ' --- Read REPORT_SETTINGS ---
    Dim schoolName As String, projName As String, schedTitle As String, schedDate As String
    Dim attLine As String, payNote As String
    Dim feeBudget As Double, constrBudget As Double
    schoolName = CStr(wsRS.Range("B3").Value)
    projName = CStr(wsRS.Range("B4").Value)
    schedTitle = CStr(wsRS.Range("B5").Value)
    schedDate = CStr(wsRS.Range("B6").Value)
    attLine = CStr(wsRS.Range("B7").Value)
    payNote = CStr(wsRS.Range("B8").Value)
    If Not IsEmpty(wsRS.Range("B9").Value) Then feeBudget = CDbl(wsRS.Range("B9").Value)
    If Not IsEmpty(wsRS.Range("B10").Value) Then constrBudget = CDbl(wsRS.Range("B10").Value)
    Dim miscBudget As Double, csBudget As Double
    If Not IsEmpty(wsRS.Range("B11").Value) Then miscBudget = CDbl(wsRS.Range("B11").Value)
    If Not IsEmpty(wsRS.Range("B12").Value) Then csBudget = CDbl(wsRS.Range("B12").Value)

    ' --- Invoice Summary options ---
    Dim showCashFlow As Boolean: showCashFlow = True
    Dim showMisc As Boolean: showMisc = True
    Dim showClientSupply As Boolean: showClientSupply = True
    Dim showYTD As Boolean: showYTD = True
    If Not IsEmpty(wsRS.Range("B16").Value) Then
        showCashFlow = (UCase(Trim(CStr(wsRS.Range("B16").Value))) = "YES")
    End If
    If Not IsEmpty(wsRS.Range("B17").Value) Then
        showMisc = (UCase(Trim(CStr(wsRS.Range("B17").Value))) = "YES")
    End If
    If Not IsEmpty(wsRS.Range("B18").Value) Then
        showClientSupply = (UCase(Trim(CStr(wsRS.Range("B18").Value))) = "YES")
    End If
    If Not IsEmpty(wsRS.Range("B19").Value) Then
        showYTD = (UCase(Trim(CStr(wsRS.Range("B19").Value))) = "YES")
    End If

    ' --- Per-contractor retention map (from COMMITTED col G) ---
    ' and retention cap map (col H, Update 4): the first Contractor row per
    ' name holding a positive number. Text (a stray label, the relocated
    ' note) never reaches a formula.
    Dim ctrRet As Object: Set ctrRet = CreateObject("Scripting.Dictionary")
    Dim ctrCap As Object: Set ctrCap = CreateObject("Scripting.Dictionary")
    If hasFees Then
        Dim crIdx As Long
        For crIdx = 1 To UBound(feesArr, 1)
            If CStr(feesArr(crIdx, 1)) = "Contractor" Then
                Dim crName As String: crName = CStr(feesArr(crIdx, 3))
                If Len(crName) > 0 And Not ctrRet.Exists(crName) Then
                    Dim crRt As String: crRt = Trim(CStr(feesArr(crIdx, 7)))
                    If Len(crRt) = 0 Then crRt = "Major Works"
                    ctrRet.Add crName, crRt
                End If
                If Len(crName) > 0 And Not ctrCap.Exists(crName) Then
                    If Not IsError(feesArr(crIdx, 8)) And Not IsEmpty(feesArr(crIdx, 8)) Then
                        If IsNumeric(feesArr(crIdx, 8)) Then
                            If CDbl(feesArr(crIdx, 8)) > 0 Then
                                ctrCap.Add crName, CDbl(feesArr(crIdx, 8))
                            End If
                        End If
                    End If
                End If
            End If
        Next crIdx
    End If

    ' --- Per-company discipline map (COMMITTED col B) — first non-blank wins.
    '     Prefixes company rows on PAYMENTS: "QS | Brown Project Services". ---
    Set mCompDisc = CreateObject("Scripting.Dictionary")
    If hasFees Then
        Dim dcIdx As Long
        For dcIdx = 1 To UBound(feesArr, 1)
            Dim dcName As String: dcName = CStr(feesArr(dcIdx, 3))
            Dim dcVal As String: dcVal = Trim(CStr(feesArr(dcIdx, 2)))
            If Len(dcName) > 0 And Len(dcVal) > 0 And Not mCompDisc.Exists(dcName) Then
                mCompDisc.Add dcName, dcVal
            End If
        Next dcIdx
    End If

    ' --- Update header/footer text (leave Left alone for logo images) ---
    ' Wrapped in OERN because PageSetup writes are flaky on some Excel builds
    ' (printer driver state, locale issues, etc). Errors are logged to
    ' ThisWorkbook.Names("LastMacroError") so silent failures are visible to
    ' diagnostic tools instead of disappearing.
    On Error Resume Next
    Err.Clear
    Application.PrintCommunication = True
    wsPay.PageSetup.CenterHeader = "&""-,Bold""&12&K274459" & schoolName & Chr(10) & projName & Chr(10) & schedTitle & " " & schedDate
    wsPay.PageSetup.RightHeader = "&""-,Regular""" & attLine & Chr(10) & payNote
    wsPay.PageSetup.CenterFooter = "MilnePM Ltd"
    wsPay.PageSetup.RightFooter = "Page &P of &N"
    If Err.Number <> 0 Then
        ThisWorkbook.Names.Add Name:="LastMacroError", _
            RefersTo:="=""PageSetup OERN: Err " & Err.Number & " | " & Replace(Err.Description, """", """""") & """", _
            Visible:=False
        Err.Clear
    End If
    On Error GoTo ErrHandler

    ' --- Pre-count rows per type|company for look-ahead ---
    Dim compRowCounts As Object
    Set compRowCounts = CreateObject("Scripting.Dictionary")
    Dim i As Long, ck As String
    For i = 1 To UBound(dataArr, 1)
        ck = CStr(dataArr(i, 1)) & "|" & CStr(dataArr(i, 2))
        If compRowCounts.Exists(ck) Then
            compRowCounts(ck) = compRowCounts(ck) + 1
        Else
            compRowCounts.Add ck, 1
        End If
    Next i
    ' Count contractor rows from CLAIM_ENTRY
    If hasClaims Then
        For i = 1 To UBound(claimArr, 1)
            ck = "Contractor|" & CStr(claimArr(i, 1))
            If compRowCounts.Exists(ck) Then
                compRowCounts(ck) = compRowCounts(ck) + 1
            Else
                compRowCounts.Add ck, 1
            End If
        Next i
    End If

    ' Reset module-level trackers
    mFirstFeeDataRow = 0: mLastFeeDataRow = 0
    mTotalFeeCommittedRow = 0
    mMiscStartRow = 0: mMiscEndRow = 0: mMiscTotalFeeRow = 0
    mCsStartRow = 0: mCsEndRow = 0: mCsTotalFeeRow = 0
    mCfStartRow = 0: mCfEndRow = 0: mCfTotalFeeRow = 0
    mContrCount = 0

    curRow = 1
    Dim pageStartRow As Long: pageStartRow = 1

    ' =====================================================================
    curSection = "InvoiceHeader"
    ' SECTION 1: INVOICE HEADER
    ' =====================================================================
    WriteInvoiceHeader wsPay, curRow

    ' =====================================================================
    curSection = "Consultants"
    ' SECTION 2: CONSULTANTS A-Z (excluding MilnePM)
    ' =====================================================================
    Dim consultants As Variant
    consultants = GetAllCompaniesForType(dataArr, feesArr, hasFees, "Consultant", "MilnePM")
    Dim compIdx As Long
    If Not IsEmpty(consultants) Then
        For compIdx = LBound(consultants) To UBound(consultants)
            If Len(consultants(compIdx)) > 0 Then
                WriteStandardSection wsPay, dataArr, feesArr, hasFees, _
                    CStr(consultants(compIdx)), "Consultant", _
                    curRow, pageStartRow, compRowCounts, True
            End If
        Next compIdx
    End If

    ' =====================================================================
    curSection = "Council"
    ' SECTION 3: COUNCIL A-Z
    ' =====================================================================
    Dim councils As Variant
    councils = GetAllCompaniesForType(dataArr, feesArr, hasFees, "Council", "")
    If Not IsEmpty(councils) Then
        For compIdx = LBound(councils) To UBound(councils)
            If Len(councils(compIdx)) > 0 Then
                WriteStandardSection wsPay, dataArr, feesArr, hasFees, _
                    CStr(councils(compIdx)), "Council", _
                    curRow, pageStartRow, compRowCounts, True
            End If
        Next compIdx
    End If

    ' =====================================================================
    curSection = "MilnePM"
    ' SECTION 4: MilnePM (always last before totals)
    ' =====================================================================
    If HasCompanyData(dataArr, "Consultant", "MilnePM") Or _
       HasCompanyFees(feesArr, hasFees, "Consultant", "MilnePM") Then
        WriteStandardSection wsPay, dataArr, feesArr, hasFees, _
            "MilnePM", "Consultant", _
            curRow, pageStartRow, compRowCounts, True
    End If

    ' =====================================================================
    curSection = "TotalConsultants"
    ' SECTION 5: TOTAL INVOICES PROCESSED
    ' =====================================================================
    If mFirstFeeDataRow > 0 Then
        ' Check if TOTAL INVOICES row fits on current page (needs 2 rows: total + gap)
        If (curRow - pageStartRow) + 2 > ROWS_PER_PAGE And curRow > pageStartRow + 2 Then
            curRow = curRow + 1  ' buffer row at bottom of page
            wsPay.HPageBreaks.Add Before:=wsPay.Cells(curRow, 1)
            WriteInvoiceHeader wsPay, curRow
            pageStartRow = curRow
        End If
        curRow = curRow + 1
        wsPay.Cells(curRow, 1).Value = "TOTAL CONSULTANT INVOICES"
        FormatRowBase wsPay, curRow
        ' Thick black outside border on the TOTAL row (A-G) + whole row bold
        With wsPay.Range(wsPay.Cells(curRow, 1), wsPay.Cells(curRow, 10))
            .Font.Bold = True
        End With
        With wsPay.Range(wsPay.Cells(curRow, 1), wsPay.Cells(curRow, 7))
            .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlMedium
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlMedium
        End With
        wsPay.Cells(curRow, 4).Formula = "=SUM(D" & mFirstFeeDataRow & ":D" & mLastFeeDataRow & ")"
        wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
        wsPay.Cells(curRow, 5).Formula = "=IF(I" & curRow & "=0,0,D" & curRow & "/I" & curRow & ")"
        wsPay.Cells(curRow, 5).NumberFormatLocal =PCT_FMT
        wsPay.Cells(curRow, 9).Formula = "=SUBTOTAL(9,I" & mFirstFeeDataRow & ":I" & mLastFeeDataRow & ")"
        wsPay.Cells(curRow, 9).NumberFormatLocal =MONEY_FMT_WHOLE
        ' Large bracketed totals render as hashes at 80-90% zoom in a fixed-width
        ' column (print is fine). Shrink-to-fit adapts to the viewer's screen.
        wsPay.Range(wsPay.Cells(curRow, 4), wsPay.Cells(curRow, 5)).ShrinkToFit = True
        wsPay.Cells(curRow, 9).ShrinkToFit = True
        wsPay.Cells(curRow, 10).Value = "Total committed"
        wsPay.Cells(curRow, 9).Interior.Color = FEE_FILL
        wsPay.Cells(curRow, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous: wsPay.Cells(curRow, 9).Borders(xlEdgeLeft).Weight = xlThin
        wsPay.Cells(curRow, 9).Borders(xlEdgeRight).LineStyle = xlContinuous: wsPay.Cells(curRow, 9).Borders(xlEdgeRight).Weight = xlThin
        wsPay.Cells(curRow, 9).Borders(xlEdgeTop).LineStyle = xlContinuous: wsPay.Cells(curRow, 9).Borders(xlEdgeTop).Weight = xlThin
        wsPay.Cells(curRow, 9).Borders(xlEdgeBottom).LineStyle = xlContinuous: wsPay.Cells(curRow, 9).Borders(xlEdgeBottom).Weight = xlThin
        mTotalFeeCommittedRow = curRow
        curRow = curRow + 1  ' advance past this row so page breaks can't overwrite it
    End If

    ' =====================================================================
    curSection = "MiscFees"
    ' SECTION 6: MISCELLANEOUS FEES (smart page break)
    ' =====================================================================
    If showMisc Then
    Dim miscRowCount As Long: miscRowCount = 0
    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = "Misc Fees" Then miscRowCount = miscRowCount + 1
    Next i
    Dim miscFeeCount As Long: miscFeeCount = CountTypeFees(feesArr, hasFees, "Misc Fees")
    Dim miscSectionSize As Long
    miscSectionSize = 1 + Application.WorksheetFunction.Max(miscRowCount, miscFeeCount + 1) + 1 + 1

    If (curRow - pageStartRow) + miscSectionSize + 1 > ROWS_PER_PAGE And curRow > pageStartRow + 2 Then
        curRow = curRow + 1  ' buffer row at bottom of page
        wsPay.HPageBreaks.Add Before:=wsPay.Cells(curRow, 1)
        WriteInvoiceHeader wsPay, curRow
        pageStartRow = curRow
    End If

    curRow = curRow + 1
    wsPay.Cells(curRow, 1).Value = "Miscellaneous Fees"
    wsPay.Cells(curRow, 1).Font.Bold = True
    FormatCompanyRow wsPay, curRow
    curRow = curRow + 1
    Dim miscStartRow As Long, miscCount As Long
    miscStartRow = curRow: miscCount = 0

    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = "Misc Fees" Then
            Dim isPaidMisc As Boolean
            isPaidMisc = (UCase(Trim(CStr(dataArr(i, 6)))) = "Y")
            Dim miscLabel As String: miscLabel = CStr(dataArr(i, 2))
            If UBound(dataArr, 2) >= 7 Then
                Dim miscCmnt As String: miscCmnt = Trim(CStr(dataArr(i, 7)))
                If Len(miscCmnt) > 0 Then miscLabel = miscLabel & " - " & miscCmnt
            End If
            wsPay.Cells(curRow, 1).Value = miscLabel
            If Not isPaidMisc Then
                wsPay.Range(wsPay.Cells(curRow, 1), wsPay.Cells(curRow, 7)).Interior.Color = UNPAID_FILL
            End If
            wsPay.Cells(curRow, 1).Font.Name = FONT_NAME
            wsPay.Cells(curRow, 1).Font.Size = FONT_SIZE
            wsPay.Cells(curRow, 2).Value = dataArr(i, 3)
            wsPay.Cells(curRow, 2).NumberFormatLocal =DATE_FMT
            wsPay.Cells(curRow, 3).NumberFormatLocal ="@"
            wsPay.Cells(curRow, 3).Value = CStr(dataArr(i, 4))
            wsPay.Cells(curRow, 4).Value = dataArr(i, 5)
            wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
            If miscCount = 0 Then
                wsPay.Cells(curRow, 5).Formula = "=D" & curRow
            Else
                wsPay.Cells(curRow, 5).Formula = "=E" & (curRow - 1) & "+D" & curRow
            End If
            wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
            wsPay.Cells(curRow, 6).Formula = "=D" & curRow & "*" & GST_RATE
            wsPay.Cells(curRow, 6).NumberFormatLocal =MONEY_FMT
            If miscCount = 0 Then
                wsPay.Cells(curRow, 7).Formula = "=F" & curRow
            Else
                wsPay.Cells(curRow, 7).Formula = "=G" & (curRow - 1) & "+F" & curRow
            End If
            wsPay.Cells(curRow, 7).NumberFormatLocal =MONEY_FMT
            If isPaidMisc Then wsPay.Cells(curRow, 11).Value = "Y"
            FormatRowBase wsPay, curRow
            FormatDataRow wsPay, curRow
            miscCount = miscCount + 1
            curRow = curRow + 1
        End If
    Next i
    Dim miscEndRow As Long: miscEndRow = curRow - 1
    mMiscStartRow = miscStartRow: mMiscEndRow = miscEndRow

    ' Misc committed fees
    Dim miscFeeRow As Long: miscFeeRow = miscStartRow
    Dim miscFeeWritten As Long: miscFeeWritten = 0
    WriteFeeBlock wsPay, feesArr, hasFees, "Misc Fees", "", miscFeeRow, miscFeeWritten, miscEndRow
    ' Ensure curRow covers fee rows if fees > invoices
    If miscFeeRow > curRow Then curRow = miscFeeRow
    ' Total fee for misc
    WriteTotalAgreedFee wsPay, miscStartRow, miscFeeRow, miscFeeWritten, "Total"
    mMiscTotalFeeRow = miscStartRow + miscFeeWritten
    If miscFeeWritten = 0 Then mMiscTotalFeeRow = miscStartRow
    ' Ensure curRow is past total fee row
    If mMiscTotalFeeRow > curRow Then curRow = mMiscTotalFeeRow

    ' Misc percentage row
    wsPay.Cells(curRow, 1).Value = "Percentage paid to date"
    wsPay.Cells(curRow, 1).Font.Name = FONT_NAME
    wsPay.Cells(curRow, 1).Font.Size = FONT_SIZE
    wsPay.Cells(curRow, 5).Formula = "=IF(I" & mMiscTotalFeeRow & "=0,0,E" & miscEndRow & "/I" & mMiscTotalFeeRow & ")"
    wsPay.Cells(curRow, 5).NumberFormatLocal =PCT_FMT
    FormatPercentageRow wsPay, curRow
    curRow = curRow + 1  ' advance past percentage row
    End If ' showMisc

    ' =====================================================================
    curSection = "ClientSupply"
    ' SECTION 7: CLIENT SUPPLIED ITEMS (smart page break)
    ' =====================================================================
    If showClientSupply Then
    Dim csRowCount As Long: csRowCount = 0
    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = "Client Supply" Then csRowCount = csRowCount + 1
    Next i
    Dim csFeeCount As Long: csFeeCount = CountTypeFees(feesArr, hasFees, "Client Supply")
    Dim csSectionSize As Long
    csSectionSize = 1 + Application.WorksheetFunction.Max(csRowCount, csFeeCount + 1) + 1 + 1

    If (curRow - pageStartRow) + csSectionSize + 1 > ROWS_PER_PAGE And curRow > pageStartRow + 2 Then
        curRow = curRow + 1  ' buffer row at bottom of page
        wsPay.HPageBreaks.Add Before:=wsPay.Cells(curRow, 1)
        WriteInvoiceHeader wsPay, curRow
        pageStartRow = curRow
    End If

    curRow = curRow + 1
    wsPay.Cells(curRow, 1).Value = "Client Supplied Items"
    wsPay.Cells(curRow, 1).Font.Bold = True
    FormatCompanyRow wsPay, curRow
    curRow = curRow + 1
    Dim csStartRow As Long, csCount As Long
    csStartRow = curRow: csCount = 0

    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = "Client Supply" Then
            Dim isPaidCS As Boolean
            isPaidCS = (UCase(Trim(CStr(dataArr(i, 6)))) = "Y")
            Dim csLabel As String: csLabel = CStr(dataArr(i, 2))
            If UBound(dataArr, 2) >= 7 Then
                Dim csCmnt As String: csCmnt = Trim(CStr(dataArr(i, 7)))
                If Len(csCmnt) > 0 Then csLabel = csLabel & " - " & csCmnt
            End If
            wsPay.Cells(curRow, 1).Value = csLabel
            If Not isPaidCS Then
                wsPay.Range(wsPay.Cells(curRow, 1), wsPay.Cells(curRow, 7)).Interior.Color = UNPAID_FILL
            End If
            wsPay.Cells(curRow, 1).Font.Name = FONT_NAME
            wsPay.Cells(curRow, 1).Font.Size = FONT_SIZE
            wsPay.Cells(curRow, 2).Value = dataArr(i, 3)
            wsPay.Cells(curRow, 2).NumberFormatLocal =DATE_FMT
            wsPay.Cells(curRow, 3).NumberFormatLocal ="@"
            wsPay.Cells(curRow, 3).Value = CStr(dataArr(i, 4))
            wsPay.Cells(curRow, 4).Value = dataArr(i, 5)
            wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
            If csCount = 0 Then
                wsPay.Cells(curRow, 5).Formula = "=D" & curRow
            Else
                wsPay.Cells(curRow, 5).Formula = "=E" & (curRow - 1) & "+D" & curRow
            End If
            wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
            wsPay.Cells(curRow, 6).Formula = "=D" & curRow & "*" & GST_RATE
            wsPay.Cells(curRow, 6).NumberFormatLocal =MONEY_FMT
            If csCount = 0 Then
                wsPay.Cells(curRow, 7).Formula = "=F" & curRow
            Else
                wsPay.Cells(curRow, 7).Formula = "=G" & (curRow - 1) & "+F" & curRow
            End If
            wsPay.Cells(curRow, 7).NumberFormatLocal =MONEY_FMT
            If isPaidCS Then wsPay.Cells(curRow, 11).Value = "Y"
            FormatRowBase wsPay, curRow
            FormatDataRow wsPay, curRow
            csCount = csCount + 1
            curRow = curRow + 1
        End If
    Next i
    Dim csEndRow As Long: csEndRow = curRow - 1
    mCsStartRow = csStartRow: mCsEndRow = csEndRow

    ' Client Supply committed fees
    Dim csFeeRow As Long: csFeeRow = csStartRow
    Dim csFeeWritten As Long: csFeeWritten = 0
    WriteFeeBlock wsPay, feesArr, hasFees, "Client Supply", "", csFeeRow, csFeeWritten, csEndRow
    If csFeeRow > curRow Then curRow = csFeeRow
    WriteTotalAgreedFee wsPay, csStartRow, csFeeRow, csFeeWritten, "Total"
    mCsTotalFeeRow = csStartRow + csFeeWritten
    If csFeeWritten = 0 Then mCsTotalFeeRow = csStartRow
    If mCsTotalFeeRow > curRow Then curRow = mCsTotalFeeRow

    ' CS percentage row
    wsPay.Cells(curRow, 1).Value = "Percentage paid to date"
    wsPay.Cells(curRow, 1).Font.Name = FONT_NAME
    wsPay.Cells(curRow, 1).Font.Size = FONT_SIZE
    wsPay.Cells(curRow, 5).Formula = "=IF(I" & mCsTotalFeeRow & "=0,0,E" & csEndRow & "/I" & mCsTotalFeeRow & ")"
    wsPay.Cells(curRow, 5).NumberFormatLocal =PCT_FMT
    FormatPercentageRow wsPay, curRow
    curRow = curRow + 1
    End If ' showClientSupply

    ' =====================================================================
    curSection = "Contractors"
    ' SECTION 8: CONTRACTORS (forced page break for section, smart breaks within)
    ' =====================================================================
    ' Contractors = A-Z union of COMMITTED rows typed Contractor and
    ' CLAIM_ENTRY col-A names (item 15: a claims-only or committed-only
    ' contractor is always listed). CLAIM_ENTRY has no TYPE column, so
    ' GetAllCompaniesForType cannot read names from it.
    Dim contractors As Variant
    Dim ctrDict As Object: Set ctrDict = CreateObject("Scripting.Dictionary")
    Dim cdIdx As Long, cdName As String
    If hasFees Then
        For cdIdx = 1 To UBound(feesArr, 1)
            If CStr(feesArr(cdIdx, 1)) = "Contractor" Then
                cdName = CStr(feesArr(cdIdx, 3))
                If Len(cdName) > 0 And Not ctrDict.Exists(cdName) Then ctrDict.Add cdName, 0
            End If
        Next cdIdx
    End If
    If hasClaims Then
        For cdIdx = 1 To UBound(claimArr, 1)
            cdName = CStr(claimArr(cdIdx, 1))
            If Len(cdName) > 0 And Not ctrDict.Exists(cdName) Then ctrDict.Add cdName, 0
        Next cdIdx
    End If
    If ctrDict.Count = 0 Then
        contractors = Empty
    Else
        Dim ctrList() As String: ReDim ctrList(0 To ctrDict.Count - 1)
        Dim cdKey As Variant: cdIdx = 0
        For Each cdKey In ctrDict.Keys: ctrList(cdIdx) = CStr(cdKey): cdIdx = cdIdx + 1: Next cdKey
        ' Bubble sort A-Z (same as GetAllCompaniesForType)
        Dim cdJ As Long, cdSwap As String
        For cdIdx = 0 To UBound(ctrList) - 1
            For cdJ = 0 To UBound(ctrList) - cdIdx - 1
                If ctrList(cdJ) > ctrList(cdJ + 1) Then
                    cdSwap = ctrList(cdJ): ctrList(cdJ) = ctrList(cdJ + 1): ctrList(cdJ + 1) = cdSwap
                End If
            Next cdJ
        Next cdIdx
        contractors = ctrList
    End If

    If Not IsEmpty(contractors) Then
        ReDim mContrStartRows(0 To UBound(contractors))
        ReDim mContrEndRows(0 To UBound(contractors))
        ReDim mContrTotalFeeRows(0 To UBound(contractors))
        mContrCount = 0

        ' Forced page break at start of contractor section
        curRow = curRow + 1
        wsPay.HPageBreaks.Add Before:=wsPay.Cells(curRow, 1)
        pageStartRow = curRow
        WriteContractorHeader wsPay, curRow, ""

        For compIdx = LBound(contractors) To UBound(contractors)
            If Len(contractors(compIdx)) = 0 Then Exit For
            Dim contrName As String: contrName = CStr(contractors(compIdx))

            ' Resolve this contractor's retention type from COMMITTED
            Dim retentionType As String
            If ctrRet.Exists(contrName) Then
                retentionType = ctrRet(contrName)
            Else
                retentionType = "Major Works"
            End If
            ' ...and its retention cap (COMMITTED col H), 0 = no extra cap
            Dim capVal As Double: capVal = 0
            If ctrCap.Exists(contrName) Then capVal = ctrCap(contrName)

            ' Look-ahead: will this contractor fit on current page?
            Dim ctrInvCount As Long: ctrInvCount = 0
            Dim ckCtr As String: ckCtr = "Contractor|" & contrName
            If compRowCounts.Exists(ckCtr) Then ctrInvCount = compRowCounts(ckCtr)
            Dim ctrFeeCnt As Long: ctrFeeCnt = CountCompanyFees(feesArr, hasFees, "Contractor", contrName)
            Dim ctrBodyRows As Long
            If ctrInvCount > ctrFeeCnt + 1 Then ctrBodyRows = ctrInvCount Else ctrBodyRows = ctrFeeCnt + 1
            If ctrBodyRows < 1 Then ctrBodyRows = 1
            Dim ctrSectionRows As Long: ctrSectionRows = 1 + ctrBodyRows + 1 + 1

            If (curRow - pageStartRow) + ctrSectionRows + 1 > ROWS_PER_PAGE And curRow > pageStartRow + 2 Then
                curRow = curRow + 1  ' buffer row at bottom of page
                wsPay.HPageBreaks.Add Before:=wsPay.Cells(curRow, 1)
                WriteContractorHeader wsPay, curRow, ""
                pageStartRow = curRow
            End If

            ' Company name row
            curRow = curRow + 1
            wsPay.Cells(curRow, 1).Value = DisplayCompanyName(contrName)
            wsPay.Cells(curRow, 1).Font.Bold = True
            FormatCompanyRow wsPay, curRow

            curRow = curRow + 1
            Dim ctrStartRow As Long, ctrCount As Long
            ctrStartRow = curRow: ctrCount = 0
            Dim retFormula As String
            ' Releases so far, kept as an Excel EXPRESSION ("2500+E9*.5") so
            ' every balance stays a sheet-auditable formula, never a number
            ' baked in VBA. Reset per contractor.
            Dim relSumExpr As String: relSumExpr = ""

            ' Write contractor progress payments from CLAIM_ENTRY (claimArr is
            ' never dimensioned when CLAIM_ENTRY is empty: ctrClaimCnt = 0 skips it)
            Dim ctrClaimCnt As Long: ctrClaimCnt = 0
            If hasClaims Then ctrClaimCnt = UBound(claimArr, 1)
            For i = 1 To ctrClaimCnt
                If CStr(claimArr(i, 1)) = contrName Then
                    Dim isPaidCtr As Boolean
                    isPaidCtr = (UCase(Trim(CStr(claimArr(i, 5)))) = "Y")
                    ' Release row = a value in CLAIM_ENTRY col G ("Retention Release $"):
                    '   > 1        a dollar amount
                    '   0 < v <= 1 a percentage of the balance currently held
                    '              (Excel stores a typed "50%" as 0.5)
                    '   negative   a correction: re-holds previously released
                    '              retention (same convention as negative
                    '              refund invoices elsewhere in the workbook)
                    ' Value of Work is always entered as per the claim — restated on a
                    ' pure release, increased when a claim and a release share a row.
                    ' The balance drops by the release, so the existing col-F formula
                    ' pays it out and it flows into the construction totals untouched.
                    Dim relAmt As Double: relAmt = 0
                    If UBound(claimArr, 2) >= 7 Then
                        If IsNumeric(claimArr(i, 7)) Then relAmt = CDbl(claimArr(i, 7))
                    End If
                    Dim isRelease As Boolean: isRelease = (relAmt <> 0)
                    If Not isPaidCtr Then
                        wsPay.Cells(curRow, 1).Value = "Invoice recommended for payment"
                        wsPay.Range(wsPay.Cells(curRow, 1), wsPay.Cells(curRow, 7)).Interior.Color = UNPAID_FILL
                    End If
                    wsPay.Cells(curRow, 2).Value = claimArr(i, 2)
                    wsPay.Cells(curRow, 2).NumberFormatLocal =DATE_FMT
                    wsPay.Cells(curRow, 3).NumberFormatLocal ="@"
                    wsPay.Cells(curRow, 3).Value = CStr(claimArr(i, 3))
                    ' Entered value used as-is on every row — release rows restate
                    ' the completed value. A blank on a release row renders as an
                    ' obviously-wrong negative balance rather than being papered over.
                    wsPay.Cells(curRow, 4).Value = claimArr(i, 4)
                    wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
                    If isRelease Then
                        Dim thisRel As String
                        If relAmt > 0 And relAmt <= 1 And ctrCount > 0 Then
                            ' Percentage of the balance held = the row above's E
                            thisRel = "E" & (curRow - 1) & "*" & Trim$(Str$(relAmt))
                        Else
                            thisRel = Trim$(Str$(relAmt))
                        End If
                        If Len(relSumExpr) = 0 Then
                            relSumExpr = thisRel
                        Else
                            relSumExpr = relSumExpr & "+" & thisRel
                        End If
                    End If
                    ' Tier expression + type cap per retention type; the
                    ' contractor's own cap (COMMITTED col H) is one more MIN
                    ' argument, so it stays visible in the sheet formula.
                    ' Without a cap the string is exactly the pre-Update-4 one.
                    Dim tierExpr As String, typeCap As String, capPart As String
                    capPart = ""
                    If capVal > 0 Then capPart = "," & Trim$(Str$(capVal))
                    Select Case retentionType
                        Case "Minor Works"
                            tierExpr = ""
                        Case "Medium Works"
                            tierExpr = "IF(D" & curRow & "<200000,D" & curRow & "*0.05," & _
                                       "10000+(D" & curRow & "-200000)*0.025)"
                            typeCap = "30000"
                        Case "NZS3910"
                            ' Full NZ industry standard: 10% first $200K, 5% next $800K, 1.75% over $1M, capped at $200K
                            tierExpr = "IF(D" & curRow & "<200000,D" & curRow & "*0.1," & _
                                       "IF(D" & curRow & "<1000000,20000+(D" & curRow & "-200000)*0.05," & _
                                       "60000+(D" & curRow & "-1000000)*0.0175))"
                            typeCap = "200000"
                        Case Else  ' Major Works (capped at $200,000) - half-rate variant
                            tierExpr = "IF(D" & curRow & "<200000,D" & curRow & "*0.05," & _
                                       "IF(D" & curRow & "<1000000,10000+(D" & curRow & "-200000)*0.025," & _
                                       "30000+(D" & curRow & "-1000000)*0.0087))"
                            typeCap = "200000"
                    End Select
                    If retentionType = "Minor Works" Then
                        retFormula = "=0"
                    Else
                        retFormula = "=ROUND(MIN(" & tierExpr & "," & typeCap & capPart & "),2)"
                    End If
                    ' Balance held = tier retention on cumulative work MINUS all
                    ' releases so far. One shape for every row (plain claims,
                    ' releases, combined claim+release), so released retention is
                    ' never re-held. Str$ (not CStr) for literals: it always renders
                    ' a period decimal, which .Formula requires regardless of locale.
                    If Len(relSumExpr) > 0 Then
                        wsPay.Cells(curRow, 5).Formula = retFormula & "-(" & relSumExpr & ")"
                    Else
                        wsPay.Cells(curRow, 5).Formula = retFormula
                    End If
                    wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
                    If ctrCount = 0 Then
                        wsPay.Cells(curRow, 6).Formula = "=D" & curRow & "-E" & curRow
                    Else
                        wsPay.Cells(curRow, 6).Formula = "=D" & curRow & "-E" & curRow & "-(D" & (curRow - 1) & "-E" & (curRow - 1) & ")"
                    End If
                    wsPay.Cells(curRow, 6).NumberFormatLocal =MONEY_FMT
                    wsPay.Cells(curRow, 7).Formula = "=F" & curRow & "*" & GST_RATE
                    wsPay.Cells(curRow, 7).NumberFormatLocal =MONEY_FMT
                    If isPaidCtr Then wsPay.Cells(curRow, 11).Value = "Y"
                    ' Write comment from CLAIM_ENTRY column F if present
                    If UBound(claimArr, 2) >= 6 Then
                        Dim ctrCmnt As String: ctrCmnt = Trim(CStr(claimArr(i, 6)))
                        If Len(ctrCmnt) > 0 Then
                            Dim ctrA As String: ctrA = Trim(CStr(wsPay.Cells(curRow, 1).Value))
                            If Len(ctrA) > 0 Then
                                wsPay.Cells(curRow, 1).Value = ctrA & " - " & ctrCmnt
                            Else
                                wsPay.Cells(curRow, 1).Value = ctrCmnt
                            End If
                        End If
                    End If
                    FormatRowBase wsPay, curRow
                    FormatDataRow wsPay, curRow
                    ctrCount = ctrCount + 1
                    curRow = curRow + 1
                End If
            Next i
            Dim ctrEndRow As Long: ctrEndRow = curRow - 1

            ' Percentage row placed right after last claim (collapses up to title if no claims).
            ' Fees are written in parallel in col I/J so the percentage row can share rows with them.
            Dim ctrPctRow As Long: ctrPctRow = curRow

            ' Contractor committed fees (col I/J, starting at ctrStartRow — parallel to claims and percentage)
            Dim ctrFeeRow As Long: ctrFeeRow = ctrStartRow
            Dim ctrFeeWritten As Long: ctrFeeWritten = 0
            WriteFeeBlock wsPay, feesArr, hasFees, "Contractor", contrName, ctrFeeRow, ctrFeeWritten, ctrEndRow
            WriteTotalAgreedFee wsPay, ctrStartRow, ctrFeeRow, ctrFeeWritten
            Dim ctrTotalFeeRow As Long
            ctrTotalFeeRow = ctrStartRow + ctrFeeWritten
            If ctrFeeWritten = 0 Then ctrTotalFeeRow = ctrStartRow

            ' Write percentage row at the reserved position
            wsPay.Cells(ctrPctRow, 1).Value = "Percentage work completed to date"
            wsPay.Cells(ctrPctRow, 1).Font.Name = FONT_NAME
            wsPay.Cells(ctrPctRow, 1).Font.Size = FONT_SIZE
            If ctrCount > 0 Then
                wsPay.Cells(ctrPctRow, 4).Formula = "=IF(I" & ctrTotalFeeRow & "=0,0,D" & ctrEndRow & "/I" & ctrTotalFeeRow & ")"
            Else
                wsPay.Cells(ctrPctRow, 4).Value = 0
            End If
            wsPay.Cells(ctrPctRow, 4).NumberFormatLocal =PCT_FMT
            FormatPercentageRow wsPay, ctrPctRow

            ' Advance curRow past both the percentage row and the fee block
            curRow = ctrPctRow + 1
            If ctrTotalFeeRow + 1 > curRow Then curRow = ctrTotalFeeRow + 1
            curRow = curRow + 1  ' gap row before next contractor

            ' Store contractor tracking
            mContrStartRows(mContrCount) = ctrStartRow
            mContrEndRows(mContrCount) = ctrEndRow
            mContrTotalFeeRows(mContrCount) = ctrTotalFeeRow
            mContrCount = mContrCount + 1
        Next compIdx
    End If

    ' =====================================================================
    curSection = "CashFlow"
    ' SECTION 9: CASH FLOW IN (forced page break, after Contractors per Update 4 item 4)
    ' =====================================================================
    Dim cfRowCount As Long: cfRowCount = 0
    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = "Cash Flow In" Then cfRowCount = cfRowCount + 1
    Next i
    Dim cfFeeCountTotal As Long: cfFeeCountTotal = CountTypeFees(feesArr, hasFees, "Cash Flow In")

    ' Track Cash Flow In total fee rows (for summary)
    Dim cfTotalFeeRows() As Long
    Dim cfCompanyCount As Long: cfCompanyCount = 0

    If cfRowCount > 0 Or cfFeeCountTotal > 0 Then
        curRow = curRow + 1
        wsPay.HPageBreaks.Add Before:=wsPay.Cells(curRow, 1)
        pageStartRow = curRow
        WriteInvoiceHeader wsPay, curRow

        ' Get Cash Flow In companies
        Dim cfCompanies As Variant
        cfCompanies = GetAllCompaniesForType(dataArr, feesArr, hasFees, "Cash Flow In", "")
        If Not IsEmpty(cfCompanies) Then
            ReDim cfTotalFeeRows(0 To UBound(cfCompanies))
            For compIdx = LBound(cfCompanies) To UBound(cfCompanies)
                If Len(cfCompanies(compIdx)) > 0 Then
                    Dim cfTFR As Long: cfTFR = 0
                    WriteStandardSection wsPay, dataArr, feesArr, hasFees, _
                        CStr(cfCompanies(compIdx)), "Cash Flow In", _
                        curRow, pageStartRow, compRowCounts, False, cfTFR
                    cfTotalFeeRows(cfCompanyCount) = cfTFR
                    cfCompanyCount = cfCompanyCount + 1
                End If
            Next compIdx
        End If
    End If

    ' =====================================================================
    curSection = "InvoiceSummary"
    ' SECTION 10: INVOICE SUMMARY (forced page break)
    ' =====================================================================
    curRow = curRow + 1  ' buffer row at bottom of page
    wsPay.HPageBreaks.Add Before:=wsPay.Cells(curRow, 1)
    pageStartRow = curRow

    ' --- Summary header (columns A, C, D, E only) ---
    WriteSummaryHeader wsPay, curRow
    curRow = curRow + 1

    ' --- Helper: build safe SUM formulas (avoids =0+I## pattern) ---
    ' Fee committed parts
    Dim feeCommParts As Long: feeCommParts = 0
    Dim feeCommFormula As String: feeCommFormula = ""
    If mTotalFeeCommittedRow > 0 Then
        feeCommFormula = "I" & mTotalFeeCommittedRow: feeCommParts = feeCommParts + 1
    End If
    If Not showMisc Then
        If mMiscTotalFeeRow > 0 Then
            If feeCommParts > 0 Then feeCommFormula = feeCommFormula & "+"
            feeCommFormula = feeCommFormula & "I" & mMiscTotalFeeRow: feeCommParts = feeCommParts + 1
        End If
    End If
    If Not showClientSupply Then
        If mCsTotalFeeRow > 0 Then
            If feeCommParts > 0 Then feeCommFormula = feeCommFormula & "+"
            feeCommFormula = feeCommFormula & "I" & mCsTotalFeeRow: feeCommParts = feeCommParts + 1
        End If
    End If
    If feeCommParts = 0 Then feeCommFormula = "0"
    feeCommFormula = "=" & feeCommFormula

    ' Fee invoices to date parts
    Dim feeInvParts As Long: feeInvParts = 0
    Dim feeInvFormula As String: feeInvFormula = ""
    If mTotalFeeCommittedRow > 0 Then
        feeInvFormula = "D" & mTotalFeeCommittedRow: feeInvParts = feeInvParts + 1
    End If
    If Not showMisc Then
        If mMiscStartRow > 0 And mMiscEndRow >= mMiscStartRow Then
            If feeInvParts > 0 Then feeInvFormula = feeInvFormula & "+"
            feeInvFormula = feeInvFormula & "SUM(D" & mMiscStartRow & ":D" & mMiscEndRow & ")"
            feeInvParts = feeInvParts + 1
        End If
    End If
    If Not showClientSupply Then
        If mCsStartRow > 0 And mCsEndRow >= mCsStartRow Then
            If feeInvParts > 0 Then feeInvFormula = feeInvFormula & "+"
            feeInvFormula = feeInvFormula & "SUM(D" & mCsStartRow & ":D" & mCsEndRow & ")"
            feeInvParts = feeInvParts + 1
        End If
    End If
    If feeInvParts = 0 Then feeInvFormula = "0"
    feeInvFormula = "=" & feeInvFormula

    ' Construction committed parts
    Dim constrCommFormula As String: constrCommFormula = ""
    Dim ci As Long
    For ci = 0 To mContrCount - 1
        If ci > 0 Then constrCommFormula = constrCommFormula & "+"
        constrCommFormula = constrCommFormula & "I" & mContrTotalFeeRows(ci)
    Next ci
    If mContrCount = 0 Then constrCommFormula = "0"
    constrCommFormula = "=" & constrCommFormula

    ' Construction invoices to date parts
    Dim constrInvFormula As String: constrInvFormula = ""
    Dim constrInvParts As Long: constrInvParts = 0
    For ci = 0 To mContrCount - 1
        If mContrEndRows(ci) >= mContrStartRows(ci) Then
            If constrInvParts > 0 Then constrInvFormula = constrInvFormula & "+"
            constrInvFormula = constrInvFormula & "SUM(F" & mContrStartRows(ci) & ":F" & mContrEndRows(ci) & ")"
            constrInvParts = constrInvParts + 1
        End If
    Next ci
    If constrInvParts = 0 Then constrInvFormula = "0"
    constrInvFormula = "=" & constrInvFormula

    ' CF committed parts
    Dim cfSumFormula As String: cfSumFormula = ""
    If cfCompanyCount > 0 Then
        Dim cfi As Long
        For cfi = 0 To cfCompanyCount - 1
            If cfi > 0 Then cfSumFormula = cfSumFormula & "+"
            cfSumFormula = cfSumFormula & "I" & cfTotalFeeRows(cfi)
        Next cfi
        cfSumFormula = "=" & cfSumFormula
    End If

    ' --- Fee invoices row ---
    Dim sumFeeRow As Long: sumFeeRow = curRow
    FormatSummaryRow wsPay, curRow, False
    wsPay.Cells(curRow, 1).Value = "Consultant invoices"
    wsPay.Cells(curRow, 3).Value = feeBudget: wsPay.Cells(curRow, 3).NumberFormatLocal =MONEY_FMT
    wsPay.Cells(curRow, 4).Formula = feeCommFormula: wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
    wsPay.Cells(curRow, 5).Formula = feeInvFormula: wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
    curRow = curRow + 1

    ' --- Misc Fees row (optional) ---
    If showMisc Then
        FormatSummaryRow wsPay, curRow, False
        wsPay.Cells(curRow, 1).Value = "Miscellaneous invoices"
        wsPay.Cells(curRow, 3).Value = miscBudget: wsPay.Cells(curRow, 3).NumberFormatLocal =MONEY_FMT
        If mMiscTotalFeeRow > 0 Then
            wsPay.Cells(curRow, 4).Formula = "=I" & mMiscTotalFeeRow
        Else
            wsPay.Cells(curRow, 4).Value = 0
        End If
        wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
        If mMiscStartRow > 0 And mMiscEndRow >= mMiscStartRow Then
            wsPay.Cells(curRow, 5).Formula = "=SUM(D" & mMiscStartRow & ":D" & mMiscEndRow & ")"
        Else
            wsPay.Cells(curRow, 5).Value = 0
        End If
        wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
        curRow = curRow + 1
    End If

    ' --- Client Supplied Items row (optional) ---
    If showClientSupply Then
        FormatSummaryRow wsPay, curRow, False
        wsPay.Cells(curRow, 1).Value = "Client supplied invoices"
        wsPay.Cells(curRow, 3).Value = csBudget: wsPay.Cells(curRow, 3).NumberFormatLocal =MONEY_FMT
        If mCsTotalFeeRow > 0 Then
            wsPay.Cells(curRow, 4).Formula = "=I" & mCsTotalFeeRow
        Else
            wsPay.Cells(curRow, 4).Value = 0
        End If
        wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
        If mCsStartRow > 0 And mCsEndRow >= mCsStartRow Then
            wsPay.Cells(curRow, 5).Formula = "=SUM(D" & mCsStartRow & ":D" & mCsEndRow & ")"
        Else
            wsPay.Cells(curRow, 5).Value = 0
        End If
        wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
        curRow = curRow + 1
    End If

    ' --- Construction invoices row ---
    Dim sumConstrRow As Long: sumConstrRow = curRow
    FormatSummaryRow wsPay, curRow, False
    wsPay.Cells(curRow, 1).Value = "Construction invoices"
    wsPay.Cells(curRow, 3).Value = constrBudget: wsPay.Cells(curRow, 3).NumberFormatLocal =MONEY_FMT
    wsPay.Cells(curRow, 4).Formula = constrCommFormula: wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
    wsPay.Cells(curRow, 5).Formula = constrInvFormula: wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
    curRow = curRow + 1

    ' --- Total Invoices row (highlighted, bold values) ---
    Dim sumTotalRow As Long: sumTotalRow = curRow
    FormatSummaryRow wsPay, curRow, True
    wsPay.Cells(curRow, 1).Value = "Total Invoices"
    wsPay.Cells(curRow, 1).Font.Bold = True
    wsPay.Cells(curRow, 3).Formula = "=SUM(C" & sumFeeRow & ":C" & sumConstrRow & ")"
    wsPay.Cells(curRow, 3).NumberFormatLocal =MONEY_FMT
    wsPay.Cells(curRow, 3).Font.Bold = True
    wsPay.Cells(curRow, 4).Formula = "=SUM(D" & sumFeeRow & ":D" & sumConstrRow & ")"
    wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
    wsPay.Cells(curRow, 4).Font.Bold = True
    wsPay.Cells(curRow, 5).Formula = "=SUM(E" & sumFeeRow & ":E" & sumConstrRow & ")"
    wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
    wsPay.Cells(curRow, 5).Font.Bold = True
    curRow = curRow + 1

    ' --- Funding drawdown (5YA) row (optional) — with spacer rows above/below ---
    If showCashFlow Then
        ' Blank spacer row above Funding drawdown
        curRow = curRow + 1

        Dim sum5YARow As Long: sum5YARow = curRow
        FormatSummaryRow wsPay, curRow, False
        wsPay.Cells(curRow, 1).Value = "Funding drawdown"
        ' Column C left blank (not $0.00)
        If cfCompanyCount > 0 Then
            wsPay.Cells(curRow, 4).Formula = cfSumFormula
        Else
            wsPay.Cells(curRow, 4).Value = 0
        End If
        wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
        If mCfStartRow > 0 And mCfEndRow >= mCfStartRow Then
            wsPay.Cells(curRow, 5).Formula = "=SUM(D" & mCfStartRow & ":D" & mCfEndRow & ")"
        Else
            wsPay.Cells(curRow, 5).Value = 0
        End If
        wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
        curRow = curRow + 1

        ' Blank spacer row below Funding drawdown
        curRow = curRow + 1

        ' --- Drawdown vs Invoices row (highlighted) ---
        FormatSummaryRow wsPay, curRow, True
        wsPay.Cells(curRow, 1).Value = "Drawdown vs Invoices"
        wsPay.Cells(curRow, 1).Font.Bold = True
        wsPay.Cells(curRow, 4).Formula = "=D" & sum5YARow & "-D" & sumTotalRow
        wsPay.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
        wsPay.Cells(curRow, 5).Formula = "=E" & sum5YARow & "-E" & sumTotalRow
        wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
        curRow = curRow + 1
    End If

    ' =====================================================================
    curSection = "SummaryYTD"
    ' SECTION 11: SUMMARY YTD (optional)
    ' =====================================================================
    If showYTD Then
        wsPay.Cells(curRow, 1).Value = "Summary YTD"
        wsPay.Cells(curRow, 1).Font.Bold = True
        wsPay.Cells(curRow, 1).Font.Name = FONT_NAME: wsPay.Cells(curRow, 1).Font.Size = FONT_SIZE
        curRow = curRow + 1
        Dim ytdStartRow As Long: ytdStartRow = curRow

        ' Determine full data range for SUMIFS
        Dim feeRangeStart As Long, feeRangeEnd As Long
        feeRangeStart = mFirstFeeDataRow
        feeRangeEnd = mCsEndRow
        If feeRangeEnd < mMiscEndRow Then feeRangeEnd = mMiscEndRow
        If feeRangeEnd < mLastFeeDataRow Then feeRangeEnd = mLastFeeDataRow

        ' Contractor range
        Dim ctrRangeStart As Long, ctrRangeEnd As Long
        ctrRangeStart = 0: ctrRangeEnd = 0
        If mContrCount > 0 Then
            ctrRangeStart = mContrStartRows(0)
            ctrRangeEnd = mContrEndRows(mContrCount - 1)
        End If

        ' Auto-detect year range from FEES_ENTRY + CLAIM_ENTRY dates
        Dim minYr As Long, maxYr As Long
        minYr = 9999: maxYr = 0
        Dim di As Long
        For di = 1 To UBound(dataArr, 1)
            If IsDate(dataArr(di, 3)) Then
                Dim dtYr As Long: dtYr = Year(CDate(dataArr(di, 3)))
                If dtYr < minYr Then minYr = dtYr
                If dtYr > maxYr Then maxYr = dtYr
            End If
        Next di
        If hasClaims Then
            For di = 1 To UBound(claimArr, 1)
                If IsDate(claimArr(di, 2)) Then
                    dtYr = Year(CDate(claimArr(di, 2)))
                    If dtYr < minYr Then minYr = dtYr
                    If dtYr > maxYr Then maxYr = dtYr
                End If
            Next di
        End If
        If minYr > maxYr Then minYr = Year(Date): maxYr = minYr

        Dim yr As Long
        For yr = minYr To maxYr
            wsPay.Range("A" & curRow & ":B" & curRow).Merge
            wsPay.Cells(curRow, 1).Value = "TOTAL PROJECT INVOICES YTD " & yr
            wsPay.Cells(curRow, 1).Font.Name = FONT_NAME: wsPay.Cells(curRow, 1).Font.Size = FONT_SIZE

            Dim ytdFormula As String: ytdFormula = "="
            ' Fee invoices portion
            If feeRangeStart > 0 And feeRangeEnd > 0 Then
                If yr = minYr Then
                    ytdFormula = ytdFormula & "(SUMIFS(D" & feeRangeStart & ":D" & feeRangeEnd & _
                        ",B" & feeRangeStart & ":B" & feeRangeEnd & ",""<""&DATE(" & (minYr + 1) & ",1,1)))"
                ElseIf yr = maxYr Then
                    ytdFormula = ytdFormula & "(SUMIFS(D" & feeRangeStart & ":D" & feeRangeEnd & _
                        ",B" & feeRangeStart & ":B" & feeRangeEnd & ","">=""&DATE(" & maxYr & ",1,1)))"
                Else
                    ytdFormula = ytdFormula & "(SUMIFS(D" & feeRangeStart & ":D" & feeRangeEnd & _
                        ",B" & feeRangeStart & ":B" & feeRangeEnd & ","">=""&DATE(" & yr & ",1,1)" & _
                        ",B" & feeRangeStart & ":B" & feeRangeEnd & ",""<""&DATE(" & (yr + 1) & ",1,1)))"
                End If
            Else
                ytdFormula = ytdFormula & "0"
            End If
            ' Contractor portion (column F = value less retentions less previous)
            If ctrRangeStart > 0 And ctrRangeEnd > 0 Then
                If yr = minYr Then
                    ytdFormula = ytdFormula & "+(SUMIFS(F" & ctrRangeStart & ":F" & ctrRangeEnd & _
                        ",B" & ctrRangeStart & ":B" & ctrRangeEnd & ",""<""&DATE(" & (minYr + 1) & ",1,1)))"
                ElseIf yr = maxYr Then
                    ytdFormula = ytdFormula & "+(SUMIFS(F" & ctrRangeStart & ":F" & ctrRangeEnd & _
                        ",B" & ctrRangeStart & ":B" & ctrRangeEnd & ","">=""&DATE(" & maxYr & ",1,1)))"
                Else
                    ytdFormula = ytdFormula & "+(SUMIFS(F" & ctrRangeStart & ":F" & ctrRangeEnd & _
                        ",B" & ctrRangeStart & ":B" & ctrRangeEnd & ","">=""&DATE(" & yr & ",1,1)" & _
                        ",B" & ctrRangeStart & ":B" & ctrRangeEnd & ",""<""&DATE(" & (yr + 1) & ",1,1)))"
                End If
            End If
            wsPay.Cells(curRow, 5).Formula = ytdFormula
            wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
            curRow = curRow + 1
        Next yr

        ' Total Invoices YTD
        wsPay.Range("A" & curRow & ":B" & curRow).Merge
        wsPay.Cells(curRow, 1).Value = "Total Invoices"
        wsPay.Cells(curRow, 1).Font.Bold = True
        wsPay.Cells(curRow, 1).Font.Name = FONT_NAME: wsPay.Cells(curRow, 1).Font.Size = FONT_SIZE
        wsPay.Cells(curRow, 5).Formula = "=SUM(E" & ytdStartRow & ":E" & (curRow - 1) & ")"
        wsPay.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
    End If

    curRow = curRow + 1  ' buffer row at bottom of Invoice Summary page

    ' =====================================================================
    ' PAGE SETUP
    ' =====================================================================
    Application.PrintCommunication = False
    With wsPay.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
        .LeftMargin = Application.InchesToPoints(0.7087)
        .RightMargin = Application.InchesToPoints(0.5118)
        .TopMargin = Application.InchesToPoints(1.5354)
        .BottomMargin = Application.InchesToPoints(0.748)
        .HeaderMargin = Application.InchesToPoints(0.5118)
        .FooterMargin = Application.InchesToPoints(0.315)
        .CenterHorizontally = True
        .PrintArea = "A1:J" & curRow
    End With
    Application.PrintCommunication = True

    ' Page Break Preview
    curSection = "Finalize"
    wsPay.Activate
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 85

    ' Clear any stale LastMacroError from a previous failed run -- on a fully
    ' successful run we want the named range to be absent.
    On Error Resume Next
    ThisWorkbook.Names("LastMacroError").Delete
    On Error GoTo ErrHandler

    If Application.Interactive Then
        MsgBox "Schedule generated successfully!", vbInformation, "Generate Schedule"
    End If
    GoTo Cleanup

ErrHandler:
    ' Capture Err state before any further VBA op can clobber it.
    Dim errNum As Long: errNum = Err.Number
    Dim errDesc As String: errDesc = Err.Description
    Dim errMsg As String
    errMsg = "Err " & errNum & " | " & errDesc & _
             " | section=" & curSection & " | curRow=" & curRow
    ' Always log to a defined name so COM-driven callers (which force
    ' Application.Interactive=False and thus suppress the MsgBox) can detect
    ' the failure. Read via ThisWorkbook.Names("LastMacroError").RefersTo.
    On Error Resume Next
    ThisWorkbook.Names.Add Name:="LastMacroError", _
        RefersTo:="=""" & Replace(errMsg, """", """""") & """", _
        Visible:=False
    On Error GoTo 0
    If Application.Interactive Then
        MsgBox "Error " & errNum & ": " & errDesc & vbNewLine & _
               "Section: " & curSection & vbNewLine & _
               "At line near row " & curRow, vbCritical
    End If
Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Public Sub ExportToPDF()
    On Error GoTo PdfErr
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("PAYMENTS")
    Dim savePath As String
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Invoice_Schedule_" & Format(Date, "yyyy-mm-dd"), _
        FileFilter:="PDF Files (*.pdf), *.pdf")
    If savePath = "False" Then Exit Sub
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=savePath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    MsgBox "PDF exported to:" & vbNewLine & savePath, vbInformation
    Exit Sub
PdfErr:
    MsgBox "Error exporting PDF: " & Err.Description, vbCritical
End Sub

' ========================= SORTING HELPERS ==================================

Private Sub SortDataByDate(arr() As Variant)
    ' Bubble sort by (Type, Date, Company)
    ' Date-first within type ensures Misc Fees, Client Supply, Cash Flow In
    ' are date-ordered. Consultant/Council/Contractor sections pick by company
    ' via WriteStandardSection so company order doesn't matter for them.
    Dim n As Long: n = UBound(arr, 1)
    Dim i As Long, j As Long, c As Long
    Dim swap As Boolean
    For i = 1 To n - 1
        For j = 1 To n - i
            swap = False
            ' Compare type first
            If CStr(arr(j, 1)) > CStr(arr(j + 1, 1)) Then
                swap = True
            ElseIf CStr(arr(j, 1)) = CStr(arr(j + 1, 1)) Then
                ' Same type: compare date first
                Dim d1 As Boolean, d2 As Boolean
                d1 = IsDate(arr(j, 3)): d2 = IsDate(arr(j + 1, 3))
                If d1 And d2 Then
                    If CDate(arr(j, 3)) > CDate(arr(j + 1, 3)) Then
                        swap = True
                    ElseIf CDate(arr(j, 3)) = CDate(arr(j + 1, 3)) Then
                        ' Same date: compare company
                        If CStr(arr(j, 2)) > CStr(arr(j + 1, 2)) Then swap = True
                    End If
                ElseIf d1 And Not d2 Then
                    swap = True  ' rows with dates after rows without
                End If
            End If
            If swap Then
                For c = 1 To UBound(arr, 2)
                    Dim tmp As Variant: tmp = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = tmp
                Next c
            End If
        Next j
    Next i
End Sub

Private Sub SortFeesByDate(arr() As Variant)
    ' Stable bubble sort of COMMITTED rows (Update 3 layout) on a
    ' lexicographic key, compared by FeeAfter:
    '   group 0 (Consultant, Council, Contractor, Cash Flow In, other):
    '     (Company col 3, Date col 6), undated before dated
    '   group 1 (pooled sections, Misc Fees and Client Supply), after group 0:
    '     (Date col 6) only, undated first, company ignored (Update 4 item 14:
    '     pooled committed items render in Date Appointed order)
    ' Equal keys never swap, so ties keep entry order. Group-0 rows keep
    ' their relative order, and a company's group-0 rows now precede its
    ' pooled rows, so the first-row-wins retention and discipline maps
    ' (built after this) read the Consultant/Contractor row first.
    Dim n As Long: n = UBound(arr, 1)
    Dim i As Long, j As Long, c As Long
    Dim swap As Boolean
    For i = 1 To n - 1
        For j = 1 To n - i
            swap = FeeAfter(arr, j)
            If swap Then
                For c = 1 To UBound(arr, 2)
                    Dim tmpF As Variant: tmpF = arr(j, c)
                    arr(j, c) = arr(j + 1, c)
                    arr(j + 1, c) = tmpF
                Next c
            End If
        Next j
    Next i
End Sub

Private Function FeeAfter(arr As Variant, j As Long) As Boolean
    ' True when COMMITTED row j must move below row j+1 (see SortFeesByDate).
    ' Strict comparisons only: equal keys return False (stable sort).
    Dim g1 As Long, g2 As Long
    g1 = 0: g2 = 0
    If CStr(arr(j, 1)) = "Misc Fees" Or CStr(arr(j, 1)) = "Client Supply" Then g1 = 1
    If CStr(arr(j + 1, 1)) = "Misc Fees" Or CStr(arr(j + 1, 1)) = "Client Supply" Then g2 = 1
    FeeAfter = False
    If g1 <> g2 Then
        FeeAfter = (g1 > g2)  ' group 0 first
        Exit Function
    End If
    If g1 = 0 Then
        ' Group 0: company first
        If CStr(arr(j, 3)) > CStr(arr(j + 1, 3)) Then
            FeeAfter = True
            Exit Function
        ElseIf CStr(arr(j, 3)) < CStr(arr(j + 1, 3)) Then
            Exit Function
        End If
    End If
    ' Same company (group 0) or pooled (group 1): date col 6, undated first
    Dim d1Has As Boolean, d2Has As Boolean
    d1Has = IsDate(arr(j, 6))
    d2Has = IsDate(arr(j + 1, 6))
    If d1Has And d2Has Then
        If CDate(arr(j, 6)) > CDate(arr(j + 1, 6)) Then FeeAfter = True
    ElseIf d1Has And Not d2Has Then
        FeeAfter = True  ' dated items after undated
    End If
End Function

' ========================= COMPANY HELPERS ==================================

Private Function GetAllCompaniesForType(dataArr As Variant, feesArr As Variant, _
    hasFees As Boolean, typFilter As String, excludeCompany As String) As Variant
    ' Returns sorted unique company names from both data and fees for a given type
    ' Excludes excludeCompany (e.g. "MilnePM")
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim i As Long, comp As String

    ' From data
    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = typFilter Then
            comp = CStr(dataArr(i, 2))
            If Len(excludeCompany) = 0 Or comp <> excludeCompany Then
                If Not dict.Exists(comp) Then dict.Add comp, 0
            End If
        End If
    Next i

    ' From fees (TYPE in col 1, COMPANY in col 3 — Update-3 COMMITTED layout)
    If hasFees Then
        For i = 1 To UBound(feesArr, 1)
            If CStr(feesArr(i, 1)) = typFilter Then
                comp = CStr(feesArr(i, 3))
                If Len(excludeCompany) = 0 Or comp <> excludeCompany Then
                    If Not dict.Exists(comp) Then dict.Add comp, 0
                End If
            End If
        Next i
    End If

    If dict.Count = 0 Then
        GetAllCompaniesForType = Empty
        Exit Function
    End If

    ' Copy to array and sort alphabetically
    Dim result() As String: ReDim result(0 To dict.Count - 1)
    Dim key As Variant, idx As Long: idx = 0
    For Each key In dict.Keys: result(idx) = CStr(key): idx = idx + 1: Next key

    ' Bubble sort A-Z
    Dim j As Long, swapStr As String
    For i = 0 To UBound(result) - 1
        For j = 0 To UBound(result) - i - 1
            If result(j) > result(j + 1) Then
                swapStr = result(j): result(j) = result(j + 1): result(j + 1) = swapStr
            End If
        Next j
    Next i

    GetAllCompaniesForType = result
End Function

' "QS | Brown Project Services" when a discipline is entered on COMMITTED
' col G for the company; the bare name otherwise. Display-only — sorting,
' lookups, and totals all keep using the bare company name.
Private Function DisplayCompanyName(compName As String) As String
    DisplayCompanyName = compName
    If mCompDisc Is Nothing Then Exit Function
    If mCompDisc.Exists(compName) Then
        DisplayCompanyName = mCompDisc(compName) & " | " & compName
    End If
End Function

Private Function HasCompanyData(dataArr As Variant, typFilter As String, compName As String) As Boolean
    Dim i As Long
    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = typFilter And CStr(dataArr(i, 2)) = compName Then
            HasCompanyData = True: Exit Function
        End If
    Next i
    HasCompanyData = False
End Function

Private Function HasCompanyFees(feesArr As Variant, hasFees As Boolean, _
    typFilter As String, compName As String) As Boolean
    If Not hasFees Then HasCompanyFees = False: Exit Function
    Dim i As Long
    For i = 1 To UBound(feesArr, 1)
        If CStr(feesArr(i, 1)) = typFilter And CStr(feesArr(i, 3)) = compName Then
            HasCompanyFees = True: Exit Function
        End If
    Next i
    HasCompanyFees = False
End Function

Private Function CountTypeFees(feesArr As Variant, hasFees As Boolean, typFilter As String) As Long
    If Not hasFees Then CountTypeFees = 0: Exit Function
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = 1 To UBound(feesArr, 1)
        If CStr(feesArr(i, 1)) = typFilter Then cnt = cnt + 1
    Next i
    CountTypeFees = cnt
End Function

Private Function CountCompanyFees(feesArr As Variant, hasFees As Boolean, _
    typFilter As String, compName As String) As Long
    If Not hasFees Then CountCompanyFees = 0: Exit Function
    Dim cnt As Long: cnt = 0
    Dim i As Long
    For i = 1 To UBound(feesArr, 1)
        If CStr(feesArr(i, 1)) = typFilter And CStr(feesArr(i, 3)) = compName Then cnt = cnt + 1
    Next i
    CountCompanyFees = cnt
End Function

' ========================= SECTION WRITERS ==================================

Private Sub WriteStandardSection(ws As Worksheet, dataArr As Variant, feesArr As Variant, _
    hasFees As Boolean, compName As String, typFilter As String, _
    ByRef curRow As Long, ByRef pageStartRow As Long, _
    compRowCounts As Object, trackAsFeeInvoice As Boolean, _
    Optional ByRef totalFeeRowOut As Long = 0)
    '
    ' Writes a company section: company name, invoices (date-sorted), committed fees,
    ' total fee, percentage row. Used for Consultant, Council, Cash Flow In.
    '
    Dim invoiceCnt As Long: invoiceCnt = 0
    Dim ckLookup As String: ckLookup = typFilter & "|" & compName
    If compRowCounts.Exists(ckLookup) Then invoiceCnt = compRowCounts(ckLookup)

    Dim feeCnt As Long: feeCnt = CountCompanyFees(feesArr, hasFees, typFilter, compName)
    Dim bodyRows As Long
    If invoiceCnt > feeCnt + 1 Then bodyRows = invoiceCnt Else bodyRows = feeCnt + 1
    If bodyRows < 1 Then bodyRows = 1

    ' Section size: 1 (company name) + bodyRows + 1 (percentage) + 1 (gap) + 1 (header row)
    Dim sectionRows As Long: sectionRows = 1 + bodyRows + 1 + 1 + 1

    ' Look-ahead page break (with 1-row buffer at bottom of page)
    If (curRow - pageStartRow) + sectionRows > ROWS_PER_PAGE And curRow > pageStartRow + 2 Then
        curRow = curRow + 1  ' buffer row at bottom of page
        ws.HPageBreaks.Add Before:=ws.Cells(curRow, 1)
        WriteInvoiceHeader ws, curRow
        pageStartRow = curRow
    End If

    ' Company name row
    curRow = curRow + 1
    ws.Cells(curRow, 1).Value = DisplayCompanyName(compName)
    ws.Cells(curRow, 1).Font.Bold = True
    FormatCompanyRow ws, curRow
    curRow = curRow + 1

    ' Invoice data rows
    Dim invoiceStartRow As Long: invoiceStartRow = curRow
    Dim invCount As Long: invCount = 0
    Dim hasUnpaid As Boolean: hasUnpaid = False
    Dim lastInvRow As Long: lastInvRow = 0
    Dim i As Long

    For i = 1 To UBound(dataArr, 1)
        If CStr(dataArr(i, 1)) = typFilter And CStr(dataArr(i, 2)) = compName Then
            Dim isPaid As Boolean
            isPaid = (UCase(Trim(CStr(dataArr(i, 6)))) = "Y")
            If Not isPaid Then
                hasUnpaid = True
                ws.Cells(curRow, 1).Value = "Invoice recommended for payment"
                ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, 7)).Interior.Color = UNPAID_FILL
            End If
            ws.Cells(curRow, 2).Value = dataArr(i, 3)
            ws.Cells(curRow, 2).NumberFormatLocal =DATE_FMT
            ws.Cells(curRow, 3).NumberFormatLocal ="@"
            ws.Cells(curRow, 3).Value = CStr(dataArr(i, 4))
            ws.Cells(curRow, 4).Value = dataArr(i, 5)
            ws.Cells(curRow, 4).NumberFormatLocal =MONEY_FMT
            ' Cumulative ex GST
            ws.Cells(curRow, 5).Formula = "=" & ws.Cells(curRow - 1, 5).Address(False, False) & "+D" & curRow
            ws.Cells(curRow, 5).NumberFormatLocal =MONEY_FMT
            ' Amount inc GST
            ws.Cells(curRow, 6).Formula = "=D" & curRow & "*" & GST_RATE
            ws.Cells(curRow, 6).NumberFormatLocal =MONEY_FMT
            ' Cumulative inc GST
            ws.Cells(curRow, 7).Formula = "=" & ws.Cells(curRow - 1, 7).Address(False, False) & "+F" & curRow
            ws.Cells(curRow, 7).NumberFormatLocal =MONEY_FMT
            If isPaid Then ws.Cells(curRow, 11).Value = "Y"
            ' Write comment from column G if present
            If UBound(dataArr, 2) >= 7 Then
                Dim cmnt As String: cmnt = Trim(CStr(dataArr(i, 7)))
                If Len(cmnt) > 0 Then
                    Dim curA As String: curA = Trim(CStr(ws.Cells(curRow, 1).Value))
                    If Len(curA) > 0 Then
                        ws.Cells(curRow, 1).Value = curA & " - " & cmnt
                    Else
                        ws.Cells(curRow, 1).Value = cmnt
                    End If
                End If
            End If
            FormatRowBase ws, curRow
            FormatDataRow ws, curRow
            ' Green fill for Cash Flow In paid rows (unpaid already has blue)
            If Not trackAsFeeInvoice And isPaid Then
                ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, 7)).Interior.Color = CASHFLOW_FILL
            End If
            ' Track fee data rows
            If trackAsFeeInvoice Then
                If mFirstFeeDataRow = 0 Then mFirstFeeDataRow = curRow
                mLastFeeDataRow = curRow
            Else
                ' Cash Flow In tracking
                If mCfStartRow = 0 Then mCfStartRow = curRow
                mCfEndRow = curRow
            End If
            lastInvRow = curRow
            invCount = invCount + 1
            curRow = curRow + 1
        End If
    Next i

    ' Percentage row — placed right after last invoice, not after fees
    Dim pctRow As Long: pctRow = curRow
    If trackAsFeeInvoice Then
        ws.Cells(pctRow, 1).Value = "Percentage paid to date"
    ElseIf InStr(1, compName, "Ministry", vbTextCompare) > 0 Then
        ws.Cells(pctRow, 1).Value = "Current drawdown as percentage of total " & compName & " funding"
    Else
        ws.Cells(pctRow, 1).Value = "Percentage received to date"
    End If
    ws.Cells(pctRow, 1).Font.Name = FONT_NAME
    ws.Cells(pctRow, 1).Font.Size = FONT_SIZE

    ' Write committed fee items (column I/J, starting at invoiceStartRow)
    Dim feeRow As Long: feeRow = invoiceStartRow
    Dim feeWritten As Long: feeWritten = 0
    Dim maxFeeRow As Long: maxFeeRow = curRow - 1  ' last invoice row
    WriteFeeBlock ws, feesArr, hasFees, typFilter, compName, feeRow, feeWritten, maxFeeRow

    ' Total fee (always present, in column I/J)
    Dim totalFeeRow As Long
    totalFeeRow = invoiceStartRow + feeWritten
    If trackAsFeeInvoice Then
        WriteTotalAgreedFee ws, invoiceStartRow, feeRow, feeWritten
    Else
        WriteTotalAgreedFee ws, invoiceStartRow, feeRow, feeWritten, "Total"
    End If

    ' Committed-fee rows must sit inside the grand-total span too. The invoice
    ' loop above only extends the trackers when it writes an invoice row, so a
    ' company with COMMITTED items but no invoices yet (common early in a
    ' project) would leave its col-I fees outside [mFirstFeeDataRow,
    ' mLastFeeDataRow] whenever it sorts before the first invoiced company or
    ' after the last — silently dropping it from "Total committed". The
    ' per-company "Total fee" row is safe to exclude: it's a nested SUBTOTAL,
    ' which the grand total's SUBTOTAL(9,...) ignores anyway.
    If trackAsFeeInvoice And feeWritten > 0 Then
        If mFirstFeeDataRow = 0 Or invoiceStartRow < mFirstFeeDataRow Then mFirstFeeDataRow = invoiceStartRow
        If totalFeeRow - 1 > mLastFeeDataRow Then mLastFeeDataRow = totalFeeRow - 1
    End If

    ' If no invoices and no fees, ensure at least one body row for the total
    If invCount = 0 And feeWritten = 0 Then
        totalFeeRow = invoiceStartRow
    End If

    ' Output the total fee row for caller tracking
    totalFeeRowOut = totalFeeRow

    ' Track cash flow total fee row (legacy single-company fallback)
    If Not trackAsFeeInvoice Then
        mCfTotalFeeRow = totalFeeRow
    End If

    ' Finish percentage formula (needs totalFeeRow which is now known)
    If invCount > 0 Then
        ws.Cells(pctRow, 5).Formula = "=IF(I" & totalFeeRow & "=0,0," & _
            ws.Cells(lastInvRow, 5).Address(False, False) & "/I" & totalFeeRow & ")"
    Else
        ws.Cells(pctRow, 5).Value = 0
    End If
    ws.Cells(pctRow, 5).NumberFormatLocal =PCT_FMT
    FormatPercentageRow ws, pctRow

    ' curRow = max of (pctRow + 1, totalFeeRow + 1) to cover both invoice+percentage and fee items
    curRow = pctRow + 1
    If totalFeeRow + 1 > curRow Then curRow = totalFeeRow + 1
End Sub

Private Sub WriteFeeBlock(ws As Worksheet, feesArr As Variant, hasFees As Boolean, _
    typFilter As String, compFilter As String, _
    ByRef feeRow As Long, ByRef feeWritten As Long, maxRow As Long)
    '
    ' Writes committed fee items in columns I/J starting at feeRow.
    ' If compFilter is empty, matches all companies for the type (for Misc/CS).
    ' feeRow advances. feeWritten counts items written.
    '
    If Not hasFees Then Exit Sub
    Dim fIdx As Long
    For fIdx = 1 To UBound(feesArr, 1)
        Dim typeMatch As Boolean: typeMatch = (CStr(feesArr(fIdx, 1)) = typFilter)
        Dim compMatch As Boolean
        If Len(compFilter) > 0 Then
            compMatch = (CStr(feesArr(fIdx, 3)) = compFilter)
        Else
            compMatch = True
        End If
        If typeMatch And compMatch Then
            ' Write fee item (COMMITTED cols: TYPE|DISCIPLINE|COMPANY|FEE ITEM|AMOUNT|DATE|RETENTION)
            ws.Cells(feeRow, 9).Value = feesArr(fIdx, 5)
            ws.Cells(feeRow, 9).NumberFormatLocal =MONEY_FMT_WHOLE
            ws.Cells(feeRow, 10).Value = CStr(feesArr(fIdx, 4))
            ws.Cells(feeRow, 9).Interior.Color = FEE_FILL
            ws.Cells(feeRow, 9).HorizontalAlignment = xlRight
            ' Borders: left+right on all, top on first
            ws.Cells(feeRow, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
            ws.Cells(feeRow, 9).Borders(xlEdgeLeft).Weight = xlThin
            ws.Cells(feeRow, 9).Borders(xlEdgeRight).LineStyle = xlContinuous
            ws.Cells(feeRow, 9).Borders(xlEdgeRight).Weight = xlThin
            If feeWritten = 0 Then
                ws.Cells(feeRow, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
                ws.Cells(feeRow, 9).Borders(xlEdgeTop).Weight = xlThin
            End If
            ' Fee-only rows: no A-G borders needed (only I/J)
            feeWritten = feeWritten + 1
            feeRow = feeRow + 1
        End If
    Next fIdx

    ' Last fee item before total: double bottom border, keep grey fill
    If feeWritten > 0 Then
        ws.Cells(feeRow - 1, 9).Borders(xlEdgeBottom).LineStyle = xlDouble
        ws.Cells(feeRow - 1, 9).Borders(xlEdgeBottom).Weight = xlThick
    End If
End Sub

Private Sub WriteTotalAgreedFee(ws As Worksheet, startRow As Long, _
    ByRef feeRow As Long, feeWritten As Long, _
    Optional totalLabel As String = "Total fee")
    '
    ' Writes the total fee row at startRow + feeWritten.
    ' Always written, even if feeWritten = 0 (shows $0).
    '
    Dim totalRow As Long: totalRow = startRow + feeWritten
    If feeWritten > 0 Then
        ws.Cells(totalRow, 9).Formula = "=SUBTOTAL(9,I" & startRow & ":I" & (totalRow - 1) & ")"
    Else
        ws.Cells(totalRow, 9).Value = 0
    End If
    ws.Cells(totalRow, 9).NumberFormatLocal =MONEY_FMT_WHOLE
    ws.Cells(totalRow, 9).HorizontalAlignment = xlRight
    ws.Cells(totalRow, 9).Font.Bold = True
    ws.Cells(totalRow, 9).Interior.Color = FEE_FILL
    ws.Cells(totalRow, 10).Value = totalLabel
    ws.Cells(totalRow, 10).Font.Bold = True
    ' Borders: left+right+bottom
    ws.Cells(totalRow, 9).Borders(xlEdgeLeft).LineStyle = xlContinuous
    ws.Cells(totalRow, 9).Borders(xlEdgeLeft).Weight = xlThin
    ws.Cells(totalRow, 9).Borders(xlEdgeRight).LineStyle = xlContinuous
    ws.Cells(totalRow, 9).Borders(xlEdgeRight).Weight = xlThin
    ws.Cells(totalRow, 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(totalRow, 9).Borders(xlEdgeBottom).Weight = xlThin
    If feeWritten = 0 Then
        ' Also add top border when there are no fee items above
        ws.Cells(totalRow, 9).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Cells(totalRow, 9).Borders(xlEdgeTop).Weight = xlThin
    End If
End Sub

' ========================= FORMATTING HELPERS ===============================

Private Sub FormatRowBase(ws As Worksheet, rowNum As Long)
    Dim c As Long
    For c = 1 To 10
        With ws.Cells(rowNum, c)
            .Font.Name = FONT_NAME: .Font.Size = FONT_SIZE: .VerticalAlignment = xlCenter
            If c >= 1 And c <= 7 Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
            End If
        End With
    Next c
End Sub

Private Sub FormatCompanyRow(ws As Worksheet, rowNum As Long)
    Dim c As Long
    For c = 1 To 10
        With ws.Cells(rowNum, c)
            .Font.Name = FONT_NAME: .Font.Size = FONT_SIZE: .VerticalAlignment = xlCenter
            If c = 1 Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
            End If
        End With
    Next c
End Sub

Private Sub FormatPercentageRow(ws As Worksheet, rowNum As Long)
    Dim c As Long
    For c = 1 To 10
        ws.Cells(rowNum, c).Font.Name = FONT_NAME
        ws.Cells(rowNum, c).Font.Size = FONT_SIZE
        ws.Cells(rowNum, c).VerticalAlignment = xlCenter
    Next c
    ' A: left + top + bottom
    ws.Cells(rowNum, 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
    ws.Cells(rowNum, 1).Borders(xlEdgeLeft).Weight = xlThin
    ws.Cells(rowNum, 1).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Cells(rowNum, 1).Borders(xlEdgeTop).Weight = xlThin
    ws.Cells(rowNum, 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(rowNum, 1).Borders(xlEdgeBottom).Weight = xlThin
    ' B: top + bottom
    ws.Cells(rowNum, 2).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Cells(rowNum, 2).Borders(xlEdgeTop).Weight = xlThin
    ws.Cells(rowNum, 2).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(rowNum, 2).Borders(xlEdgeBottom).Weight = xlThin
    ' C: top + bottom
    ws.Cells(rowNum, 3).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Cells(rowNum, 3).Borders(xlEdgeTop).Weight = xlThin
    ws.Cells(rowNum, 3).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(rowNum, 3).Borders(xlEdgeBottom).Weight = xlThin
    ' D: top + bottom
    ws.Cells(rowNum, 4).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Cells(rowNum, 4).Borders(xlEdgeTop).Weight = xlThin
    ws.Cells(rowNum, 4).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(rowNum, 4).Borders(xlEdgeBottom).Weight = xlThin
    ' E: full box around percentage value
    ws.Cells(rowNum, 5).Borders(xlEdgeLeft).LineStyle = xlContinuous
    ws.Cells(rowNum, 5).Borders(xlEdgeLeft).Weight = xlThin
    ws.Cells(rowNum, 5).Borders(xlEdgeRight).LineStyle = xlContinuous
    ws.Cells(rowNum, 5).Borders(xlEdgeRight).Weight = xlThin
    ws.Cells(rowNum, 5).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Cells(rowNum, 5).Borders(xlEdgeTop).Weight = xlThin
    ws.Cells(rowNum, 5).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(rowNum, 5).Borders(xlEdgeBottom).Weight = xlThin
End Sub

Private Sub FormatDataRow(ws As Worksheet, rowNum As Long)
    ws.Cells(rowNum, 2).HorizontalAlignment = xlCenter
    ws.Cells(rowNum, 3).HorizontalAlignment = xlCenter
    ws.Cells(rowNum, 4).HorizontalAlignment = xlRight
    ws.Cells(rowNum, 5).HorizontalAlignment = xlRight
    ws.Cells(rowNum, 6).HorizontalAlignment = xlRight
    ws.Cells(rowNum, 7).HorizontalAlignment = xlRight
End Sub

' ========================= HEADER WRITERS ===================================

Private Sub WriteInvoiceHeader(ws As Worksheet, rowNum As Long)
    WriteHeader ws, rowNum, "COMPANY", "DATE", "INV No.", _
                "AMOUNT" & vbLf & "ex GST", _
                "CUMULATIVE" & vbLf & "ex GST", _
                "AMOUNT" & vbLf & "inc GST", _
                "CUMULATIVE" & vbLf & "inc GST", _
                "", "COMMITTED" & vbLf & "AMOUNT"
End Sub

Private Sub WriteContractorHeader(ws As Worksheet, rowNum As Long, contrName As String)
    Dim colAText As String
    If Len(contrName) > 0 Then
        colAText = "Contractor - " & contrName
    Else
        colAText = "CONTRACTOR"
    End If
    WriteHeader ws, rowNum, colAText, "DATE", _
                "PROGRESS PAYMENT NO.", "VALUE OF WORK COMPLETED", "RETENTIONS", _
                "VALUE LESS RETENTIONS LESS PREVIOUS", "PAYMENT DUE INCL GST", "", "COMMITTED" & vbLf & "AMOUNT"
    ws.Rows(rowNum).RowHeight = 43.2  ' 72 pixels for contractor headers
End Sub

Private Sub WriteSummaryHeader(ws As Worksheet, rowNum As Long)
    ' Summary uses columns A:B (merged), C, D, E
    ws.Range("A" & rowNum & ":B" & rowNum).Merge
    Dim sumCols As Variant: sumCols = Array(1, 3, 4, 5)
    Dim sumHeaders As Variant: sumHeaders = Array("INVOICE SUMMARY", "APPROVED BUDGET", "TOTAL COMMITTED", "INVOICES TO DATE")
    Dim sc As Long
    For sc = LBound(sumCols) To UBound(sumCols)
        Dim col As Long: col = CLng(sumCols(sc))
        With ws.Cells(rowNum, col)
            .Value = CStr(sumHeaders(sc))
            .Font.Name = FONT_NAME: .Font.Size = FONT_SIZE: .Font.Bold = True
            .Font.Color = BRAND_NAVY
            .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
            .Interior.Color = HEADER_FILL
        End With
    Next sc
    ' Borders on merged A:B range (top border is brand aqua accent)
    With ws.Range("A" & rowNum & ":B" & rowNum)
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
    End With
    ' Borders on C, D, E (top border is brand aqua accent)
    Dim nc As Long
    For nc = 3 To 5
        With ws.Cells(rowNum, nc)
            .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
        End With
    Next nc
    ws.Rows(rowNum).RowHeight = 24
End Sub

Private Sub FormatSummaryRow(ws As Worksheet, rowNum As Long, isHighlight As Boolean)
    ' Format a summary data row — A:B merged, C, D, E with borders
    ws.Range("A" & rowNum & ":B" & rowNum).Merge
    With ws.Range("A" & rowNum & ":B" & rowNum)
        .Font.Name = FONT_NAME: .Font.Size = FONT_SIZE: .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
        If isHighlight Then
            .Interior.Color = HEADER_FILL
            .Font.Color = BRAND_NAVY
            .Font.Bold = True
        End If
    End With
    Dim nc As Long
    For nc = 3 To 5
        With ws.Cells(rowNum, nc)
            .Font.Name = FONT_NAME: .Font.Size = FONT_SIZE: .VerticalAlignment = xlCenter
            .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
            .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
            .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
            .HorizontalAlignment = xlRight
            .ShrinkToFit = True   ' summary totals fit at any zoom; these rows never wrap
            If isHighlight Then
                .Interior.Color = HEADER_FILL
                .Font.Color = BRAND_NAVY
            End If
        End With
    Next nc
End Sub

Private Sub WriteHeader(ws As Worksheet, rowNum As Long, ParamArray headers() As Variant)
    Dim i As Long, col As Long: col = 1
    For i = LBound(headers) To UBound(headers)
        With ws.Cells(rowNum, col)
            If col >= 1 And col <= 7 Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
            ElseIf col = 9 Then
                .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeLeft).Weight = xlThin
                .Borders(xlEdgeRight).LineStyle = xlContinuous: .Borders(xlEdgeRight).Weight = xlThin
                .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous: .Borders(xlEdgeBottom).Weight = xlThin
            End If
            If Len(CStr(headers(i))) > 0 Then
                .Value = CStr(headers(i))
                .Font.Name = FONT_NAME: .Font.Size = FONT_SIZE: .Font.Bold = True
                .Font.Color = BRAND_NAVY
                .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
                .Interior.Color = HEADER_FILL
            End If
        End With
        col = col + 1
    Next i
    ws.Rows(rowNum).RowHeight = 24
End Sub
