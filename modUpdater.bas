Attribute VB_Name = "modUpdater"
'=============================================================================
' MODULE: modUpdater
' PURPOSE: Self-update mechanism for the Invoice Schedule macro.
'          Fetches modGenerateSchedule.bas from a public GitHub URL and swaps
'          in the server's version whenever it differs from the installed one.
'          Treats GitHub as source of truth (allows rollbacks, not just forward
'          updates). MODULE_VERSION is purely metadata for display, not control.
'
' DESIGN:  This module is INTENTIONALLY IMMUTABLE. It never updates itself.
'          If we ever need to evolve it, that requires a separate manual
'          re-install. Keep this module's scope narrow and predictable.
'
' SHIPPED: Bundled with Update 2 (2026-05). One-time install per workbook.
'=============================================================================

Option Explicit

' Bump only if we ever ship a new version of the updater itself (rare).
Public Const UPDATER_VERSION As String = "1.3"

' Source of truth for the latest modGenerateSchedule.bas.
Private Const UPDATE_URL As String = "https://raw.githubusercontent.com/JJ-San/scheduler-bas-update/main/modGenerateSchedule.bas"
Private Const TARGET_MODULE As String = "modGenerateSchedule"
Private Const TARGET_SUB As String = "Sub GenerateSchedule"
Private Const BUTTON_CAPTION As String = "Check for Updates"
Private Const BUTTON_NAME As String = "btnCheckUpdates"

' Where the human-visible version label is shown on REPORT_SETTINGS. modUpdater
' writes here DIRECTLY after a successful update so the cell refreshes
' immediately, without needing the just-imported modGenerateSchedule to be
' compiled (Application.Run can't find subs in newly-imported modules until
' VBA recompiles). modGenerateSchedule.WriteVersionLabel also writes here on
' every macro run — same cell, two writers, eventually consistent.
' If the label cell ever moves, update this constant.
Private Const VERSION_LABEL_CELL As String = "D10"


'=============================================================================
' PUBLIC: CheckForUpdates
'   Wired to the button on REPORT_SETTINGS. Full update flow with friendly
'   user-facing messages and a backup/rollback path if anything goes wrong.
'=============================================================================
Public Sub CheckForUpdates()
    Dim installedVer As String
    Dim remoteText As String
    Dim remoteVer As String
    Dim backupPath As String
    Dim tmpPath As String

    ' 1. Trust access check — cheap, fail fast before any network call.
    If Not VBProjectAccessible() Then
        MsgBox "Excel needs a one-time permission to install updates. " & _
               "Please refer to the installation manual for setup steps, " & _
               "then click Check for Updates again.", _
               vbInformation, "One-time setup needed"
        Exit Sub
    End If

    installedVer = ReadInstalledVersion()

    remoteText = FetchRemote(UPDATE_URL)
    If Len(remoteText) = 0 Then
        MsgBox "Couldn't reach the update server. " & _
               "Check your internet and try again in a few minutes.", _
               vbExclamation, "Can't reach the server"
        Exit Sub
    End If

    remoteVer = ParseVersion(remoteText)
    If Len(remoteVer) = 0 Then
        MsgBox "The update file doesn't look right. Try again later. " & _
               "If it keeps happening, please let your IT or developer know.", _
               vbCritical, "Hmm, something's off"
        Exit Sub
    End If

    ' Equality compare, not semver. GitHub is source of truth: any mismatch
    ' (forward update OR rollback) prompts the user to apply the server's version.
    If installedVer = remoteVer Then
        MsgBox "Your workbook is on the latest version (" & installedVer & ").", _
               vbInformation, "You're up to date"
        Exit Sub
    End If

    If MsgBox("A different version of the schedule generator is ready." & vbCrLf & vbCrLf & _
              "Your version:  " & installedVer & vbCrLf & _
              "New version:   " & remoteVer & vbCrLf & vbCrLf & _
              "Install it now?", _
              vbYesNo + vbQuestion, "Update available") <> vbYes Then
        Exit Sub
    End If

    ' Backup + swap + verify, with rollback on any failure between here and success.
    backupPath = BackupModule()
    tmpPath = WriteTempBas(remoteText)

    On Error GoTo Rollback
    RemoveTargetModules
    ThisWorkbook.VBProject.VBComponents.Import tmpPath
    If Not VerifyImport() Then
        Err.Raise vbObjectError + 1, "CheckForUpdates", "Verify failed after import"
    End If
    On Error GoTo 0

    ' Refresh the on-sheet version label immediately so the user sees the new
    ' version while still looking at the success dialog. Direct write — does NOT
    ' depend on the just-imported modGenerateSchedule being compiled (it isn't
    ' yet, so Application.Run "WriteVersionLabel" silently fails here).
    On Error Resume Next
    ThisWorkbook.Worksheets("REPORT_SETTINGS").Range(VERSION_LABEL_CELL).Value = _
        "Workbook Version: " & remoteVer
    On Error GoTo 0

    MsgBox "Your workbook is now on version " & remoteVer & "." & vbCrLf & vbCrLf & _
           "Press Ctrl+S to save and keep this update.", _
           vbInformation, "Done!"
    Exit Sub

Rollback:
    On Error Resume Next
    RemoveTargetModules
    ThisWorkbook.VBProject.VBComponents.Import backupPath
    On Error GoTo 0
    MsgBox "Something went wrong, so your original version is still in place. " & _
           "Try again in a moment. If it keeps failing, " & _
           "let your IT or developer know.", _
           vbCritical, "Couldn't update"
End Sub


'=============================================================================
' PUBLIC: InstallButton
'   JOSIAH-SIDE TOOL — not run by end users.
'   Called from tmp/prep_update2_workbook.py during workbook prep.
'   Places the "Check for Updates" Form button on REPORT_SETTINGS.
'   Idempotent — removes any existing button with our name first.
'=============================================================================
Public Sub InstallButton()
    Dim ws As Worksheet
    Dim btn As Shape
    Dim sh As Shape
    Dim anchor As Range

    Set ws = ThisWorkbook.Worksheets("REPORT_SETTINGS")

    ' Idempotent: drop any prior button by our name.
    For Each sh In ws.Shapes
        If sh.Name = BUTTON_NAME Then
            sh.Delete
            Exit For
        End If
    Next sh

    Set anchor = ws.Range("C21")
    Set btn = ws.Shapes.AddFormControl( _
        Type:=xlButtonControl, _
        Left:=anchor.Left, _
        Top:=anchor.Top, _
        Width:=120, _
        Height:=28)

    btn.Name = BUTTON_NAME
    btn.OLEFormat.Object.Caption = BUTTON_CAPTION
    btn.OnAction = "'" & ThisWorkbook.Name & "'!CheckForUpdates"

    ' Guarded — silent under COM automation, vocal when run manually via Alt+F8.
    If Application.Interactive Then
        MsgBox "Update button installed on REPORT_SETTINGS at " & anchor.Address(False, False) & "." & vbCrLf & _
               "Save the workbook to keep it.", _
               vbInformation, "Install Complete"
    End If
End Sub


'=============================================================================
' PRIVATE HELPERS
'=============================================================================

' True if we can read the VBProject (needed to swap modules).
' False if "Trust access to the VBA project object model" is disabled.
Private Function VBProjectAccessible() As Boolean
    Dim n As Long
    On Error Resume Next
    n = ThisWorkbook.VBProject.VBComponents.Count
    VBProjectAccessible = (Err.Number = 0 And n > 0)
    On Error GoTo 0
End Function

' Reads MODULE_VERSION from the currently-loaded modGenerateSchedule.
' Returns "0.0" if not found (treats legacy workbooks as out-of-date).
Private Function ReadInstalledVersion() As String
    Dim cmp As Object
    Dim code As String
    On Error Resume Next
    Set cmp = ThisWorkbook.VBProject.VBComponents(TARGET_MODULE)
    On Error GoTo 0
    If cmp Is Nothing Then
        ReadInstalledVersion = "0.0"
        Exit Function
    End If
    code = cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines)
    ReadInstalledVersion = ParseVersion(code)
    If Len(ReadInstalledVersion) = 0 Then ReadInstalledVersion = "0.0"
End Function

' Parses Public Const MODULE_VERSION As String = "X.Y" out of a text blob.
' Tolerant of whitespace; case-insensitive on keywords.
Private Function ParseVersion(text As String) As String
    Dim re As Object
    Dim m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "Public\s+Const\s+MODULE_VERSION\s+As\s+String\s*=\s*""([^""]+)"""
    re.IgnoreCase = True
    re.Global = False
    Set m = re.Execute(text)
    If m.Count > 0 Then
        ParseVersion = m(0).SubMatches(0)
    Else
        ParseVersion = ""
    End If
End Function

' Fetches a URL via MSXML2.XMLHTTP. Returns "" on any failure.
' Cache-buster query param bypasses GitHub raw CDN cache (~5 min).
Private Function FetchRemote(url As String) As String
    Dim http As Object
    Dim bustedUrl As String
    On Error GoTo HttpFail
    ' ServerXMLHTTP uses WinHTTP (not WinINet) — does NOT share IE's local cache,
    ' so stale-response bugs from previous calls don't haunt us. The cache-buster
    ' query string below is belt-and-braces for the GitHub CDN edge layer.
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    bustedUrl = url & "?t=" & CStr(CLng((Now - DateSerial(1970, 1, 1)) * 86400))
    http.Open "GET", bustedUrl, False
    http.setRequestHeader "Cache-Control", "no-cache"
    http.send
    If http.Status = 200 Then
        FetchRemote = StripBOM(http.responseText)
    Else
        FetchRemote = ""
    End If
    Exit Function
HttpFail:
    FetchRemote = ""
End Function

' Drops a UTF-8 BOM if present at the start of the text.
' Attribute VB_Name must be the very first line of an imported .bas — a leading
' BOM would corrupt the import.
Private Function StripBOM(text As String) As String
    If Len(text) > 0 Then
        If AscW(Left$(text, 1)) = &HFEFF Then
            StripBOM = Mid$(text, 2)
            Exit Function
        End If
    End If
    StripBOM = text
End Function

' Exports current modGenerateSchedule to %TEMP%\modGenerateSchedule_backup_<TS>.bas
Private Function BackupModule() As String
    Dim path As String
    path = Environ$("TEMP") & "\modGenerateSchedule_backup_" & _
           Format(Now, "yyyymmdd_hhnnss") & ".bas"
    ThisWorkbook.VBProject.VBComponents(TARGET_MODULE).Export path
    BackupModule = path
End Function

' Writes downloaded text to %TEMP%\modGenerateSchedule_new.bas.
' Normalises line endings so the on-disk file isn't malformed in text editors.
Private Function WriteTempBas(text As String) As String
    Dim path As String
    Dim f As Integer
    path = Environ$("TEMP") & "\modGenerateSchedule_new.bas"
    text = Replace(text, vbCrLf, vbLf)
    text = Replace(text, vbLf, vbCrLf)
    f = FreeFile
    Open path For Output As #f
    Print #f, text;
    Close #f
    WriteTempBas = path
End Function

' Removes all VBComponents whose Name starts with TARGET_MODULE and Type=1.
' Two-pass to avoid skipping items when mutating a live collection.
Private Sub RemoveTargetModules()
    Dim cmp As Object
    Dim toRemove As Collection
    Dim i As Long
    Set toRemove = New Collection
    For Each cmp In ThisWorkbook.VBProject.VBComponents
        If cmp.Type = 1 Then
            If Left$(cmp.Name, Len(TARGET_MODULE)) = TARGET_MODULE Then
                toRemove.Add cmp
            End If
        End If
    Next cmp
    For i = 1 To toRemove.Count
        ThisWorkbook.VBProject.VBComponents.Remove toRemove(i)
    Next i
End Sub

' Confirms the new modGenerateSchedule was imported and contains GenerateSchedule.
Private Function VerifyImport() As Boolean
    Dim cmp As Object
    Dim code As String
    On Error Resume Next
    Set cmp = ThisWorkbook.VBProject.VBComponents(TARGET_MODULE)
    On Error GoTo 0
    If cmp Is Nothing Then Exit Function
    code = cmp.CodeModule.Lines(1, cmp.CodeModule.CountOfLines)
    VerifyImport = (InStr(code, TARGET_SUB) > 0)
End Function
