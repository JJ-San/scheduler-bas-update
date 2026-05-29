Attribute VB_Name = "modUpdater"
'=============================================================================
' MODULE: modUpdater
' PURPOSE: Self-update mechanism for the Invoice Schedule macro.
'          Fetches the latest modGenerateSchedule.bas from a public GitHub URL,
'          compares versions, and swaps in the new module if newer.
'
' DESIGN:  This module is INTENTIONALLY IMMUTABLE. It never updates itself.
'          If we ever need to evolve it, that requires a separate manual
'          re-install. Keep this module's scope narrow and predictable.
'
' SHIPPED: Bundled with Update 2 (2026-05). One-time install per workbook.
'=============================================================================

Option Explicit

' Bump only if we ever ship a new version of the updater itself (rare).
Public Const UPDATER_VERSION As String = "1.1"

' Source of truth for the latest modGenerateSchedule.bas.
Private Const UPDATE_URL As String = "https://raw.githubusercontent.com/JJ-San/scheduler-bas-update/main/modGenerateSchedule.bas"
Private Const TARGET_MODULE As String = "modGenerateSchedule"
Private Const TARGET_SUB As String = "Sub GenerateSchedule"
Private Const BUTTON_CAPTION As String = "Check for Updates"
Private Const BUTTON_NAME As String = "btnCheckUpdates"


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
        MsgBox "Excel needs permission to update the macro." & vbCrLf & vbCrLf & _
               "To enable: File > Options > Trust Center > Trust Center Settings >" & vbCrLf & _
               "Macro Settings > tick ""Trust access to the VBA project object model""." & vbCrLf & vbCrLf & _
               "Then click Check for Updates again.", _
               vbInformation, "Update Setup Required"
        Exit Sub
    End If

    installedVer = ReadInstalledVersion()

    remoteText = FetchRemote(UPDATE_URL)
    If Len(remoteText) = 0 Then
        MsgBox "Could not reach the update server." & vbCrLf & _
               "Check your internet connection and try again later.", _
               vbExclamation, "Update Unavailable"
        Exit Sub
    End If

    remoteVer = ParseVersion(remoteText)
    If Len(remoteVer) = 0 Then
        MsgBox "The update file from the server doesn't look right." & vbCrLf & _
               "Please contact support.", _
               vbCritical, "Update File Invalid"
        Exit Sub
    End If

    If CompareSemver(installedVer, remoteVer) >= 0 Then
        MsgBox "You're up to date." & vbCrLf & vbCrLf & _
               "Current version: " & installedVer, _
               vbInformation, "No Update Available"
        Exit Sub
    End If

    If MsgBox("An update is available." & vbCrLf & vbCrLf & _
              "Installed version:  " & installedVer & vbCrLf & _
              "Latest version:     " & remoteVer & vbCrLf & vbCrLf & _
              "Update now?", _
              vbYesNo + vbQuestion, "Update Available") <> vbYes Then
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

    ' Refresh the on-sheet version label immediately so users see the new
    ' version without having to click Generate Schedule. Late-bound via
    ' Application.Run so a future modGenerateSchedule without WriteVersionLabel
    ' (renamed/removed) doesn't break the update — label just won't refresh.
    On Error Resume Next
    Application.Run "WriteVersionLabel"
    On Error GoTo 0

    MsgBox "Updated to version " & remoteVer & "." & vbCrLf & vbCrLf & _
           "Please save your workbook now (Ctrl+S) to keep the update.", _
           vbInformation, "Update Complete"
    Exit Sub

Rollback:
    On Error Resume Next
    RemoveTargetModules
    ThisWorkbook.VBProject.VBComponents.Import backupPath
    On Error GoTo 0
    MsgBox "The update failed and was rolled back." & vbCrLf & _
           "Your original macro is intact. Please try again later or contact support.", _
           vbCritical, "Update Failed"
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

' Compares semver-style "X.Y" strings. -1 if a<b, 0 equal, 1 if a>b.
' Parts compared as integers so "2.10" > "2.9" works correctly.
Private Function CompareSemver(a As String, b As String) As Long
    Dim partsA() As String, partsB() As String
    Dim i As Long, maxLen As Long
    Dim ai As Long, bi As Long

    partsA = Split(a, ".")
    partsB = Split(b, ".")
    maxLen = WorksheetFunction.Max(UBound(partsA), UBound(partsB))

    For i = 0 To maxLen
        ai = 0: bi = 0
        If i <= UBound(partsA) Then ai = CLng(Val(partsA(i)))
        If i <= UBound(partsB) Then bi = CLng(Val(partsB(i)))
        If ai < bi Then CompareSemver = -1: Exit Function
        If ai > bi Then CompareSemver = 1: Exit Function
    Next i
    CompareSemver = 0
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
