Attribute VB_Name = "TestLoggerReport"
Option Explicit

' Tests for the issue report. Run RunLoggerReportTests from the VBA editor
' (F5) in the DEV copy of GCV Logger. Uses a temporary folder and a fixed
' clock; creates real zip files there. Do NOT ship this module.

Private passedCount As Long
Private failedCount As Long
Private currentTest As String

Public Sub RunLoggerReportTests()
    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "Logger report tests"
    On Error GoTo Crash

    Test_FilterLines_KeepsOnlyOneSource
    Test_AddinFromInfo_ReadsTheFields
    Test_ScopePrompt_NumbersEveryAddin
    Test_ReportText_ListsVersionsAndDescription
    Test_BuildReport_ZipsReportAndLogs

Finish:
    GcvLog.ResetLoggerForTests
    LoggerUi.UseTestMessages False
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "Logger report tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

Private Sub Test_FilterLines_KeepsOnlyOneSource()
    Dim text As String
    StartTest "Filtering keeps only the chosen add-in's lines"
    text = "[student-lookup] - INFO: a" & vbCrLf & "[updater] - INFO: b" & vbCrLf & _
           "[student-lookup] - WARN: c" & vbCrLf & "[student-lookup-x] - INFO: d" & vbCrLf
    AssertEqual "[student-lookup] - INFO: a" & vbCrLf & "[student-lookup] - WARN: c" & vbCrLf, _
                FilterLines(text, "student-lookup"), "two lines"
    AssertEqual "", FilterLines("", "updater"), "empty"
End Sub

Private Sub Test_AddinFromInfo_ReadsTheFields()
    Dim addin As Object
    StartTest "Add-in info is split into its fields"
    Set addin = AddinFromInfo("gcv/student-lookup|3.0.0|Gaza-children-village/excel-student-lookup|GCV Student Lookup")
    AssertEqual "gcv/student-lookup", addin("id"), "id"
    AssertEqual "3.0.0", addin("version"), "version"
    AssertEqual "GCV Student Lookup", addin("name"), "name"
    AssertEqual "student-lookup", addin("source"), "log source"
End Sub

Private Sub Test_ScopePrompt_NumbersEveryAddin()
    Dim prompt As String
    StartTest "The choice lists 0 = all and every add-in by number"
    prompt = ScopePrompt(TwoAddins())
    AssertEqual True, InStr(prompt, "0 = ") > 0, "all"
    AssertEqual True, InStr(prompt, "1 = GCV Logger") > 0, "first"
    AssertEqual True, InStr(prompt, "2 = GCV Student Lookup") > 0, "second"
End Sub

Private Sub Test_ReportText_ListsVersionsAndDescription()
    Dim text As String
    StartTest "report.txt lists versions, scope and description"
    text = ReportText("GCV Student Lookup", "Rows stay empty", TwoAddins(), DateSerial(2026, 11, 19) + TimeSerial(10, 30, 0))
    AssertEqual True, InStr(text, "Created: 2026-11-19 10:30") > 0, "date"
    AssertEqual True, InStr(text, "About: GCV Student Lookup") > 0, "scope"
    AssertEqual True, InStr(text, "Rows stay empty") > 0, "description"
    AssertEqual True, InStr(text, "GCV Student Lookup 3.0.0  (gcv/student-lookup)") > 0, "version"
    AssertEqual True, InStr(ReportText("x", "  ", TwoAddins(), Now), "(none)") > 0, "no description"
End Sub

Private Sub Test_BuildReport_ZipsReportAndLogs()
    Dim logs As String
    Dim target As String
    Dim result As String
    Dim at As Date
    StartTest "The report is zipped into the target folder"
    at = DateSerial(2026, 11, 19) + TimeSerial(10, 30, 0)
    logs = TempFolder("logs")
    target = TempFolder("target")
    GcvLog.ResetLoggerForTests
    GcvLog.SetLogFolderForTests logs
    WriteText logs & "\2026-11-18.log", "[updater] - INFO: 2026-11-18 - 09:00:00 x" & vbCrLf
    WriteText logs & "\2026-11-19.log", ""                     ' empty file must not break the report
    LoggerUi.UseTestMessages True
    result = BuildReport("", "all", "test", TwoAddins(), target, at)
    AssertEqual target & "\GCV report 2026-11-19 103000.zip", result, "zip path"
    AssertEqual True, CreateObject("Scripting.FileSystemObject").FileExists(result), "zip exists"
    GcvLog.ResetLoggerForTests
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function TwoAddins() As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add AddinFromInfo("gcv/logger|1.0.0|owner/excel-logger|GCV Logger")
    result.Add AddinFromInfo("gcv/student-lookup|3.0.0|owner/excel-student-lookup|GCV Student Lookup")
    Set TwoAddins = result
End Function

Private Function TempFolder(ByVal name As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    TempFolder = fso.GetSpecialFolder(2).Path & "\gcv-logger-report-tests"
    If Not fso.FolderExists(TempFolder) Then fso.CreateFolder TempFolder
    TempFolder = TempFolder & "\" & name & "-" & Format$(Now, "hhnnss") & Int(Rnd() * 10000)
    fso.CreateFolder TempFolder
End Function

Private Sub WriteText(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("Scripting.FileSystemObject").CreateTextFile(path, True, False)
    stream.Write text
    stream.Close
End Sub

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    If CStr(expected) = CStr(actual) Then
        passedCount = passedCount + 1
    Else
        failedCount = failedCount + 1
        Debug.Print "FAIL: " & currentTest & " / " & label & " - expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
    End If
End Sub
