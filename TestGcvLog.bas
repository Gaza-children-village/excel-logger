Attribute VB_Name = "TestGcvLog"
Option Explicit

' Tests for the GCV logger. Run RunLoggerTests (Alt+F8) in the DEV workbook;
' details in the Immediate window (Ctrl+G). Uses a temporary log folder and a
' fixed clock. Do NOT ship this module in production workbooks.

Private passedCount As Long
Private failedCount As Long
Private currentTest As String
Private savedLevel As String

Public Sub RunLoggerTests()
    passedCount = 0
    failedCount = 0
    Debug.Print String$(60, "-")
    Debug.Print "Logger tests"
    savedLevel = GcvLog.GetLogLevel()
    On Error GoTo Crash

    Test_FormatLogLine_FollowsTheAgreedLayout
    Test_CleanText_KeepsLinesAsciiAndOnOneLine
    Test_Levels_ProductionWritesInfoWarnErrorButNotDebug
    Test_Levels_DebugLevelWritesEverything
    Test_File_AppendsLinesToTheDailyFile
    Test_LogError_AddsTheCurrentErrorAndKeepsIt
    Test_LogInfo_KeepsTheCallersErr
    Test_Failure_SwitchesFileLoggingOffWithoutRaising
    Test_DeleteOldLogs_RemovesOnlyFilesOlderThan30Days
    Test_ExportLogs_CopiesRecentDaysIntoOneFile

Finish:
    GcvLog.ResetLoggerForTests
    GcvLog.SetLogLevel savedLevel
    Debug.Print "Passed: " & passedCount & "   Failed: " & failedCount
    MsgBox "Passed: " & passedCount & vbCrLf & "Failed: " & failedCount & vbCrLf & _
           "Details in the Immediate window (Ctrl+G).", _
           IIf(failedCount = 0, vbInformation, vbCritical), "Logger tests"
    Exit Sub

Crash:
    failedCount = failedCount + 1
    Debug.Print "CRASH in " & currentTest & ": " & Err.Number & " " & Err.Description
    Resume Finish
End Sub

Private Sub Test_FormatLogLine_FollowsTheAgreedLayout()
    StartTest "Log line follows [source] - LEVEL: date - time message"
    AssertEqual "[student-lookup] - INFO: 2026-11-14 - 10:22:03 Lookup finished", _
                FormatLogLine("INFO", "student-lookup", "Lookup finished", DateSerial(2026, 11, 14) + TimeSerial(10, 22, 3)), _
                "layout"
End Sub

Private Sub Test_CleanText_KeepsLinesAsciiAndOnOneLine()
    StartTest "Text is kept on one ASCII line"
    AssertEqual "a | b | c", CleanText("a" & vbCrLf & "b" & vbLf & "c"), "line breaks"
    AssertEqual "\u0627x", CleanText(ChrW$(&H627) & "x"), "Arabic escaped"
    AssertEqual "a b", CleanText("a" & vbTab & "b"), "tab"
End Sub

Private Sub Test_Levels_ProductionWritesInfoWarnErrorButNotDebug()
    Dim folder As String
    StartTest "Production level writes INFO, WARN, ERROR but not DEBUG"
    folder = UseTempLog("levels-prod")
    GcvLog.SetLogLevel vbNullString
    LogDebug "test", "debug line"
    LogInfo "test", "info line"
    LogWarn "test", "warn line"
    LogError "test", "error line"
    AssertEqual False, FileHas("debug line"), "no DEBUG"
    AssertEqual True, FileHas("[test] - INFO:"), "INFO"
    AssertEqual True, FileHas("[test] - WARN:"), "WARN"
    AssertEqual True, FileHas("[test] - ERROR:"), "ERROR"
End Sub

Private Sub Test_Levels_DebugLevelWritesEverything()
    StartTest "DEBUG level writes everything"
    UseTempLog "levels-debug"
    GcvLog.SetLogLevel "DEBUG"
    LogDebug "test", "debug line"
    AssertEqual True, FileHas("[test] - DEBUG: 2026-11-14 - 09:30:00 debug line"), "DEBUG written"
End Sub

Private Sub Test_File_AppendsLinesToTheDailyFile()
    Dim content As String
    StartTest "Lines are appended to the daily file"
    UseTempLog "append"
    GcvLog.SetLogLevel vbNullString
    LogInfo "test", "first"
    LogInfo "test", "second"
    AssertEqual True, Len(Dir$(LogFileFor(TestDay()))) > 0, "file named after the day"
    content = ReadAll(LogFileFor(TestDay()))
    AssertEqual True, InStr(content, "first") > 0 And InStr(content, "second") > InStr(content, "first"), "in order"
End Sub

Private Sub Test_LogError_AddsTheCurrentErrorAndKeepsIt()
    StartTest "LogError adds the current error and leaves it in place"
    UseTempLog "error"
    GcvLog.SetLogLevel vbNullString
    On Error Resume Next
    Err.Raise 1004, "test", "Something failed"
    LogError "test", "Step failed"
    AssertEqual 1004, Err.Number, "Err kept"
    On Error GoTo 0
    AssertEqual True, FileHas("Step failed (error 1004: Something failed)"), "error text in the line"
End Sub

Private Sub Test_LogInfo_KeepsTheCallersErr()
    StartTest "Logging does not clear the caller's error"
    UseTempLog "keep-err"
    On Error Resume Next
    Err.Raise 5
    LogInfo "test", "while handling an error"
    AssertEqual 5, Err.Number, "still error 5"
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Test_Failure_SwitchesFileLoggingOffWithoutRaising()
    Dim n As Long
    StartTest "A write failure switches file logging off and does not raise"
    GcvLog.ResetLoggerForTests
    GcvLog.SetClockForTests TestDay() + TimeSerial(9, 30, 0)
    GcvLog.SetLogFolderForTests "Q:\no such drive\gcv-logs"
    On Error Resume Next
    LogInfo "test", "cannot be written"
    n = Err.Number
    On Error GoTo 0
    AssertEqual 0, n, "no error for the caller"
    AssertEqual True, GcvLog.IsFileLoggingOff(), "file logging off"
End Sub

Private Sub Test_DeleteOldLogs_RemovesOnlyFilesOlderThan30Days()
    Dim folder As String
    StartTest "Old log files are deleted after 30 days"
    folder = UseTempLog("retention")
    WriteFile folder & "\" & Format$(TestDay() - 40, "yyyy-mm-dd") & ".log", "old"
    WriteFile folder & "\" & Format$(TestDay() - 5, "yyyy-mm-dd") & ".log", "recent"
    WriteFile folder & "\notes.log", "not a daily file"
    GcvLog.DeleteOldLogs TestDay()
    AssertEqual 0, Len(Dir$(folder & "\" & Format$(TestDay() - 40, "yyyy-mm-dd") & ".log")), "old deleted"
    AssertEqual True, Len(Dir$(folder & "\" & Format$(TestDay() - 5, "yyyy-mm-dd") & ".log")) > 0, "recent kept"
    AssertEqual True, Len(Dir$(folder & "\notes.log")) > 0, "other files kept"
End Sub

Private Sub Test_ExportLogs_CopiesRecentDaysIntoOneFile()
    Dim folder As String
    Dim target As String
    Dim lines As Long
    StartTest "Export copies recent days into one file"
    folder = UseTempLog("export")
    WriteFile folder & "\" & Format$(TestDay() - 1, "yyyy-mm-dd") & ".log", "[x] - INFO: yesterday"
    WriteFile folder & "\" & Format$(TestDay(), "yyyy-mm-dd") & ".log", "[x] - INFO: today"
    WriteFile folder & "\" & Format$(TestDay() - 20, "yyyy-mm-dd") & ".log", "[x] - INFO: too old"
    target = folder & "\export.txt"
    lines = GcvLog.ExportLogs(target, 7)
    AssertEqual 2, lines, "two lines"
    AssertEqual True, InStr(ReadAll(target), "yesterday") < InStr(ReadAll(target), "today"), "oldest first"
    AssertEqual 0, InStr(ReadAll(target), "too old"), "older days left out"
End Sub

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function TestDay() As Date
    TestDay = DateSerial(2026, 11, 14)
End Function

Private Function UseTempLog(ByVal name As String) As String
    UseTempLog = Environ$("TEMP") & "\gcv-logger-tests\" & name & "-" & Format$(Now, "hhnnss") & Int(Rnd() * 10000)
    MakeTestFolder UseTempLog
    GcvLog.ResetLoggerForTests
    GcvLog.SetLogFolderForTests UseTempLog
    GcvLog.SetClockForTests TestDay() + TimeSerial(9, 30, 0)
End Function

Private Function FileHas(ByVal text As String) As Boolean
    If Len(Dir$(LogFileFor(TestDay()))) = 0 Then Exit Function
    FileHas = InStr(ReadAll(LogFileFor(TestDay())), text) > 0
End Function

Private Function ReadAll(ByVal path As String) As String
    Dim fileNumber As Integer
    Dim content As String
    fileNumber = FreeFile
    Open path For Binary Access Read As #fileNumber
    content = Space$(LOF(fileNumber))
    Get #fileNumber, , content
    Close #fileNumber
    ReadAll = content
End Function

Private Sub WriteFile(ByVal path As String, ByVal text As String)
    Dim fileNumber As Integer
    fileNumber = FreeFile
    Open path For Output As #fileNumber
    Print #fileNumber, text
    Close #fileNumber
End Sub

Private Sub MakeTestFolder(ByVal path As String)
    Dim parts() As String
    Dim current As String
    Dim i As Long
    parts = Split(path, "\")
    current = parts(0)
    For i = 1 To UBound(parts)
        current = current & "\" & parts(i)
        If Len(Dir$(current, vbDirectory)) = 0 Then MkDir current
    Next i
End Sub

Private Sub StartTest(ByVal name As String)
    currentTest = name
End Sub

Private Sub AssertEqual(ByVal expected As Variant, ByVal actual As Variant, ByVal label As String)
    Dim same As Boolean
    If IsNumeric(expected) And IsNumeric(actual) And VarType(expected) <> vbString And VarType(actual) <> vbString Then
        same = (CDbl(expected) = CDbl(actual))
    Else
        same = (CStr(expected) = CStr(actual))
    End If
    If same Then
        passedCount = passedCount + 1
    Else
        failedCount = failedCount + 1
        Debug.Print "FAIL: " & currentTest & " / " & label & " - expected [" & CStr(expected) & "] got [" & CStr(actual) & "]"
    End If
End Sub
