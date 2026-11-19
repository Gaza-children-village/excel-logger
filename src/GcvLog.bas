Attribute VB_Name = "GcvLog"
Option Explicit

' ============================================================================
' GCV logger (package gcv/logger) - records what the modules DO, never the
' data they handle: do not put names, IDs or other personal data in messages.
'
' Line format:
'   [source] - LEVEL: yyyy-mm-dd - hh:nn:ss message
'   [student-lookup] - INFO: 2026-11-14 - 10:22:03 Lookup finished: 38 found
'
' Files: %LOCALAPPDATA%\GCV\ExcelLogs\<workbook>\<yyyy-mm-dd>.log, kept 30 days.
'        Inside an add-in (GCV Tools): ...\ExcelLogs\GCV Tools\, whatever the version.
' Level: in a workbook, the hidden name GCV_LOG_LEVEL; in an add-in, the
' registry (HKCU\Software\VB and VBA Program Settings\GCV Tools\Logger\Level),
' because an add-in is opened read-only. DEV: "DEBUG" (everything, also in the
' Immediate window). Production (not set): INFO, WARN and ERROR.
'   Alt+F8 > GcvLogUseDebugLevel / GcvLogUseProductionLevel
' Export for support: Alt+F8 > GcvLogExport (last 7 days -> a file on the Desktop)
'
' Logging never stops the calling module: if a file cannot be written, file
' logging is switched off for the session and a note goes to the Immediate
' window. The caller's Err object is left exactly as it was.
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

Public Const LOG_LEVEL_NAME As String = "GCV_LOG_LEVEL"
Private Const KEEP_DAYS As Long = 30
Private Const EXPORT_DAYS As Long = 7
Private Const LEVEL_DEBUG As Long = 10
Private Const LEVEL_INFO As Long = 20
Private Const LEVEL_WARN As Long = 30
Private Const LEVEL_ERROR As Long = 40
Private Const MB_ICONINFO As Long = &H40
Private Const MB_ICONERROR As Long = &H10
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000
Private Const ADDIN_LOG_NAME As String = "GCV Tools"
Private Const SETTINGS_APP As String = "GCV Tools"
Private Const SETTINGS_SECTION As String = "Logger"
Private Const SETTINGS_KEY As String = "Level"

Private mFileLoggingOff As Boolean
Private mFolderForTests As String
Private mClockForTests As Date
Private mCleanedOn As String
Private mFso As Object
Private mLocalAppData As String

' ----------------------------------------------------------------------------
' Public API
' ----------------------------------------------------------------------------

Public Sub LogDebug(ByVal source As String, ByVal message As String)
    LogWrite "DEBUG", source, message
End Sub

Public Sub LogInfo(ByVal source As String, ByVal message As String)
    LogWrite "INFO", source, message
End Sub

Public Sub LogWarn(ByVal source As String, ByVal message As String)
    LogWrite "WARN", source, message
End Sub

' Adds the current VBA error (if any) to the message.
Public Sub LogError(ByVal source As String, ByVal message As String)
    If Err.Number <> 0 Then message = message & " (error " & Err.Number & ": " & Err.Description & ")"
    LogWrite "ERROR", source, message
End Sub

' Entry point for callers that bind at run time:
'   Application.Run "'Book.xlsm'!GcvLog.LogWrite", "INFO", "updater", "message"
Public Sub LogWrite(ByVal levelName As String, ByVal source As String, ByVal message As String)
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedSource As String
    Dim logLine As String

    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source

    levelName = UCase$(Trim$(levelName))
    If LevelValue(levelName) >= MinimumLevel() Then
        logLine = FormatLogLine(levelName, source, message, Clock())
        If MinimumLevel() <= LEVEL_DEBUG Then Debug.Print logLine
        AppendLine logLine
    End If

    RestoreErr savedNumber, savedDescription, savedSource
End Sub

' ----------------------------------------------------------------------------
' Macros (Alt+F8)
' ----------------------------------------------------------------------------

Public Sub GcvLogUseDebugLevel()
    SetLogLevel "DEBUG"
    LogInfo "logger", "Log level set to DEBUG (all entries)."
End Sub

Public Sub GcvLogUseProductionLevel()
    SetLogLevel vbNullString
    LogInfo "logger", "Log level set to production (INFO, WARN, ERROR)."
End Sub

Public Sub GcvLogExport()
    Dim target As String
    Dim lines As Long

    target = DesktopFolder() & "\GCV logs - " & SafeName(WorkbookBaseName()) & " - " & _
             Format$(Now, "yyyy-mm-dd hhnn") & ".txt"
    lines = ExportLogs(target, EXPORT_DAYS)
    LogInfo "logger", "Logs exported: " & lines & " line(s) from the last " & EXPORT_DAYS & " days."
    If lines < 0 Then
        ShowMessage FromCodePoints("062A 0639 0630 0631 0020 062D 0641 0638 0020 0645 0644 0641 0020 0627 0644 0633 062C 0644 0627 062A 0020 0639 0644 0649 0020 0633 0637 062D 0020 0627 0644 0645 0643 062A 0628 002E") & vbCrLf & target, MB_ICONERROR
    Else
        ShowMessage FromCodePoints("062A 0645 0020 062D 0641 0638 0020 0645 0644 0641 0020 0627 0644 0633 062C 0644 0627 062A 0020 0639 0644 0649 0020 0633 0637 062D 0020 0627 0644 0645 0643 062A 0628 002E 0020 0623 0631 0633 0644 0647 0020 0625 0644 0649 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 003A") & vbCrLf & target, MB_ICONINFO
    End If
End Sub

' ""   -> production (INFO and above); "DEBUG", "INFO", "WARN", "ERROR" otherwise.
Public Sub SetLogLevel(ByVal levelName As String)
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedSource As String

    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source
    If ThisWorkbook.IsAddin Then
        ' DeleteSetting raises when the value does not exist; nothing to delete then.
        On Error Resume Next
        DeleteSetting SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY
        On Error GoTo 0
        If Len(levelName) > 0 Then SaveSetting SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, UCase$(levelName)
    Else
        ' Names(...) raises when the name does not exist; nothing to delete then.
        On Error Resume Next
        ThisWorkbook.Names(LOG_LEVEL_NAME).Delete
        On Error GoTo 0
        If Len(levelName) > 0 Then
            ThisWorkbook.Names.Add Name:=LOG_LEVEL_NAME, RefersTo:="=""" & UCase$(levelName) & """", Visible:=False
        End If
    End If
    RestoreErr savedNumber, savedDescription, savedSource
End Sub

' ----------------------------------------------------------------------------
' Formatting
' ----------------------------------------------------------------------------

Public Function FormatLogLine(ByVal levelName As String, ByVal source As String, _
                              ByVal message As String, ByVal at As Date) As String
    FormatLogLine = "[" & CleanText(source) & "] - " & UCase$(levelName) & ": " & _
                    Format$(at, "yyyy-mm-dd") & " - " & Format$(at, "hh:nn:ss") & " " & CleanText(message)
End Function

' One line, ASCII only: line breaks become " | ", other characters outside
' ASCII become \uXXXX (so lines can be written fast with Print #).
Public Function CleanText(ByVal text As String) As String
    Dim i As Long
    Dim code As Long
    Dim parts() As String

    text = Replace(Replace(Replace(text, vbCrLf, " | "), vbCr, " | "), vbLf, " | ")
    If Len(text) = 0 Then Exit Function
    ReDim parts(1 To Len(text))
    For i = 1 To Len(text)
        code = AscW(Mid$(text, i, 1)) And &HFFFF&
        If code < 32 Then
            parts(i) = " "
        ElseIf code > 126 Then
            parts(i) = "\u" & Right$("000" & Hex$(code), 4)
        Else
            parts(i) = Mid$(text, i, 1)
        End If
    Next i
    CleanText = Join(parts, vbNullString)
End Function

Private Function LevelValue(ByVal levelName As String) As Long
    Select Case levelName
        Case "DEBUG": LevelValue = LEVEL_DEBUG
        Case "WARN", "WARNING": LevelValue = LEVEL_WARN
        Case "ERROR": LevelValue = LEVEL_ERROR
        Case Else: LevelValue = LEVEL_INFO
    End Select
End Function

Private Function MinimumLevel() As Long
    Dim value As String
    value = GetLogLevel()
    ' The name is absent in production workbooks: that means "INFO".
    If Len(value) = 0 Then MinimumLevel = LEVEL_INFO Else MinimumLevel = LevelValue(value)
End Function

' The stored level ("DEBUG", "INFO", ...) or "" when none is set (production).
' Pass it back to SetLogLevel to restore it. Leaves Err unchanged.
Public Function GetLogLevel() As String
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedSource As String
    Dim definedName As Name

    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source
    If ThisWorkbook.IsAddin Then
        GetLogLevel = Trim$(UCase$(GetSetting(SETTINGS_APP, SETTINGS_SECTION, SETTINGS_KEY, vbNullString)))
    Else
        On Error Resume Next
        Set definedName = ThisWorkbook.Names(LOG_LEVEL_NAME)
        On Error GoTo 0
        If Not definedName Is Nothing Then _
            GetLogLevel = Trim$(UCase$(Replace(Replace(definedName.RefersTo, "=", vbNullString), """", vbNullString)))
    End If
    RestoreErr savedNumber, savedDescription, savedSource
End Function

' ----------------------------------------------------------------------------
' Files
' ----------------------------------------------------------------------------

Public Function LogFolder() As String
    If Len(mFolderForTests) > 0 Then
        LogFolder = mFolderForTests
    Else
        LogFolder = LocalAppData() & "\GCV\ExcelLogs\" & LogName()
    End If
End Function

Public Function LogFileFor(ByVal day As Date) As String
    LogFileFor = LogFolder() & "\" & Format$(day, "yyyy-mm-dd") & ".log"
End Function

Public Function IsFileLoggingOff() As Boolean
    IsFileLoggingOff = mFileLoggingOff
End Function

' Files are handled with Scripting.FileSystemObject: unlike Open/Dir$/MkDir it
' accepts Unicode paths (Arabic Windows user names) and does not disturb a
' Dir$ loop of the caller.
Private Sub AppendLine(ByVal logLine As String)
    Dim stream As Object
    Dim errText As String

    If mFileLoggingOff Then Exit Sub
    On Error GoTo Failed
    MakeFolder LogFolder()
    Set stream = Fso().OpenTextFile(LogFileFor(Clock()), 8, True, 0)   ' append, create, ASCII
    stream.WriteLine logLine
    stream.Close
    Set stream = Nothing
    If mCleanedOn <> Format$(Clock(), "yyyy-mm-dd") Then
        mCleanedOn = Format$(Clock(), "yyyy-mm-dd")
        DeleteOldLogs Clock()
    End If
    Exit Sub

Failed:
    ' Deliberate: a logging failure must never stop the module that logs.
    errText = Err.Number & ": " & Err.Description
    mFileLoggingOff = True
    On Error GoTo -1
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Debug.Print "GcvLog: file logging switched off for this session (" & errText & ") - folder: " & LogFolder()
End Sub

' Deletes log files older than KEEP_DAYS (file names are yyyy-mm-dd.log).
' Leaves the caller's Err unchanged.
Public Sub DeleteOldLogs(ByVal today As Date)
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedSource As String
    Dim file As Object
    Dim old As Collection
    Dim item As Variant
    Dim fileName As String

    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source
    Set old = New Collection
    ' Cleaning up is best effort: a file that cannot be read or deleted (e.g.
    ' open elsewhere) is skipped; it must not switch logging off.
    On Error Resume Next
    If Fso().FolderExists(LogFolder()) Then
        For Each file In Fso().GetFolder(LogFolder()).Files
            fileName = file.Name
            If fileName Like "####-##-##.log" Then
                If DateSerial(CInt(Left$(fileName, 4)), CInt(Mid$(fileName, 6, 2)), CInt(Mid$(fileName, 9, 2))) _
                   < DateAdd("d", -KEEP_DAYS, DateValue(today)) Then old.Add file.Path
            End If
        Next file
        For Each item In old
            Fso().DeleteFile item, True
            If Err.Number <> 0 Then
                Debug.Print "GcvLog: could not delete old log " & item & " (" & Err.Description & ")"
                Err.Clear
            End If
        Next item
    End If
    On Error GoTo 0
    RestoreErr savedNumber, savedDescription, savedSource
End Sub

' Copies the last `days` daily files into one file; returns the number of
' lines written, or -1 when the file could not be written. Leaves Err unchanged.
Public Function ExportLogs(ByVal target As String, ByVal days As Long) As Long
    Dim savedNumber As Long
    Dim savedDescription As String
    Dim savedSource As String
    Dim writer As Object
    Dim reader As Object
    Dim offset As Long
    Dim source As String
    Dim count As Long

    savedNumber = Err.Number
    savedDescription = Err.Description
    savedSource = Err.Source
    On Error GoTo Failed
    Set writer = Fso().OpenTextFile(target, 2, True, 0)      ' write, create, ASCII
    writer.WriteLine "GCV logs - " & SafeName(WorkbookBaseName()) & " - exported " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    For offset = days - 1 To 0 Step -1
        source = LogFileFor(DateAdd("d", -offset, Clock()))
        If Fso().FileExists(source) Then
            Set reader = Fso().OpenTextFile(source, 1, False, 0)
            Do Until reader.AtEndOfStream
                writer.WriteLine reader.ReadLine
                count = count + 1
            Loop
            reader.Close
            Set reader = Nothing
        End If
    Next offset
    writer.Close
    ExportLogs = count
    RestoreErr savedNumber, savedDescription, savedSource
    Exit Function

Failed:
    Debug.Print "GcvLog: export failed (" & Err.Number & ": " & Err.Description & ")"
    On Error GoTo -1
    On Error Resume Next
    If Not reader Is Nothing Then reader.Close
    If Not writer Is Nothing Then writer.Close
    On Error GoTo 0
    ExportLogs = -1
    RestoreErr savedNumber, savedDescription, savedSource
End Function

Private Sub MakeFolder(ByVal path As String)
    Dim parts() As String
    Dim current As String
    Dim i As Long

    parts = Split(path, "\")
    current = parts(0)
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 Then
            current = current & "\" & parts(i)
            If Not Fso().FolderExists(current) Then Fso().CreateFolder current
        End If
    Next i
End Sub

Private Function Fso() As Object
    If mFso Is Nothing Then Set mFso = CreateObject("Scripting.FileSystemObject")
    Set Fso = mFso
End Function

' %LOCALAPPDATA% in Unicode (Environ$ would turn Arabic user names into "?").
Private Function LocalAppData() As String
    If Len(mLocalAppData) = 0 Then
        mLocalAppData = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%LOCALAPPDATA%")
    End If
    LocalAppData = mLocalAppData
End Function

Private Function SafeName(ByVal text As String) As String
    Dim i As Long
    Dim ch As String

    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch Like "[A-Za-z0-9 ._()-]" Then
            SafeName = SafeName & ch
        Else
            SafeName = SafeName & "_"
        End If
    Next i
    If Len(Trim$(SafeName)) = 0 Then SafeName = "workbook"
End Function

' One folder for every GCV Tools version; one folder per workbook otherwise.
Private Function LogName() As String
    If ThisWorkbook.IsAddin Then
        LogName = ADDIN_LOG_NAME
    Else
        LogName = SafeName(WorkbookBaseName())
    End If
End Function

Private Function WorkbookBaseName() As String
    Dim dotPos As Long
    WorkbookBaseName = ThisWorkbook.Name
    dotPos = InStrRev(WorkbookBaseName, ".")
    If dotPos > 1 Then WorkbookBaseName = Left$(WorkbookBaseName, dotPos - 1)
End Function

Private Function DesktopFolder() As String
    DesktopFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function

Private Function Clock() As Date
    If mClockForTests <> 0 Then
        Clock = mClockForTests
    Else
        Clock = Now
    End If
End Function

Private Sub RestoreErr(ByVal number As Long, ByVal description As String, ByVal source As String)
    If number = 0 Then
        Err.Clear
    Else
        Err.Number = number
        Err.Description = description
        Err.Source = source
    End If
End Sub

Private Sub ShowMessage(ByVal text As String, ByVal icon As Long)
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(FromCodePoints("0633 062C 0644 0627 062A 0020 0627 0644 0648 062D 062F 0627 062A")), icon Or MB_RIGHT Or MB_RTLREADING
End Sub

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function

' ----------------------------------------------------------------------------
' Tests only
' ----------------------------------------------------------------------------

Public Sub SetLogFolderForTests(ByVal folder As String)
    mFolderForTests = folder
End Sub

Public Sub SetClockForTests(ByVal at As Date)
    mClockForTests = at
End Sub

Public Sub ResetLoggerForTests()
    mFolderForTests = vbNullString
    mClockForTests = 0
    mFileLoggingOff = False
    mCleanedOn = vbNullString
End Sub
