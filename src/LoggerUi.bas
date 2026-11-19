Attribute VB_Name = "LoggerUi"
Option Explicit

' ============================================================================
' GCV Logger add-in: start-up, ribbon group "logs" (show logs, report an
' issue) and the issue report.
'
' Report an issue:
'   1. choose: all GCV add-ins (default) or one of them
'   2. optional short description
'   3. report.txt (versions, Excel, Windows, description) + the last 14 days
'      of logs (only that add-in's lines when one was chosen) are zipped to
'      the Desktop: "GCV report yyyy-mm-dd hhnn.zip"
'   4. an Outlook e-mail to SUPPORT_EMAIL is opened with the zip attached; when
'      Outlook is not available, the Desktop folder is opened instead.
' Logs hold what the add-ins did, never personal data.
' ============================================================================

Private Declare PtrSafe Function MessageBoxW Lib "user32" (ByVal hWnd As LongPtr, _
    ByVal lpText As LongPtr, ByVal lpCaption As LongPtr, ByVal uType As Long) As Long

' Address that receives issue reports (leave "" to let the secretary type it).
Public Const SUPPORT_EMAIL As String = "it.team@thegazachildrenvillage.org"

Public Const REPORT_DAYS As Long = 14
Private Const LOG_SOURCE As String = "logger"
Private Const LOADER_FILE As String = "GCV Loader.xlam"
Private Const MB_ICONINFO As Long = &H40
Private Const MB_ICONERROR As Long = &H10
Private Const MB_RIGHT As Long = &H80000
Private Const MB_RTLREADING As Long = &H100000

Private mMessagesForTests As Collection

' ----------------------------------------------------------------------------
' Add-in life cycle (ThisWorkbook and GcvAddin call these)
' ----------------------------------------------------------------------------

Public Sub AddinOpened()
    On Error Resume Next                         ' start-up must never show a VBA error
    LogInfo LOG_SOURCE, GcvAddin.ADDIN_NAME & " " & GcvAddin.ADDIN_VERSION & " opened."
End Sub

Public Sub AddinClosing()
End Sub

Public Sub AddinSelfCheck()
    GcvLog.CleanText vbNullString
    FilterLines vbNullString, "x"
End Sub

' ----------------------------------------------------------------------------
' Ribbon callbacks (customUI14.xml)
' ----------------------------------------------------------------------------

Public Sub Ribbon_ShowLogs(control As IRibbonControl)
    ShowLogs
End Sub

Public Sub Ribbon_ReportIssue(control As IRibbonControl)
    ReportIssue
End Sub

' Opens the log folder with today's file selected (or just the folder).
Public Sub ShowLogs()
    Dim today As String
    LogInfo LOG_SOURCE, "Show logs."
    today = GcvLog.LogFileFor(Date)
    If CreateObject("Scripting.FileSystemObject").FileExists(today) Then
        OpenExplorer "/select,""" & today & """"
    Else
        OpenExplorer """" & GcvLog.LogFolder() & """"
    End If
End Sub

Public Sub ReportIssue()
    Dim addins As Collection
    Dim choice As Variant
    Dim description As Variant
    Dim scopeSource As String
    Dim scopeName As String
    Dim result As String

    On Error GoTo Failed
    Set addins = InstalledAddins()
    choice = Application.InputBox(ScopePrompt(addins), T_Title(), 0, Type:=1)
    If VarType(choice) = vbBoolean Then Exit Sub                  ' Cancel
    If choice < 0 Or choice > addins.Count Then choice = 0
    If choice = 0 Then
        scopeName = FromCodePoints("062C 0645 064A 0639 0020 0627 0644 0623 062F 0648 0627 062A")
    Else
        scopeSource = addins(choice)("source")
        scopeName = addins(choice)("name")
    End If

    description = Application.InputBox(FromCodePoints("0635 0641 0020 0627 0644 0645 0634 0643 0644 0629 0020 0628 0627 062E 062A 0635 0627 0631 0020 0028 0627 062E 062A 064A 0627 0631 064A 0029 003A"), T_Title(), "", Type:=2)
    If VarType(description) = vbBoolean Then Exit Sub             ' Cancel

    LogInfo LOG_SOURCE, "Issue report started (" & IIf(Len(scopeSource) = 0, "all add-ins", scopeSource) & ")."
    result = BuildReport(scopeSource, scopeName, CStr(description), addins, DesktopFolder(), Now)
    LogInfo LOG_SOURCE, "Issue report saved: " & Mid$(result, InStrRev(result, "\") + 1)
    If Not OpenEmailDraft(result, scopeName, CStr(description)) Then
        OpenExplorer "/select,""" & result & """"
        Notify FromCodePoints("062A 0645 0020 062D 0641 0638 0020 0627 0644 062A 0642 0631 064A 0631 0020 0639 0644 0649 0020 0633 0637 062D 0020 0627 0644 0645 0643 062A 0628 002E 0020 0623 0631 0633 0644 0647 0020 0625 0644 0649 0020 0627 0644 0645 0633 0624 0648 0644 0020 0627 0644 062A 0642 0646 064A 003A") & vbCrLf & result, MB_ICONINFO
    End If
    Exit Sub

Failed:
    LogError LOG_SOURCE, "Issue report failed"
    Notify FromCodePoints("062A 0639 0630 0631 0020 0625 0646 0634 0627 0621 0020 0627 0644 062A 0642 0631 064A 0631 003A") & vbCrLf & Err.Description, MB_ICONERROR
End Sub

' ----------------------------------------------------------------------------
' Building the report (tested)
' ----------------------------------------------------------------------------

' Writes report.txt and the logs into a new folder, zips it into `targetFolder`
' and returns the zip path (or the folder path when zipping is not possible).
Public Function BuildReport(ByVal scopeSource As String, ByVal scopeName As String, _
                            ByVal description As String, ByVal addins As Collection, _
                            ByVal targetFolder As String, ByVal at As Date) As String
    Dim fso As Object
    Dim baseName As String
    Dim work As String
    Dim offset As Long
    Dim logPath As String
    Dim text As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    baseName = "GCV report " & Format$(at, "yyyy-mm-dd hhnnss")
    work = fso.GetSpecialFolder(2).Path & "\" & baseName           ' 2 = temporary folder
    If fso.FolderExists(work) Then fso.DeleteFolder work, True
    fso.CreateFolder work
    fso.CreateFolder work & "\logs"

    WriteUtf8 work & "\report.txt", ReportText(scopeName, description, addins, at)
    For offset = REPORT_DAYS - 1 To 0 Step -1
        logPath = GcvLog.LogFileFor(DateAdd("d", -offset, at))
        text = vbNullString
        If fso.FileExists(logPath) Then
            ' ReadAll raises on an empty file
            If fso.GetFile(logPath).Size > 0 Then text = fso.OpenTextFile(logPath, 1, False, 0).ReadAll
            If Len(scopeSource) > 0 Then text = FilterLines(text, scopeSource)
            If Len(text) > 0 Then WriteUtf8 work & "\logs\" & fso.GetFileName(logPath), text
        End If
    Next offset

    BuildReport = ZipFolder(work, targetFolder & "\" & baseName & ".zip")
    If Len(BuildReport) = 0 Then
        ' No zip (PowerShell blocked): hand over the folder itself.
        fso.CopyFolder work, targetFolder & "\" & baseName, True
        BuildReport = targetFolder & "\" & baseName
    End If
    On Error Resume Next
    fso.DeleteFolder work, True
    On Error GoTo 0
End Function

' Lines of one source ("[student-lookup] - ...") only; "" keeps nothing.
Public Function FilterLines(ByVal text As String, ByVal source As String) As String
    Dim lines() As String
    Dim i As Long
    Dim result As String
    Dim prefix As String

    If Len(text) = 0 Then Exit Function
    prefix = "[" & source & "] "
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        If Left$(lines(i), Len(prefix)) = prefix Then result = result & lines(i) & vbCrLf
    Next i
    FilterLines = result
End Function

Public Function ReportText(ByVal scopeName As String, ByVal description As String, _
                           ByVal addins As Collection, ByVal at As Date) As String
    Dim text As String
    Dim addin As Variant

    text = "GCV Tools - issue report" & vbCrLf & _
           "Created: " & Format$(at, "yyyy-mm-dd hh:nn") & vbCrLf & _
           "About: " & scopeName & vbCrLf & vbCrLf & _
           "Description:" & vbCrLf & IIf(Len(Trim$(description)) = 0, "(none)", description) & vbCrLf & vbCrLf & _
           "Installed add-ins:" & vbCrLf
    For Each addin In addins
        text = text & "  " & addin("name") & " " & addin("version") & "  (" & addin("id") & ")" & vbCrLf
    Next addin
    text = text & "  GCV Loader " & LoaderVersion() & vbCrLf & vbCrLf & _
           "Excel: " & Application.Version & " build " & Application.Build & vbCrLf & _
           "Windows: " & Application.OperatingSystem & vbCrLf & _
           "Logs: last " & REPORT_DAYS & " days" & vbCrLf
    ReportText = text
End Function

' Numbered list for the choice: 0 = all (default).
Public Function ScopePrompt(ByVal addins As Collection) As String
    Dim i As Long
    ScopePrompt = FromCodePoints("0639 0646 0020 0623 064A 0020 0623 062F 0627 0629 0020 062A 0631 064A 062F 0020 0627 0644 0625 0628 0644 0627 063A 061F 0020 0627 0643 062A 0628 0020 0627 0644 0631 0642 0645 0020 062B 0645 0020 0627 0636 063A 0637 0020 004F 004B 003A") & vbCrLf & vbCrLf & _
                  "0 = " & FromCodePoints("062C 0645 064A 0639 0020 0627 0644 0623 062F 0648 0627 062A 0020 0028 0627 0644 0623 0641 0636 0644 0020 0625 0646 0020 0644 0645 0020 062A 0643 0646 0020 0645 062A 0623 0643 062F 0627 064B 0029")
    For i = 1 To addins.Count
        ScopePrompt = ScopePrompt & vbCrLf & i & " = " & addins(i)("name")
    Next i
End Function

' Every GCV add-in opened by the loader: Dictionaries name/version/id/source.
' Always contains this logger, also when the loader is not running (DEV).
Public Function InstalledAddins() As Collection
    Dim result As Collection
    Dim files As String
    Dim fileName As Variant
    Dim info As String

    Set result = New Collection
    files = LoadedAddinFiles()
    If Len(files) = 0 Then files = GcvAddin.ADDIN_FILE
    For Each fileName In Split(files, ";")
        If StrComp(fileName, GcvAddin.ADDIN_FILE, vbTextCompare) = 0 Then
            info = GcvAddin.GcvAddinInfo()
        Else
            info = RunText("'" & fileName & "'!GcvAddin.GcvAddinInfo")
        End If
        If Len(info) > 0 Then result.Add AddinFromInfo(info)
    Next fileName
    Set InstalledAddins = result
End Function

' "gcv/student-lookup|3.0.0|owner/repo|GCV Student Lookup"
Public Function AddinFromInfo(ByVal info As String) As Object
    Dim parts() As String
    Dim addin As Object
    parts = Split(info, "|")
    Set addin = CreateObject("Scripting.Dictionary")
    addin("id") = parts(0)
    addin("version") = parts(1)
    addin("name") = IIf(UBound(parts) >= 3, parts(UBound(parts)), parts(0))
    addin("source") = Mid$(parts(0), InStr(parts(0), "/") + 1)
    Set AddinFromInfo = addin
End Function

' ----------------------------------------------------------------------------
' Helpers
' ----------------------------------------------------------------------------

Private Function LoadedAddinFiles() As String
    If Not IsWorkbookOpen(LOADER_FILE) Then Exit Function
    LoadedAddinFiles = RunText("'" & LOADER_FILE & "'!GcvLoader.LoadedAddins")
End Function

Private Function LoaderVersion() As String
    If IsWorkbookOpen(LOADER_FILE) Then LoaderVersion = RunText("'" & LOADER_FILE & "'!GcvLoader.LoaderVersion")
    If Len(LoaderVersion) = 0 Then LoaderVersion = "(not running)"
End Function

Private Function RunText(ByVal macro As String) As String
    On Error Resume Next
    RunText = CStr(Application.Run(macro))
    If Err.Number <> 0 Then RunText = vbNullString
    Err.Clear
    On Error GoTo 0
End Function

Private Function IsWorkbookOpen(ByVal workbookName As String) As Boolean
    Dim wb As Workbook
    On Error Resume Next
    Set wb = Workbooks(workbookName)
    On Error GoTo 0
    IsWorkbookOpen = Not wb Is Nothing
End Function

' PowerShell's Compress-Archive (Windows 10 and newer). Returns "" on failure.
Private Function ZipFolder(ByVal folder As String, ByVal zipPath As String) As String
    Dim command As String
    Dim exitCode As Long

    On Error GoTo Failed
    command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & _
              """Compress-Archive -Path '" & Replace(folder, "'", "''") & "\*' -DestinationPath '" & _
              Replace(zipPath, "'", "''") & "' -Force"""
    exitCode = CreateObject("WScript.Shell").Run(command, 0, True)      ' hidden, wait
    If exitCode = 0 And CreateObject("Scripting.FileSystemObject").FileExists(zipPath) Then ZipFolder = zipPath
    If Len(ZipFolder) = 0 Then LogWarn LOG_SOURCE, "Zip failed (exit code " & exitCode & "); folder handed over instead."
    Exit Function
Failed:
    LogWarn LOG_SOURCE, "Zip failed: " & Err.Description
End Function

Private Function OpenEmailDraft(ByVal attachment As String, ByVal scopeName As String, _
                                ByVal description As String) As Boolean
    Dim outlook As Object
    Dim mail As Object

    If Not mMessagesForTests Is Nothing Then Exit Function
    If Right$(LCase$(attachment), 4) <> ".zip" Then Exit Function    ' a folder cannot be attached
    On Error GoTo NoOutlook
    Set outlook = CreateObject("Outlook.Application")
    Set mail = outlook.CreateItem(0)                                  ' 0 = mail item
    mail.To = SUPPORT_EMAIL
    mail.Subject = "GCV Tools - " & scopeName & " - " & Format$(Now, "yyyy-mm-dd")
    mail.Body = FromCodePoints("0627 0644 0648 0635 0641 003A") & vbCrLf & description & vbCrLf & vbCrLf & _
                FromCodePoints("0627 0644 062A 0642 0631 064A 0631 0020 0648 0627 0644 0633 062C 0644 0627 062A 0020 0645 0631 0641 0642 0629 002E")
    mail.Attachments.Add attachment
    mail.Display
    LogInfo LOG_SOURCE, "Issue report e-mail opened in Outlook."
    OpenEmailDraft = True
    Exit Function
NoOutlook:
    LogInfo LOG_SOURCE, "Outlook not available (" & Err.Description & "); report left on the Desktop."
End Function

Private Sub WriteUtf8(ByVal path As String, ByVal text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2                        ' text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile path, 2              ' overwrite
    stream.Close
End Sub

' WScript.Shell.Run keeps Unicode paths (Arabic user names); VBA's Shell does not.
Private Sub OpenExplorer(ByVal arguments As String)
    CreateObject("WScript.Shell").Run "explorer.exe " & arguments, 1, False
End Sub

Private Function DesktopFolder() As String
    DesktopFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")
End Function

' Tests collect messages instead of showing them.
Public Sub UseTestMessages(ByVal enabled As Boolean)
    If enabled Then Set mMessagesForTests = New Collection Else Set mMessagesForTests = Nothing
End Sub

Private Sub Notify(ByVal text As String, ByVal icon As Long)
    If Not mMessagesForTests Is Nothing Then
        mMessagesForTests.Add text
        Exit Sub
    End If
    MessageBoxW Application.hWnd, StrPtr(text), StrPtr(T_Title()), icon Or MB_RIGHT Or MB_RTLREADING
End Sub

Public Function T_Title() As String
    T_Title = FromCodePoints("0623 062F 0648 0627 062A 0020 0642 0631 064A 0629 0020 0627 0644 0623 0637 0641 0627 0644 0020 002D 0020 0627 0644 0625 0628 0644 0627 063A 0020 0639 0646 0020 0645 0634 0643 0644 0629")
End Function

Private Function FromCodePoints(ByVal codePoints As String) As String
    Dim parts() As String
    Dim i As Long
    parts = Split(codePoints, " ")
    For i = LBound(parts) To UBound(parts)
        FromCodePoints = FromCodePoints & ChrW$(CLng("&H" & parts(i)))
    Next i
End Function
