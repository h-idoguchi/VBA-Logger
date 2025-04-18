VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' CLASS_MODULE:   VBA Logger - allows 'log4VBA' style Logging in VBA
'''
'''                            - use Macro 'Test' from 'TestLogging' for testing and as an example
'''
''' Date                Developer                          Action
''' -----------------------------------------------------------------------------------
''' 28/08/08    Christian Bolterauer    Created
'''2016FEB25    Kevin Conner            Reduce to single class file for easy import
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Public iLogLevel As Integer
Public bUseLogPrefix As Boolean
Public bConsole As Boolean
Public bBuffer As Boolean
Public bToLogFile As Boolean
Public LogFilePath As String
Public TraceFilePath As String
Public bDelLogFileAtSetup As Boolean
Public PropsFileName As String
Public strLogbuffer As String

'ModuleName
Public ModuleName As String

'Define log levels
Public Enum LogLEVEL
  DISABLED = 0
  BASIC = 1 'like Debug.Print
  FATAL = 2
  WARN = 3
  INFO = 4
  FINE = 5
  FINER = 6
  FINEST = 7
  ALL = 8
End Enum

' The defaults
Const DEFAULT_LOG_LEVEL = LogLEVEL.INFO
Const DEFAULT_LOG_Console = True
Const DEFAULT_LOG_Buffer = False
Const DEFAULT_LOG_FILE = False
Const DEFAULT_PROPSFILE_NAME = "vba_log.properties"
'


'Class Konstructor
Private Sub Class_Initialize()
        On Error GoTo ErrHandler:
        Dim localpath As String
        bUseLogPrefix = True
        bDelLogFileAtSetup = True

        'default
        ModuleName = ""

        'set default location of props file to directory of this Logger and add default name
        localpath = getParentFolder(Application.VBE.ActiveVBProject.Filename) 'set path to location of file containing this Logger
        PropsFileName = localpath & "\" & DEFAULT_PROPSFILE_NAME

        'make sure defaults are set
        Call setLoggingParams(DEFAULT_LOG_LEVEL, DEFAULT_LOG_Console, DEFAULT_LOG_Buffer, DEFAULT_LOG_FILE)
        'set default log file path
        LogFilePath = localpath & "\" & "vba_logger.log"
        TraceFilePath = localpath & "\" & "vba_trace.log"
        
        'check if params can be set from a CustomDocumentProperties and overwrite defaults if available
        Call getLogParamsFromCustomDocumentProperties
        
        'set log file
        Call setLogFile(LogFilePath, False)

Exit Sub

ErrHandler:
  Debug.Print "Error in Logger.Class_Initialize & "; " & Err.Number & "; " & Err.Description"

End Sub

'set logging parameters
Public Sub setLoggingParams(level As LogLEVEL, toConsole As Boolean, toBuffer As Boolean, toLogFile As Boolean, Optional deleteExistingLogFile As Boolean)
         Dim delLogfile As Boolean

         iLogLevel = level
         bConsole = toConsole
         bBuffer = toBuffer
         bToLogFile = toLogFile

         If IsMissing(deleteExistingLogFile) Then
                delLogfile = False
         Else
                delLogfile = deleteExistingLogFile
         End If
         ' delete currently set Logfile if set
         If (delLogfile) Then deleteLogFile


End Sub

'The main log procedure
Public Sub log(sLogText As String, Optional level As LogLEVEL, Optional slogpoint As String)
        If level = 0 Then
            level = iLogLevel
        End If
        If (Me.iLogLevel > LogLEVEL.DISABLED And Me.iLogLevel >= level) Then
                If IsMissing(slogpoint) Then
                        Call WriteLog(sLogText, level, "")
                Else
                        Call WriteLog(sLogText, level, slogpoint)
                End If
        End If
End Sub

Private Sub WriteLog(sLogText As String, level As LogLEVEL, slogpoint As String)

    Dim LogMessage As String
    
    LogMessage = getLogPrefix(level, slogpoint) & sLogText
    
    ' write to console
    If Me.bConsole Then Debug.Print (LogMessage)
    ' write to Buffer
    If Me.bBuffer Then addline (LogMessage)
    ' write to file
    If Me.bToLogFile Then writeToLogFile (LogMessage)

End Sub

' get LogLevelName for Integer value
Public Function getLogLevelName(level As Integer)
        Dim myLevelName As String

        Select Case level
        Case LogLEVEL.DISABLED:
                myLevelName = "DISABLED"
        Case LogLEVEL.BASIC:
                myLevelName = "BASIC"
        Case LogLEVEL.INFO:
                myLevelName = "INFO:"
        Case LogLEVEL.WARN:
                myLevelName = "WARN:"
        Case LogLEVEL.FATAL:
                myLevelName = "FATAL:"
        Case LogLEVEL.FINE:
                myLevelName = "FINE:"
        Case LogLEVEL.FINER:
                myLevelName = "FINER:"
        Case LogLEVEL.FINEST:
                myLevelName = "FINEST:"
        Case LogLEVEL.ALL:
                myLevelName = "ALL:"

        Case Else
                myLevelName = "Level is not defined:"
  End Select

        getLogLevelName = myLevelName

End Function

Private Function getLogPrefix(level As LogLEVEL, logpoint As String)
        Dim sDateTime As String
        Dim myLevelPrefix As String
        Dim mySubModul As String
        Dim iLevel As Integer

  If Not (bUseLogPrefix) Or level = LogLEVEL.BASIC Then   'when level = LogLEVEL.BASIC no prefix to simulate Debug.Print
         getLogPrefix = ""
         Exit Function
        End If

        iLevel = level ' to Integer
        myLevelPrefix = getLogLevelName(iLevel)

        If (Len(Me.ModuleName) > 0 And Len(logpoint) > 0) Then
                mySubModul = "[" & Me.ModuleName & "::" & logpoint & "]"
        ElseIf (Len(logpoint) > 0) Then
                mySubModul = "[" & logpoint & "]"
        ElseIf (Len(Me.ModuleName) > 0) Then
                mySubModul = "[" & Me.ModuleName & "]"
        Else
                mySubModul = ""
        End If

        sDateTime = CStr(Now())
        'ToDo provide different output styles ..
        'getLogPrefix = myLevelPrefix & " (" & sDateTime & ") - "
        getLogPrefix = "(" & sDateTime & ")" & mySubModul & "-" & myLevelPrefix & "  "

End Function

Private Sub writeToLogFile(logmsg As String)
    Dim fileNum As Integer
    On Error GoTo ErrHandler:
    If Len(Me.LogFilePath) = 0 Then
       Debug.Print "Error: Log file path is empty."
       Exit Sub
    End If
    
    fileNum = FreeFile ' next file number
    Open Me.LogFilePath For Append As #fileNum ' creates the file if it doesn't exist
    Print #fileNum, logmsg ' write information at the end of the text file
    Close #fileNum ' close the file
Exit Sub

ErrHandler:
    Debug.Print "Error writing to Logfile: " & Me.LogFilePath & " " & Err.Number & " " & Err.Description

End Sub

Public Sub writeLogBuffer(Optional myfilePath As String)
    On Error GoTo ErrHandler:
    Dim mytracefile As String
    Dim lines() As String
    Dim line As Variant

    If Len(myfilePath) > 0 Then
        mytracefile = myfilePath
    ElseIf Len(Me.TraceFilePath) > 0 Then
        mytracefile = Me.TraceFilePath
    ElseIf Len(Me.LogFilePath) > 0 Then
        mytracefile = Me.LogFilePath
    Else
        Me.log "Error: Trace file path is empty.", FATAL
        Exit Sub
    End If
    
    'write to trace file
    Dim fileNum As Integer
    fileNum = FreeFile ' next file number
    Open mytracefile For Append As #fileNum ' creates the file if it doesn't exist
    lines = VBA.Split(Me.strLogbuffer, VBA.vbLf)
    For Each line In lines
       Print #fileNum, line ' write Logbuffer to text file
    Next line
    Close #fileNum ' close the file
Exit Sub

ErrHandler:
        Debug.Print "Error writing to Tracefile: " & mytracefile & " " & Err.Number & " " & Err.Description
End Sub

'delete log file currently set
Public Sub deleteLogFile()
  On Error GoTo ErrHandler:
  If (FileExists(Me.LogFilePath)) Then
         Kill (Me.LogFilePath)
  End If
Exit Sub

ErrHandler:
  Debug.Print "Error deleting Logfile " & Me.LogFilePath & " " & Err.Number & " " & Err.Description

End Sub
'set logfilepath
'- will delete an existing log file if bDelLogFileAtSetup is set to true
Public Sub setLogFile(filePath As String, delExitingFile As Boolean)
  On Error GoTo ErrHandler:

  Me.LogFilePath = filePath
  'delete if set to true
  If (delExitingFile) Then Call deleteLogFile
  If (bToLogFile) Then Debug.Print "Logfile set to: " & LogFilePath
Exit Sub

ErrHandler:
  Debug.Print "Error setLogFile " & LogFilePath & " " & Err.Number & " " & Err.Description

End Sub

Public Function getLogParamsFromCustomDocumentProperties() As Boolean
    Dim prop As Object
    On Error GoTo ErrHandler

    For Each prop In ThisWorkbook.CustomDocumentProperties
        Select Case UCase(prop.Name)
            Case "LOG_LEVEL"
                Call setLogLevel(CStr(prop.Value))
            Case "LOG_TO_CONSOLE"
                bConsole = valIsTrue(CStr(prop.Value))
            Case "LOG_TO_BUFFER"
                bBuffer = valIsTrue(CStr(prop.Value))
            Case "LOG_TO_FILE"
                bToLogFile = valIsTrue(CStr(prop.Value))
            Case "LOG_FILE_PATH"
                If Len(CStr(prop.Value)) > 0 Then LogFilePath = CStr(prop.Value)
            Case "TRACE_FILE_PATH"
                If Len(CStr(prop.Value)) > 0 Then TraceFilePath = CStr(prop.Value)
        End Select
    Next prop

    getLogParamsFromCustomDocumentProperties = True
    Exit Function

ErrHandler:
    Debug.Print "Error getLogParamsFromCustomDocumentProperties: " & Err.Number & " " & Err.Description
    getLogParamsFromCustomDocumentProperties = False
End Function

Private Sub setLogLevel(level As String)
  Dim mylevel
  mylevel = UCase(level)

  Select Case mylevel
        Case "DISABLED":
           iLogLevel = LogLEVEL.DISABLED
        Case "BASIC":
           iLogLevel = LogLEVEL.BASIC
        Case "INFO":
           iLogLevel = LogLEVEL.INFO
        Case "WARN":
           iLogLevel = LogLEVEL.WARN
        Case "FATAL":
           iLogLevel = LogLEVEL.FATAL
        Case "FINE":
           iLogLevel = LogLEVEL.FINE
        Case "FINER":
           iLogLevel = LogLEVEL.FINER
        Case "FINEST":
           iLogLevel = LogLEVEL.FINEST
        Case "ALL":
           iLogLevel = LogLEVEL.ALL
  End Select
End Sub

Public Function getLogBuffer() As String
   getLogBuffer = strLogbuffer
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Utils
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'-- extract full path of parent folder of file mypath
Public Function getParentFolder(mypath As String) As String
  Dim pos As Integer

  pos = InStrRev(mypath, "\")
  If (pos <> 0) Then
        getParentFolder = Strings.Left(mypath, pos - 1)
        Exit Function
  End If
  getParentFolder = ""
End Function

'-- Check File Exists --
Public Function FileExists(path As String) As Boolean
  FileExists = (Dir(path) <> "")
End Function

'-- isNothing --
Function checkIsNothing(obj As Object)
  If (obj Is Nothing) Then
         checkIsNothing = True
  Else
         checkIsNothing = False
  End If
End Function

'-- parameters ---
Function readParameter(line As String) As String()
  Dim txtarr() As String
  Dim proparray(2) As String

  txtarr = VBA.Split(line, "=")
  If (UBound(txtarr) > 0) Then
        proparray(0) = VBA.Trim(txtarr(0))
        proparray(1) = VBA.Trim(txtarr(1))
        readParameter = proparray
  Else
        readParameter = proparray
  End If
End Function

'-- check text coded boolean value (if read from text file) --
Function valIsTrue(boolval As String) As Boolean
        If ("TRUE" = VBA.UCase(boolval)) Then
                valIsTrue = True
                Exit Function
        End If
        valIsTrue = False
End Function

Public Sub addline(logmsg As String)
   If (Len(strLogbuffer) > 0) Then
         strLogbuffer = strLogbuffer & vbLf & logmsg
   Else
         strLogbuffer = logmsg  'avoid empty line when strLogbuffer=""
   End If
End Sub

Private Sub Class_Terminate()

End Sub
''' デフォルト値をもとにカスタムドキュメントプロパティを生成／更新する
Public Sub createCustomDocumentProperties(Optional overwriteExisting As Boolean = False)
    Dim props      As Object
    Dim prop       As Object
    Dim defaultVal As Variant

    On Error GoTo ErrHandler
    Set props = ThisWorkbook.CustomDocumentProperties

    ' 既存プロパティを削除（上書きモードの場合）
    If overwriteExisting Then
        For Each prop In props
            Select Case UCase(prop.Name)
                Case "LOG_LEVEL", "LOG_TO_CONSOLE", "LOG_TO_BUFFER", _
                     "LOG_TO_FILE",  "LOG_FILE_PATH",   "TRACE_FILE_PATH"
                    prop.Delete
            End Select
        Next prop
    End If

    ' 各プロパティを追加（存在しなければ）
    If Not propertyExists("LOG_LEVEL") Then
        ' デフォルトのログレベルは定数 DEFAULT_LOG_LEVEL（数値）なので、
        ' 名前で扱いたい場合は getLogLevelName で文字列化し、末尾の「:」を除去
        defaultVal = Replace(getLogLevelName(DEFAULT_LOG_LEVEL), ":", "")
        props.Add Name:="LOG_LEVEL", LinkToContent:=False, _
                  Type:=msoPropertyTypeText, Value:=defaultVal
    End If

    If Not propertyExists("LOG_TO_CONSOLE") Then
        props.Add Name:="LOG_TO_CONSOLE", LinkToContent:=False, _
                  Type:=msoPropertyTypeBoolean, Value:=DEFAULT_LOG_Console
    End If

    If Not propertyExists("LOG_TO_BUFFER") Then
        props.Add Name:="LOG_TO_BUFFER", LinkToContent:=False, _
                  Type:=msoPropertyTypeBoolean, Value:=DEFAULT_LOG_Buffer
    End If

    If Not propertyExists("LOG_TO_FILE") Then
        props.Add Name:="LOG_TO_FILE", LinkToContent:=False, _
                  Type:=msoPropertyTypeBoolean, Value:=DEFAULT_LOG_FILE
    End If

    If Not propertyExists("LOG_FILE_PATH") Then
        props.Add Name:="LOG_FILE_PATH", LinkToContent:=False, _
                  Type:=msoPropertyTypeText, Value:=Me.LogFilePath
    End If

    If Not propertyExists("TRACE_FILE_PATH") Then
        props.Add Name:="TRACE_FILE_PATH", LinkToContent:=False, _
                  Type:=msoPropertyTypeText, Value:=Me.TraceFilePath
    End If

    Exit Sub

ErrHandler:
    Debug.Print "Error in createCustomDocumentProperties: " & Err.Number & " - " & Err.Description
End Sub

''' 指定した名前のカスタムプロパティが存在するかチェック
Private Function propertyExists(propName As String) As Boolean
    Dim prop As Object
    On Error Resume Next
    propertyExists = False
    For Each prop In ThisWorkbook.CustomDocumentProperties
        If UCase(prop.Name) = UCase(propName) Then
            propertyExists = True
            Exit For
        End If
    Next prop
    On Error GoTo 0
End Function

