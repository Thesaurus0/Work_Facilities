VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum enLogLevel
    Off = 0
    Debug__ = 1
    Info_ = 2
    Warning_ = 3
    Error_ = 4
    Fatal_ = 5
End Enum
Public Enum enLogMode
    Append = 1
    Overwrite = 2
End Enum

Private Const BUSINESS_ERROR_NUMBER As Integer = 2000

Private sLogFilePath As String
Private sUserName As String
Private fso As Object
Private Const LOG_FILE_MAX_SIZE As Long = 2000000

Private iLogLevel As enLogLevel
Private iLogMode As enLogMode

Property Let Level(val As enLogLevel)
    iLogLevel = val
End Property
'-----------------------------------------------------------
Property Let Mode(val As enLogMode)
    iLogMode = val
    
    If iLogMode = Overwrite Then ClearLogFile
End Property
'-----------------------------------------------------------

'-----------------------------------------------------------
Private Sub Class_Initialize()
    sUserName = Environ("UserName")
    
    iLogLevel = Off
    iLogMode = Append
    
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
    
    sLogFilePath = Left(ThisWorkbook.FullName, InStr(1, ThisWorkbook.FullName, ".xl", vbTextCompare)) & "log"
    
    CheckFileSize
End Sub
Private Sub Class_Terminate()
    Set fso = Nothing
End Sub
'-----------------------------------------------------------

'-----------------------------------------------------------
Function Debug_(sMsg As String, Optional sSource As String = "")
    Call Write2LogFile(enLogLevel.Debug__, "DEBUG", sMsg, sSource)
End Function
Function Info(sMsg As String, Optional sSource As String = "")
    Call Write2LogFile(enLogLevel.Info_, "INFO", sMsg, sSource)
End Function
Function Warning(sMsg As String, Optional sSource As String = "")
    Call Write2LogFile(enLogLevel.Warning_, "WARNING", sMsg, sSource)
End Function
Function Error(sMsg As String, Optional sSource As String = "")
    Call Write2LogFile(enLogLevel.Error_, "ERROR", sMsg, sSource)
End Function
Function Fatal(sMsg As String, Optional sSource As String = "")
    Call Write2LogFile(enLogLevel.Fatal_, "FATAL", sMsg, sSource)
End Function
'-----------------------------------------------------------

'-----------------------------------------------------------
Private Function Write2LogFile(iLogInfotype As enLogLevel, sLevelTag As String, Optional sMsg As String = "", Optional sSource As String = "")
    Dim fileNum As Integer
    
    If Skip(iLogInfotype) Then Exit Function
   
    fileNum = FreeFile
    Open sLogFilePath For Append As #fileNum
    
    If Len(sSource) > 0 Then
        Print #fileNum, Format(Now, "yyyy-mm-dd HH:MM:SS") & vbTab & sUserName & vbTab & sLevelTag & vbTab & sSource & vbTab & sMsg '
    Else
        Print #fileNum, Format(Now, "yyyy-mm-dd HH:MM:SS") & vbTab & sUserName & vbTab & sLevelTag & vbTab & sMsg
    End If
    
    Close #fileNum
End Function
Private Function Skip(iLogInfoLvl As enLogLevel) As Boolean
    Skip = False
    
    If iLogInfoLvl = Off Then
        Skip = True
    Else
        If iLogInfoLvl < iLogLevel Then Skip = True
    End If
End Function
Private Function ClearLogFile()
    Dim fileNum As Integer
    
    fileNum = FreeFile
    Open sLogFilePath For Output As #fileNum
    Close #fileNum
End Function
Private Function CheckFileSize()
    If fso.FileExists(sLogFilePath) Then
        If FileLen(sLogFilePath) > BUSINESS_ERROR_NUMBER Then
            BackupFile sLogFilePath, sLogFilePath & "." & Format(Now, "yyyymmddhhmmss")
            Kill sLogFilePath
        End If
    End If
End Function
Private Function BackupFile(sOrigFilePath As String, sBackUpFilePath As String, Optional bOverWriteWhenBackupFileExists As Boolean = True)
    If Len(Trim(sOrigFilePath)) <= 0 Then Err.Raise vbObjectError + BUSINESS_ERROR_NUMBER, "BackupFile", "sOrigFilePath is blank"
    If Len(Trim(sBackUpFilePath)) <= 0 Then Err.Raise vbObjectError + BUSINESS_ERROR_NUMBER, "BackupFile", "sBackUpFilePath is blank"

    If Not fso.FileExists(sOrigFilePath) Then Err.Raise vbObjectError + BUSINESS_ERROR_NUMBER, "BackupFile", "The source file does not exists. " & vbCrLf & sOrigFilePath
    
    If fso.FileExists(sBackUpFilePath) Then
        If bOverWriteWhenBackupFileExists Then
            
        Else
            Err.Raise vbObjectError + BUSINESS_ERROR_NUMBER, "BackupFile", "The backup file already exists. " & vbCrLf & sBackUpFilePath
        End If
    End If
    
    Call fso.CopyFile(sOrigFilePath, sBackUpFilePath, True)
End Function
