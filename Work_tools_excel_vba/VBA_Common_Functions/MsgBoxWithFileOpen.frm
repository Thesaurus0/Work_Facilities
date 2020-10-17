VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MsgBoxWithFileOpen 
   Caption         =   "Process Completed"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10575
   OleObjectBlob   =   "MsgBoxWithFileOpen.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "MsgBoxWithFileOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sHeaderMsg As String
Private sFilePath As String

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String _
    , ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String _
    , ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Private Function OpenFile(asFileFullPath As String)
    Dim lReturnVal As LongPtr
    Dim msg As String
    
    Const SW_HIDE = 0&   '{隐藏}
    Const SW_SHOWNORMAL = 1&   '{用最近的大小和位置显示, 激活}
    Const SW_SHOWMINIMIZED = 2&   '{最小化, 激活}
    Const SW_SHOWMAXIMIZED = 3&   '{最大化, 激活}
    Const SW_SHOWNOACTIVATE = 4&   '{用最近的大小和位置显示, 不激活}
    Const SW_SHOW = 5&   '{同 SW_SHOWNORMAL}
    Const SW_MINIMIZE = 6&   '{最小化, 不激活}
    Const SW_SHOWMINNOACTIVE = 7&   '{同 SW_MINIMIZE}
    Const SW_SHOWNA = 8&   '{同 SW_SHOWNOACTIVATE}
    Const SW_RESTORE = 9&   '{同 SW_SHOWNORMAL}
    Const SW_SHOWDEFAULT = 10&   '{同 SW_SHOWNORMAL}
    
    Const ERROR_FILE_NOT_FOUND = 2&
    Const ERROR_PATH_NOT_FOUND = 3&
    Const SE_ERR_ACCESSDENIED = 5&
    Const SE_ERR_OOM = 8&
    Const SE_ERR_DLLNOTFOUND = 32&
    Const SE_ERR_SHARE = 26&
    Const SE_ERR_ASSOCINCOMPLETE = 27&
    Const SE_ERR_DDETIMEOUT = 28&
    Const SE_ERR_DDEFAIL = 29&
    Const SE_ERR_DDEBUSY = 30&
    Const SE_ERR_NOASSOC = 31&
    Const ERROR_BAD_FORMAT = 11&

    lReturnVal = ShellExecute(Application.hwnd, "Open", asFileFullPath, "", "C:\", SW_SHOWMAXIMIZED)
    
    If lReturnVal <= 32 Then
        Select Case lReturnVal
            Case ERROR_FILE_NOT_FOUND
                msg = "File not found"
            Case ERROR_PATH_NOT_FOUND
                msg = "Path not found"
            Case SE_ERR_ACCESSDENIED
                msg = "Access denied"
            Case SE_ERR_OOM
                msg = "Out of memory"
            Case SE_ERR_DLLNOTFOUND
                msg = "DLL not found"
            Case SE_ERR_SHARE
                msg = "A sharing violation occurred"
            Case SE_ERR_ASSOCINCOMPLETE
                msg = "Incomplete or invalid file association"
            Case SE_ERR_DDETIMEOUT
                msg = "DDE Time out"
            Case SE_ERR_DDEFAIL
                msg = "DDE transaction failed"
            Case SE_ERR_DDEBUSY
                msg = "DDE busy"
            Case SE_ERR_NOASSOC
                msg = "No association for file extension"
            Case ERROR_BAD_FORMAT
                msg = "Invalid EXE file or error in EXE image"
            Case Else
                msg = "Unknown error"
        End Select
        
        Err.Raise vbObjectError + 2000, "", msg
    End If
End Function
'-----------------------------------------------------------
Property Let HeaderMsg(val As String)
    sHeaderMsg = val
End Property
'-----------------------------------------------------------
Property Let FilePath(val As String)
    sFilePath = val
End Property
'-----------------------------------------------------------

Private Sub UserForm_Initialize()
    Me.lblHeaderMsg.Caption = "Process completed, report was generated as below:"
End Sub

Private Sub UserForm_Terminate()
    sFilePath = ""
End Sub

Private Sub UserForm_Activate()
    Me.lblHeaderMsg.Top = 10
    Me.lblHeaderMsg.Left = 18
    Me.lblHeaderMsg.Width = IIf(Len(sHeaderMsg) > Len(sFilePath) + 60, 500, 550)
    
    Me.lblFilePath.Width = 350
    
    Me.lblFilePath.Text = sFilePath
    Me.lblHeaderMsg.Caption = sHeaderMsg
    
    AdjustControlPosition
    
    SetFocus Me.lblFilePath
End Sub
 
Private Sub cbOk_Click()
    Unload Me
End Sub

Private Sub cbOpenFile_Click()
    If Dir(sFilePath) = "" Then
        MsgBox "File does not exists", vbCritical
    Else
        If InStr(1, sFilePath, ".xl", vbTextCompare) > 0 Then
            Dim wb As Workbook
            Set wb = Application.Workbooks.Open(sFilePath)
            wb.Activate
            Set wb = Nothing
        Else
            OpenFile sFilePath
        End If
        
        Unload Me
    End If
End Sub

Private Sub cbOpenFolder_Click()
    Dim sFolder As String
    sFolder = Left(sFilePath, InStrRev(sFilePath, Application.PathSeparator))
    
    If Dir(sFolder, vbDirectory) <> "" Then
        OpenFile sFolder
        Unload Me
    Else
        MsgBox "Folder does not exists", vbCritical
    End If
End Sub

Private Function AdjustControlPosition()
    Me.imgIcon.Left = Me.lblHeaderMsg.Left
    Me.lblFilePath.Left = Me.imgIcon.Left + Me.imgIcon.Width + 10
    
    Me.lblFilePath.Top = Me.lblHeaderMsg.Top + Me.lblHeaderMsg.Height + 10
    
    Me.cbOpenFile.Top = Me.lblFilePath.Top + (Me.lblFilePath.Height - Me.cbOpenFile.Height) / 2
    Me.cbOpenFile.Left = Me.lblFilePath.Left + Me.lblFilePath.Width + 10
    
    Me.cbOpenFolder.Top = Me.cbOpenFile.Top
    Me.cbOpenFolder.Left = Me.cbOpenFile.Left + Me.cbOpenFile.Width + 3
    
    Me.cbOk.Top = Me.lblFilePath.Top + Me.lblFilePath.Height + 15
    
    Me.Height = Me.cbOk.Top + Me.cbOk.Height + 45
    Me.Width = IIf(Me.lblHeaderMsg.Left + Me.lblHeaderMsg.Width + 30 > Me.cbOpenFolder.Left + Me.cbOpenFolder.Width + 30 _
                , Me.lblHeaderMsg.Left + Me.lblHeaderMsg.Width + 30 _
                , Me.cbOpenFolder.Left + Me.cbOpenFolder.Width + 30)
                
    Me.cbOk.Left = (Me.Width - Me.cbOk.Width) / 2
    Me.imgIcon.Top = Me.cbOpenFile.Top
End Function

Private Function SetFocus(control)
    control.SelStart = 0
    control.SelLength = Len(control.Value)
    control.SetFocus
End Function
