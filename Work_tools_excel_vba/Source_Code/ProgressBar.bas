VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private objApp As Object                              '指向当前的文档，如ThisDocument或ThisWorkBook
'Private uForm As Object                               '进度条窗体
Private lbl1 As Object                                '显示标签文字 MSForms.Label
Private lbl2 As Object                                '显示进度 MSForms.Label
'Private FormName As String

'窗体风格
Private Const GWL_STYLE As Long = (-16)
Private Const WS_CAPTION As Long = &HC00000
Private BarLength As Long  '= 300                 '进度条长度

#If Win64 Then
    Private Declare PtrSafe Function DrawMenuBar Lib "User32" (ByVal hwnd As Long) As Long
    Private Declare PtrSafe Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Sub Class_Initialize()
    '在Office会出现窗体名不能重用的BUG，即使用VBA创建窗体后，删除，再创建同名窗体会报错“文件/路径访问错误”
    '微软已经确认了该BUG的存在 http://support.microsoft.com/kb/244238/zh-cn
    '需要每次启动的时候，创建一个新名称的窗体
    t = Timer
    ms = t - Int(t)                                       '计算毫秒
    BarLength = frmProgressBar.Width
    'FormName = "FORM" & format(Now, "ddhhmmss") & Replace(ms, ".", "")
End Sub

'创建进度条
Public Sub ShowBar()
    CreateProgressBar
End Sub

'销毁进度条
Public Sub DestroyBar()
    Unload frmProgressBar   'uForm
    'RemoveModule FormName
    'Set frmProgressBar = Nothing    'uForm
   ' Set objApp = Nothing
End Sub

'设置进度条进度
Public Sub ChangeProcessBarValue(value As Double, Optional Message As String = "")
    On Error Resume Next
    lbl1.Width = Int(value * BarLength)                   '显示进度条
    lbl2.Caption = IIf(Message = "", Format(value, "已经完成 0.00%"), Message)
    DoEvents                                              '转让控制权给操作系统
End Sub

'阻塞进程
Public Sub SleepBar(ms As Long)
    Sleep ms
End Sub

'创建进度条对象
Private Sub CreateProgressBar()
'    Dim UsForm As Object
'    If InStr(1, Application.Name, "Word") > 0 Then
'        Set objApp = ThisDocument
'    ElseIf InStr(1, Application.Name, "Excel") > 0 Then
'        Set objApp = ThisWorkbook
'    ElseIf InStr(1, Application.Name, "PowerPoint") > 0 Then
'        Set objApp = ActivePresentation
'    ElseIf InStr(1, Application.Name, "Access") > 0 Then
'        Set objApp = Application.VBE                      'Access
'    End If
'    '创建一个窗体。不能中断运行。
'    RemoveModule FormName
'    Set UsForm = objApp.VBProject.VBComponents.Add(3)
'    With UsForm
'        '由于该窗体还未运行，相当于处于设计状态
'        '对于该窗体的属性，需要用Properties属性访问
'        .Properties("Caption") = "进度"
'        .Properties("Name") = FormName
'        .Properties("Height") = 30
'        .Properties("Width") = BarLength
'        .Properties("BackColor") = RGB(240, 240, 240)
'        .Properties("SpecialEffect") = fmSpecialEffectFlat
'        .Properties("BorderStyle") = fmBorderStyleNone    '要在该窗体上创建控件，则需要访问.Designer设计器对象
'    End With
'
'    '加载并显示该窗体。注意与平时加载显示窗体的不同
'    Set uForm = VBA.UserForms.Add(FormName)
'
'    With uForm                                            '用于显示进度
    With frmProgressBar
        Set lbl1 = .Controls.Add("Forms.Label.1", "Label1", True)
        With lbl1
            .Left = 0
            .Top = 0
            .Height = frmProgressBar.Width  'uForm
            .Width = 0
            .Caption = ""
            .BackColor = RGB(128, 128, 255)
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleOpaque
            .BorderColor = .BackColor
            .ZOrder 1
        End With
    
        '用于显示文字
        Set lbl2 = .Controls.Add("Forms.Label.1", "Label1", True)
        With lbl2
            .Left = 0
            .Top = 9
            .Height = 12
            .Width = BarLength
            .Caption = ""
            .TextAlign = fmTextAlignLeft
            .Font.size = 10
            .Font.Bold = False
            .Font.Italic = False
            .Font.Name = "宋体"
            .ForeColor = RGB(255, 255, 255)
            .BorderStyle = fmBorderStyleNone
            .BackStyle = fmBackStyleTransparent
            .ZOrder 0
        End With
        RemoveFormCaption frmProgressBar 'uForm
        frmProgressBar.Show vbModeless  'uForm
    End With
End Sub

Private Sub RemoveModule(n As String)                 '移除具有指定名称的模块
    On Error Resume Next
    objApp.VBProject.VBComponents.Remove objApp.VBProject.VBComponents(n)
    objApp.Save
End Sub

Private Sub RemoveFormCaption(form As Object)
    If val(Application.Version) < 9 Then
        hwnd = FindWindow("ThunderXFrame", form.Caption)
    Else
        hwnd = FindWindow("ThunderDFrame", form.Caption)
    End If
    IStyle = GetWindowLong(hwnd, GWL_STYLE)
    IStyle = IStyle And Not WS_CAPTION
    SetWindowLong hwnd, GWL_STYLE, IStyle
    DrawMenuBar hwnd
End Sub

Private Sub Class_Terminate()

End Sub
