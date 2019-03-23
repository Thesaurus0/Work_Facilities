VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCompareTwoMacroFiles 
   Caption         =   "Compare 2 Macro Files"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11850
   OleObjectBlob   =   "FrmCompareTwoMacroFiles.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FrmCompareTwoMacroFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private sNewSelectedFile As String
Private iWbSelected As Integer

Private Sub btnIterateWbs_Left_Click()
    Dim sFile As String
    
    If Workbooks.Count <= 0 Then Exit Sub
    
    If iWbSelected > Workbooks.Count Then iWbSelected = 1
    
    sFile = Workbooks(iWbSelected).FullName
    iWbSelected = iWbSelected + 1
    tbFilePath_Left.value = sFile
    Call fSetFocus(tbFilePath_Left)
    sNewSelectedFile = sFile
End Sub

Private Sub btnIterateWbs_Right_Click()
    Dim sFile As String
    
    If iWbSelected > Workbooks.Count Then iWbSelected = 1
    
    sFile = Workbooks(iWbSelected).FullName
    iWbSelected = iWbSelected + 1
    
    tbFilePath_Right.value = sFile
    Call fSetFocus(tbFilePath_Right)
    sNewSelectedFile = sFile
End Sub

Private Sub btnSelectFile_Left_Click()
    Dim sFile As String
    
    sFile = fSelectFileDialog(sNewSelectedFile, "Excel File=*.xlsx;*.xls;*.xls*", "Left Macro")
    
    If Len(Trim(sFile)) > 0 Then tbFilePath_Left.value = sFile
    Call fSetFocus(tbFilePath_Left)
    sNewSelectedFile = sFile
End Sub

Private Sub btnSelectFile_Right_Click()
    Dim sFile As String
    
    sFile = fSelectFileDialog(sNewSelectedFile, "Excel File=*.xlsx;*.xls;*.xls*", "right Macro")
    
    If Len(Trim(sFile)) > 0 Then tbFilePath_Right.value = sFile
    Call fSetFocus(tbFilePath_Right)
    sNewSelectedFile = sFile
End Sub

Private Sub btnSwap_Click()
    Dim sTmp As String
    
    sTmp = tbFilePath_Right.value
    tbFilePath_Right.value = tbFilePath_Left.value
    tbFilePath_Left.value = sTmp
End Sub

Private Sub cbCancel_Click()
    gsRtnValueOfForm = CONST_CANCEL
    Unload Me
End Sub

Private Sub cbOK_Click()
    If Not fValidateUserInput() Then Exit Sub
    
    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).value = tbFilePath_Left.value
    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).value = tbFilePath_Right.value
    
    gsRtnValueOfForm = CONST_SUCCESS
    Unload Me
End Sub

Private Sub cbReset_Click()
'    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).Value = ""
'    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).Value = ""
'    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyOpened).Value = ""
'    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyOpened).Value = ""
    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyExported).value = ""
    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyExported).value = ""
End Sub

Private Sub UserForm_Initialize()
    gsRtnValueOfForm = ""
    
    tbFilePath_Left.value = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).value)
    tbFilePath_Right.value = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).value)
    
    Call fSetFocus(tbFilePath_Left)
    
    sNewSelectedFile = tbFilePath_Left.value
    iWbSelected = 1
End Sub

Function fValidateUserInput() As Boolean
    fValidateUserInput = False
    
    If Not fFilesUserInputCheck(tbFilePath_Left, "Macro On Left") Then Exit Function
    If Not fFilesUserInputCheck(tbFilePath_Right, "Macro On Right") Then Exit Function
    
    If UCase(Trim(tbFilePath_Left.value)) = UCase(Trim(tbFilePath_Right.value)) Then
        fMsgBox "The two files are same, please check"
        Call fSetFocus(tbFilePath_Left)
        Exit Function
    End If
    
    fValidateUserInput = True
End Function
