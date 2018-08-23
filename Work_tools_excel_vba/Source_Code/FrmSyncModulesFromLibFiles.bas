VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmSyncModulesFromLibFiles 
   Caption         =   "Sync with Common Lib"
   ClientHeight    =   9150.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15915
   OleObjectBlob   =   "FrmSyncModulesFromLibFiles.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FrmSyncModulesFromLibFiles"
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
    tbFilePath_TargetMacro.value = sFile
    Call fSetFocus(tbFilePath_TargetMacro)
    sNewSelectedFile = sFile
End Sub

'Private Sub btnIterateWbs_Right_Click()
'    Dim sFile As String
'
'    If iWbSelected > Workbooks.Count Then iWbSelected = 1
'
'    sFile = Workbooks(iWbSelected).FullName
'    iWbSelected = iWbSelected + 1
'
'    tbFilePath_Right.value = sFile
'    Call fSetFocus(tbFilePath_Right)
'    sNewSelectedFile = sFile
'End Sub

Private Sub btnSelectFile_TargetMacro_Click()
    Dim sFile As String
    
    sFile = fSelectFileDialog(sNewSelectedFile, "Excel File=*.xlsm;*.xls", "Target Macro")
    
    If Len(Trim(sFile)) > 0 Then tbFilePath_TargetMacro.value = sFile
    Call fSetFocus(tbFilePath_TargetMacro)
    sNewSelectedFile = sFile
End Sub

Private Sub btnSelectFile_CommonLibFolder_Click()
    Dim sFile As String
    
    sFile = fSelectFolderDialog(sNewSelectedFile, "Excel File=*.xlsx;*.xls;*.xls*", "right Macro")
    
    If Len(Trim(sFile)) > 0 Then tbFilePath_Right.value = sFile
    Call fSetFocus(tbFilePath_Right)
    sNewSelectedFile = sFile
End Sub



'Private Sub btnSwap_Click()
'    Dim sTmp As String
'
'    sTmp = tbFilePath_Right.value
'    tbFilePath_Right.value = tbFilePath_TargetMacro.value
'    tbFilePath_TargetMacro.value = sTmp
'End Sub

Private Sub cbCancel_Click()
    gsRtnValueOfForm = CONST_CANCEL
    Unload Me
End Sub

Private Sub cbOK_Click()
    If Not fValidateUserInput() Then Exit Sub
    
    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).value = tbFilePath_TargetMacro.value
    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).value = tbFilePath_Right.value
    
    gsRtnValueOfForm = CONST_SUCCESS
    Unload Me
End Sub

Private Sub obByFiles_Click()
    If obByFiles.value Then
        Call fDisableUserFormControl(btnSelectFile_CommonLibFolder)
        Call fDisableUserFormControl(tbCommonLibFolder)
        
        Call fEnableUserFormControl(tbFilePath_CommonLibFiles)
        Call fEnableUserFormControl(btnSelectFile_CommonLibFiles)
    End If
End Sub

Private Sub obByFolder_Click()
    If obByFolder.value Then
        Call fEnableUserFormControl(btnSelectFile_CommonLibFolder)
        Call fEnableUserFormControl(tbCommonLibFolder)
        
        Call fDisableUserFormControl(tbFilePath_CommonLibFiles)
        Call fDisableUserFormControl(btnSelectFile_CommonLibFiles)
    End If
End Sub
 

'Private Sub cbReset_Click()
''    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).Value = ""
''    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).Value = ""
''    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyOpened).Value = ""
''    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyOpened).Value = ""
'    ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyExported).value = ""
'    ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyExported).value = ""
'End Sub

Private Sub UserForm_Initialize()
    gsRtnValueOfForm = ""
    
    tbFilePath_TargetMacro.value = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_TargetMacroToSyncWithCommLib).value)
    tbFilePath_CommonLibFiles.value = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_CommonLibFilesSelected).value)
    
    obByFiles.value = True
    Call fDisableUserFormControl(tbCommonLibFolder)
    Call fDisableUserFormControl(btnSelectFile_CommonLibFolder)
    
    Call fSetFocus(tbFilePath_TargetMacro)
    sNewSelectedFile = tbFilePath_TargetMacro.value
    iWbSelected = 1
End Sub

Function fValidateUserInput() As Boolean
    fValidateUserInput = False
    
    If Not fFilesUserInputCheck(tbFilePath_TargetMacro, "Macro On Left") Then Exit Function
    If Not fFilesUserInputCheck(tbFilePath_Right, "Macro On Right") Then Exit Function
    
    If UCase(Trim(tbFilePath_TargetMacro.value)) = UCase(Trim(tbFilePath_Right.value)) Then
        fMsgBox "The two files are same, please check"
        Call fSetFocus(tbFilePath_TargetMacro)
        Exit Function
    End If
    
    fValidateUserInput = True
End Function
