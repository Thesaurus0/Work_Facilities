VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCloneMacro 
   Caption         =   "Clone Macro"
   ClientHeight    =   4995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14430
   OleObjectBlob   =   "FrmCloneMacro.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "FrmCloneMacro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sNewSelectedFile As String
Private iWbSelected As Integer

Private Sub btnCopySourceFullName_Click()
    If Len(Trim(tbFilePath_Left.value)) > 0 Then
        tbFilePath_Right.value = tbFilePath_Left.value
    End If
End Sub

Private Sub btnCopySourceName_Click()
    If Len(Trim(tbFilePath_Left.value)) > 0 Then
        tbFileNameUnderSameFolder.value = fGetFileNetName(tbFilePath_Left.value, False)
    End If
End Sub

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
    
    sFile = fSelectFileDialog(sNewSelectedFile, "Excel File=*.xlsx;*.xls;*.xls*", "Sourece Macro")
    
    If Len(Trim(sFile)) > 0 Then tbFilePath_Left.value = sFile
    Call fSetFocus(tbFilePath_Left)
    sNewSelectedFile = sFile
End Sub

Private Sub btnSelectFile_Right_Click()
    Dim sFile As String
    
    sFile = fSelectSaveAsFileDialog(sNewSelectedFile, "Excel File(*.xlsm),*.xlsm", "New Macro")
    
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
    
    Call fSetValue(RANGE_CloneMacro_Source, tbFilePath_Left.value)
    
    If obSourceFolder.value Then
        Call fSetValue(RANGE_CloneMacro_Target, fGetFileParentFolder(tbFilePath_Left.value) & Trim(tbFileNameUnderSameFolder.value) & ".xlsm")
    Else
        Call fSetValue(RANGE_CloneMacro_Target, Trim(tbFilePath_Right.value))
    End If
    
    gsRtnValueOfForm = CONST_SUCCESS
    Unload Me
End Sub


Private Sub obManuallySpecify_Click()
    Call fEnableUserFormControl(tbFilePath_Right)
    Call fEnableUserFormControl(btnSelectFile_Right)
    Call fEnableUserFormControl(btnIterateWbs_Right)
    Call fEnableUserFormControl(btnCopySourceFullName)
     
    Call fDisableUserFormControl(tbFileNameUnderSameFolder)
    Call fDisableUserFormControl(btnCopySourceName)
End Sub

Private Sub obSourceFolder_Click()
    
    
    Call fEnableUserFormControl(tbFileNameUnderSameFolder)
    Call fEnableUserFormControl(btnCopySourceName)
     
    Call fDisableUserFormControl(tbFilePath_Right)
    Call fDisableUserFormControl(btnSelectFile_Right)
    Call fDisableUserFormControl(btnIterateWbs_Right)
    Call fDisableUserFormControl(btnCopySourceFullName)
End Sub
  
Private Sub UserForm_Initialize()
    gsRtnValueOfForm = ""
    
    tbFilePath_Left.value = fGetValue(RANGE_CloneMacro_Source)
    tbFilePath_Right.value = fGetValue(RANGE_CloneMacro_Target)
    tbFileNameUnderSameFolder.value = fGetFileNetName(tbFilePath_Right.value)
    obSourceFolder.value = True
    Call fDisableUserFormControl(btnSelectFile_Left)
    
    Call fSetFocus(tbFilePath_Left)
    
    sNewSelectedFile = tbFilePath_Left.value
    iWbSelected = 1
End Sub

Function fValidateUserInput() As Boolean
    Dim sFile As String
        
    fValidateUserInput = False
    
    If Not fFilesUserInputCheck(tbFilePath_Left, "Macro On Left") Then Exit Function
    'If Not fFilesUserInputCheck(tbFilePath_Right, "Macro On Right") Then Exit Function
    
    If obSourceFolder.value Then
        sFile = Trim(tbFileNameUnderSameFolder.value)
        
        If Len(sFile) <= 0 Then
            fMsgBox "please input a file name."
            Call fSetFocus(tbFileNameUnderSameFolder)
            Exit Function
        End If
        
        If UCase(Right(sFile, 5)) = ".XLSM" Then
            sFile = Left(sFile, Len(sFile) - 5)
            tbFileNameUnderSameFolder = sFile
        End If
    Else
        sFile = Trim(tbFilePath_Right.value)
        If Len(sFile) <= 0 Then
            fMsgBox "please input the new macro file name"
            Call fSetFocus(tbFilePath_Right)
            Exit Function
        End If
        
        fGetFSO
        If Not gFSO.FolderExists(fGetFileParentFolder(sFile)) Then
            fMsgBox "the parent folder of the new macro does not exist."
            Call fSetFocus(tbFilePath_Right)
            Exit Function
        End If
    End If
    
    If UCase(Trim(tbFilePath_Left.value)) = UCase(Trim(tbFilePath_Right.value)) Then
        fMsgBox "The two files are same, please check"
        Call fSetFocus(tbFilePath_Left)
        Exit Function
    End If
    
    fValidateUserInput = True
End Function
