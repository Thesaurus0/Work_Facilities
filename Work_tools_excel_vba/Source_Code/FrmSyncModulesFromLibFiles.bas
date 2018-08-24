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
Option Base 1

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

Private Sub btnSelectFile_CommonLibFiles_Click()
    Dim arrFiles
    Dim sDefaultFile As String
    
    If Len(Trim(tbFilePath_CommonLibFiles.Text)) > 0 Then
        sDefaultFile = Split(tbFilePath_CommonLibFiles.Text, vbCrLf)(0)
    End If
    arrFiles = fSelectMultipleFileDialog(sDefaultFile, "VBA source code file=*.bas;*.cls;*.frm", "Common Lib Files")
    
    If ArrLen(arrFiles) > 0 Then
        tbFilePath_CommonLibFiles.Text = Join(arrFiles, vbCrLf)
    
        Call fSetFocus(tbFilePath_CommonLibFiles)
        sNewSelectedFile = arrFiles(LBound(arrFiles))
    End If
    
    Erase arrFiles
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
    Dim sFolder As String
    
    sFolder = fSelectFolderDialog(tbCommonLibFolder.Text, "Common Lib Folder")
    
    If Len(Trim(sFolder)) > 0 Then
        tbCommonLibFolder.value = sFolder
        
        Dim arrFiles()
        arrFiles = fGetAllFilesUnderFolder(sFolder)
        
        tbFilePath_CommonLibFiles.Text = Join(arrFiles, vbCrLf)
        Erase arrFiles
    End If
    
    Call fSetFocus(tbCommonLibFolder)
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
    
    Call fSetSavedValue(RANGE_TargetMacroToSyncWithCommLib, tbFilePath_TargetMacro.value)
    Call fSetSavedValue(RANGE_CommonLibFolderSelected, tbCommonLibFolder.value)
    Call fSetSavedValue(RANGE_CommonLibFilesSelected, tbFilePath_CommonLibFiles.value)
    
    gsRtnValueOfForm = CONST_SUCCESS
    Unload Me
End Sub


Private Sub Frame2_Click()

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
 

Private Sub tbCommonLibFolder_AfterUpdate()
    Dim arrFiles
    
    If Len(Trim(tbCommonLibFolder.Text)) > 0 Then
        If Not fFolderExists(Trim(tbCommonLibFolder.Text)) Then
            fMsgBox "The folder you specified does not exist."
            Call fSetFocus(tbCommonLibFolder)
            Exit Sub
        End If
        
        arrFiles = fGetAllFilesUnderFolder(Trim(tbCommonLibFolder.Text))
        
        tbFilePath_CommonLibFiles = Join(arrFiles, vbCrLf)
    Else
        tbFilePath_CommonLibFiles.Text = ""
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
    
    tbCommonLibFolder.value = fGetSavedValue(RANGE_CommonLibFolderSelected)
    tbFilePath_CommonLibFiles.value = fGetSavedValue(RANGE_CommonLibFilesSelected)
    
    obByFiles.value = True
    
    If fGetSavedValue(RANGE_SyncWithCommLibWhichFunction) = "SYNC_WITH_COMMON_LIB" Then
        Call fDisableUserFormControl(tbFilePath_TargetMacro)
        Call fDisableUserFormControl(btnSelectFile_TargetMacro)
        Call fDisableUserFormControl(btnIterateWbs_Left)
        Call fDisableUserFormControl(tbCommonLibFolder)
        Call fDisableUserFormControl(btnSelectFile_CommonLibFolder)
        Call fDisableUserFormControl(tbFilePath_CommonLibFiles)
        Call fDisableUserFormControl(btnSelectFile_CommonLibFiles)

        Call fDisableUserFormControl(obByFolder)
        Call fDisableUserFormControl(obByFiles)
        
        tbFilePath_TargetMacro.value = fGetSavedValue(RANGE_TargetMacroToSyncWithCommLib)
    Else 'COMPARE_WITH_COMMON_LIB
        Call fEnableUserFormControl(tbFilePath_TargetMacro)
        Call fEnableUserFormControl(btnSelectFile_TargetMacro)
        Call fEnableUserFormControl(btnIterateWbs_Left)
        Call fEnableUserFormControl(tbFilePath_CommonLibFiles)
        Call fEnableUserFormControl(btnSelectFile_CommonLibFiles)
        
        Call fDisableUserFormControl(tbCommonLibFolder)
        Call fDisableUserFormControl(btnSelectFile_CommonLibFolder)
        
        Call fSetFocus(tbFilePath_TargetMacro)
        
        tbFilePath_TargetMacro.value = ActiveWorkbook.FullName
    End If
    
    sNewSelectedFile = tbFilePath_TargetMacro.value
    iWbSelected = 1
End Sub

Function fValidateUserInput() As Boolean
    fValidateUserInput = False
    
    If Not fFilesUserInputCheck(tbFilePath_TargetMacro, "Target Macro") Then Exit Function
    
    Dim arrFiles
    If obByFolder.value Then
        If Not fFolderExists(tbCommonLibFolder.Text) Then
            fMsgBox "The common lib folder you specified does not exist, please check."
            Call fSetFocus(tbCommonLibFolder)
            Exit Function
        End If
        
        arrFiles = fGetAllFilesUnderFolder(tbCommonLibFolder.Text)
        tbFilePath_CommonLibFiles.Text = Join(arrFiles, vbCrLf)
        
        Erase arrFiles
    End If
    
    If obByFiles.value Then
        Dim i As Long
        Dim sFile As String
        
        arrFiles = Split(tbFilePath_CommonLibFiles.Text, vbCrLf)
        
        For i = LBound(arrFiles) To UBound(arrFiles)
            sFile = arrFiles(i)
            
            If Not fFileExists(sFile) Then
                fMsgBox "File below does not exists " & vbCr & sFile
                Call fSetFocus(tbFilePath_CommonLibFiles)
                Exit Function
            End If
        Next
        
        Erase arrFiles
    End If
'
'    If UCase(Trim(tbFilePath_TargetMacro.value)) = UCase(Trim(tbFilePath_Right.value)) Then
'        fMsgBox "The two files are same, please check"
'        Call fSetFocus(tbFilePath_TargetMacro)
'        Exit Function
'    End If
    
    fValidateUserInput = True
End Function
