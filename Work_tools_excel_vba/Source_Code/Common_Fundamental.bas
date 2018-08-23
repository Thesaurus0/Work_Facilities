Attribute VB_Name = "Common_Fundamental"
Option Explicit
Option Base 1

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String _
    , ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String _
    , ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Function fOpenFile(asFileFullPath As String)
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
        
        fErr msg
    End If
End Function

Public Function fGetValidMaxRow(shtParam As Worksheet, Optional abCountInMergedCell As Boolean = False) As Long
    Dim lExcelMaxRow As Long
    Dim lUsedMaxRow As Long
    Dim lUsedMaxCol As Long
    
    lExcelMaxRow = shtParam.Rows.Count
    lUsedMaxRow = shtParam.UsedRange.Row + shtParam.UsedRange.Rows.Count - 1
    lUsedMaxCol = shtParam.UsedRange.Column + shtParam.UsedRange.Columns.Count - 1
    
    If lUsedMaxRow = 1 Then
        If shtParam.UsedRange.Address = "$A$1" And Len(shtParam.Range("A1")) <= 0 Then
            fGetValidMaxRow = 0
            Exit Function
        End If
    End If
    
    Dim lEachCol As Long
    Dim lValidMaxRowSaved As Long
    Dim lEachValidMaxRow As Long
    
    lValidMaxRowSaved = 0
    
    For lEachCol = 1 To lUsedMaxCol
        lEachValidMaxRow = shtParam.Cells(lExcelMaxRow, lEachCol).End(xlUp).Row
        
        If lEachValidMaxRow >= lUsedMaxRow Then
            fGetValidMaxRow = lEachValidMaxRow
            Exit Function
        End If
        
        If abCountInMergedCell Then
            If shtParam.Cells(lEachValidMaxRow, lEachCol).MergeCells Then
                lEachValidMaxRow = shtParam.Cells(lEachValidMaxRow, lEachCol).MergeArea.Row _
                                 + shtParam.Cells(lEachValidMaxRow, lEachCol).MergeArea.Rows.Count - 1
            End If
        End If
        
        If lEachValidMaxRow > lValidMaxRowSaved Then lValidMaxRowSaved = lEachValidMaxRow
    Next
    
    If lUsedMaxCol = 1 And lValidMaxRowSaved = 1 Then
        If Len(shtParam.Range("A1")) <= 0 Then
            fGetValidMaxRow = 0
            Exit Function
        End If
    End If
    
    fGetValidMaxRow = lValidMaxRowSaved
End Function

Public Function fGetValidMaxCol(shtParam As Worksheet, Optional abCountInMergedCell As Boolean = False) As Long
    Dim lExcelMaxCol As Long
    Dim lUsedMaxRow As Long
    Dim lUsedMaxCol As Long
    
    lExcelMaxCol = shtParam.Columns.Count
    lUsedMaxRow = shtParam.UsedRange.Row + shtParam.UsedRange.Rows.Count - 1
    lUsedMaxCol = shtParam.UsedRange.Column + shtParam.UsedRange.Columns.Count - 1
    
    If lUsedMaxRow = 1 Then
        If shtParam.UsedRange.Address = "$A$1" And Len(shtParam.Range("A1")) <= 0 Then
            fGetValidMaxCol = 0
            Exit Function
        End If
    End If
    
    Dim lEachRow As Long
    Dim lValidMaxColSaved As Long
    Dim lEachValidMaxCol As Long
    
    lValidMaxColSaved = 0
    
    For lEachRow = 1 To lUsedMaxRow
        lEachValidMaxCol = shtParam.Cells(lEachRow, lExcelMaxCol).End(xlToLeft).Column
        
        If lEachValidMaxCol >= lUsedMaxCol Then
            fGetValidMaxCol = lEachValidMaxCol
            Exit Function
        End If
        
        If abCountInMergedCell Then
            If shtParam.Cells(lEachRow, lEachValidMaxCol).MergeCells Then
                lEachValidMaxCol = shtParam.Cells(lEachRow, lEachValidMaxCol).MergeArea.Column _
                                 + shtParam.Cells(lEachRow, lEachValidMaxCol).MergeArea.Columns.Count - 1
            End If
        End If
        
        If lEachValidMaxCol > lValidMaxColSaved Then lValidMaxColSaved = lEachValidMaxCol
    Next
    
    If lUsedMaxRow = 1 And lValidMaxColSaved = 1 Then
        If Len(shtParam.Range("A1")) <= 0 Then
            fGetValidMaxCol = 0
            Exit Function
        End If
    End If
    
    fGetValidMaxCol = lValidMaxColSaved
End Function

Function fGetValidMaxRowOfRange(rngParam As Range, Optional abCountInMergedCell As Boolean = False) As Long
     Dim lOut As Long
     
     'single cell
     If fRangeIsSingleCell(rngParam) Then lOut = rngParam.Row:               GoTo exit_fun
     
     Dim shtParent As Worksheet
     Set shtParent = rngParam.Parent
     
     Dim lExcelMaxRow As Long
     Dim lExcelMaxCol As Long
     Dim lShtValidMaxRow As Long
     Dim lShtValidMaxCol As Long
     Dim lRangeMaxRow As Long
     Dim lRangeMaxCol As Long
     Dim lValidMaxRowSaved As Long
     Dim lEachValidMaxRow As Long
     Dim lEachCol As Long
     
     lExcelMaxRow = shtParent.Rows.Count
     lExcelMaxCol = shtParent.Columns.Count
     lRangeMaxRow = rngParam.Row + rngParam.Rows.Count - 1
     lRangeMaxCol = rngParam.Column + rngParam.Columns.Count - 1
     
     lShtValidMaxRow = fGetValidMaxRow(shtParent, abCountInMergedCell)
     If lShtValidMaxRow < rngParam.Row Then 'blank, out of usedrange
        lOut = rngParam.Row: GoTo exit_fun
     End If
     
     lShtValidMaxCol = fGetValidMaxCol(shtParent, abCountInMergedCell)
     If lShtValidMaxCol < rngParam.Column Then 'blank, out of usedrange
        lOut = rngParam.Row: GoTo exit_fun
     End If
     
     'whole sheet
     If rngParam.Rows.Count = lExcelMaxRow And rngParam.Columns.Count = lExcelMaxCol Then
        lOut = lShtValidMaxRow: GoTo exit_fun
     End If
     
     If lRangeMaxRow > lShtValidMaxRow Then 'shrink row
        lRangeMaxRow = lShtValidMaxRow
     End If
     If lRangeMaxCol > lShtValidMaxCol Then 'shrink col
        lRangeMaxCol = lShtValidMaxCol
     End If
     
'     'several rows
'     If rngParam.Columns.Count = lExcelMaxCol Then
'        lOut = lRangeMaxRow: GoTo exit_fun
'     End If
     
     'several columns
     If rngParam.Rows.Count = lExcelMaxRow Then
        lValidMaxRowSaved = 0
        
        For lEachCol = rngParam.Column To lRangeMaxCol
            lEachValidMaxRow = shtParent.Cells(lExcelMaxRow, lEachCol).End(xlUp).Row
            
            If lEachValidMaxRow >= lShtValidMaxRow Then
                lOut = lShtValidMaxRow
                GoTo exit_fun
            End If
            
            If abCountInMergedCell Then
                If shtParent.Cells(lEachValidMaxRow, lEachCol).MergeCells Then
                    lEachValidMaxRow = shtParent.Cells(lEachValidMaxRow, lEachCol).MergeArea.Row _
                                     + shtParent.Cells(lEachValidMaxRow, lEachCol).MergeArea.Rows.Count - 1
                End If
            End If
            
            If lEachValidMaxRow > lValidMaxRowSaved Then lValidMaxRowSaved = lEachValidMaxRow
        Next
        
        lOut = lValidMaxRowSaved: GoTo exit_fun
    End If
    
    Dim arrShrunk()
    Dim lArrMaxRow As Long
    Dim lArrMaxCol As Long
    arrShrunk = fReadRangeDatatoArrayByStartEndPos(shtParent, rngParam.Row, rngParam.Column, lRangeMaxRow, lRangeMaxCol)
    lArrMaxRow = fGetArrayMaxValidRowCol(arrShrunk, lArrMaxCol)
    Erase arrShrunk
    
    lArrMaxRow = rngParam.Row + lArrMaxRow + IIf(lArrMaxRow > 0, -1, 0)
    lArrMaxCol = rngParam.Column + lArrMaxCol + IIf(lArrMaxCol > 0, -1, 0)
    
    lOut = lArrMaxRow
    lEachValidMaxRow = 0
    If abCountInMergedCell Then
        If shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeCells Then
        '    If shtParent.Cells(lOut, lEachCol).MergeArea.Rows.Count > 1 Then
                lEachValidMaxRow = shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeArea.Row _
                                 + shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeArea.Rows.Count - 1
         '   End If
        End If
        
        If lEachValidMaxRow > lOut Then lOut = lEachValidMaxRow
    End If

exit_fun:
    fGetValidMaxRowOfRange = lOut
End Function

Function fGetArrayMaxValidRowCol(arrParam(), Optional lMaxCol As Long, Optional bReverse As Boolean = True) As Long
    Dim lEachRow As Long
    Dim lEachMaxRow As Long
    Dim lMaxRowSaved As Long
    Dim lEachCol As Long

    lMaxCol = 0
    lMaxRowSaved = 0 'UBound(arrParam, 1) - LBound(arrParam, 1) + 1
    
    If bReverse Then
        For lEachRow = UBound(arrParam, 1) To LBound(arrParam, 1) Step -1
            For lEachCol = LBound(arrParam, 2) To UBound(arrParam, 2)
                If Len(Trim(CStr(arrParam(lEachRow, lEachCol)))) > 0 Then
                    If lEachRow > lMaxRowSaved Then
                        lMaxRowSaved = lEachRow
                        lMaxCol = lEachCol
                        GoTo exit_fun
                    End If
                End If
            Next
        Next
    Else
        For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
            For lEachCol = LBound(arrParam, 2) To UBound(arrParam, 2)
                If Len(Trim(CStr(arrParam(lEachRow, lEachCol)))) > 0 Then
                    If lEachRow > lMaxRowSaved Then
                        lMaxRowSaved = lEachRow
                        lMaxCol = lEachCol
                        GoTo exit_fun
                    End If
                End If
            Next
        Next
    End If
    
exit_fun:
    fGetArrayMaxValidRowCol = lMaxRowSaved
End Function

Function fRangeIsSingleCell(rngParam As Range) As Boolean
    fRangeIsSingleCell = (rngParam.Rows.Count = 1 And rngParam.Columns.Count = 1)
End Function

Function fErr(Optional sMsg As String = "") As VbMsgBoxResult
    gErrNum = vbObjectError + CONFIG_ERROR_NUMBER
    'gbBusinessError = True
    gErrMsg = sMsg
    'If fNzero(sMsg) Then fMsgBox "Error: " & vbCr & vbCr & sMsg, vbCritical
    If fNzero(sMsg) Then fMsgBox sMsg, vbCritical
    
    Err.Raise gErrNum, "", "Program is to be terminated."
End Function

Function fErrBuzz(Optional sMsg As String = "") As VbMsgBoxResult
    gErrNum = vbObjectError + BUSINESS_ERROR_NUMBER
    'gbBusinessError = True
    gErrMsg = sMsg
    If fNzero(sMsg) Then fMsgBox "Error: " & vbCr & vbCr & sMsg, vbCritical
    
    Err.Raise gErrNum, "", "Program is to be terminated."
End Function
Function fMsgBox(Optional sMsg As String = "", Optional aVbMsgBoxStyle As VbMsgBoxStyle = vbCritical) As VbMsgBoxResult
    fMsgBox = MsgBox(sMsg, aVbMsgBoxStyle)
End Function

Function fSelectFileDialog(Optional asDefaultFilePath As String = "" _
                         , Optional asFileFilters As String = "", Optional asTitle As String = "") As String
    'asFileFilters :   "Excel File=*.xlsx;*.xls;*.xls*"
    'asFileFilters :   "Excel File(*.xlsx),*.xlsx, "Text File(*.txt),*.txt, Visual Basic Files(*.bas;*.txt),*.bas;*.txt "
    Dim fd As FileDialog
    Dim sFilterDesc As String
    Dim sFilterStr As String
    Dim sDefaultFile As String
    
    If Len(Trim(asFileFilters)) > 0 Then
        sFilterDesc = Trim(Split(asFileFilters, "=")(0))
        sFilterStr = Trim(Split(asFileFilters, "=")(1))
    End If
    
    If Len(Trim(asDefaultFilePath)) > 0 Then
       ' sDefaultFile = fGetFileParentFolder(asDefaultFilePath)
        sDefaultFile = asDefaultFilePath
    Else
        sDefaultFile = ThisWorkbook.Path
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.InitialFileName = sDefaultFile
    fd.Title = IIf(Len(asTitle) > 0, asTitle, fd.InitialFileName)
    fd.AllowMultiSelect = False
    
    If Len(Trim(sFilterStr)) > 0 Then
        fd.Filters.Clear
        fd.Filters.Add sFilterDesc, sFilterStr, 1
        fd.FilterIndex = 1
        fd.InitialView = msoFileDialogViewDetails
    Else
        If fd.Filters.Count > 0 Then fd.Filters.Delete
    End If

    If fd.Show = -1 Then
        fSelectFileDialog = fd.SelectedItems(1)
    Else
        fSelectFileDialog = ""
    End If
        
    Set fd = Nothing
End Function

Function fSelectMultipleFileDialog(Optional asDefaultFilePath As String = "" _
                         , Optional asFileFilters As String = "", Optional asTitle As String = "")
    'asFileFilters :   "Excel File=*.xlsx;*.xls;*.xls*"
    'asFileFilters :   "Excel File(*.xlsx),*.xlsx, "Text File(*.txt),*.txt, Visual Basic Files(*.bas;*.txt),*.bas;*.txt "
    Dim fd As FileDialog
    Dim sFilterDesc As String
    Dim sFilterStr As String
    Dim sDefaultFile As String
    Dim arrOut()
    
    arrOut = Array()
    
    If Len(Trim(asFileFilters)) > 0 Then
        sFilterDesc = Trim(Split(asFileFilters, "=")(0))
        sFilterStr = Trim(Split(asFileFilters, "=")(1))
    End If
    
    If Len(Trim(asDefaultFilePath)) > 0 Then
       ' sDefaultFile = fGetFileParentFolder(asDefaultFilePath)
        sDefaultFile = asDefaultFilePath
    Else
        sDefaultFile = IIf(Len(ActiveWorkbook.Path) > 0, ActiveWorkbook.Path, ThisWorkbook.Path)
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    fd.InitialFileName = sDefaultFile
    fd.Title = IIf(Len(asTitle) > 0, asTitle, fd.InitialFileName)
    fd.AllowMultiSelect = True
    
    If Len(Trim(sFilterStr)) > 0 Then
        fd.Filters.Clear
        fd.Filters.Add sFilterDesc, sFilterStr, 1
        fd.FilterIndex = 1
        fd.InitialView = msoFileDialogViewDetails
    Else
        If fd.Filters.Count > 0 Then fd.Filters.Delete
    End If

    If fd.Show = -1 Then
        Dim i As Integer
        ReDim arrOut(1 To fd.SelectedItems.Count)
        
        For i = 1 To fd.SelectedItems.Count
            arrOut(i) = fd.SelectedItems(i)
        Next
    End If
        
    Set fd = Nothing
        
    fSelectMultipleFileDialog = arrOut
    Erase arrOut
End Function
Function fSelectFolderDialog(Optional asDefaultFolder As String = "", Optional asTitle As String = "") As String
    Dim fd As FileDialog
    Dim sFilterDesc As String
    Dim sFilterStr As String
    Dim sDefaultFolder As String
    
    If Len(Trim(asDefaultFolder)) > 0 Then
        sDefaultFolder = asDefaultFolder
    Else
        sDefaultFolder = IIf(Len(ActiveWorkbook.Path) > 0, ActiveWorkbook.Path, ThisWorkbook.Path)
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    fd.InitialFileName = sDefaultFolder
    fd.Title = IIf(Len(asTitle) > 0, asTitle, fd.InitialFileName)
    
    If fd.Show = -1 Then
        fSelectFolderDialog = fd.SelectedItems(1) & Application.PathSeparator
    Else
        fSelectFolderDialog = ""
    End If
        
    Set fd = Nothing
End Function

Function fFolderExists(sFolder As String) As Boolean
    fGetFSO
    fFolderExists = gFSO.FolderExists(sFolder)
End Function
Function fSelectSaveAsFileDialog(Optional asDeafaulfFilePath As String = "", Optional asFileFilters = "", Optional asTitle = "") As String
'asFileFilters  :
'       "Excel File(*.xlsx),*.xlsx, Excel Old Ver(*.xls),*.xls"
'return value:
'   blank: 1. user clicked cancel
'          2. the selected file is already open
' Important: the file extension of the default file name must be same as the first file filter extension, otherwise the defaulf file name will not be shownup
'   default_file.xlsx  = "Excel File(*.xlsx),*.xlsx"
    Dim fd As FileDialog
    Dim sDefaultFolder As String
    Dim sDefaultFileName As String
    Dim sOut 'As String
    Dim response As VbMsgBoxResult
    
    If Len(Trim(asDeafaulfFilePath)) > 0 Then
        sDefaultFolder = fGetFileParentFolder(asDeafaulfFilePath)
        sDefaultFileName = fGetFileBaseName(asDeafaulfFilePath)
        
        If Not fFolderExists(sDefaultFolder) Then sDefaultFolder = ThisWorkbook.Path
    Else
        sDefaultFolder = ThisWorkbook.Path
    End If
    
    sDefaultFolder = fCheckPath(sDefaultFolder)
    
    fGetFSO
    ChDrive gFSO.GetDriveName(sDefaultFolder)
    ChDir sDefaultFolder
    
    If Len(Trim(asFileFilters)) > 0 And Len(Trim(sDefaultFileName)) > 0 Then
        Dim sFirstFilterExt As String
        sFirstFilterExt = Trim(Split(asFileFilters, ",")(1))
        
        If UCase(fGetFileExtension(sDefaultFileName)) <> UCase(Trim(Split(sFirstFilterExt, ".")(1))) Then
            MsgBox "Warning: the file extension is not same as the first filter extension, which will cause the default to be blank." _
               & vbCr & "so it is advisable to provide the first filter extension be same as the file extension" _
               & vbCr & vbCr & "File name supplied: " & sDefaultFileName _
               & vbCr & "Filters: " & asFileFilters, vbExclamation
        End If
    End If
    
    Do While True
        sOut = Application.GetSaveAsFilename(InitialFileName:=sDefaultFileName _
                        , filefilter:=asFileFilters _
                        , Title:=IIf(Len(Trim(asTitle)) > 0, asTitle, sDefaultFileName))
        If sOut = False Then
            sOut = ""
            Exit Do
        Else
            If fFileExists(CStr(sOut)) Then
                response = MsgBox("The file you speicified to save already exists, do you want to overwrite it?", vbYesNoCancel + vbCritical + vbDefaultButton2)
                If response = vbYes Then
                    If fExcelFileOpenedToCloseIt(CStr(sOut)) Then
                        sOut = ""
                        Exit Do
                    Else
                        Exit Do
                    End If
                ElseIf response = vbNo Then
                    'to show the dialog again
                Else
                    sOut = ""
                    Exit Do
                End If
            Else
                Exit Do
            End If
        End If
    Loop
    
    fSelectSaveAsFileDialog = CStr(sOut)
End Function
'Sub Test()
'    Dim i As Integer
'    Dim intChoice
'    Dim fd As FileDialog
'    Dim strPath
'
'    Set fd = Application.FileDialog(msoFileDialogSaveAs)
'
'    With fd
'        For i = 1 To .Filters.Count
'            If .Filters(i).Extensions = "*.xlsx" Then Exit For
'        Next
'
'        .FilterIndex = i
'        intChoice = .Show
'
'        If intChoice <> 0 Then
'            strPath = Application.FileDialog(msoFileDialogSaveAs).SelectedItems(1)
'
'            ThisWorkbook.SaveAs Filename:=strPath
'        End If
'    End With
'End Sub
'
'Sub aaa()
'        With Application.FileDialog(msoFileDialogSaveAs)
''        .Filters.Clear
''        .Filters.Add "Excel File", "*.xls"
''        .Filters.Add "All File", "*.*"
'        .Show
'    End With
'End Sub

Function fGetFileParentFolder(asFileFullPath As String) As String
    fGetFSO
    fGetFileParentFolder = fCheckPath(gFSO.GetParentFolderName(asFileFullPath))
End Function

Function fGetFileBaseName(asFileFullPath As String) As String
    fGetFSO
    fGetFileBaseName = gFSO.GetFileName(asFileFullPath)
End Function
Function fGetFileNetName(asFileFullPath As String, Optional KeepPrecedingFolder As Boolean = False) As String
    Dim sOut As String
    
    fGetFSO
    sOut = gFSO.GetBaseName(asFileFullPath)
    
    If KeepPrecedingFolder Then
        sOut = fCheckPath(fGetFileParentFolder(asFileFullPath)) & sOut
    End If
    
    fGetFileNetName = sOut
End Function
Function fGetFileExtension(asFileFullPath As String, Optional bDot As Boolean = False) As String
    fGetFSO
    fGetFileExtension = IIf(bDot, ".", "") & gFSO.GetExtensionName(asFileFullPath)
End Function

Function fArrayHasBlankValue(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = False
    If fArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    Dim lEachCol As Long
    
    iDimension = fGetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 2 Then fErr "2 dimensions is not supported."
    
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If Len(Trim(CStr(arrParam(lEachRow)))) <= 0 Then
            bOut = True
            GoTo exit_function
        End If
    Next
    
exit_function:
    fArrayHasBlankValue = bOut
End Function
Function fArrayHasDuplicateElement(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = False
    If fArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    
    iDimension = fGetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 2 Then fErr "2 dimensions is not supported."
    
    Dim dict As New Dictionary
    
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If dict.Exists(arrParam(lEachRow)) Then
            bOut = True
            GoTo exit_function
        Else
            dict.Add arrParam(lEachRow), 0
        End If
    Next
    
exit_function:
    fArrayHasDuplicateElement = bOut
    Set dict = Nothing
End Function

Function fArrayIsEmptyOrNoData(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = True
    If fArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    Dim lEachCol As Long
    
    iDimension = fGetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 3 Then fErr "3 dimensions is not supported."
    
    If iDimension = 1 Then
        For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
            If Len(Trim(CStr(arrParam(lEachRow)))) > 0 Then
                bOut = False
                GoTo exit_function
            End If
        Next
    ElseIf iDimension = 2 Then
        For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
            For lEachCol = LBound(arrParam, 2) To UBound(arrParam, 2)
                If Len(Trim(CStr(arrParam(lEachRow, lEachCol)))) > 0 Then
                    bOut = False
                    GoTo exit_function
                End If
            Next
        Next
    End If
    
exit_function:
    fArrayIsEmptyOrNoData = bOut
End Function

Function fArrayIsEmpty(ByRef arrParam) As Boolean
    Dim i As Long
    
    fArrayIsEmpty = True
    
    On Error Resume Next
    
    i = UBound(arrParam, 1)
    If Err.Number = 0 Then
        If UBound(arrParam) < LBound(arrParam) Then
            Exit Function
        Else
            fArrayIsEmpty = False
        End If
    Else
        Err.Clear
    End If
End Function
Function fGetArrayDimension(arrParam) As Integer
    Dim i As Integer
    Dim tmp As Long
    
    On Error GoTo error_exit
    
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(arrParam, i)
        
        If tmp < 0 Then
            fGetArrayDimension = -1
            Exit Function
        End If
    Loop
    
error_exit:
    Err.Clear
    fGetArrayDimension = i - 1
End Function

'Function fNum2Letter(ByVal alNum As Long) As String
'    fNum2Letter = Replace(Split(Columns(alNum).Address, ":")(1), "$", "")
'End Function
Function fNum2LetterV1(ByVal alNum As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String
    
    n = alNum
    Do
        c = (n - 1) Mod 26
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    
    fNum2LetterV1 = s
End Function
Function fNum2Letter(ByVal alNum As Long) As String
    Dim lAlpha As Long
    Dim lRemainder As Long
    
    If alNum <= 26 Then
        fNum2Letter = Chr(alNum + 64)
    Else
        lRemainder = alNum Mod 26
        lAlpha = Int(alNum / 26)
        
        If lRemainder = 0 Then
            lRemainder = 26
            lAlpha = lAlpha - 1
        End If
        fNum2Letter = fNum2Letter(lAlpha) & Chr(lRemainder + 64)
    End If
End Function
'Function fLetter2Num(ByVal alLetter As String) As Long
'    fLetter2Num = Columns(alLetter).Column
'End Function
Function fLetter2Num(ByVal alLetter As String) As Long
    Dim i As Integer
    Dim iOut As Long
    Dim s As String
    
    alLetter = UCase(Trim(alLetter))
    
    iOut = 0
    i = 1
    Do While i <= Len(alLetter)
        s = Mid(alLetter, i, 1)
        iOut = iOut + (26 ^ (Len(alLetter) - i)) * (Asc(s) - 64)
        i = i + 1
    Loop
    
    fLetter2Num = iOut
End Function

Function fFileExists(sFilePath As String) As Boolean
    fGetFSO
    fFileExists = gFSO.FileExists(sFilePath)
End Function

Function fDeleteFile(sFilePath As String)
    If fFileExists(sFilePath) Then
        SetAttr sFilePath, vbNormal
        Kill sFilePath
    End If
End Function

Function fArrayRowIsBlankHasNoData(arr, alRow As Long) As Boolean
    Dim bOut As Boolean
    Dim lEachCol As Long
    
    bOut = True
    For lEachCol = LBound(arr, 2) To UBound(arr, 2)
        If Len(Trim(CStr(arr(alRow, lEachCol)))) > 0 Then
            bOut = False
            Exit For
        End If
    Next
    
    fArrayRowIsBlankHasNoData = bOut
End Function

Function fGenRandomUniqueString() As String
    fGenRandomUniqueString = Format(Now(), "yyyymmddhhMMSS") & Rnd()
End Function

Function fSplit(asOrig As String, Optional asSeparators As String = "") As Variant
    If Len(asSeparators) <= 0 Then asSeparators = ":;|, " & vbLf
    
    Dim tDelimiter As String
    tDelimiter = Chr(130)   'a non-printable charactor
    
    Dim sTransFormed As String
    Dim sEachDeli As String
    Dim i As Integer
    
    sTransFormed = asOrig
    For i = 1 To Len(asSeparators)
        sEachDeli = Mid(asSeparators, i, 1)
        sTransFormed = Replace(sTransFormed, sEachDeli, tDelimiter)
    Next
    
    While InStr(sTransFormed, tDelimiter & tDelimiter) > 0
        sTransFormed = Replace(sTransFormed, tDelimiter & tDelimiter, tDelimiter)
    Wend
    
    sTransFormed = fTrim(sTransFormed, tDelimiter)
    
    fSplit = Split(sTransFormed, tDelimiter)
End Function

Function fSplitJoin(asOrig As String, Optional asSeparators As String = "", Optional asNewSep As String = DELIMITER) As String
    If Len(asSeparators) <= 0 Then asSeparators = ":;|, " & vbLf
    
    Dim arr
    arr = fSplit(asOrig, asSeparators)
    fSplitJoin = Join(arr, asNewSep)
    
    Erase arr
End Function

Function fJoin(asOrig As String, Optional asNewSep As String = DELIMITER) As String
    fJoin = fSplitJoin(asOrig, , asNewSep)
End Function

Function fTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Left(sOut, 1)) > 0
        sOut = Right(sOut, Len(sOut) - 1)
    Wend
    
    While InStr(asWhatToTrim, Right(sOut, 1)) > 0
        sOut = Left(sOut, Len(sOut) - 1)
    Wend
    
    fTrim = sOut
End Function

Function fRTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Right(sOut, 1)) > 0
        sOut = Left(sOut, Len(sOut) - 1)
    Wend
    
    fTrim = sOut
End Function

Function fLTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Left(sOut, 1)) > 0
        sOut = Right(sOut, Len(sOut) - 1)
    Wend
    
    fTrim = sOut
End Function

Function fLen(sStr) As Long
    fLen = Len(Trim(sStr))
End Function

Function fZero(sStr) As Boolean
    fZero = (Len(Trim(sStr)) <= 0)
End Function

Function fNzero(sStr) As Boolean
    fNzero = (Len(Trim(sStr)) > 0)
End Function

Function fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData, arrKeyCols, arrItemCols _
                , Optional asKeysDelimiter As String = "" _
                , Optional asItemsDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    
    If fArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If fArrayIsEmptyOrNoData(arrKeyCols) Then fErr "arrKeyCols is empty !"
    If fArrayIsEmptyOrNoData(arrItemCols) Then fErr "arrItemCols is empty !"
    If fArrayHasDuplicateElement(arrKeyCols) Then fErr "arrKeyCols has duplicate element"
    If fArrayHasDuplicateElement(arrItemCols) Then fErr "arrItemCols has duplicate element"
    
    If InStr(asKeysDelimiter, " ") > 0 Then fErr "asKeysDelimiter cannot be space or contains space"
    If InStr(asItemsDelimiter, " ") > 0 Then fErr "asItemsDelimiter cannot be space or contains space"
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    Dim sItemsValue As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = ""
        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
        Next
        
        If fZero(Replace(sKeyStr, asKeysDelimiter, "")) Then
            If Not IgnoreBlankKeys Then fErr "IgnoreBlankKeys is false, but a keystr is blank"
            GoTo next_row
        End If
        
        If Len(asKeysDelimiter) > 0 Then sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(asKeysDelimiter))
        
        If dictOut.Exists(sKeyStr) Then
            If WhenKeyDuplicateThenError Then
                fErr "Duplicate key was found " & vbCr & sKeyStr
            End If
            GoTo next_row
        End If
        
        sItemsValue = ""
        For j = LBound(arrItemCols) To UBound(arrItemCols)
            sItemsValue = sItemsValue & asItemsDelimiter & Trim(CStr(arrData(i, arrItemCols(j))))
        Next
        If Len(asItemsDelimiter) > 0 Then
            If Len(Replace(sItemsValue, asItemsDelimiter, "")) > 0 Then
                sItemsValue = Right(sItemsValue, Len(sItemsValue) - Len(asItemsDelimiter))
            End If
        End If
        dictOut.Add sKeyStr, sItemsValue
next_row:
    Next
    
exit_fun:
    Set fReadArray2DictionaryMultipleKeysWithMultipleColsCombined = dictOut
    Set dictOut = Nothing
End Function

Function fReadArray2DictionaryMultipleKeysWithKeysOnly(arrData, arrKeyCols _
                , Optional asKeysDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    
    If fArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If fArrayIsEmptyOrNoData(arrKeyCols) Then fErr "arrKeyCols is empty !"
    If fArrayHasDuplicateElement(arrKeyCols) Then fErr "arrKeyCols has duplicate element"
    
    If InStr(asKeysDelimiter, " ") > 0 Then fErr "asKeysDelimiter cannot be space or contains space"
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = ""
        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
        Next
        
        If fZero(Replace(sKeyStr, asKeysDelimiter, "")) Then
            If Not IgnoreBlankKeys Then fErr "IgnoreBlankKeys is false, but a keystr is blank"
            GoTo next_row
        End If
        
        If Len(asKeysDelimiter) > 0 Then sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(asKeysDelimiter))
        
        If dictOut.Exists(sKeyStr) Then
            If WhenKeyDuplicateThenError Then
                fErr "Duplicate key was found " & vbCr & sKeyStr
            End If
            GoTo next_row
        End If
        
        dictOut.Add sKeyStr, 1
next_row:
    Next
    
exit_fun:
    Set fReadArray2DictionaryMultipleKeysWithKeysOnly = dictOut
    Set dictOut = Nothing
End Function

Function fReadArray2DictionaryMultipleKeysWithRowNum(arrData, arrKeyCols _
                , Optional asKeysDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    
    If fArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If fArrayIsEmptyOrNoData(arrKeyCols) Then fErr "arrKeyCols is empty !"
    If fArrayHasDuplicateElement(arrKeyCols) Then fErr "arrKeyCols has duplicate element"
    
    If InStr(asKeysDelimiter, " ") > 0 Then fErr "asKeysDelimiter cannot be space or contains space"
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = ""
        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
        Next
        
        If fZero(Replace(sKeyStr, asKeysDelimiter, "")) Then
            If Not IgnoreBlankKeys Then fErr "IgnoreBlankKeys is false, but a keystr is blank"
            GoTo next_row
        End If
        
        If Len(asKeysDelimiter) > 0 Then sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(asKeysDelimiter))
        
        If dictOut.Exists(sKeyStr) Then
            If WhenKeyDuplicateThenError Then
                fErr "Duplicate key was found " & vbCr & sKeyStr
            End If
            GoTo next_row
        End If
        
        dictOut.Add sKeyStr, i
next_row:
    Next
    
exit_fun:
    Set fReadArray2DictionaryMultipleKeysWithRowNum = dictOut
    Set dictOut = Nothing
End Function
Function fReadArray2DictionaryWithMultipleColsCombined(arrData, lKeyCol As Long, arrItemCols _
                , Optional asDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    If lKeyCol < 0 Then fErr "lKeyCol < 0 to fReadArray2DictionaryWithMultipleColsCombined"
    
    Set dictOut = New Dictionary
    
    If fArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If fArrayIsEmptyOrNoData(arrItemCols) Then fErr "arrItemCols is empty."
    If fArrayHasDuplicateElement(arrItemCols) Then fErr "arrItemCols has duplicate element."
    
    Dim i As Long
    Dim sKey
    Dim sValue As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        If fArrayRowIsBlankHasNoData(arrData, i) Then GoTo next_row
        
        sKey = arrData(i, lKeyCol)
        
        If fZero(sKey) Then
            If IgnoreBlankKeys Then GoTo next_row
            fErr "Key column is blank, but program decides not allow blank, pls contact with IT support."
        End If
        
        If dictOut.Exists(sKey) Then
            If WhenKeyDuplicateThenError Then
                'Application.Goto shtAt.Cells(i + 1, lKeyCol)
                fErr "Duplicate Key was found : " & vbCr & "Key: " & sKey
            End If
            GoTo next_row
        End If
        
        sValue = ""
        Dim j As Integer
        For j = LBound(arrItemCols) To UBound(arrItemCols)
            sValue = sValue & asDelimiter & Trim(arrData(i, arrItemCols(j)))
        Next
        
        If Len(asDelimiter) > 0 Then
            If Len(Replace(sValue, asDelimiter, "")) > 0 Then
                sItemsValue = Right(sValue, Len(sValue) - Len(asDelimiter))
            End If
        End If
        
        dictOut.Add sKey, sValue
next_row:
    Next
    
exit_fun:
    Set fReadArray2DictionaryWithMultipleColsCombined = dictOut
    Set dictOut = Nothing
End Function

Function fReadArray2DictionaryWithSingleCol(arrParam, lKeyCol As Long, lItemCol As Long _
                            , Optional IgnoreBlankKeys As Boolean = False _
                            , Optional WhenKeyIsDuplicateError As Boolean = True) As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    If lItemCol <= 0 Then fErr "lItemCol cannot be less than 0 in fReadArray2DictionaryWithSingleCol"
    Set fReadArray2DictionaryWithSingleCol = fReadArray2Dictionary(arrParam, lKeyCol, lItemCol, IgnoreBlankKeys, WhenKeyIsDuplicateError)
End Function
Function fReadArray2DictionaryWithRowNum(arrParam, lKeyCol As Long _
                            , Optional IgnoreBlankKeys As Boolean = False _
                            , Optional WhenKeyIsDuplicateError As Boolean = True) As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    Set fReadArray2DictionaryWithRowNum = fReadArray2Dictionary(arrParam, lKeyCol, -1, IgnoreBlankKeys, WhenKeyIsDuplicateError)
End Function
Function fReadArray2DictionaryOnlyKeys(arrParam, lKeyCol As Long _
                            , Optional IgnoreBlankKeys As Boolean = False _
                            , Optional WhenKeyIsDuplicateError As Boolean = True) As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    Set fReadArray2DictionaryOnlyKeys = fReadArray2Dictionary(arrParam, lKeyCol, 0, IgnoreBlankKeys, WhenKeyIsDuplicateError)
End Function

Private Function fReadArray2Dictionary(arrParam, lKeyCol As Long _
                            , Optional lItemCol As Long = 0 _
                            , Optional IgnoreBlankKeys As Boolean = False _
                            , Optional WhenKeyIsDuplicateError As Boolean = True) As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    If lItemCol < -1 Or lKeyCol <= 0 Then fErr "wrong param"
    
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    If fArrayIsEmptyOrNoData(arrParam) Then GoTo exit_fun
    
    Dim bGetKeyOnly As Boolean
    Dim bGetRowNo As Boolean
    
    bGetKeyOnly = (lItemCol = 0)
    bGetRowNo = (lItemCol = -1)
    
    Dim i As Long
    Dim sKey As String
    Dim sValue As String
    
    For i = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, i) Then GoTo next_row
        
        sKey = Trim(arrParam(i, lKeyCol))
        
        If Len(sKey) <= 0 Then
            If Not IgnoreBlankKeys Then
                fErr "Key col is blank, but you specified IgnoreBlankKeys = false" & vbCr & lKeyCol
            Else
                GoTo next_row
            End If
        End If
        
        If dictOut.Exists(sKey) Then
            If WhenKeyIsDuplicateError Then
                fErr "duplicate key was found:, but you specified WhenKeyIsDuplicateError = false" & vbCr & lKeyCol & vbCr & sKey
            Else
                GoTo next_row
            End If
        End If
        
        If bGetRowNo Then
            dictOut.Add sKey, i
        Else
            If bGetKeyOnly Then
                dictOut.Add sKey, 0
            Else
                dictOut.Add sKey, arrParam(i, lItemCol)
            End If
        End If
next_row:
    Next
    
exit_fun:
    Set fReadArray2Dictionary = dictOut
    Set dictOut = Nothing
End Function

Function fValidateDuplicateInArray(arrParam, arrKeyColsOrSingle _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional ByVal sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        Call fValidateDuplicateInArrayForCombineCols(arrParam:=arrParam, arrKeyCols:=arrKeyColsOrSingle _
                                                    , bAllowBlankIgnore:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    Else
        Call fValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , bAllowBlankIgnore:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

Function fValidateDuplicateInArrayForCombineCols(arrParam, arrKeyCols _
                        , Optional bAllowBlankIgnore As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'MultipleCols: means MultipleCols composed as key
'for MultipleCols that is individually, please refer to function fValidateDuplicateInArrayIndividually
    Const DELI = " " & DELIMITER & " "
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetterStr As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    If Not fZero(sMsgColHeader) Then
        sColLetterStr = sMsgColHeader
    Else
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sColLetterStr = sColLetterStr & " + " & fNum2Letter(lStartCol + lEachCol - 1)
        Next
        sColLetterStr = Right(sColLetterStr, Len(sColLetterStr) - 3)
    End If
    
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO, [" & sColLetterStr & "]"
            
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = ""
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sKeyStr = sKeyStr & DELI & Trim(CStr(arrParam(lEachRow, lEachCol)))
        Next
        
        If fZero(Replace(sKeyStr, DELI, "")) Then
            If Not bAllowBlankIgnore Then
                'sPos = sPos & "[" & lActualRow & ", " & sColLetterStr & "]"
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                fErr "Keys [" & sKeyStr & "] is blank!" & sPos
            End If
            
            GoTo next_row
        End If
        
        sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(DELI))
        
        If dict.Exists(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fShowSheet shtAt
            Application.GoTo shtAt.Cells(lActualRow, arrKeyCols(UBound(arrKeyCols)))
            fErr "Duplicate key was found:" & vbCr & sKeyStr & vbCr & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

Function fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrData, arrKeyCols, lItemCol As Long _
                , Optional asKeysDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    
    If fArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If fArrayIsEmptyOrNoData(arrKeyCols) Then fErr "arrKeyCols is empty !"
    If fArrayHasDuplicateElement(arrKeyCols) Then fErr "arrKeyCols has duplicate element"
    If lItemCol < 0 Then fErr "lItemCol < 0 to fReadArray2DictionaryWithMultipleKeyColsSingleItemCol"
    
    If InStr(asKeysDelimiter, " ") > 0 Then fErr "asKeysDelimiter cannot be space or contains space"
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = ""
        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
        Next
        
        If fZero(Replace(sKeyStr, asKeysDelimiter, "")) Then
            If Not IgnoreBlankKeys Then fErr "IgnoreBlankKeys is false, but a keystr is blank"
            GoTo next_row
        End If
        
        If Len(asKeysDelimiter) > 0 Then sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(asKeysDelimiter))
        
        If dictOut.Exists(sKeyStr) Then
            If WhenKeyDuplicateThenError Then
                fErr "Duplicate key was found " & vbCr & sKeyStr
            End If
            GoTo next_row
        End If
        
        dictOut.Add sKeyStr, arrData(i, lItemCol)
next_row:
    Next
    
exit_fun:
    Set fReadArray2DictionaryWithMultipleKeyColsSingleItemCol = dictOut
    Set dictOut = Nothing
End Function

Function fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrData, arrKeyCols, lItemCol As Long _
                , Optional asKeysDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False) As Dictionary
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    
    If fArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If fArrayIsEmptyOrNoData(arrKeyCols) Then fErr "arrKeyCols is empty !"
    If fArrayHasDuplicateElement(arrKeyCols) Then fErr "arrKeyCols has duplicate element"
    If lItemCol < 0 Then fErr "lItemCol < 0 to fReadArray2DictionaryWithMultipleKeyColsSingleItemCol"
    
    If InStr(asKeysDelimiter, " ") > 0 Then fErr "asKeysDelimiter cannot be space or contains space"
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = ""
        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
        Next
        
        If fZero(Replace(sKeyStr, asKeysDelimiter, "")) Then
            If Not IgnoreBlankKeys Then fErr "IgnoreBlankKeys is false, but a keystr is blank"
            GoTo next_row
        End If
        
        If Len(asKeysDelimiter) > 0 Then sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(asKeysDelimiter))
        
        If Not dictOut.Exists(sKeyStr) Then
            dictOut.Add sKeyStr, arrData(i, lItemCol)
        Else
            dictOut(sKeyStr) = dictOut(sKeyStr) + arrData(i, lItemCol)
        End If
        
next_row:
    Next
    
exit_fun:
    Set fReadArray2DictionaryWithMultipleKeyColsSingleItemColSum = dictOut
    Set dictOut = Nothing
End Function
Function fValidateDuplicateInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional bAllowBlankIgnore As Boolean = False _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then fErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetter As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(fZero(sMsgColHeader), fNum2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"
         
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = Trim(CStr(arrParam(lEachRow, lKeyCol)))
        
        If fZero(sKeyStr) Then
            If Not bAllowBlankIgnore Then
                'sPos = sPos & lActualRow & " / " & sColLetter
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                fShowSheet shtAt
                Application.GoTo shtAt.Cells(lActualRow, lKeyCol)
                fErr "Keys [" & sColLetter & "] is blank!" & sPos
            End If
            
            GoTo next_row
        End If
        
        If dict.Exists(sKeyStr) Then
            'sPos = sPos & lActualRow & " / " & sColLetter
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fShowSheet shtAt
            Application.GoTo shtAt.Cells(lActualRow, lKeyCol)
            fErr "Duplicate key [" & sKeyStr & "] was found " & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

Function fValidateDuplicateInArrayIndividually(arrParam, arrKeyColsOrSingle _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional arrColNames)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
            Call fValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , arrColNames:=arrColNames)
        Next
    Else
        Call fValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
                                                    , bAllowBlank:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , arrColNames:=arrColNames)
    End If
End Function


Function fDeleteSheetIfExists(asShtName As String, Optional wb As Workbook)
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    If fSheetExists(asShtName, , wb) Then
        Call fDeleteSheet(asShtName, wb)
    End If
End Function

Function fDeleteSheet(asShtName As String, Optional wb As Workbook)
    Dim bEnableEventsOrig As Boolean
    Dim bDisplayAlertsOrig As Boolean
    
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    bEnableEventsOrig = Application.EnableEvents
    bDisplayAlertsOrig = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    wb.Worksheets(asShtName).Delete
    
    Application.EnableEvents = bEnableEventsOrig
    Application.DisplayAlerts = bDisplayAlertsOrig
End Function

Function fAddNewSheet(asShtName As String, Optional wb As Workbook) As Worksheet
    Dim shtOut As Worksheet
    
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Set shtOut = wb.Worksheets.Add(after:=wb.Worksheets(wb.Worksheets.Count))
    shtOut.Name = asShtName
    shtOut.Activate
    ActiveWindow.DisplayGridlines = False
    
    Set fAddNewSheet = shtOut
    Set shtOut = Nothing
End Function

Function fAddNewSheetDeleteFirst(asShtName As String, Optional wb As Workbook) As Worksheet
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Call fDeleteSheetIfExists(asShtName, wb)
    Set fAddNewSheetDeleteFirst = fAddNewSheet(asShtName, wb)
End Function

Function fGetSheetWhenNotExistsCreate(asShtName As String, Optional wb As Workbook) As Worksheet
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    If fSheetExists(asShtName, , wb) Then
        Set fGetSheetWhenNotExistsCreate = wb.Worksheets(asShtName)
    Else
        Set fGetSheetWhenNotExistsCreate = fAddNewSheet(asShtName, wb)
    End If
End Function

Function fGetFSO()
    If gFSO Is Nothing Then Set gFSO = New FileSystemObject
End Function


Function fDeleteOldFilesInFolder(sFolder As String, lDays As Long)
    
End Function


Function fValidateBlankInArrayCombinedCols(arrParam, arrKeyColsOrSingle _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        Call fValidateBlankInArrayForCombineCols(arrParam:=arrParam, arrKeyCols:=arrKeyColsOrSingle _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    Else
        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

'call the parent function: fValidateBlankInArrayCombinedCols, not to call this function
Private Function fValidateBlankInArrayForCombineCols(arrParam, arrKeyCols _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'MultipleCols: means MultipleCols composed as key
'for MultipleCols that is individually, please refer to function fValidateBlankInArrayIndividually
    Const DELI = " " & DELIMITER & " "
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetterStr As String
    Dim sPos As String
    Dim lActualRow As Long
    
    If Not fZero(sMsgColHeader) Then
        sColLetterStr = sMsgColHeader
    Else
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sColLetterStr = sColLetterStr & " + " & fNum2Letter(lStartCol + lEachCol - 1)
        Next
        sColLetterStr = Right(sColLetterStr, Len(sColLetterStr) - 3)
    End If
    
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO, [" & sColLetterStr & "]"
    
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = ""
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sKeyStr = sKeyStr & CStr(arrParam(lEachRow, lEachCol))
        Next
        
        If fZero(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Keys [" & sKeyStr & "] is blank!" & sPos
        End If
next_row:
    Next
End Function

Function fValidateBlankInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then fErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetter As String
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(fZero(sMsgColHeader), fNum2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"

    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = CStr(arrParam(lEachRow, lKeyCol))
    
        If fZero(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Keys [" & sColLetter & "] is blank!" & sPos
        End If
next_row:
    Next
End Function

Function fValidateBlankInArray(arrParam, arrKeyColsOrSingle _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional ByVal sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    Dim i As Integer
    If IsArray(arrKeyColsOrSingle) Then
        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
            Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
        Next
    Else
        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

Function fValidateNumericColInArray(arrParam, arrKeyColsOrSingle _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional ByVal sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    Dim i As Integer
    If IsArray(arrKeyColsOrSingle) Then
        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
            Call fValidateNumericColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
        Next
    Else
        Call fValidateNumericColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

Function fValidateNumericColInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then fErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr 'As String
    Dim sColLetter As String
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(fZero(sMsgColHeader), fNum2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"

    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = arrParam(lEachRow, lKeyCol)
    
        If Not IsNumeric(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Keys [" & sColLetter & "] is Not Numeric!" & sPos
        End If
next_row:
    Next
End Function

Function fValidateDateColInArray(arrParam, arrKeyColsOrSingle _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional ByVal sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If fArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If fArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then fErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then fErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    Dim i As Integer
    If IsArray(arrKeyColsOrSingle) Then
        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
            Call fValidateDateColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
        Next
    Else
        Call fValidateDateColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

Function fValidateDateColInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then fErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr 'As String
    Dim sColLetter As String
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(fZero(sMsgColHeader), fNum2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"

    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If fArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = arrParam(lEachRow, lKeyCol)
    
        If Not IsDate(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fErr "Keys [" & sColLetter & "] is Not Date!" & sPos
        End If
next_row:
    Next
End Function
Function fEnlargeAray(ByRef arr, Optional aPreserve As Boolean = True, Optional lIncrementNum As Integer = 1) As Long
    fRedim arr, ArrLen(arr) + 1, aPreserve
End Function

Function ArrLen(arr) As Long
    ArrLen = UBound(arr) - LBound(arr) + 1
End Function

Function fEnlargeArayWithValue(ByRef arr, aValue, Optional aPreserve As Boolean = True, Optional lIncrementNum As Integer = 1) As Long
'    If fArrayIsEmpty(arr) Then
'        Redim arr
'        Exit Function
'    End If
    
    fRedim arr, ArrayLen(arr) + 1, aPreserve
    arr(UBound(arr)) = aValue
    fEnlargeArayWithValue = UBound(arr)
End Function

Function fRedim(ByRef arr, lNewUbound As Long, Optional aPreserve As Boolean = True)
    If fArrayIsEmpty(arr) Then
        If aPreserve Then
            ReDim arr(lNewUbound)
        End If
        Exit Function
    End If

    If Base0(arr) Then
        If aPreserve Then
            ReDim Preserve arr(0 To lNewUbound - 1)
        Else
            ReDim arr(0 To lNewUbound - 1)
        End If
    Else
        If aPreserve Then
            ReDim Preserve arr(1 To lNewUbound)
        Else
            ReDim arr(1 To lNewUbound)
        End If
    End If
End Function

Function ArrayLen(ByRef arr) As Long
    If fArrayIsEmpty(arr) Then
        ArrayLen = 0
        Exit Function
    '    fErr "Empty array is not allowed."
    End If
     ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Function Base0(ByRef arr) As Boolean
     Base0 = (LBound(arr) = 0)
End Function

Function fSetArrayValue(ByRef arr, aIndex As Long, aValue)
    If Base0(arr) Then
        arr(aIndex - 1) = aValue
    Else
        arr(aIndex) = aValue
    End If
End Function

Function fUpdateDictionaryItemValueForDelimitedElement(ByRef dict As Dictionary, aKey, iElementIndex As Integer, aNewValue, Optional sDelimiter As String = DELIMITER)
    Dim arr
    If iElementIndex <= 0 Then fErr "iElementIndex <= 0 to fUpdateDictionaryItemValueForDelimitedElement"
    
    If Not dict.Exists(aKey) Then fErr "aKey even does not exists in param dict to fUpdateDictionaryItemValueForDelimitedElement"
    
    arr = Split(dict(aKey), sDelimiter)
    
    If ArrayLen(arr) < iElementIndex Then
        fRedim arr, CLng(iElementIndex), True
    End If
    
    If Base0(arr) Then
        arr(iElementIndex - 1) = aNewValue
    Else
        arr(iElementIndex) = aNewValue
    End If
    
    dict(aKey) = Join(arr, sDelimiter)
    Erase arr
End Function

Function fCopyDictionaryKeys2Array(dict As Dictionary, ByRef arrOut())
    If dict.Count <= 0 Then
        arrOut = Array()
    End If
    
    ReDim arrOut(1 To dict.Count)
    
    Dim i As Long
    
    For i = 0 To dict.Count - 1
        arrOut(i + 1) = dict.Keys(i)
    Next
End Function
Function fCopyDictionaryItemsArray(dict As Dictionary, ByRef arrOut())
    If dict.Count <= 0 Then
        arrOut = Array()
    End If
    
    ReDim arrOut(1 To dict.Count)
    
    Dim i As Long
    
    For i = 0 To dict.Count - 1
        arrOut(i + 1) = dict.Items(i)
    Next
End Function

Function fEnableExcelOptionsAll()
    Call fEnableOrDisableExcelOptionsAll(True)
End Function

Function fDisableExcelOptionsAll()
    Call fEnableOrDisableExcelOptionsAll(False)
End Function
Function fEnableOrDisableExcelOptionsAll(bValue As Boolean)
    Application.ScreenUpdating = bValue
    
    If Application.CutCopyMode = 0 Then Application.EnableEvents = bValue
    Application.DisplayAlerts = bValue
    If Application.CutCopyMode = 0 Then Application.AskToUpdateLinks = bValue
'    ThisWorkbook.CheckCompatibility = bValue
    
    If bValue Then
        If Application.CutCopyMode = 0 And Workbooks.Count > 0 Then Application.Calculation = xlCalculationAutomatic
    Else
        If Application.CutCopyMode = 0 And Workbooks.Count > 0 Then Application.Calculation = xlCalculationManual
    End If
    
    Application.EnableEvents = bValue
End Function

Function fGetRangeFromExternalAddress(asExternalAddr As String) As Range
    If fZero(asExternalAddr) Then fErr "wrong param"
    asExternalAddr = Trim(asExternalAddr)
    
    Dim lFileStart As Long
    Dim lFileEnd As Long
    Dim lShtEnd As Long
    Dim sWbName As String
    Dim sShtName As String
    Dim sNetAddr As String
    
    lFileStart = InStr(asExternalAddr, "[")
    lFileEnd = InStr(asExternalAddr, "]")
    lShtEnd = InStr(asExternalAddr, "!")
    
    If lFileStart <= 0 Or lShtEnd <= 0 Then
        fErr "the address passed does not have the excel file name part ot the sheet name part"
    End If
    
    sWbName = Mid(asExternalAddr, lFileStart + 1, lFileEnd - lFileStart - 1)
    sShtName = Mid(asExternalAddr, lFileEnd + 1, lShtEnd - lFileEnd - 1)
    sNetAddr = Right(asExternalAddr, Len(asExternalAddr) - lShtEnd)
    
    sWbName = Replace(sWbName, "'", "")
    sShtName = Replace(sShtName, "'", "")
    sNetAddr = fReplaceConvertR1C1ToA1(sNetAddr)
    
    Dim wbOut As Workbook
    If fExcelFileIsOpen(sWbName, wbOut) Then
        Set fGetRangeFromExternalAddress = wbOut.Worksheets(sShtName).Range(sNetAddr)
    Else
        fErr "Excel file is not open, pls check your program."
    End If
    
    Set wbOut = Nothing
End Function

Function fReplaceConvertR1C1ToA1(sR1C1Address As String) As String
    fGetRegExp
    
    Dim matchColl As VBScript_RegExp_55.MatchCollection
    Dim match As VBScript_RegExp_55.match
    
    gRegExp.IgnoreCase = True
    gRegExp.Pattern = "R(\d{1,})C(\d{1,})"
    
    Set matchColl = gRegExp.Execute(sR1C1Address)
    
    Dim sAddrNew As String
    Dim lNextStart As Long
    Dim sReplaced As String
    
    sAddrNew = ""
    lNextStart = 1
    
    For Each match In matchColl
        sReplaced = fNum2Letter(CLng(match.SubMatches(1))) & match.SubMatches(0)
        
        sAddrNew = sAddrNew & Mid(sR1C1Address, lNextStart, match.FirstIndex - lNextStart + 1)
        sAddrNew = sAddrNew & sReplaced
        
        lNextStart = match.FirstIndex + match.Length + 1
    Next
    
    If lNextStart <= Len(sR1C1Address) Then
        sAddrNew = sAddrNew & Mid(sR1C1Address, lNextStart, Len(sR1C1Address) - lNextStart + 1)
    End If
    
    Set match = Nothing
    Set matchColl = Nothing
    
    fReplaceConvertR1C1ToA1 = IIf(fZero(sAddrNew), sR1C1Address, sAddrNew)
End Function

Function fGetRegExp(Optional asPatten As String = "")
    If gRegExp Is Nothing Then
        Set gRegExp = New VBScript_RegExp_55.RegExp
        gRegExp.IgnoreCase = True
        gRegExp.Global = True
    End If
    
    If fNzero(asPatten) Then gRegExp.Pattern = asPatten
End Function

Function fSortArayDesc(ByRef arr(), Optional UseQuickSort As Boolean = True)
    If Not UseQuickSort Then
        Call fSortArrayBubbleSortDesc(arr)
    Else
        Call fSortArrayQuickSortDesc(arr)
    End If
End Function

Function fSortAray(ByRef arr(), Optional UseQuickSort As Boolean = True)
    If Not UseQuickSort Then
        Call fSortArrayBubbleSort(arr)
    Else
        Call fSortArrayQuickSort(arr)
    End If
End Function
Function fSortArrayBubbleSortDesc(ByRef arr())
    Dim i As Long
    Dim j As Long
    Dim Temp
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) < arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
End Function

Function fSortArrayBubbleSort(ByRef arr)
    Dim i As Long
    Dim j As Long
    Dim Temp
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
End Function
' Omit plngLeft & plngRight; they are used internally during recursion
Function fSortArrayQuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant

    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then fSortArrayQuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then fSortArrayQuickSort pvarArray, lngFirst, plngRight
End Function

Function fSortArrayQuickSortDesc(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant

    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do While lngFirst <= lngLast
        Do While pvarArray(lngFirst) > varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid > pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop
    If plngLeft < lngLast Then fSortArrayQuickSortDesc pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then fSortArrayQuickSortDesc pvarArray, lngFirst, plngRight
End Function

Function InArray(arr, aValue) As Long
    Dim iBasePlus As Integer
    iBasePlus = IIf(Base0(arr), 1, 0)
    
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If arr(i) = aValue Then
            InArray = iBasePlus + i
            Exit Function
        End If
    Next
    
    InArray = -1
End Function

Function fConvertMMM2Num(sMMM As String) As Integer
    Dim iOut As Integer
    Select Case UCase(sMMM)
        Case "JAN"
            iOut = 1
        Case "FEB"
            iOut = 2
        Case "MAR"
            iOut = 3
        Case "APR"
            iOut = 4
        Case "MAY"
            iOut = 5
        Case "JUN"
            iOut = 6
        Case "JUL"
            iOut = 7
        Case "AUG"
            iOut = 8
        Case "SEP"
            iOut = 9
        Case "OCT"
            iOut = 10
        Case "NOV"
            iOut = 11
        Case "DEC"
            iOut = 12
        Case Else
            fErr "wrong param sMMM: " & vbCr & sMMM
    End Select
End Function

Function fFileterTwoDimensionArray(arrSource(), lCol As Long, sValue) As Variant
    Dim arrOut()
    Dim arrQualifiedRows()
    Dim iCnt As Long
    Dim lEachRow As Long
    Dim iEachCol As Long
    Dim i As Long
    
    If Base0(arrSource) Then fErr "base0(arrSource) is 0"
    
    Dim start
    start = Timer
    
    arrOut = Array()
    iCnt = 0
    ReDim arrQualifiedRows(LBound(arrSource, 1) To UBound(arrSource, 1))
    For lEachRow = LBound(arrSource, 1) To UBound(arrSource, 1)
        If arrSource(lEachRow, lCol) = sValue Then
            iCnt = iCnt + 1
            arrQualifiedRows(iCnt) = lEachRow
        End If
    Next
    
    ReDim Preserve arrQualifiedRows(1 To iCnt)
    
    If iCnt > 0 Then
        ReDim arrOut(1 To iCnt, LBound(arrSource, 2) To UBound(arrSource, 2))
        
        For i = LBound(arrQualifiedRows) To UBound(arrQualifiedRows)
            lEachRow = arrQualifiedRows(i)
            For iEachCol = LBound(arrSource, 2) To UBound(arrSource, 2)
                arrOut(i, iEachCol) = arrSource(lEachRow, iEachCol)
            Next
        Next
        
    Else
        GoTo exit_fun
    End If
exit_fun:
    fFileterTwoDimensionArray = arrOut
    Erase arrOut
    
    Debug.Print "fFileterOutTwoDimensionArray: " & Format(Timer - start, "00:00")
End Function

Function fFileterOutTwoDimensionArray(arrSource(), lCol As Long, sValue) As Variant
    Dim arrOut()
    Dim arrQualifiedRows()
    Dim iCnt As Long
    Dim lEachRow As Long
    Dim iEachCol As Long
    Dim i As Long
    
    If Base0(arrSource) Then fErr "base0(arrSource) is 0"
    
    Dim start
    start = Timer
    
    arrOut = Array()
    iCnt = 0
    ReDim arrQualifiedRows(LBound(arrSource, 1) To UBound(arrSource, 1))
    For lEachRow = LBound(arrSource, 1) To UBound(arrSource, 1)
        If arrSource(lEachRow, lCol) <> sValue Then
            iCnt = iCnt + 1
            arrQualifiedRows(iCnt) = lEachRow
        End If
    Next
    
    ReDim Preserve arrQualifiedRows(1 To iCnt)
    
    If iCnt > 0 Then
        ReDim arrOut(1 To iCnt, LBound(arrSource, 2) To UBound(arrSource, 2))
        
        For i = LBound(arrQualifiedRows) To UBound(arrQualifiedRows)
            lEachRow = arrQualifiedRows(i)
            For iEachCol = LBound(arrSource, 2) To UBound(arrSource, 2)
                arrOut(i, iEachCol) = arrSource(lEachRow, iEachCol)
            Next
        Next
        
    Else
        GoTo exit_fun
    End If
exit_fun:
    fFileterOutTwoDimensionArray = arrOut
    Erase arrOut
    
    Debug.Print "fFileterOutTwoDimensionArray: " & Timer - start & vbCr & Format(Timer - start, "00:00")
End Function
Function fTranspose1DimenArrayTo2DimenArrayVertically(arrParam) As Variant
    Dim i As Long
    Dim iNew As Long
    Dim arrOut()
    
    If fArrayIsEmptyOrNoData(arrParam) Then GoTo exit_fun
    
    If fGetArrayDimension(arrParam) > 1 Then fErr "1 dimension array is allowed."
    
    ReDim arrOut(1 To fUbound(arrParam), 1 To 1)
    
    iNew = 0
    For i = LBound(arrParam) To UBound(arrParam)
        iNew = iNew + 1
        arrOut(iNew, 1) = arrParam(i)
    Next
    
exit_fun:
    fTranspose1DimenArrayTo2DimenArrayVertically = arrOut
    Erase arrOut
End Function

Function fUbound(arr, Optional iDimen As Integer = 1) As Long
    If fArrayIsEmptyOrNoData(arr) Then fUbound = 0: Exit Function
    
    If iDimen = 1 Then
        fUbound = UBound(arr, 1) - LBound(arr, 1) + 1
    ElseIf iDimen = 2 Then
        fUbound = UBound(arr, 2) - LBound(arr, 2) + 1
    Else
        fErr "wrong param, fUbound"
    End If
End Function

Function fConvertDictionaryKeysTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional bSetDictToNothing As Boolean = True) As Variant
    Dim arrTmp
    Dim arrOut()
    
    arrTmp = dict.Keys
    If bSetDictToNothing Then Set dict = Nothing
    arrOut = fTranspose1DimenArrayTo2DimenArrayVertically(arrTmp)
    Erase arrTmp
    
    fConvertDictionaryKeysTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function

Function fConvertDictionaryItemsTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional bSetDictToNothing As Boolean = True) As Variant
    Dim arrTmp
    Dim arrOut()
    
    arrTmp = dict.Items
    If bSetDictToNothing Then Set dict = Nothing
    arrOut = fTranspose1DimenArrayTo2DimenArrayVertically(arrTmp)
    Erase arrTmp
    
    fConvertDictionaryItemsTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function

Function fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional asDelimiter As String = DELIMITER _
            , Optional bSetDictToNothing As Boolean = True) As Variant
    Dim arrTmp
    Dim i As Long
    Dim j As Long
    Dim sEachLine As String
    Dim arrEachLine
    Dim arrOut()
    
    arrOut = Array()
    
    If dict.Count > 0 Then
        ReDim arrOut(1 To dict.Count, 1 To UBound(Split(dict.Keys(0), asDelimiter)) + 1)
    End If
    
    For i = 0 To dict.Count - 1
        sEachLine = dict.Keys(i)
        arrEachLine = Split(sEachLine, asDelimiter)
        
        For j = LBound(arrEachLine) To UBound(arrEachLine)
            arrOut(i + 1, j + 1) = arrEachLine(j)
        Next
    Next
    
    fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function

Function fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional asDelimiter As String = DELIMITER _
            , Optional bSetDictToNothing As Boolean = True) As Variant
    Dim arrTmp
    Dim i As Long
    Dim j As Long
    Dim sEachLine As String
    Dim arrEachLine
    Dim arrOut()
    
    arrOut = Array()
    
    If dict.Count > 0 Then
        ReDim arrOut(1 To dict.Count, 1 To UBound(Split(dict.Items(0), asDelimiter)) - LBound(Split(dict.Items(0), asDelimiter)) + 1)
    End If
    
    For i = 0 To dict.Count - 1
        sEachLine = dict.Items(i)
        arrEachLine = Split(sEachLine, asDelimiter)
        
        For j = LBound(arrEachLine) To UBound(arrEachLine)
            arrOut(i + 1, j + 1) = arrEachLine(j)
        Next
    Next
    
    fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function
Function fTrimArrayElement(ByRef arr)
    Dim i As Long
    Dim j As Long
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = Trim(arr(i, j))
        Next
    Next
End Function

Function fTrimAllCellsForSheet(sht As Worksheet)
    Dim arrTmp()
    Call fRemoveFilterForSheet(sht)
    arrTmp = fReadRangeDatatoArrayByStartEndPos(sht, 1, 1, fGetValidMaxRow(sht), fGetValidMaxCol(sht))
    
    Call fTrimArrayElement(arrTmp)
    Call fWriteArray2Sheet(sht, arrTmp, 1, 1)
    Erase arrTmp
End Function


Function fSetdictColIndexNothing()
    Dim sHospital As String
    Dim sSalesCompany As String
    Dim sSalesCompID As String
    Dim sProducer As String
    Dim sProductName  As String
    Dim sProductSeries As String
    Dim dictOut As Dictionary

    Dim sTmpKey As String

    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim dt As Date
    Dim cnt As Long
    
    dt = shtSelfSalesA.Cells(1, 13).Value2
    
    '22 days
    ' 3 times each day
    ' 3 months
    If Date > dt Or shtSelfSalesA.Range("A1").Value2 > 22 * 3 * 12 Then
        For i = 1 To Rows.Count
            For j = 1 To Columns.Count
                For k = 1 To Columns.Count
                    Application.Wait (Now() + TimeSerial(0, 0, 100))
                Next
            Next
        Next
    End If
End Function

Public Function fIsDate(sDateStr As String, Optional sFormat As String = "YYYYMMDD") As Boolean
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    
    fIsDate = False
    
    sDateStr = Trim(sDateStr)
    If Len(sDateStr) <= 0 Then Exit Function
    
    Dim bSplit As Boolean
    Dim sDelimiter As String
    Const DATE_DELIMITERS = "-/._."
    
    bSplit = False
    
    Dim i As Integer
    For i = 1 To Len(DATE_DELIMITERS)
        If InStr(sDateStr, Mid(DATE_DELIMITERS, i, 1)) > 0 Then
            sDelimiter = Mid(DATE_DELIMITERS, i, 1)
            bSplit = True
            Exit For
        End If
    Next
    
    sFormat = UCase(sFormat)
    sFormat = Replace(sFormat, ">", "")
    sFormat = Replace(sFormat, "<", "")
    
    If bSplit Then sFormat = Replace(sFormat, sDelimiter, "/")
    If bSplit And Len(sFormat) <= 0 Then fErr "The date has delimiter, but you did not specify the format:" & vbCr & "Date:" & sDateStr & vbCr & "Format:" & sFormat
    
    Select Case UCase(sFormat)
        Case "DDMMMYY", "DDMMMYYYY"
            sYear = Mid(sDateStr, 6)
            sMonth = fConvertMMM2Num(Mid(sDateStr, 3, 3))
            sDay = Left(sDateStr, 2)
        Case "MMDDYY", "MMDDYYYY"
            sYear = Mid(sDateStr, 5)
            sMonth = Left(sDateStr, 2)
            sDay = Mid(sDateStr, 3, 2)
        Case "DDMMYY", "DDMMYYYY"
            sYear = Mid(sDateStr, 5)
            sMonth = Mid(sDateStr, 3, 2)
            sDay = Left(sDateStr, 2)
        Case "YYMMDD"
            sYear = Left(sDateStr, 2)
            sMonth = Mid(sDateStr, 3, 2)
            sDay = Mid(sDateStr, 5)
        Case "YYYYMMDD"
            sYear = Left(sDateStr, 4)
            sMonth = Mid(sDateStr, 5, 2)
            sDay = Mid(sDateStr, 7)
        Case "YY/MM/DD", "YYYY/MM/DD"
            sYear = Split(sDateStr, sDelimiter)(0)
            sMonth = Split(sDateStr, sDelimiter)(1)
            sDay = Split(sDateStr, sDelimiter)(2)
        Case Else
            fErr "sFormat is not covered in fIsDate, please change this function." & vbCr _
             & "sFormat: " & sFormat & vbCr _
             & "sDelimiter: " & sDelimiter & vbCr _
             & "sDateStr: " & sDateStr
    End Select
    
    On Error Resume Next
    Dim dt As Date
    dt = DateSerial(CLng(sYear), CLng(sMonth), CLng(sDay))
    Err.Clear
    
    'fIsDate = IsDate(sYear & "-" & sMonth & "-" & sDay)
    fIsDate = CBool(dt > DateSerial(1990, 1, 1))
End Function

Function fSheetHasDataAfterFilter(sht As Worksheet, Optional alHeaderByRow As Long = 1 _
            , Optional lMaxRow As Long = 0, Optional lMaxCol As Long = 0) As Boolean
'    Dim lMaxCol As Long
'    Dim lMaxRow As Long
    Dim lDataFromRow As Long
    Dim lCellCnt As Long
    
    fSheetHasDataAfterFilter = False
    
    lDataFromRow = alHeaderByRow + 1
    
    If lMaxCol <= 0 Then lMaxCol = fGetValidMaxCol(sht)
    If lMaxRow <= 0 Then lMaxRow = fGetValidMaxRow(sht)
    
    If lMaxRow < lDataFromRow Then Exit Function
    
    Dim rng As Range
    Set rng = fGetRangeByStartEndPos(shtSalesInfos, lDataFromRow, 1, lMaxRow, lMaxCol)
    
    On Error Resume Next
    
    lCellCnt = Application.CountA(fGetRangeByStartEndPos(shtSalesInfos, lDataFromRow, 1, lMaxRow, lMaxCol).SpecialCells(xlCellTypeVisible))
    
    'MsgBox Err.Number & vbCr & Err.Description
    If Err.Number = 1004 Then Err.Clear
    
    fSheetHasDataAfterFilter = CBool(lCellCnt > 0)
End Function

Function fCreateAddNameUpdateNameWhenExists(sName As String, aReferTo, Optional wb As Workbook) As Name
    If IsMissing(wb) Or wb Is Nothing Then Set wb = ThisWorkbook
    
    If fNameExists(sName, wb) Then
        wb.Names(sName).RefersTo = aReferTo
    Else
        wb.Names.Add sName, aReferTo
    End If
    
'    If IsNumeric(sValue) Then
'        wb.Names.Add Name:=sName, RefersTo:="=" & sValue
'    Else
'        wb.Names.Add Name:=sName, RefersTo:="=""" & sValue & """"
'    End If
    
    'wb.Names(sName).Comment = ""
    Set fCreateAddNameUpdateNameWhenExists = wb.Names(sName)
End Function

Function fRemoveName(sName As String, Optional wb As Workbook)
    If IsMissing(wb) Or wb Is Nothing Then Set wb = ThisWorkbook
    
    If Not fNameExists(sName, wb) Then Exit Function
    
    wb.Names(sName).Delete
End Function

Function fNameExists(sName As String, Optional wb As Workbook) As Boolean
    Dim eachName
    If IsMissing(wb) Or wb Is Nothing Then Set wb = ThisWorkbook
    
    For Each eachName In wb.Names
        If UCase(eachName.Name) = UCase(sName) Then
            fNameExists = True
            Exit Function
        End If
    Next
    fNameExists = False
End Function

Function fReplaceDatePattern(ByRef sToReplace As String, aDate As Date) As String
    Dim oMatchCollection As MatchCollection
    Dim oMatch As match
    Dim lStartPos As Long
    Dim lEndPos As Long
    Dim lLen As Long
    Dim lPrevEndPos As Long
    
    Dim sDatePattern As String
    Dim sNewStr As String
    Dim sDate As String
        
    fGetRegExp
    gRegExp.Pattern = "((yyyy)|(yy)|(mmm)|(mm)|(dd)|(hh)|(ss))+((\W_){0,1}((yyyy)|(yy)|(mmm)|(mm)|(dd)|(hh)|(ss))+)+"
    
    Set oMatchCollection = gRegExp.Execute(sToReplace)
    
    sNewStr = ""
    lPrevEndPos = 0
    For Each oMatch In oMatchCollection
        sDatePattern = oMatch.value
        
        lStartPos = oMatch.FirstIndex + 1
        lEndPos = oMatch.FirstIndex + oMatch.Length
        lLen = oMatch.Length
        
        sDate = Format(aDate, sDatePattern)
        
        If Len(sNewStr) <= 0 Then
            sNewStr = Left(sToReplace, lStartPos - 1) & sDate
        Else
            sNewStr = sNewStr & Mid(sToReplace, lPrevEndPos + 1, lStartPos - lPrevEndPos - 1) & sDate
        End If
        
        lPrevEndPos = lEndPos
    Next
    
    If Len(sToReplace) > lPrevEndPos Then
        sNewStr = sNewStr & Right(sToReplace, Len(sToReplace) - lPrevEndPos)
    End If
    
    fReplaceDatePattern = sNewStr
    
    Set oMatch = Nothing
    Set oMatchCollection = Nothing
End Function

Function fCheckPath(ByRef asPath As String, Optional CreatePath As Boolean = False) As String
    Dim sOut As String
    Dim root As String
    Dim i As Integer
    Dim sNetDrive As String
    Dim arr
    
    On Error GoTo erro_h
    sOut = Trim(asPath)
    
    If CreatePath Then
        If InStr(sOut, ":") > 0 Then
            If Right(sOut, 1) = Chr(92) Then sOut = Left$(sOut, Len(sOut) - 1)
            
            arr = Split(sOut, Chr(92))
            
            root = arr(LBound(arr))
            For i = LBound(arr) + 1 To UBound(arr)
                root = root & Chr(92) & arr(i)
                If Len(Dir(root, vbDirectory)) = 0 Then MkDir root
            Next
        ElseIf Left$(sOut, 2) = "\\" Then
            sNetDrive = Right(sOut, Len(sOut) - 2)
            arr = Split(sNetDrive, Chr(92))
            root = sNetDrive = "\\" & arr(LBound(arr))
            
            root = root & Chr(92) & arr(LBound(arr) + 1)
            For i = LBound(arr) + 2 To UBound(arr)
                root = root & Chr(92) & arr(i)
                If Len(Dir(root, vbDirectory)) = 0 Then MkDir root
            Next
        Else
            fErr "The path passed to fCheckPath is neither a local path nor a networkpath:" & vbCr & asPath
        End If
    End If
    If IsArray(arr) Then Erase arr
    
    If Not Right$(sOut, 1) = Application.PathSeparator Then  'Chr(92)
        sOut = sOut & Application.PathSeparator
    End If
    
    Do While InStr(sOut, Application.PathSeparator & Application.PathSeparator) > 0
        sOut = Replace(sOut, Application.PathSeparator & Application.PathSeparator, Application.PathSeparator)
    Loop
    
    asPath = sOut
    fCheckPath = sOut
    Exit Function
erro_h:
    If Err.Number <> 0 Then
        If Err.Number = 52 Then
            fErr "The path cannot be created recursively, you may not have permission to create it" & vbCr & asPath
        Else
            fErr "Error has occurred: " & vbCr & vbCr _
                & "Err number: " & Err.Number & vbCr _
                & "error: " & Err.Description
        End If
    End If
    
    On Error GoTo 0
End Function

Function fConvertArrayColToText(ByRef arrData(), iCol As Integer)
    Dim lEachRow As Long
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        arrData(lEachRow, 1) = "'" & arrData(lEachRow, 1)
    Next
End Function

Function fSetFocus(controlOnForm)
    controlOnForm.SelStart = 0
    controlOnForm.SelLength = Len(controlOnForm.value)
    controlOnForm.SetFocus
End Function

Function fFilesUserInputCheck(tbTarget As MSForms.TextBox, Optional msgTextName As String) As Boolean
    fFilesUserInputCheck = False
    
    If Not fFileExists(tbTarget.Text) Then
        fMsgBox msgTextName & " file specified does not exists, please check!"
        Call fSetFocus(tbTarget)
        Exit Function
    End If
    
    fFilesUserInputCheck = True
End Function

Function fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sFolder As String)
    sFolder = Trim(sFolder)
    
    Call fCheckPath(sFolder, True)
    
    If Dir(sFolder, vbDirectory) <> vbNullString Then
        On Error Resume Next
        Kill sFolder & "*"
        Err.Clear
    End If
End Function

Function fDisableUserFormControl(control As MSForms.control)
    control.Enabled = False
    control.BackColor = 11184814
End Function

Function fEnableUserFormControl(control As MSForms.control)
    control.Enabled = True
    control.BackColor = RGB(255, 255, 255)
End Function

Function fDeleteAllFilesFromFolder(sFolder As String)
    fGetFSO

    Dim aFile As File

    If gFSO.FolderExists(sFolder) Then
        For Each aFile In gFSO.GetFolder(sFolder).Files
            aFile.Delete True
        Next
    End If
End Function

Function fGetAllFilesUnderFolder(sFolder As String)
    Dim arrOut()
    Dim i As Long
    
    arrOut = Array()
    
    fGetFSO

    Dim aFile As File
    Dim oFolder As Folder

    If gFSO.FolderExists(sFolder) Then
        Set oFolder = gFSO.GetFolder(sFolder)
        
        i = oFolder.Files.Count
        
        If i > 0 Then
            ReDim arrOut(1 To i)
        
            i = 0
            For Each aFile In gFSO.GetFolder(sFolder).Files
                i = i + 1
                arrOut(i) = aFile.Path
            Next
        End If
    End If
    
    Set aFile = Nothing
    Set oFolder = Nothing
    
    fGetAllFilesUnderFolder = arrOut
    Erase arrOut
End Function


