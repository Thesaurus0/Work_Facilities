VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Enum enParamType
    None = 0
    ArrayParam
    StringParam
    ArrayButEmpty
    ArrayButOnlyOneElement
    ArrayButHasDuplicateElement
End Enum

Dim dblStartTime As Double
Private fso As FileSystemObject
Private Const BUSINESS_ERROR_NUMBER As Integer = 2000

Private Sub Class_Terminate()
    Set fso = Nothing
End Sub
 
Function StartTimer()
    dblStartTime = Timer
End Function
Function HowLong(Optional ByVal asPrefix As String = "") As Single
    HowLong = Timer - dblStartTime
    Debug.Print IIf(Len(asPrefix) > 0, asPrefix & vbTab, "") & Format(HowLong, "0000.00000000")
End Function

Function OpenFile(asFileFullPath As String)
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
        
        RaiseErr msg
    End If
End Function

Public Function GetMaxRow(shtParam As Worksheet, Optional abCountInMergedCell As Boolean = False) As Long
    Dim lExcelMaxRow As Long
    Dim lUsedMaxRow As Long
    Dim lUsedMaxCol As Long

    Dim lEachCol As Long
    Dim lMaxRow As Long
    Dim lEachValidMaxRow As Long
     
'    If abCountInMergedCell Then
'       abCountInMergedCell = (Not Sheet1.Cells.Find(what:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious _
'                                                 , searchformat:=True, MatchCase:=False, matchbyte:=False) Is Nothing)
'    End If
      
    lExcelMaxRow = shtParam.Rows.Count
    lUsedMaxRow = shtParam.UsedRange.row + shtParam.UsedRange.Rows.Count - 1
    lUsedMaxCol = shtParam.UsedRange.Column + shtParam.UsedRange.Columns.Count - 1
 
    lMaxRow = 0
    If lUsedMaxRow = 1 Then
        lMaxRow = 1
        GoTo final_check_blank
    Else
        For lEachCol = 1 To lUsedMaxCol
            lEachValidMaxRow = shtParam.Cells(lExcelMaxRow, lEachCol).End(xlUp).row
     
            If lEachValidMaxRow >= lUsedMaxRow Then GoTo final_check_blank
            
            If abCountInMergedCell Then
                If shtParam.Cells(lEachValidMaxRow, lEachCol).MergeCells Then
                    lEachValidMaxRow = shtParam.Cells(lEachValidMaxRow, lEachCol).MergeArea.row _
                                     + shtParam.Cells(lEachValidMaxRow, lEachCol).MergeArea.Rows.Count - 1
                End If
            End If
     
            If lEachValidMaxRow > lMaxRow Then lMaxRow = lEachValidMaxRow
            
            If lMaxRow >= lUsedMaxRow Then GoTo final_check_blank
        Next
    End If
   
final_check_blank:
    If lMaxRow = 1 And lUsedMaxCol = 1 Then
        If Len(Trim(CStr(shtParam.Range("A1")))) <= 0 Then
            lMaxRow = 0
        End If
    ElseIf lMaxRow = 1 Then
'        If WorksheetFunction.CountA(shtParam.Range(shtParam.Cells(1, 1), shtParam.Cells(1, lUsedMaxCol))) <= 0 Then
'            GetMaxRow = 0
'            Exit Function
'        End If
        Dim arr()
        Dim i As Integer
        arr = shtParam.Range(shtParam.Cells(1, 1), shtParam.Cells(lMaxRow, lUsedMaxCol)).Value
        
        For i = LBound(arr, 2) To UBound(arr, 2)
            If Len(Trim(arr(1, i))) <= 0 Then
                lMaxRow = 0
                Exit For
            End If
        Next
        Erase arr
    End If
    
    GetMaxRow = lMaxRow
End Function
 
'Public Function GetMaxRow(shtParam As Worksheet) As Long
'    Dim lMaxRow As Long
'    Dim rgLastCell As Range
'
'    Set rgLastCell = shtParam.Cells.Find("*", After:=shtParam.Cells(1, 1), SearchOrder:=xlByRows, LookIn:=xlValues, SearchDirection:=xlPrevious)
'    lMaxRow = rgLastCell.row
'
'    If lMaxRow = 1 Then
'        If WorksheetFunction.CountA(shtParam.Range(shtParam.Cells(1, 1), rgLastCell)) <= 0 Then
'            lMaxRow = 0
'        End If
'    End If
'
'    GetMaxRow = lMaxRow
'End Function

Public Function GetMaxCol(shtParam As Worksheet, Optional abCountInMergedCell As Boolean = False) As Long
    Dim lExcelMaxCol As Long
    Dim lUsedMaxRow As Long
    Dim lUsedMaxCol As Long

    Dim lEachRow As Long
    Dim lMaxCol As Long
    Dim lEachValidMaxCol As Long
    
    lExcelMaxCol = shtParam.Columns.Count
    lUsedMaxRow = shtParam.UsedRange.row + shtParam.UsedRange.Rows.Count - 1
    lUsedMaxCol = shtParam.UsedRange.Column + shtParam.UsedRange.Columns.Count - 1

    lMaxCol = 0
    If lUsedMaxRow = 1 Then
        lMaxCol = 1
        GoTo final_check_blank
    Else
        For lEachRow = 1 To lUsedMaxRow
            lEachValidMaxCol = shtParam.Cells(lEachRow, lExcelMaxCol).End(xlToLeft).Column
    
            If lEachValidMaxCol >= lUsedMaxCol Then GoTo final_check_blank
            
            If abCountInMergedCell Then
                If shtParam.Cells(lEachRow, lEachValidMaxCol).MergeCells Then
                    lEachValidMaxCol = shtParam.Cells(lEachRow, lEachValidMaxCol).MergeArea.Column _
                                     + shtParam.Cells(lEachRow, lEachValidMaxCol).MergeArea.Columns.Count - 1
                End If
            End If
    
            If lEachValidMaxCol > lMaxCol Then lMaxCol = lEachValidMaxCol
            
            If lMaxCol >= lUsedMaxCol Then GoTo final_check_blank
        Next
    End If
  
final_check_blank:
    If lMaxCol = 1 And lUsedMaxRow = 1 Then
        If Len(Trim(CStr(shtParam.Range("A1")))) <= 0 Then
            lMaxCol = 0
        End If
    ElseIf lMaxCol = 1 Then
'        If WorksheetFunction.CountA(shtParam.Range(shtParam.Cells(1, 1), shtParam.Cells(lUsedMaxRow, 1))) <= 0 Then
'            GetMaxCol = 0
'            Exit Function
'        End If
        Dim arr()
        Dim i As Integer
        arr = shtParam.Range(shtParam.Cells(1, 1), shtParam.Cells(lUsedMaxRow, lMaxCol)).Value
        
        For i = LBound(arr, 1) To UBound(arr, 1)
            If Len(Trim(arr(i, 1))) <= 0 Then
                lMaxCol = 0
                Exit For
            End If
        Next
        Erase arr
    End If
    
    GetMaxCol = lMaxCol
End Function

'Public Function GetMaxCol(shtParam As Worksheet) As Long
'    Dim lMaxCol As Long
'    Dim rgLastCell As Range
'
'    Set rgLastCell = shtParam.Cells.Find("*", After:=shtParam.Cells(1, 1), SearchOrder:=xlByColumns, LookAt:=xlPart, LookIn:=xlValues, SearchDirection:=xlPrevious, MatchCase:=False, matchbyte:=False)
'    lMaxCol = rgLastCell.Column
'
'    If lMaxCol = 1 Then
'        If WorksheetFunction.CountA(shtParam.Range(shtParam.Cells(1, 1), rgLastCell)) <= 0 Then
'            lMaxCol = 0
'        End If
'    End If
'
'    GetMaxCol = lMaxCol
'End Function
'
'Function GetMaxRowOfRange(rngParam As Range, Optional abCountInMergedCell As Boolean = False) As Long
'     Dim lOut As Long
'
'     'single cell
'     If RangeIsSingleCell(rngParam) Then lOut = rngParam.row:               GoTo exit_fun
'
'     Dim shtParent As Worksheet
'     Set shtParent = rngParam.Parent
'
'     Dim lExcelMaxRow As Long
'     Dim lExcelMaxCol As Long
'     Dim lShtValidMaxRow As Long
'     Dim lShtValidMaxCol As Long
'     Dim lRangeMaxRow As Long
'     Dim lRangeMaxCol As Long
'     Dim lValidMaxRowSaved As Long
'     Dim lEachValidMaxRow As Long
'     Dim lEachCol As Long
'
'     lExcelMaxRow = shtParent.Rows.Count
'     lExcelMaxCol = shtParent.Columns.Count
'     lRangeMaxRow = rngParam.row + rngParam.Rows.Count - 1
'     lRangeMaxCol = rngParam.Column + rngParam.Columns.Count - 1
'
'     lShtValidMaxRow = GetMaxRow(shtParent, abCountInMergedCell)
'     If lShtValidMaxRow < rngParam.row Then 'blank, out of usedrange
'        lOut = rngParam.row: GoTo exit_fun
'     End If
'
'     lShtValidMaxCol = GetMaxCol(shtParent, abCountInMergedCell)
'     If lShtValidMaxCol < rngParam.Column Then 'blank, out of usedrange
'        lOut = rngParam.row: GoTo exit_fun
'     End If
'
'     'whole sheet
'     If rngParam.Rows.Count = lExcelMaxRow And rngParam.Columns.Count = lExcelMaxCol Then
'        lOut = lShtValidMaxRow: GoTo exit_fun
'     End If
'
'     If lRangeMaxRow > lShtValidMaxRow Then 'shrink row
'        lRangeMaxRow = lShtValidMaxRow
'     End If
'     If lRangeMaxCol > lShtValidMaxCol Then 'shrink col
'        lRangeMaxCol = lShtValidMaxCol
'     End If
'
''     'several rows
''     If rngParam.Columns.Count = lExcelMaxCol Then
''        lOut = lRangeMaxRow: GoTo exit_fun
''     End If
'
'     'several columns
'     If rngParam.Rows.Count = lExcelMaxRow Then
'        lValidMaxRowSaved = 0
'
'        For lEachCol = rngParam.Column To lRangeMaxCol
'            lEachValidMaxRow = shtParent.Cells(lExcelMaxRow, lEachCol).End(xlUp).row
'
'            If lEachValidMaxRow >= lShtValidMaxRow Then
'                lOut = lShtValidMaxRow
'                GoTo exit_fun
'            End If
'
'            If abCountInMergedCell Then
'                If shtParent.Cells(lEachValidMaxRow, lEachCol).MergeCells Then
'                    lEachValidMaxRow = shtParent.Cells(lEachValidMaxRow, lEachCol).MergeArea.row _
'                                     + shtParent.Cells(lEachValidMaxRow, lEachCol).MergeArea.Rows.Count - 1
'                End If
'            End If
'
'            If lEachValidMaxRow > lValidMaxRowSaved Then lValidMaxRowSaved = lEachValidMaxRow
'        Next
'
'        lOut = lValidMaxRowSaved: GoTo exit_fun
'    End If
'
'    Dim arrShrunk()
'    Dim lArrMaxRow As Long
'    Dim lArrMaxCol As Long
'    arrShrunk = ReadRangeDatatoArrayByStartEndPos(shtParent, rngParam.row, rngParam.Column, lRangeMaxRow, lRangeMaxCol)
'    lArrMaxRow = GetArrayMaxValidRowCol(arrShrunk, lArrMaxCol)
'    Erase arrShrunk
'
'    lArrMaxRow = rngParam.row + lArrMaxRow + IIf(lArrMaxRow > 0, -1, 0)
'    lArrMaxCol = rngParam.Column + lArrMaxCol + IIf(lArrMaxCol > 0, -1, 0)
'
'    lOut = lArrMaxRow
'    lEachValidMaxRow = 0
'    If abCountInMergedCell Then
'        If shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeCells Then
'        '    If shtParent.Cells(lOut, lEachCol).MergeArea.Rows.Count > 1 Then
'                lEachValidMaxRow = shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeArea.row _
'                                 + shtParent.Cells(lArrMaxRow, lArrMaxCol).MergeArea.Rows.Count - 1
'         '   End If
'        End If
'
'        If lEachValidMaxRow > lOut Then lOut = lEachValidMaxRow
'    End If
'
'exit_fun:
'    GetMaxRowOfRange = lOut
'End Function

Function GetArrayMaxValidRowCol(arrParam(), Optional lMaxCol As Long, Optional bReverse As Boolean = True) As Long
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
    GetArrayMaxValidRowCol = lMaxRowSaved
End Function

Function RangeIsSingleCell(rngParam As Range) As Boolean
    RangeIsSingleCell = (rngParam.Rows.Count = 1 And rngParam.Columns.Count = 1)
End Function
 
Function RaiseErr(Optional sMsg As String = "", Optional alErrNum As Long = 0) As VbMsgBoxResult
    Dim lErrNum As Long
    
    lErrNum = IIf(alErrNum <> 0, alErrNum, vbObjectError + BUSINESS_ERROR_NUMBER)
      
    If Nzero(sMsg) Then MsgBox sMsg, vbCritical
    
    Err.Raise lErrNum, "", "Program is to be terminated."
End Function
  
Function SelectFileDialog(Optional asDefaultFilePath As String = "" _
                         , Optional asFileFilters As String = "", Optional asTitle As String = "") As String
    'asFileFilters :   Excel File= *.xlsx;*.xls
    Dim fd As FileDialog
    Dim sFilterDesc As String
    Dim sFilterStr As String
    Dim sDefaultFile As String
    
    If Len(Trim(asFileFilters)) > 0 Then
        sFilterDesc = Trim(Split(asFileFilters, "=")(0))
        sFilterStr = Trim(Split(asFileFilters, "=")(1))
    End If
    
    If Len(Trim(asDefaultFilePath)) > 0 Then
       ' sDefaultFile = GetFileParentFolder(asDefaultFilePath)
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
        SelectFileDialog = fd.SelectedItems(1)
    Else
        SelectFileDialog = ""
    End If
        
    Set fd = Nothing
End Function

Function SelectSaveAsFileDialog(Optional asDeafaulfFilePath As String = "", Optional asFileFilters = "", Optional asTitle = "") As String
'asFileFilters  : "Excel File(*.xlsx),*.xlsx, Excel Old Ver(*.xls),*.xls"
' Important: the default file name must be same as the file filter extension
'   default_file.xlsx  = "Excel File(*.xlsx),*.xlsx"
    Dim fd As FileDialog
    Dim sDefaultFolder As String
    Dim sDefaultFile As String
    Dim sOut 'As String
    
    GetFSO
    
    If Len(Trim(asDeafaulfFilePath)) > 0 Then
        sDefaultFolder = GetFileParentFolder(asDeafaulfFilePath)
        sDefaultFile = GetFileBaseName(asDeafaulfFilePath)
        
        If Not fso.FolderExists(sDefaultFolder) Then sDefaultFolder = ThisWorkbook.Path
    Else
        sDefaultFolder = ThisWorkbook.Path
    End If
    
    sDefaultFolder = CheckPath(sDefaultFolder)
    
    ChDrive fso.GetDriveName(sDefaultFolder)
    ChDir sDefaultFolder
    sOut = Application.GetSaveAsFilename(InitialFileName:=sDefaultFile _
                        , filefilter:=asFileFilters _
                        , Title:=IIf(Len(Trim(asTitle)) > 0, asTitle, sDefaultFile))
    If sOut = False Then sOut = ""
    SelectSaveAsFileDialog = sOut
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

Function GetFileParentFolder(asFileFullPath As String) As String
    GetFSO
    GetFileParentFolder = fso.GetParentFolderName(asFileFullPath)
End Function

Function GetFileBaseName(asFileFullPath As String) As String
    GetFSO
    GetFileBaseName = fso.GetFileName(asFileFullPath)
End Function
Function GetFileNetName(asFileFullPath As String) As String
    GetFSO
    GetFileNetName = fso.GetBaseName(asFileFullPath)
End Function
Function GetFileExtension(asFileFullPath As String, Optional bDot As Boolean = False) As String
    GetFSO
    GetFileExtension = IIf(bDot, ".", "") & fso.GetExtensionName(asFileFullPath)
End Function

'Function GetFileNamePart(asFileFullPath As String _
'                        , Optional ByRef sParentFolder As String _
'                        , Optional ByRef sFileBaseName As String _
'                        , Optional ByRef sFileExtension As String _
'                        , Optional ByRef sFileNetName As String) As String
'    If Len(Trim(asFileFullPath)) <= 0 Then Exit Function
'
'    sParentFolder = fso.GetParentFolderName(asFileFullPath)
'    sFileBaseName = fso.GetFileName(asFileFullPath)
'    sFileExtension = fso.GetExtensionName(asFileFullPath)
'    sFileNetName = fso.GetBaseName(asFileFullPath)
'
'    Set fso = Nothing
'End Function

Function ArrayHasBlankValue(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = False
    If ArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    Dim lEachCol As Long
    
    iDimension = GetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 2 Then RaiseErr "2 dimensions is not supported."
    
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If Len(Trim(CStr(arrParam(lEachRow)))) <= 0 Then
            bOut = True
            GoTo exit_function
        End If
    Next
    
exit_function:
    ArrayHasBlankValue = bOut
End Function
Function ArrayHasDuplicateElement(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = False
    If ArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    
    iDimension = GetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 2 Then RaiseErr "2 dimensions is not supported."
    
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
    Set dict = Nothing
    ArrayHasDuplicateElement = bOut
End Function

Function ArrayIsEmptyOrNoData(ByRef arrParam) As Boolean
    Dim bOut As Boolean
    
    bOut = True
    If ArrayIsEmpty(arrParam) Then GoTo exit_function
    
    Dim iDimension As Integer
    Dim lEachRow As Long
    Dim lEachCol As Long
    
    iDimension = GetArrayDimension(arrParam)
    If iDimension <= 0 Then GoTo exit_function
    If iDimension >= 3 Then RaiseErr "3 dimensions is not supported."
    
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
    ArrayIsEmptyOrNoData = bOut
End Function

Function ArrayIsEmpty(ByRef arrParam) As Boolean
    Dim i As Long
    
    ArrayIsEmpty = True
    
    On Error Resume Next
    
    i = UBound(arrParam, 1)
    If Err.Number = 0 Then
        If UBound(arrParam) < LBound(arrParam) Then
            Exit Function
        Else
            ArrayIsEmpty = False
        End If
    Else
        Err.Clear
    End If
End Function
Function GetArrayDimension(arrParam) As Integer
    Dim i As Integer
    Dim tmp As Long
    
    On Error GoTo error_exit
    
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(arrParam, i)
        
        If tmp < 0 Then
            GetArrayDimension = -1
            Exit Function
        End If
    Loop
    
error_exit:
    Err.Clear
    GetArrayDimension = i - 1
End Function

'Function Num2Letter(ByVal alNum As Long) As String
'    Num2Letter = Replace(Split(Columns(alNum).Address, ":")(1), "$", "")
'End Function
Function Num2LetterV1(ByVal alNum As Long) As String
    Dim n As Long
    Dim c As Byte
    Dim s As String
    
    n = alNum
    Do
        c = (n - 1) Mod 26
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    
    Num2LetterV1 = s
End Function
Function Num2Letter(ByVal alNum As Long) As String
    Dim lAlpha As Long
    Dim lRemainder As Long
    
    If alNum <= 26 Then
        Num2Letter = Chr(alNum + 64)
    Else
        lRemainder = alNum Mod 26
        lAlpha = Int(alNum / 26)
        
        If lRemainder = 0 Then
            lRemainder = 26
            lAlpha = lAlpha - 1
        End If
        Num2Letter = Num2Letter(lAlpha) & Chr(lRemainder + 64)
    End If
End Function
'Function Letter2Num(ByVal alLetter As String) As Long
'    Letter2Num = Columns(alLetter).Column
'End Function
Function Letter2Num(ByVal alLetter As String) As Long
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
    
    Letter2Num = iOut
End Function

Function FileExists(sFilePath As String) As Boolean
    GetFSO
    FileExists = fso.FileExists(sFilePath)
End Function

Function DeleteFile(sFilePath As String)
    If FileExists(sFilePath) Then
        SetAttr sFilePath, vbNormal
        Kill sFilePath
    End If
End Function
 
Function ArrayRowIsBlankHasNoData(arr, alRow As Long, Optional iTillCol As Integer = 0) As Boolean
    Dim bOut As Boolean
    Dim lEachCol As Long
    
    If iTillCol <= 0 Then iTillCol = UBound(arr, 2)
    
    bOut = True
    For lEachCol = LBound(arr, 2) To iTillCol
        If Len(Trim(CStr(arr(alRow, lEachCol)))) > 0 Then
            bOut = False
            Exit For
        End If
    Next
    
    ArrayRowIsBlankHasNoData = bOut
End Function
Function OneRowOfArrayIsBlank(arr, alRow As Long, Optional iTillCol As Integer = 0) As Boolean
    OneRowOfArrayIsBlank = ArrayRowIsBlankHasNoData(arr, alRow, iTillCol)
End Function
Function GenRandomUniqueString() As String
    GenRandomUniqueString = Format(Now(), "yyyymmddhhMMSS") & Rnd()
End Function

Function Split_(asOrig As String, Optional asSeparators As String = "") As Variant
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
    
    sTransFormed = Trim_(sTransFormed, tDelimiter)
    
    Split_ = Split(sTransFormed, tDelimiter)
End Function

Function SplitJoin(asOrig As String, Optional asSeparators As String = "", Optional asNewSep As String = "|") As String
    If Len(asSeparators) <= 0 Then asSeparators = ":;|, " & vbLf
    
    Dim arr
    arr = Split_(asOrig, asSeparators)
    SplitJoin = Join(arr, asNewSep)
    
    Erase arr
End Function

Function Join_(asOrig As String, Optional asNewSep As String = "|") As String
    Join_ = SplitJoin(asOrig, , asNewSep)
End Function

Function Trim_(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Left(sOut, 1)) > 0
        sOut = Right(sOut, Len(sOut) - 1)
    Wend
    
    While InStr(asWhatToTrim, Right(sOut, 1)) > 0
        sOut = Left(sOut, Len(sOut) - 1)
    Wend
    
    Trim_ = sOut
End Function

Function RTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Right(sOut, 1)) > 0
        sOut = Left(sOut, Len(sOut) - 1)
    Wend
    
    RTrim = sOut
End Function

Function LTrim(asOrig As String, Optional asWhatToTrim As String = " " & vbTab & vbCr & vbLf) As String
    Dim sOut As String
    
    sOut = Trim(asOrig)
    While InStr(asWhatToTrim, Left(sOut, 1)) > 0
        sOut = Right(sOut, Len(sOut) - 1)
    Wend
    
    LTrim = sOut
End Function

Function Len_(sStr) As Long
    Len_ = Len(Trim(sStr))
End Function

Function Zero(sStr) As Boolean
    Zero = (Len(Trim(sStr)) <= 0)
End Function

Function Nzero(sStr) As Boolean
    Nzero = CBool(Len(Trim(sStr)) > 0)
End Function

Private Function GetParamType(arrParam, Optional bThrowError As Boolean = False, Optional sMsgPrefix As String = "", Optional sMsgSuffix As String = "") As enParamType
    GetParamType = enParamType.None
    
    If IsArray(arrParam) Then
        If ArrayIsEmptyOrNoData(arrParam) Then
            GetParamType = enParamType.ArrayButEmpty
            If bThrowError Then RaiseErr sMsgPrefix & " is blank " & IIf(sMsgSuffix = "", "", " passed to " & sMsgSuffix)
        ElseIf ArrayHasDuplicateElement(arrParam) Then
            GetParamType = enParamType.ArrayButHasDuplicateElement
            If bThrowError Then RaiseErr sMsgPrefix & " has duplicate element " & IIf(sMsgSuffix = "", "", " passed to " & sMsgSuffix)
        ElseIf UBound(arrParam) = LBound(arrParam) Then
            GetParamType = enParamType.ArrayButOnlyOneElement
        Else
            GetParamType = enParamType.ArrayParam
        End If
    Else
        GetParamType = enParamType.StringParam
        If Len(Trim(CStr(arrParam))) <= 0 Then
            If bThrowError Then RaiseErr sMsgPrefix & " is blank " & IIf(sMsgSuffix = "", "", " passed to " & sMsgSuffix)
        End If
    End If
End Function
Private Function JoinOneRowOfArray(arrData(), aiRow As Long, Optional arrCols, Optional paramType As enParamType = enParamType.None, Optional sDelimiter As String = "|") As String
'===================================================================
'arrCols:
'     not provided:   all columns
'     single number:  only one coumn
'     aarray(1,2,3):  1, 2, 3
'===================================================================
    Dim sOut As String
    Dim j As Integer
    
    If ArrayIsEmpty(arrData) Then GoTo exit_func
    
    'Dim paramType As enParamType
    
    sOut = ""
    If IsMissing(arrCols) Then
        For j = LBound(arrData, 2) To UBound(arrData, 2)
            sOut = sOut & sDelimiter & Trim(CStr(arrData(aiRow, j)))
        Next
    Else
        If paramType = enParamType.None Then paramType = GetParamType(arrCols, True, "arrCols", "JoinOneRowOfArray")
        
        If paramType = StringParam Then
            sOut = Trim(CStr(arrData(aiRow, arrCols)))
        ElseIf paramType = ArrayButOnlyOneElement Then
            sOut = Trim(CStr(arrData(aiRow, arrCols(LBound(arrCols)))))
        Else
            For j = LBound(arrCols) To UBound(arrCols)
                sOut = sOut & sDelimiter & Trim(CStr(arrData(aiRow, arrCols(j))))
            Next
        End If
    End If
    
    If Len(Trim(Replace(sOut, sDelimiter, ""))) <= 0 Then sOut = "": GoTo exit_func
    
    If Len(sDelimiter) > 0 Then sOut = Right(sOut, Len(sOut) - 1)
    
exit_func:
    JoinOneRowOfArray = sOut
End Function
Function CopyOneRowOfArray2OneDimArray(arrData(), alRow As Long, Optional arrCols, Optional paramType As enParamType = enParamType.None, Optional iLBoundAs0or1 As Integer = 1)
'===================================================================
'arrCols:
'     not provided:   all columns
'     single number:  only one coumn
'     aarray(1,2,3):  1, 2, 3
'===================================================================
    Dim arrOut()
    
    If iLBoundAs0or1 <> 0 And iLBoundAs0or1 <> 1 Then RaiseErr "iLBoundAs0or1 can only be 0 or 1"
     
    
    If ArrayIsEmpty(arrData) Then GoTo exit_fun
     
    Dim j As Integer
    Dim iSourceDim As Integer
     
    If IsMissing(arrCols) Then
        iSourceDim = LBound(arrData, 2)
        'ReDim arrOut(iLBoundAs0or1 To (UBound(arrData, 2) - LBound(arrData, 2) + 1 + iLBoundAs0or1 - 1))
        ReDim arrOut(iLBoundAs0or1 To (UBound(arrData, 2) + iLBoundAs0or1 - iSourceDim))
    
        For j = LBound(arrData, 2) To UBound(arrData, 2)
            arrOut(iLBoundAs0or1 + j - iSourceDim) = arrData(alRow, j)
        Next
    Else
        If paramType = enParamType.None Then paramType = GetParamType(arrCols, True, "arrCols", "CopyOneRowOfArray2OneDimArray")
        
        If paramType = StringParam Then
            arrOut(iLBoundAs0or1) = arrData(alRow, arrCols)
        ElseIf paramType = ArrayButOnlyOneElement Then
            arrOut(iLBoundAs0or1) = arrData(alRow, arrCols(LBound(arrCols)))
        Else
            iSourceDim = LBound(arrCols)
            ReDim arrOut(iLBoundAs0or1 To (UBound(arrCols) + iLBoundAs0or1 - iSourceDim))
        
            For j = LBound(arrCols) To UBound(arrCols)
                arrOut(iLBoundAs0or1 + j - iSourceDim) = arrData(alRow, arrCols(j))
            Next
        End If
    End If
exit_fun:
    CopyOneRowOfArray2OneDimArray = arrOut
End Function

Function ConvertArray2DictionaryMultipleKeysWithMultipleCols(arrData(), arrKeyCols, arrItemCols _
                , Optional sheetForMsgBox As Worksheet = Nothing _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True _
                , Optional asKeysDelimiter As String = "|") As Dictionary
'==================================================================================================
'arrKeyCols:     either array or integer
'
'arrItemCols:    either array or integer
'         -1:    only row number
'          0:    only keys
'        > 0:    single value
'      array:    array()
'item value:     array lbound starts from 1
'==================================================================================================
    Dim dictOut As Dictionary
    Dim sht
    
    Set dictOut = New Dictionary
    
    If ArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    
    Dim keyType As enParamType
    Dim itemType As enParamType
    Dim bOnlyKeys As Boolean
    Dim bOnlyRowNum As Boolean
    
    keyType = GetParamType(arrKeyCols, True, "arrKeyCols", "ConvertArray2DictionaryMultipleKeysWithMultipleCols")
    itemType = GetParamType(arrItemCols, True, "arrItemCols", "ConvertArray2DictionaryMultipleKeysWithMultipleCols")
    
    If keyType = StringParam Then
        If arrKeyCols < 0 Then RaiseErr "arrKeyCols <= 0 to ConvertArray2DictionaryMultipleKeysWithMultipleCols"
    ElseIf keyType = ArrayButOnlyOneElement Then
        If arrKeyCols(LBound(arrKeyCols)) < 0 Then RaiseErr "arrKeyCols <= 0 to ConvertArray2DictionaryMultipleKeysWithMultipleCols"
    End If
       
    If itemType = StringParam Then
        If arrItemCols <= -2 Then RaiseErr "arrItemCols <= -2 to ConvertArray2DictionaryMultipleKeysWithMultipleCols"
        
        bOnlyKeys = CBool(arrItemCols = 0)
        bOnlyRowNum = CBool(arrItemCols = -1)
    ElseIf itemType = ArrayButOnlyOneElement Then
        If arrItemCols(LBound(arrItemCols)) <= -2 Then RaiseErr "arrItemCols <= -2 to ConvertArray2DictionaryMultipleKeysWithMultipleCols"
        
        bOnlyKeys = CBool(arrItemCols(LBound(arrItemCols)) = 0)
        bOnlyRowNum = CBool(arrItemCols(LBound(arrItemCols)) = -1)
    End If
    
    Set sht = sheetForMsgBox
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    Dim sItemsValue As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = JoinOneRowOfArray(arrData, i, arrKeyCols, keyType, asKeysDelimiter)
'
'        If keyType = StringParam Then
'            sKeyStr = Trim(CStr(arrData(i, arrKeyCols)))
'        ElseIf keyType = ArrayButOnlyOneElement Then
'            sKeyStr = Trim(CStr(arrData(i, arrKeyCols(LBound(arrKeyCols)))))
'        Else
'            'sKeyStr = JoinOneRowOfArray(arrData, i, arrKeyCols, asKeysDelimiter)
'            For j = LBound(arrKeyCols) To UBound(arrKeyCols)
'                sKeyStr = sKeyStr & sDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
'            Next
'        End If
'
'        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
'            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
'        Next
        
        If Len(sKeyStr) <= 0 Then
            If Not IgnoreBlankKeys Then
                If Not OneRowOfArrayIsBlank(arrData, i) Then RaiseErr "Key columns is blank at row " & (sht.headerbyrow) & " in sheet [" & sheetForMsgBox.Name & "]"
            End If
            
            GoTo next_row
        End If
        
        If dictOut.Exists(sKeyStr) Then
            If WhenKeyDuplicateThenError Then
                ActivateSheet sheetForMsgBox
                RaiseErr "Duplicate key was found at row " & (sht.headerbyrow) & " in sheet [" & sheetForMsgBox.Name & "]" & vbCr & sKeyStr
            End If
            GoTo next_row
        End If
        
        If bOnlyRowNum Then
            dictOut.Add sKeyStr, i
        Else
            If bOnlyKeys Then
                dictOut.Add sKeyStr, i
            Else
                If itemType = StringParam Then
                    dictOut.Add sKeyStr, arrData(i, arrItemCols)
                ElseIf itemType = ArrayButOnlyOneElement Then
                    dictOut.Add sKeyStr, arrData(i, arrItemCols(LBound(arrItemCols)))
                Else
                    dictOut.Add sKeyStr, CopyOneRowOfArray2OneDimArray(arrData, i, 1)
                End If
            End If
        End If
next_row:
    Next
    
exit_fun:
    Set ConvertArray2DictionaryMultipleKeysWithMultipleCols = dictOut
    Set dictOut = Nothing
End Function

Function ConvertArray2DictionaryWithRowNum(arrData(), arrKeyCols _
                , Optional sheetForMsgBox As Worksheet = Nothing _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True _
                , Optional asKeysDelimiter As String = "|") As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    Set ConvertArray2DictionaryWithRowNum = ConvertArray2DictionaryMultipleKeysWithMultipleCols(arrData, arrKeyCols, -1 _
                , sheetForMsgBox, IgnoreBlankKeys, WhenKeyDuplicateThenError, asKeysDelimiter)
End Function
Function ConvertArray2DictionaryWithKeysOnly(arrData(), arrKeyCols _
                , Optional sheetForMsgBox As Worksheet = Nothing _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True _
                , Optional asKeysDelimiter As String = "|") As Dictionary
    Set ConvertArray2DictionaryWithKeysOnly = ConvertArray2DictionaryMultipleKeysWithMultipleCols(arrData, arrKeyCols, 0 _
                , sheetForMsgBox, IgnoreBlankKeys, WhenKeyDuplicateThenError, asKeysDelimiter)
End Function
  
Function ConvertArray2DictionaryWithSingleCol(arrData(), iKeyCol, iItemCol _
                , Optional sheetForMsgBox As Worksheet = Nothing _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True _
                , Optional asKeysDelimiter As String = "|") As Dictionary
'==========================================================================
'lItemCol
'         -1: the item is row number
'          0: get key only, not care the item value, 0 as default
'         >0: the item is specified column
'==========================================================================
    Set ConvertArray2DictionaryWithSingleCol = ConvertArray2DictionaryMultipleKeysWithMultipleCols(arrData, iKeyCol, iItemCol _
                , sheetForMsgBox, IgnoreBlankKeys, WhenKeyDuplicateThenError, asKeysDelimiter)
End Function
 
'Private Function ReadArray2Dictionary(arrParam, lKeyCol As Long _
'                            , Optional lItemCol As Long = 0 _
'                            , Optional IgnoreBlankKeys As Boolean = False _
'                            , Optional WhenKeyIsDuplicateError As Boolean = True) As Dictionary
''==========================================================================
''lItemCol
''         -1: the item is row number
''          0: get key only, not care the item value, 0 as default
''         >0: the item is specified column
''==========================================================================
'    If lItemCol < -1 Or lKeyCol <= 0 Then RaiseErr "wrong param"
'
'    Dim dictOut As Dictionary
'
'    Set dictOut = New Dictionary
'    If ArrayIsEmptyOrNoData(arrParam) Then GoTo exit_fun
'
'    Dim bGetKeyOnly As Boolean
'    Dim bGetRowNo As Boolean
'
'    bGetKeyOnly = (lItemCol = 0)
'    bGetRowNo = (lItemCol = -1)
'
'    Dim i As Long
'    Dim sKey As String
'    Dim sValue As String
'
'    For i = LBound(arrParam, 1) To UBound(arrParam, 1)
'        If ArrayRowIsBlankHasNoData(arrParam, i) Then GoTo next_row
'
'        sKey = Trim(arrParam(i, lKeyCol))
'
'        If Len(sKey) <= 0 Then
'            If Not IgnoreBlankKeys Then
'                RaiseErr "Key col is blank, but you specified IgnoreBlankKeys = false" & vbCr & lKeyCol
'            Else
'                GoTo next_row
'            End If
'        End If
'
'        If dictOut.Exists(sKey) Then
'            If WhenKeyIsDuplicateError Then
'                RaiseErr "duplicate key was found:, but you specified WhenKeyIsDuplicateError = false" & vbCr & lKeyCol & vbCr & sKey
'            Else
'                GoTo next_row
'            End If
'        End If
'
'        If bGetRowNo Then
'            dictOut.Add sKey, i
'        Else
'            If bGetKeyOnly Then
'                dictOut.Add sKey, 0
'            Else
'                dictOut.Add sKey, arrParam(i, lItemCol)
'            End If
'        End If
'next_row:
'    Next
'
'exit_fun:
'    Set ReadArray2Dictionary = dictOut
'    Set dictOut = Nothing
'End Function

Function ValidateDuplicateInArray(arrParam, arrKeyColsOrSingle _
                        , Optional bAllowBlank As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional ByVal sMsgColHeader As String)
'arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
    If ArrayIsEmptyOrNoData(arrParam) Then Exit Function
    
    If IsArray(arrKeyColsOrSingle) Then
        If ArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then RaiseErr "Wrong param: arrKeyColsOrSingle"
    Else
        If arrKeyColsOrSingle <= 0 Then RaiseErr "Wrong param: arrKeyColsOrSingle"
    End If
    
    If IsArray(arrKeyColsOrSingle) Then
        Call ValidateDuplicateInArrayForCombineCols(arrParam:=arrParam, arrKeyCols:=arrKeyColsOrSingle _
                                                    , bAllowBlankIgnore:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    Else
        Call ValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
                                                    , bAllowBlankIgnore:=bAllowBlank _
                                                    , shtAt:=shtAt _
                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
                                                    , sMsgColHeader:=sMsgColHeader)
    End If
End Function

Function ValidateDuplicateInArrayForCombineCols(arrParam, arrKeyCols _
                        , Optional bAllowBlankIgnore As Boolean = False _
                        , Optional shtAt As Worksheet _
                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                        , Optional sMsgColHeader As String)
'MultipleCols: means MultipleCols composed as key
'for MultipleCols that is individually, please refer to Function ValidateDuplicateInArrayIndividually
    Const DELI = " " & "|" & " "
    
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetterStr As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    If Not Zero(sMsgColHeader) Then
        sColLetterStr = sMsgColHeader
    Else
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sColLetterStr = sColLetterStr & " + " & Num2Letter(lStartCol + lEachCol - 1)
        Next
        sColLetterStr = Right(sColLetterStr, Len(sColLetterStr) - 3)
    End If
    
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO, [" & sColLetterStr & "]"
            
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If ArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = ""
        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            lEachCol = arrKeyCols(i)
            sKeyStr = sKeyStr & DELI & Trim(CStr(arrParam(lEachRow, lEachCol)))
        Next
        
        If Zero(Replace(sKeyStr, DELI, "")) Then
            If Not bAllowBlankIgnore Then
                'sPos = sPos & "[" & lActualRow & ", " & sColLetterStr & "]"
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                RaiseErr "Keys [" & sKeyStr & "] is blank!" & sPos
            End If
            
            GoTo next_row
        End If
        
        sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(DELI))
        
        If dict.Exists(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            ShowSheet shtAt
            Application.Goto shtAt.Cells(lActualRow, arrKeyCols(UBound(arrKeyCols)))
            RaiseErr "Duplicate key was found:" & vbCr & sKeyStr & vbCr & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

Function ReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrData, arrKeyCols, lItemCol As Long _
                , Optional asKeysDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False _
                , Optional WhenKeyDuplicateThenError As Boolean = True) As Dictionary
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    
    If ArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If ArrayIsEmptyOrNoData(arrKeyCols) Then RaiseErr "arrKeyCols is empty !"
    If ArrayHasDuplicateElement(arrKeyCols) Then RaiseErr "arrKeyCols has duplicate element"
    If lItemCol < 0 Then RaiseErr "lItemCol < 0 to ReadArray2DictionaryWithMultipleKeyColsSingleItemCol"
    
    If InStr(asKeysDelimiter, " ") > 0 Then RaiseErr "asKeysDelimiter cannot be space or contains space"
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = ""
        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
        Next
        
        If Zero(Replace(sKeyStr, asKeysDelimiter, "")) Then
            If Not IgnoreBlankKeys Then RaiseErr "IgnoreBlankKeys is false, but a keystr is blank"
            GoTo next_row
        End If
        
        If Len(asKeysDelimiter) > 0 Then sKeyStr = Right(sKeyStr, Len(sKeyStr) - Len(asKeysDelimiter))
        
        If dictOut.Exists(sKeyStr) Then
            If WhenKeyDuplicateThenError Then
                RaiseErr "Duplicate key was found " & vbCr & sKeyStr
            End If
            GoTo next_row
        End If
        
        dictOut.Add sKeyStr, arrData(i, lItemCol)
next_row:
    Next
    
exit_fun:
    Set ReadArray2DictionaryWithMultipleKeyColsSingleItemCol = dictOut
    Set dictOut = Nothing
End Function

Function ReadArray2DictionaryWithMultipleKeyColsSingleItemColSum(arrData, arrKeyCols, lItemCol As Long _
                , Optional asKeysDelimiter As String = "" _
                , Optional IgnoreBlankKeys As Boolean = False) As Dictionary
    Dim dictOut As Dictionary
    
    Set dictOut = New Dictionary
    
    If ArrayIsEmptyOrNoData(arrData) Then GoTo exit_fun
    If ArrayIsEmptyOrNoData(arrKeyCols) Then RaiseErr "arrKeyCols is empty !"
    If ArrayHasDuplicateElement(arrKeyCols) Then RaiseErr "arrKeyCols has duplicate element"
    If lItemCol < 0 Then RaiseErr "lItemCol < 0 to ReadArray2DictionaryWithMultipleKeyColsSingleItemCol"
    
    If InStr(asKeysDelimiter, " ") > 0 Then RaiseErr "asKeysDelimiter cannot be space or contains space"
    
    Dim i As Long
    Dim j As Integer
    Dim sKeyStr As String
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sKeyStr = ""
        For j = LBound(arrKeyCols) To UBound(arrKeyCols)
            sKeyStr = sKeyStr & asKeysDelimiter & Trim(CStr(arrData(i, arrKeyCols(j))))
        Next
        
        If Zero(Replace(sKeyStr, asKeysDelimiter, "")) Then
            If Not IgnoreBlankKeys Then RaiseErr "IgnoreBlankKeys is false, but a keystr is blank"
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
    Set ReadArray2DictionaryWithMultipleKeyColsSingleItemColSum = dictOut
    Set dictOut = Nothing
End Function
Function ValidateDuplicateInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional bAllowBlankIgnore As Boolean = False _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then RaiseErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetter As String
    Dim dict As Dictionary
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(Zero(sMsgColHeader), Num2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"
         
    Set dict = New Dictionary
    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If ArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = Trim(CStr(arrParam(lEachRow, lKeyCol)))
        
        If Zero(sKeyStr) Then
            If Not bAllowBlankIgnore Then
                'sPos = sPos & lActualRow & " / " & sColLetter
                sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
                ShowSheet shtAt
                Application.Goto shtAt.Cells(lActualRow, lKeyCol)
                RaiseErr "Keys [" & sColLetter & "] is blank!" & sPos
            End If
            
            GoTo next_row
        End If
        
        If dict.Exists(sKeyStr) Then
            'sPos = sPos & lActualRow & " / " & sColLetter
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            ShowSheet shtAt
            Application.Goto shtAt.Cells(lActualRow, lKeyCol)
            RaiseErr "Duplicate key [" & sKeyStr & "] was found " & sPos
        Else
            dict.Add sKeyStr, 0
        End If
next_row:
    Next
    
    Set dict = Nothing
End Function

'Function ValidateDuplicateInArrayIndividually(arrParam, arrKeyColsOrSingle _
'                        , Optional bAllowBlank As Boolean = False _
'                        , Optional shtAt As Worksheet _
'                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
'                        , Optional arrColNames)
''arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
'    If ArrayIsEmptyOrNoData(arrParam) Then Exit Function
'
'    If IsArray(arrKeyColsOrSingle) Then
'        If ArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    Else
'        If arrKeyColsOrSingle <= 0 Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    End If
'
'    Dim i As Long
'    If IsArray(arrKeyColsOrSingle) Then
'        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
'            Call ValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
'                                                    , bAllowBlank:=bAllowBlank _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , arrColNames:=arrColNames)
'        Next
'    Else
'        Call ValidateDuplicateInArrayForSingleCol(arrParam:=arrParam, arrKeyColsOrSingle:=arrKeyColsOrSingle _
'                                                    , bAllowBlank:=bAllowBlank _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , arrColNames:=arrColNames)
'    End If
'End Function


Function DeleteSheetIfExists(asShtName As String, Optional wb As Workbook)
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    If SheetExists(asShtName, , wb) Then
        Call DeleteSheet(asShtName, wb)
    End If
End Function

Function DeleteSheet(asShtName As String, Optional wb As Workbook)
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

Function AddNewSheet(asShtName As String, Optional wb As Workbook) As Worksheet
    Dim shtOut As Worksheet
    
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Set shtOut = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    shtOut.Name = asShtName
    shtOut.Activate
    ActiveWindow.DisplayGridlines = False
    
    Set AddNewSheet = shtOut
    Set shtOut = Nothing
End Function

Function AddNewSheetDeleteFirst(asShtName As String, Optional wb As Workbook) As Worksheet
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    Call DeleteSheetIfExists(asShtName, wb)
    Set AddNewSheetDeleteFirst = AddNewSheet(asShtName, wb)
End Function

Function GetSheetWhenNotExistsCreate(asShtName As String, Optional wb As Workbook) As Worksheet
    asShtName = Trim(asShtName)
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    If SheetExists(asShtName, , wb) Then
        Set GetSheetWhenNotExistsCreate = wb.Worksheets(asShtName)
    Else
        Set GetSheetWhenNotExistsCreate = AddNewSheet(asShtName, wb)
    End If
End Function

Function GetFSO()
    If fso Is Nothing Then Set fso = New FileSystemObject
End Function

Function DeleteAllFilesInFolder(sFolder As String)
    GetFSO
    
    Dim aFile As File
    
    If fso.FolderExists(sFolder) Then
        For Each aFile In fso.GetFolder(sFolder).Files
            aFile.Delete True
        Next
    End If
End Function
 
'
'Function ValidateBlankInArrayCombinedCols(arrParam, arrKeyColsOrSingle _
'                        , Optional shtAt As Worksheet _
'                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
'                        , Optional sMsgColHeader As String)
''arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
'    If ArrayIsEmptyOrNoData(arrParam) Then Exit Function
'
'    If IsArray(arrKeyColsOrSingle) Then
'        If ArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    Else
'        If arrKeyColsOrSingle <= 0 Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    End If
'
'    If IsArray(arrKeyColsOrSingle) Then
'        Call fValidateBlankInArrayForCombineCols(arrParam:=arrParam, arrKeyCols:=arrKeyColsOrSingle _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'    Else
'        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'    End If
'End Function
'
''call the parent function: fValidateBlankInArrayCombinedCols, not to call this function
'Private Function ValidateBlankInArrayForCombineCols(arrParam, arrKeyCols _
'                        , Optional shtAt As Worksheet _
'                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
'                        , Optional sMsgColHeader As String)
''MultipleCols: means MultipleCols composed as key
''for MultipleCols that is individually, please refer to Function ValidateBlankInArrayIndividually
'    Const DELI = " " & DELIMITER & " "
'
'    Dim lEachRow As Long
'    Dim lEachCol As Long
'    Dim i As Long
'    Dim sKeyStr As String
'    Dim sColLetterStr As String
'    Dim sPos As String
'    Dim lActualRow As Long
'
'    If Not Zero(sMsgColHeader) Then
'        sColLetterStr = sMsgColHeader
'    Else
'        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
'            lEachCol = arrKeyCols(i)
'            sColLetterStr = sColLetterStr & " + " & Num2Letter(lStartCol + lEachCol - 1)
'        Next
'        sColLetterStr = Right(sColLetterStr, Len(sColLetterStr) - 3)
'    End If
'
'    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
'         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO, [" & sColLetterStr & "]"
'
'    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
'        If ArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
'
'        lActualRow = (lHeaderAtRow + lEachRow)
'
'        sKeyStr = ""
'        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
'            lEachCol = arrKeyCols(i)
'            sKeyStr = sKeyStr & CStr(arrParam(lEachRow, lEachCol))
'        Next
'
'        If Zero(sKeyStr) Then
'            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
'            RaiseErr "Keys [" & sKeyStr & "] is blank!" & sPos
'        End If
'next_row:
'    Next
'End Function

Function ValidateBlankInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then RaiseErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim sColLetter As String
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(Zero(sMsgColHeader), Num2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"

    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If ArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = CStr(arrParam(lEachRow, lKeyCol))
    
        If Zero(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            RaiseErr "Keys [" & sColLetter & "] is blank!" & sPos
        End If
next_row:
    Next
End Function
'
'Function ValidateBlankInArray(arrParam, arrKeyColsOrSingle _
'                        , Optional shtAt As Worksheet _
'                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
'                        , Optional ByVal sMsgColHeader As String)
''arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
'    If ArrayIsEmptyOrNoData(arrParam) Then Exit Function
'
'    If IsArray(arrKeyColsOrSingle) Then
'        If ArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    Else
'        If arrKeyColsOrSingle <= 0 Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    End If
'
'    Dim i As Integer
'    If IsArray(arrKeyColsOrSingle) Then
'        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
'            Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'        Next
'    Else
'        Call fValidateBlankInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'    End If
'End Function
'
'Function ValidateNumericColInArray(arrParam, arrKeyColsOrSingle _
'                        , Optional shtAt As Worksheet _
'                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
'                        , Optional ByVal sMsgColHeader As String)
''arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
'    If ArrayIsEmptyOrNoData(arrParam) Then Exit Function
'
'    If IsArray(arrKeyColsOrSingle) Then
'        If ArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    Else
'        If arrKeyColsOrSingle <= 0 Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    End If
'
'    Dim i As Integer
'    If IsArray(arrKeyColsOrSingle) Then
'        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
'            Call fValidateNumericColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'        Next
'    Else
'        Call fValidateNumericColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'    End If
'End Function

Function ValidateNumericColInArrayForSingleCol(arrParam, lKeyCol As Long _
                                            , Optional shtAt As Worksheet _
                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
                                            , Optional sMsgColHeader As String)
    If lKeyCol <= 0 Then RaiseErr "Wrong param: lKeyCol"
    
    Dim lEachRow As Long
    Dim i As Long
    Dim sKeyStr 'As String
    Dim sColLetter As String
    Dim sPos As String
    Dim lActualRow As Long
    
    sColLetter = IIf(Zero(sMsgColHeader), Num2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
        
    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"

    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
        If ArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
        
        lActualRow = (lHeaderAtRow + lEachRow)
        
        sKeyStr = arrParam(lEachRow, lKeyCol)
    
        If Not IsNumeric(sKeyStr) Then
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            RaiseErr "Keys [" & sColLetter & "] is Not Numeric!" & sPos
        End If
next_row:
    Next
End Function
'
'Function ValidateDateColInArray(arrParam, arrKeyColsOrSingle _
'                        , Optional shtAt As Worksheet _
'                        , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
'                        , Optional ByVal sMsgColHeader As String)
''arrKeyColsOrSingle : should be start from 1, since two dimension array is starting from 1
'    If ArrayIsEmptyOrNoData(arrParam) Then Exit Function
'
'    If IsArray(arrKeyColsOrSingle) Then
'        If ArrayIsEmptyOrNoData(arrKeyColsOrSingle) Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    Else
'        If arrKeyColsOrSingle <= 0 Then RaiseErr "Wrong param: arrKeyColsOrSingle"
'    End If
'
'    Dim i As Integer
'    If IsArray(arrKeyColsOrSingle) Then
'        For i = LBound(arrKeyColsOrSingle) To UBound(arrKeyColsOrSingle)
'            Call fValidateDateColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'        Next
'    Else
'        Call fValidateDateColInArrayForSingleCol(arrParam:=arrParam, lKeyCol:=CLng(arrKeyColsOrSingle) _
'                                                    , shtAt:=shtAt _
'                                                    , lHeaderAtRow:=lHeaderAtRow, lStartCol:=lStartCol _
'                                                    , sMsgColHeader:=sMsgColHeader)
'    End If
'End Function
'
'Function ValidateDateColInArrayForSingleCol(arrParam, lKeyCol As Long _
'                                            , Optional shtAt As Worksheet _
'                                            , Optional lHeaderAtRow As Long = 1, Optional lStartCol As Long _
'                                            , Optional sMsgColHeader As String)
'    If lKeyCol <= 0 Then RaiseErr "Wrong param: lKeyCol"
'
'    Dim lEachRow As Long
'    Dim i As Long
'    Dim sKeyStr 'As String
'    Dim sColLetter As String
'    Dim sPos As String
'    Dim lActualRow As Long
'
'    sColLetter = IIf(Zero(sMsgColHeader), Num2Letter(lStartCol + lKeyCol - 1), sMsgColHeader)
'
'    sPos = vbCr & vbCr & "sheet     : " & shtAt.Name _
'         & vbCr & vbCr & "Row, Column: " & " ACTUAL_ROW_NO,  [" & sColLetter & "]"
'
'    For lEachRow = LBound(arrParam, 1) To UBound(arrParam, 1)
'        If ArrayRowIsBlankHasNoData(arrParam, lEachRow) Then GoTo next_row
'
'        lActualRow = (lHeaderAtRow + lEachRow)
'
'        sKeyStr = arrParam(lEachRow, lKeyCol)
'
'        If Not IsDate(sKeyStr) Then
'            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
'            RaiseErr "Keys [" & sColLetter & "] is Not Date!" & sPos
'        End If
'next_row:
'    Next
'End Function
Function EnlargeAray(ByRef arr, Optional aPreserve As Boolean = True, Optional lIncrementNum As Integer = 1) As Long
    Redim_ arr, ArrLen(arr) + 1, aPreserve
End Function


Function EnlargeArayWithValue(ByRef arr, aValue, Optional aPreserve As Boolean = True, Optional lIncrementNum As Integer = 1) As Long
'    If ArrayIsEmpty(arr) Then
'        Redim arr
'        Exit Function
'    End If
    
    Redim_ arr, ArrayLen(arr) + 1, aPreserve
    arr(UBound(arr)) = aValue
    EnlargeArayWithValue = UBound(arr)
End Function

Function Redim_(ByRef arr, lNewUbound As Long, Optional aPreserve As Boolean = True)
    If ArrayIsEmpty(arr) Then
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
    If ArrayIsEmpty(arr) Then
        ArrayLen = 0
        Exit Function
    '    RaiseErr "Empty array is not allowed."
    End If
     ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

Private Function Base0(ByRef arr) As Boolean
     Base0 = (LBound(arr) = 0)
End Function

'Function SetArrayValue(ByRef arr, aIndex As Long, aValue)
'    If Base0(arr) Then
'        arr(aIndex - 1) = aValue
'    Else
'        arr(aIndex) = aValue
'    End If
'End Function

Function UpdateDictionaryItemValueForDelimitedElement(ByRef dict As Dictionary, aKey, iElementIndex As Integer, aNewValue, Optional sDelimiter As String = "|")
    Dim arr
    If iElementIndex <= 0 Then RaiseErr "iElementIndex <= 0 to fUpdateDictionaryItemValueForDelimitedElement"
    
    If Not dict.Exists(aKey) Then RaiseErr "aKey even does not exists in param dict to fUpdateDictionaryItemValueForDelimitedElement"
    
    arr = Split(dict(aKey), sDelimiter)
    
    If ArrayLen(arr) < iElementIndex Then
        Redim_ arr, CLng(iElementIndex), True
    End If
    
    If Base0(arr) Then
        arr(iElementIndex - 1) = aNewValue
    Else
        arr(iElementIndex) = aNewValue
    End If
    
    dict(aKey) = Join(arr, sDelimiter)
    Erase arr
End Function

Function CopyDictionaryKeys2Array(dict As Dictionary, ByRef arrOut())
    If dict.Count <= 0 Then
        arrOut = Array()
    End If
    
    ReDim arrOut(1 To dict.Count)
    
    Dim i As Long
    Dim aKey
    
    i = 0
    For Each aKey In dict.Keys
        i = i + 1
        arrOut(i) = aKey
    Next
End Function
Function CopyDictionaryItemsArray(dict As Dictionary, ByRef arrOut())
    If dict.Count <= 0 Then
        arrOut = Array()
    End If
    
    ReDim arrOut(1 To dict.Count)
    
    Dim i As Long
    Dim aKey
    
    i = 0
    For Each aKey In dict.Items
        i = i + 1
        arrOut(i) = aKey
    Next
End Function

Function EnableExcelOptionsAll()
    Call EnableOrDisableExcelOptionsAll(True)
End Function

Function DisableExcelOptionsAll()
    Call EnableOrDisableExcelOptionsAll(False)
End Function
Private Function EnableOrDisableExcelOptionsAll(bValue As Boolean)
    Application.ScreenUpdating = bValue
    
    If Application.CutCopyMode = 0 Then Application.EnableEvents = bValue
    Application.DisplayAlerts = bValue
    If Application.CutCopyMode = 0 Then Application.AskToUpdateLinks = bValue
'    ThisWorkbook.CheckCompatibility = bValue
    
'    If bValue Then
'        If Application.CutCopyMode = 0 Then Application.Calculation = xlCalculationAutomatic
'    Else
'        If Application.CutCopyMode = 0 Then Application.Calculation = xlCalculationManual
'    End If
    
    Application.EnableEvents = bValue
End Function

Function GetRangeFromExternalAddress(asExternalAddr As String) As Range
    If Zero(asExternalAddr) Then RaiseErr "wrong param"
    asExternalAddr = UCase(Trim(asExternalAddr))
    
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
        RaiseErr "the address passed does not have the excel file name part ot the sheet name part"
    End If
    
    sWbName = Mid(asExternalAddr, lFileStart + 1, lFileEnd - lFileStart - 1)
    sShtName = Mid(asExternalAddr, lFileEnd + 1, lShtEnd - lFileEnd - 1)
    sNetAddr = Right(asExternalAddr, Len(asExternalAddr) - lShtEnd)
    
    sWbName = Replace(sWbName, "'", "")
    sShtName = Replace(sShtName, "'", "")
    'sNetAddr = ReplaceConvertR1C1ToA1(sNetAddr)
    If sNetAddr Like "R*C" Then
        Dim row As Long
        Dim col As Integer
        
        row = Split(Split(sNetAddr, "C")(0), "R")(1)
        col = Split(sNetAddr, "C")(1)
        sNetAddr = Cells(row, col).addreess(external:=False, ReferenceStyle:=xlA1)
    End If
    
    Dim wbOut As Workbook
    If ExcelFileIsOpen(sWbName, wbOut) Then
        Set GetRangeFromExternalAddress = wbOut.Worksheets(sShtName).Range(sNetAddr)
    Else
        RaiseErr "Excel file is not open, pls check your program."
    End If
    
    Set wbOut = Nothing
End Function

Function ExcelFileIsOpen(sExcelFileName As String, Optional wbOut As Workbook) As Boolean
    On Error Resume Next
    Set wbOut = Workbooks(GetFileBaseName(sExcelFileName))
    Err.Clear

    ExcelFileIsOpen = (Not wbOut Is Nothing)
End Function
Function ExactExcelFileIsopen(sExcelFileName As String, Optional ByRef wbOut As Workbook) As Boolean
    Dim bOut As Boolean

    bOut = False

    On Error Resume Next
    Set wbOut = Workbooks(GetFileBaseName(sExcelFileName))
    Err.Clear

    If wbOut Is Nothing Then GoTo exit_fun

    If UCase(wbOut.FullName) <> UCase(Trim(sExcelFileName)) Then
        Set wbOut = Nothing
        GoTo exit_fun
    Else
        bOut = True: GoTo exit_fun
    End If

exit_fun:
    ExactExcelFileIsopen = bOut
End Function
'Function ReplaceConvertR1C1ToA1(sR1C1Address As String) As String
'    GetRegExp
'
'    Dim matchColl As VBScript_RegExp_55.MatchCollection
'    Dim match As VBScript_RegExp_55.match
'
'    mRegExp.IgnoreCase = True
'    mRegExp.Pattern = "R(\d{1,})C(\d{1,})"
'
'    Set matchColl = mRegExp.Execute(sR1C1Address)
'
'    Dim sAddrNew As String
'    Dim lNextStart As Long
'    Dim sReplaced As String
'
'    sAddrNew = ""
'    lNextStart = 1
'
'    For Each match In matchColl
'        sReplaced = Num2Letter(CLng(match.SubMatches(1))) & match.SubMatches(0)
'
'        sAddrNew = sAddrNew & Mid(sR1C1Address, lNextStart, match.FirstIndex - lNextStart + 1)
'        sAddrNew = sAddrNew & sReplaced
'
'        lNextStart = match.FirstIndex + match.length + 1
'    Next
'
'    If lNextStart <= Len(sR1C1Address) Then
'        sAddrNew = sAddrNew & Mid(sR1C1Address, lNextStart, Len(sR1C1Address) - lNextStart + 1)
'    End If
'
'    Set match = Nothing
'    Set matchColl = Nothing
'
'    ReplaceConvertR1C1ToA1 = IIf(Zero(sAddrNew), sR1C1Address, sAddrNew)
'End Function



Function SortArayDesc(ByRef arr(), Optional UseQuickSort As Boolean = True)
    If Not UseQuickSort Then
        Call SortArrayBubbleSortDesc(arr)
    Else
        Call SortArrayQuickSortDesc(arr)
    End If
End Function

Function SortAray(ByRef arr(), Optional UseQuickSort As Boolean = True)
    If Not UseQuickSort Then
        Call SortArrayBubbleSort(arr)
    Else
        Call SortArrayQuickSort(arr)
    End If
End Function
Function SortArrayBubbleSortDesc(ByRef arr())
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

Function SortArrayBubbleSort(ByRef arr)
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
Function SortArrayQuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
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
    If plngLeft < lngLast Then SortArrayQuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then SortArrayQuickSort pvarArray, lngFirst, plngRight
End Function

Function SortArrayQuickSortDesc(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
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
    If plngLeft < lngLast Then SortArrayQuickSortDesc pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then SortArrayQuickSortDesc pvarArray, lngFirst, plngRight
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

Function ConvertMMM2Num(sMMM As String) As Integer
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
            RaiseErr "wrong param sMMM: " & vbCr & sMMM
    End Select
End Function

Function FileterTwoDimensionArray(arrSource(), lCol As Long, sValue) As Variant
    Dim arrOut()
    Dim arrQualifiedRows()
    Dim iCnt As Long
    Dim lEachRow As Long
    Dim iEachCol As Long
    Dim i As Long
    
    If Base0(arrSource) Then RaiseErr "base0(arrSource) is 0"
    
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
    FileterTwoDimensionArray = arrOut
    Erase arrOut
    
    Debug.Print "fFileterOutTwoDimensionArray: " & Format(Timer - start, "00:00")
End Function

Function FileterOutTwoDimensionArray(arrSource(), lCol As Long, sValue) As Variant
    Dim arrOut()
    Dim arrQualifiedRows()
    Dim iCnt As Long
    Dim lEachRow As Long
    Dim iEachCol As Long
    Dim i As Long
    
    If Base0(arrSource) Then RaiseErr "base0(arrSource) is 0"
    
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
    FileterOutTwoDimensionArray = arrOut
    Erase arrOut
    
    Debug.Print "fFileterOutTwoDimensionArray: " & Timer - start & vbCr & Format(Timer - start, "00:00")
End Function
Function Transpose1DimenArrayTo2DimenArrayVertically(arrParam) As Variant
    Dim i As Long
    Dim iNew As Long
    Dim arrOut()
    
    If ArrayIsEmptyOrNoData(arrParam) Then GoTo exit_fun
    
    If GetArrayDimension(arrParam) > 1 Then RaiseErr "1 dimension array is allowed."
    
    ReDim arrOut(1 To ArrLen(arrParam), 1 To 1)
    
    iNew = 0
    For i = LBound(arrParam) To UBound(arrParam)
        iNew = iNew + 1
        arrOut(iNew, 1) = arrParam(i)
    Next
    
exit_fun:
    Transpose1DimenArrayTo2DimenArrayVertically = arrOut
    Erase arrOut
End Function

 

Function ConvertDictionaryKeysTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional bSetDictToNothing As Boolean = True) As Variant
    Dim arrTmp
    Dim arrOut()
    
    arrTmp = dict.Keys
    If bSetDictToNothing Then Set dict = Nothing
    arrOut = Transpose1DimenArrayTo2DimenArrayVertically(arrTmp)
    Erase arrTmp
    
    ConvertDictionaryKeysTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function

Function ConvertDictionaryItemsTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional bSetDictToNothing As Boolean = True) As Variant
    Dim arrTmp
    Dim arrOut()
    
    arrTmp = dict.Items
    If bSetDictToNothing Then Set dict = Nothing
    arrOut = Transpose1DimenArrayTo2DimenArrayVertically(arrTmp)
    Erase arrTmp
    
    ConvertDictionaryItemsTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function

Function ConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional asDelimiter As String = "|" _
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
    
    Dim aKey
    i = 0
    For Each aKey In dict.Keys
        sEachLine = aKey
        arrEachLine = Split(sEachLine, asDelimiter)
        
        i = i + 1
        For j = LBound(arrEachLine) To UBound(arrEachLine)
            arrOut(i, j + 1) = arrEachLine(j)
        Next
    Next
    
    ConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function

Function ConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(ByRef dict As Dictionary, Optional asDelimiter As String = "|" _
            , Optional bSetDictToNothing As Boolean = True) As Variant
    Dim arrTmp
    Dim i As Long
    Dim j As Long
    Dim sEachLine As String
    Dim arrEachLine
    Dim arrOut()
    Dim aKey
    
    arrOut = Array()
    
    If dict.Count > 0 Then
        ReDim arrOut(1 To dict.Count, 1 To UBound(Split(dict.Items(0), asDelimiter)) - LBound(Split(dict.Items(0), asDelimiter)) + 1)
    End If
    
    i = 0
    For Each aKey In dict.Items
        sEachLine = aKey
        arrEachLine = Split(sEachLine, asDelimiter)
        
        i = i + 1
        For j = LBound(arrEachLine) To UBound(arrEachLine)
            arrOut(i, j + 1) = arrEachLine(j)
        Next
    Next
    
    ConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste = arrOut
    Erase arrOut
End Function
Function TrimArrayElement(ByRef arr)
    Dim i As Long
    Dim j As Long
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            arr(i, j) = Trim(arr(i, j))
        Next
    Next
End Function

Function GetRangeByStartEndPos(shtParam As Worksheet, alStartRow As Long, Optional alStartCol As Long = 1, Optional alEndRow As Long = 0, Optional alEndCol As Long = 0) As Range
    If alEndRow <= 0 Then alEndRow = GetMaxRow(shtParam)
    If alEndCol <= 0 Then alEndCol = GetMaxCol(shtParam)

    If alEndRow < alStartRow Then Set GetRangeByStartEndPos = Nothing: Exit Function
    If alEndCol < alStartCol Then Set GetRangeByStartEndPos = Nothing: Exit Function

    With shtParam
        Set GetRangeByStartEndPos = .Range(.Cells(alStartRow, alStartCol), .Cells(alEndRow, alEndCol))
    End With
End Function
Function ReadRangeDatatoArrayByStartEndPos(shtParam As Worksheet, alStartRow As Long, Optional alStartCol As Long = 1, Optional alEndRow As Long = 0, Optional alEndCol As Long = 0) As Variant
    Dim rg As Range
    
    Set rg = GetRangeByStartEndPos(shtParam, alStartRow, alStartCol, alEndRow, alEndCol)
    
    If rg Is Nothing Then
        ReadRangeDatatoArrayByStartEndPos = Array()
    Else
        ReadRangeDatatoArrayByStartEndPos = ReadRangeDataToArray(rg)
    End If
    Set rg = Nothing
End Function

Function ReadRangeDataToArray(rngParam As Range) As Variant
    Dim arrOut()

    If (rngParam.Rows.Count = 1 And rngParam.Columns.Count = 1) Then
        ReDim arrOut(1 To 1, 1 To 1)
        arrOut(1, 1) = rngParam.Value
    Else
        arrOut = rngParam.Value
    End If

    ReadRangeDataToArray = arrOut
    Erase arrOut
End Function
   
Function RemoveFilterForSheet(sht As Worksheet, Optional ByVal asDegree As String = "SHOW_ALL_DATA")
    asDegree = UCase(Trim(asDegree))

    Dim rgActiveCell As Range
    Set rgActiveCell = ActiveCell

    If sht.FilterMode Then  'advanced filter
        sht.ShowAllData
    End If

    If sht.AutoFilterMode Then  'auto filter
        If Zero(asDegree) Or asDegree = "SHOW_ALL_DATA" Then
            sht.AutoFilter.ShowAllData
        Else
            sht.AutoFilterMode = False
        End If
    End If

    rgActiveCell.Select
    Set rgActiveCell = Nothing

    'Call fGotoCell(sht.Range("A2"), True)
End Function
 
Public Function IsDate(sDateStr As String, Optional sFormat As String = "YYYYMMDD") As Boolean
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String
    
    IsDate = False
    
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
    If bSplit And Len(sFormat) <= 0 Then RaiseErr "The date has delimiter, but you did not specify the format:" & vbCr & "Date:" & sDateStr & vbCr & "Format:" & sFormat
    
    Select Case UCase(sFormat)
        Case "DDMMMYY", "DDMMMYYYY"
            sYear = Mid(sDateStr, 6)
            sMonth = ConvertMMM2Num(Mid(sDateStr, 3, 3))
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
            RaiseErr "sFormat is not covered in fIsDate, please change this function." & vbCr _
             & "sFormat: " & sFormat & vbCr _
             & "sDelimiter: " & sDelimiter & vbCr _
             & "sDateStr: " & sDateStr
    End Select
    
    On Error Resume Next
    Dim dt As Date
    dt = DateSerial(CLng(sYear), CLng(sMonth), CLng(sDay))
    Err.Clear
    
    'fIsDate = IsDate(sYear & "-" & sMonth & "-" & sDay)
    IsDate = CBool(dt > DateSerial(1990, 1, 1))
End Function
  
Function CreateAddNameUpdateNameWhenExists(sName As String, aReferTo, Optional wb As Workbook) As Name
    If IsMissing(wb) Or wb Is Nothing Then Set wb = ThisWorkbook
    
    If NameExists(sName, wb) Then
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
    Set CreateAddNameUpdateNameWhenExists = wb.Names(sName)
End Function

Function RemoveName(sName As String, Optional wb As Workbook)
    If IsMissing(wb) Or wb Is Nothing Then Set wb = ThisWorkbook
    
    If Not NameExists(sName, wb) Then Exit Function
    
    wb.Names(sName).Delete
End Function

Function NameExists(sName As String, Optional wb As Workbook) As Boolean
    Dim eachName
    If IsMissing(wb) Or wb Is Nothing Then Set wb = ThisWorkbook
    
    For Each eachName In wb.Names
        If UCase(eachName.Name) = UCase(sName) Then
            NameExists = True
            Exit Function
        End If
    Next
    NameExists = False
End Function

Function CheckPath(asPath As String, Optional CreatePath As Boolean = False) As String
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
            RaiseErr "The path passed to CheckPath is neither a local path nor a networkpath:" & vbCr & asPath
        End If
    End If
    If IsArray(arr) Then Erase arr
    
    If Not Right$(sOut, 1) = Application.PathSeparator Then  'Chr(92)
        sOut = sOut & Application.PathSeparator
    End If
    
    Do While InStr(sOut, Application.PathSeparator & Application.PathSeparator) > 0
        sOut = Replace(sOut, Application.PathSeparator & Application.PathSeparator, Application.PathSeparator)
    Loop
    
    CheckPath = sOut
    Exit Function
erro_h:
    If Err.Number <> 0 Then
        If Err.Number = 52 Then
            RaiseErr "The path cannot be created recursively, you may not have permission to create it" & vbCr & asPath
        Else
            RaiseErr "Error has occurred: " & vbCr & vbCr _
                & "Err number: " & Err.Number & vbCr _
                & "error: " & Err.Description
        End If
    End If
    
    On Error GoTo 0
End Function

Function ConvertArrayColToText(ByRef arrData(), iCol As Integer)
    Dim lEachRow As Long
    
    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
        arrData(lEachRow, 1) = "'" & arrData(lEachRow, 1)
    Next
End Function

Function ArrLen(arrParam, Optional iDimension As Integer = 1) As Long
    If iDimension > 2 Then RaiseErr "wrong param in arrlen"
    
    If ArrayIsEmpty(arrParam) Then ArrLen = 0: Exit Function
    
    If iDimension <= 0 Then iDimension = 1
    
    ArrLen = UBound(arrParam, iDimension) - LBound(arrParam, iDimension) + 1
End Function

Function PasteAppendSimpleDictionaryToSheet(dict As Dictionary, sht As Worksheet, Optional alPasteFromRow As Long = 1, Optional alPasteFromCol As Long = 1)
    If dict.Count <= 0 Then Exit Function
    
    If alPasteFromRow <= 0 Then alPasteFromRow = GetMaxRow(sht) + 1
    
    Dim arr()
    Dim aKey
    Dim i As Long
    ReDim arr(1 To dict.Count, 1 To 1)
    
    i = 0
    For Each aKey In dict.Keys
        i = i + 1
        arr(i, 1) = aKey
    Next

    sht.Cells(alPasteFromRow, alPasteFromCol).Resize(ArrLen(arr, 1), ArrLen(arr, 2)).Value = arr
    Erase arr
    
    ReDim arr(1 To dict.Count, 1 To 1)
    i = 0
    For Each aKey In dict.Items
        i = i + 1
        arr(i, 1) = aKey
    Next

    sht.Cells(alPasteFromRow, alPasteFromCol + 1).Resize(ArrLen(arr, 1), ArrLen(arr, 2)).Value = arr
    Erase arr
End Function
  
Function DeleteTrailingBlankRowsFromSheet(sht As Worksheet)
    Dim lValidMax As Long
    Dim lUsedMax As Long
    
    lUsedMax = sht.UsedRange.row + sht.UsedRange.Rows.Count - 1
    lValidMax = GetMaxRow(sht)
    
    If lUsedMax > lValidMax Then
        sht.Rows(lValidMax + 1 & ":" & lUsedMax).Delete Shift:=xlUp
    End If
End Function


Function VeryHideSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetVeryHidden
End Function

Function HideSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetHidden
End Function

Function ShowSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetVisible
    sht.Activate
End Function

Function ActivateSheet(ByRef sht As Worksheet)
    Dim bEnableEventsOrig As Boolean
    Dim bDisplayAlertsOrig As Boolean
      
    bEnableEventsOrig = Application.EnableEvents
    bDisplayAlertsOrig = Application.DisplayAlerts
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    sht.Visible = xlSheetVisible
    sht.Activate
    
    Application.EnableEvents = bEnableEventsOrig
    Application.DisplayAlerts = bDisplayAlertsOrig
End Function
Function SheetExists(asShtName As String, Optional ByRef shtOut As Worksheet, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    Dim bOut As Boolean

    If wb Is Nothing Then Set wb = ThisWorkbook

    bOut = False
    asShtName = UCase(Trim(asShtName))

    For Each sht In wb.Worksheets
        If UCase(sht.Name) = asShtName Then
            bOut = True
            Set shtOut = sht
            Exit For
        End If
    Next

    Set sht = Nothing
    SheetExists = bOut
End Function

Function SheetExistsByCodeName(asShtCodeName As String, Optional ByRef shtOut As Worksheet, Optional wb As Workbook _
                , Optional abPromptErrMsgIfNotFound As Boolean = False) As Boolean
    Dim sht As Worksheet
    Dim bOut As Boolean

    If wb Is Nothing Then Set wb = ThisWorkbook

    bOut = False
    asShtCodeName = UCase(Trim(asShtCodeName))

    For Each sht In wb.Worksheets
        If UCase(sht.CodeName) = asShtCodeName Then
            bOut = True
            Set shtOut = sht
            Exit For
        End If
    Next

    Set sht = Nothing
    SheetExistsByCodeName = bOut

    If abPromptErrMsgIfNotFound Then
        If Not bOut Then RaiseErr "The workbook " & wb.Name & " does not have a sheet whose code name is " & asShtCodeName
    End If
End Function

Function GetSheetByCodeName(asShtCodeName As String, Optional wb As Workbook) As Worksheet
    Dim sht As Worksheet
    Dim shtOut As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook

    asShtCodeName = UCase(Trim(asShtCodeName))

    For Each sht In wb.Worksheets
        If UCase(sht.CodeName) = asShtCodeName Then
            Set shtOut = sht
            Exit For
        End If
    Next

    Set GetSheetByCodeName = shtOut
    Set shtOut = Nothing
    Set sht = Nothing
End Function

Function BackupFile(sOrigFilePath As String, sBackUpFilePath As String, Optional bOverWriteWhenBackupFileExists As Boolean = True)
    If Len(Trim(sOrigFilePath)) <= 0 Then RaiseErr "sOrigFilePath is blank"
    If Len(Trim(sBackUpFilePath)) <= 0 Then RaiseErr "sBackUpFilePath is blank"

    If Not fso.FileExists(sOrigFilePath) Then RaiseErr "The source file does not exists. " & vbCrLf & sOrigFilePath
    
    If fso.FileExists(sBackUpFilePath) Then
        If bOverWriteWhenBackupFileExists Then
            
        Else
            RaiseErr "The backup file already exists. " & vbCrLf & sBackUpFilePath
        End If
    End If
    
    Call fso.CopyFile(sOrigFilePath, sBackUpFilePath, True)
End Function

Function ReadWholeSheetData2Array(shtToRead As Worksheet, Optional alDataFromRow As Long = 2, Optional alMaxCol As Long = 0)
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    Dim arrDataOut()
    
    arrDataOut = Array()
    
    lMaxRow = GetMaxRow(shtToRead)
    If lMaxRow < alDataFromRow Then GoTo exit_fun

    If alMaxCol > 0 Then
        lMaxCol = alMaxCol
    Else
        lMaxCol = GetMaxCol(shtToRead)
    End If

    arrDataOut = ReadRangeDatatoArrayByStartEndPos(shtToRead, alDataFromRow, 1, lMaxRow, lMaxCol)
    
exit_fun:
    ReadWholeSheetData2Array = arrDataOut
    Erase arrDataOut
End Function

Function FindInWorksheet(rngToFindIn As Range, sWhatToFind As String _
                    , Optional abNotFoundThenError As Boolean = True _
                    , Optional abAllowMultiple As Boolean = False) As Range
    If Len(Trim(sWhatToFind)) <= 0 Then RaiseErr "Wrong param sWhatToFind to FindInWorksheet " & sWhatToFind

    Dim rngOut  As Range
    Dim rngFound As Range
    Dim lFoundCnt As Long
    Dim sFirstAddress As String

    Set rngFound = rngToFindIn.Find(what:=sWhatToFind _
                                    , After:=rngToFindIn.Cells(rngToFindIn.Rows.Count, rngToFindIn.Columns.Count) _
                                    , LookIn:=xlValues _
                                    , LookAt:=xlWhole _
                                    , SearchOrder:=xlByRows _
                                    , SearchDirection:=xlNext _
                                    , MatchCase:=False _
                                    , matchbyte:=False)
    Set rngOut = rngFound

    If rngFound Is Nothing Then
        If abNotFoundThenError Then
            RaiseErr """" & sWhatToFind & """ cannot be found in sheet " & rngToFindIn.Parent.Name & "[" & rngToFindIn.Address & "], pls check your program."
        Else
            GoTo exit_function
        End If
    Else
        If Not abAllowMultiple Then
            sFirstAddress = rngFound.Address
            lFoundCnt = 1

            Do While True
                Set rngFound = rngToFindIn.Find(what:=sWhatToFind _
                                            , After:=rngFound _
                                            , LookIn:=xlValues _
                                            , LookAt:=xlWhole _
                                            , SearchOrder:=xlByRows _
                                            , SearchDirection:=xlNext _
                                            , MatchCase:=False _
                                            , matchbyte:=False)
                If rngFound Is Nothing Then Exit Do
                If rngFound.Address = sFirstAddress Then Exit Do

                lFoundCnt = lFoundCnt + 1
            Loop

            If lFoundCnt > 1 Then
                RaiseErr lFoundCnt & " copies of """ & sWhatToFind & """ were found in sheet " & rngToFindIn.Parent.Name & ", pls check your program."
            End If
        End If
    End If
exit_function:
    Set FindInWorksheet = rngOut
    Set rngOut = Nothing
    Set rngFound = Nothing
End Function


