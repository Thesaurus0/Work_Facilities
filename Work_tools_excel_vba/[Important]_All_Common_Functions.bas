
#If VBA7 And Win64 Then  'Win64
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Public mRibbonObj As IRibbonUI

'=============================================================
Sub subRefreshRibbon()
    fGetRibbonReference.Invalidate
End Sub
Sub ERP_UI_Onload(ribbon As IRibbonUI)
  Set mRibbonObj = ribbon

  fCreateAddNameUpdateNameWhenExists "nmRibbonPointer", ObjPtr(ribbon)
  'Names("nmRibbonPointer").RefersTo = ObjPtr(ribbon)

  mRibbonObj.ActivateTab "ERP_2010"
  ThisWorkbook.Saved = True
End Sub
Function fGetRibbonReference() As IRibbonUI
    If Not mRibbonObj Is Nothing Then Set fGetRibbonReference = mRibbonObj: Exit Function

    Dim objRibbon As Object
    Dim lRibPointer As LongPtr

    lRibPointer = [nmRibbonPointer]

    CopyMemory objRibbon, lRibPointer, LenB(lRibPointer)

    Set fGetRibbonReference = objRibbon
    Set mRibbonObj = objRibbon
    Set objRibbon = Nothing
End Function
'---------------------------------------------------------------------
Sub Button_onAction(control As IRibbonControl)
    Call fGetControlAttributes(control, "ACTION")
End Sub
Sub Button_getImage(control As IRibbonControl, ByRef imageMso)
    Call fGetControlAttributes(control, "IMAGE", imageMso)
End Sub
Sub Button_getLabel(control As IRibbonControl, ByRef label)
    Call fGetControlAttributes(control, "LABEL", label)
End Sub
Sub Button_getSize(control As IRibbonControl, ByRef size)
    Call fGetControlAttributes(control, "SIZE", size)
End Sub

'================== toggle button common function===========================================
Sub ToggleButtonToSwitchSheet_onAction(control As IRibbonControl, pressed As Boolean)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)

    If Not sht Is Nothing Then
        fToggleSheetVisibleFromUIRibbonControl pressed, sht, control
    End If
    Set sht = Nothing
End Sub

Sub ToggleButtonToSwitchSheet_getPressed(control As IRibbonControl, ByRef returnedVal)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)

    If sht Is Nothing Then
        returnedVal = False
    Else
        returnedVal = (sht.Visible = xlSheetVisible And ActiveSheet Is sht)
    End If
End Sub
Function fGetSheetByUIRibbonTag(ByVal asButtonTag As String) As Worksheet
    Dim sht As Worksheet

    If fSheetExistsByCodeName(asButtonTag, sht) Then
        Set fGetSheetByUIRibbonTag = sht
    Else
        MsgBox "The button's Tag is not corresponding to any worksheet in this workbook, please check the customUI.xml you prepared," _
            & " The design thought is that the button's tag is the name of a sheet, so that the common function ToggleButtonToSwitchSheet_onAction/getPressed can get a worksheet." _
            & vbCr & vbCr & "asButtonTag: " & asButtonTag

    End If
    Set sht = Nothing
End Function
Function fToggleSheetVisibleFromUIRibbonControl(ByVal pressed As Boolean, sht As Worksheet, control As IRibbonControl)
    If pressed Then
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        End If
    Else
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        Else
            If fWorkbookHasMoreThanOneSheetVisible(ThisWorkbook) Then
                fVeryHideSheet sht
            End If
        End If
    End If

    'fGetRibbonReference.InvalidateControl (control.id)
    fGetRibbonReference.Invalidate
End Function

'---------------------------------------------------------------------

'==========================dev prod switch===================================
Sub btnSwitchDevProd_onAction(control As IRibbonControl, pressed As Boolean)
    sub_SwitchDevProdMode
End Sub

Sub btnSwitchDevProd_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
Sub btnSwitchDevProd_getVisible(control As IRibbonControl, ByRef returnedVal)
    'returnedVal = fIsDev()
    returnedVal = True
End Sub
Sub grpDevFacilities_getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
'---------------------------------------------------------------------

'================ dev facilities ==============================================
Sub btnListAllFunctions_onAction(control As IRibbonControl)
    sub_ListAllFunctionsOfThisWorkbook
End Sub
Sub btnExportSourceCode_onAction(control As IRibbonControl)
    sub_ExportModulesSourceCodeToFolder
End Sub
Sub btnGenNumberList_onAction(control As IRibbonControl)
    sub_GenNumberList
End Sub
Sub btnGenAlphabetList_onAction(control As IRibbonControl)
    sub_GenAlpabetList
End Sub
Sub btnListAllActiveXOnCurrSheet_onAction(control As IRibbonControl)
    Sub_ListActiveXControlOnActiveSheet
End Sub
Sub btnResetOnError_onAction(control As IRibbonControl)
    sub_ResetOnError_Initialize
End Sub
'------------------------------------------------------------------------------

Function fGetControlAttributes(control As IRibbonControl, sType As String, Optional ByRef val)
    If Not (sType = "LABEL" Or sType = "IMAGE" Or sType = "SIZE" Or sType = "ACTION") Then
        fErr "wrong param to fGetControlAttributes: " & vbCr & "sType=" & sType & vbCr & "control=" & control.id
    End If

    Select Case control.id
        Case "btnCalSummaryAmount"
            Select Case sType
                Case "LABEL":   val = "计算入帐出帐结果"
                Case "IMAGE":   val = "FunctionWizard"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_CalculateBillInOut
            End Select
        Case "tbtnShowSummaryAmount"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "入帐出帐汇总表"
                Case "IMAGE":   val = "ChartShowData"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
        Case "tbtnShowshtBillIn"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "入帐表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
        Case "tbtnShowshtBillOut"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "出帐表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select

        Case "btnSummaryBusinessData"
            Select Case sType
                Case "LABEL":   val = "汇总明细"
                Case "IMAGE":   val = "FunctionWizard"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_SummarizeBusinssDetail
            End Select
        Case "tbtnShowshtBusinessDetails"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "明细表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select

        Case "tbtnShowshtBusinessSumm"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "汇总表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
        Case "btnClearData"
            Select Case sType
                Case "LABEL":   val = "清除明细表中数据"
                Case "IMAGE":   val = "ReviewRejectChange"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""

                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_ClearBuzDetails
            End Select

    End Select

End Function

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

    lReturnVal = ShellExecute(Application.hwnd, "Open", asFileFullPath, "", "C:\", SW_SHOWNORMAL)

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

Function fGetFileParentFolder(asFileFullPath As String) As String
    fGetFSO
    fGetFileParentFolder = gFSO.GetParentFolderName(asFileFullPath)
End Function

Function fGetFileBaseName(asFileFullPath As String) As String
    fGetFSO
    fGetFileBaseName = gFSO.GetFileName(asFileFullPath)
End Function
Function fGetFileNetName(asFileFullPath As String) As String
    fGetFSO
    fGetFileNetName = gFSO.GetBaseName(asFileFullPath)
End Function
Function fGetFileExtension(asFileFullPath As String, Optional bDot As Boolean = False) As String
    fGetFSO
    fGetFileExtension = IIf(bDot, ".", "") & gFSO.GetExtensionName(asFileFullPath)
End Function

'Function fGetFileNamePart(asFileFullPath As String _
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
    fGenRandomUniqueString = format(Now(), "yyyymmddhhMMSS") & Rnd()
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
            Application.Goto shtAt.Cells(lActualRow, arrKeyCols(UBound(arrKeyCols)))
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
                Application.Goto shtAt.Cells(lActualRow, lKeyCol)
                fErr "Keys [" & sColLetter & "] is blank!" & sPos
            End If

            GoTo next_row
        End If

        If dict.Exists(sKeyStr) Then
            'sPos = sPos & lActualRow & " / " & sColLetter
            sPos = Replace(sPos, "ACTUAL_ROW_NO", lActualRow)
            fShowSheet shtAt
            Application.Goto shtAt.Cells(lActualRow, lKeyCol)
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

Function fDeleteAllFilesInFolder(sFolder As String)
    fGetFSO

    Dim aFile As File

    If gFSO.FolderExists(sFolder) Then
        For Each aFile In gFSO.GetFolder(sFolder).Files
            aFile.Delete True
        Next
    End If
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
    fRedim arr, arrlen(arr) + 1, aPreserve
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
        If Application.CutCopyMode = 0 Then Application.Calculation = xlCalculationAutomatic
    Else
        If Application.CutCopyMode = 0 Then Application.Calculation = xlCalculationManual
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

        lNextStart = match.FirstIndex + match.length + 1
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

    Debug.Print "fFileterOutTwoDimensionArray: " & format(Timer - start, "00:00")
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

    Debug.Print "fFileterOutTwoDimensionArray: " & Timer - start & vbCr & format(Timer - start, "00:00")
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
        sDatePattern = oMatch.Value

        lStartPos = oMatch.FirstIndex + 1
        lEndPos = oMatch.FirstIndex + oMatch.length
        lLen = oMatch.length

        sDate = format(aDate, sDatePattern)

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

Function fCheckPath(asPath As String, Optional CreatePath As Boolean = False) As String
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

Dim dblStartTime As Double
Dim bCopyType As String
Dim sCopyStageRangeAddr As String
'Dim sCopiedAddress As String

Function fStartTimer()
    dblStartTime = Timer
End Function
Function fHowLong(Optional ByVal asPrefix As String = "")
    Debug.Print IIf(Len(asPrefix) > 0, asPrefix & vbTab, "") & format(Timer - dblStartTime, "0000.00000000")
End Function

Function fKeepCopyContent()
    Dim myData As DataObject
    Dim sCopiedStr As String
    Const PASTE_START_CELL = "G1"

    Dim shtActiveOrig As Worksheet

    If Application.CutCopyMode = xlCopy Then
        Set shtActiveOrig = ActiveSheet
        bCopyType = "COPY_RANGE"
    ElseIf Application.CutCopyMode = xlCut Then
        Set shtActiveOrig = ActiveSheet
        bCopyType = "CUT_RANGE"
    Else
        Set myData = New DataObject
        myData.GetFromClipboard

        On Error Resume Next
        sCopiedStr = myData.GetText()

        If Err.Number <> 0 Then
            bCopyType = "NOTHING"
        Else
            bCopyType = "COPY_OTHERS"
        End If
        On Error GoTo 0

        Set myData = Nothing
    End If

    If bCopyType = "COPY_RANGE" Or bCopyType = "CUT_RANGE" Then
        shtDataStage.Activate

        shtDataStage.Range(PASTE_START_CELL).PasteSpecial xlPasteAll
        sCopyStageRangeAddr = Selection.Address(external:=True)
        fHideSheet shtDataStage

        shtActiveOrig.Activate
    ElseIf bCopyType = "COPY_OTHERS" Then
        shtDataStage.Range(PASTE_START_CELL).Value = sCopiedStr
        sCopyStageRangeAddr = ""
    ElseIf bCopyType = "NOTHING" Then
        shtDataStage.Range(PASTE_START_CELL).ClearComments
        shtDataStage.Range(PASTE_START_CELL).ClearContents
        shtDataStage.Range(PASTE_START_CELL).ClearFormats
        'shtDataStage.Range(PASTE_START_CELL).ClearHyperlinks
        shtDataStage.Range(PASTE_START_CELL).ClearNotes
        shtDataStage.Range(PASTE_START_CELL).ClearOutline
        sCopyStageRangeAddr = ""
    Else
        fErr "bCopyType"
    End If

    Set shtActiveOrig = Nothing
End Function

Function fCopyFromKept()
    Dim myData As DataObject
    Const PASTE_START_CELL = "G1"

    If bCopyType = "COPY_RANGE" Then
        shtDataStage.Range(sCopyStageRangeAddr).Copy
    ElseIf bCopyType = "CUT_RANGE" Then
        shtDataStage.Range(sCopyStageRangeAddr).Cut
    ElseIf bCopyType = "COPY_OTHERS" Then
        Set myData = New DataObject
        myData.SetText CStr(shtDataStage.Range(PASTE_START_CELL).Value)
        myData.PutInClipboard
        Set myData = Nothing
    ElseIf bCopyType = "NOTHING" Then
    Else
        fErr "bCopyType"
    End If

    sCopyStageRangeAddr = ""
End Function

'Function fGetCopyAddress()
'    sCopiedAddress = Application.Selection.Address(external:=True)
'    Application.Selection.Copy
'
'    MsgBox sCopiedAddress
'End Function

'======================================================================================================
Sub Sub_ListActiveXControlOnActiveSheet()
    Dim obj As Object
    Dim sStr As String

    For Each obj In ActiveSheet.DrawingObjects
        sStr = sStr & vbCr & obj.Name
    Next

    Set obj = Nothing

    MsgBox sStr
End Sub

Sub sub_ExportModulesSourceCodeToFolder()
    Dim sFolder As String
    Dim sMsg As String
    Dim i As Integer
    Dim iCnt As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent

    Set vbProj = ThisWorkbook.VBProject

    iCnt = vbProj.VBComponents.Count

    fGetFSO

    For i = 1 To 1
        If i = 1 Then
            sFolder = ThisWorkbook.Path & "\" & "Source_Code"
        Else
        End If

        sMsg = sMsg & vbCr & vbCr & sFolder

        If Not gFSO.FolderExists(sFolder) Then gFSO.CreateFolder (sFolder)

        'call fCheckPath(sfolder, true)
        fDeleteAllFilesInFolder (sFolder)

        iCnt = 0
        For Each vbComp In vbProj.VBComponents
            If UCase(vbComp.Name) Like "SHEET*" Then GoTo Next_mod
            If vbComp.Type = 1 Or vbComp.Type = 3 Or vbComp.Type = 100 Then
                vbComp.Export sFolder & "\" & vbComp.Name & ".bas"
            End If

Next_mod:
        Next
    Next

    MsgBox "Done"
End Sub

Sub sub_ListAllFunctionsOfThisWorkbook()
    Dim shtOutput As Worksheet
    If Not fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtOutput) Then Exit Sub

    Dim arrModules()
    Dim arrFunctions()

    arrModules = fGetListAllModulesOfThisWorkbook()
    arrFunctions = fGetListAllSubFunctionsInThisWorkbook(arrModules)

    Call fWriteArray2Sheet(shtOutput, arrFunctions)

    Erase arrModules: Erase arrFunctions

    shtOutput.Cells(1, 1) = "Type"
    shtOutput.Cells(1, 2) = "Modules"
    shtOutput.Cells(1, 3) = "Functions"

    Call fAutoFilterAutoFitSheet(shtOutput)
    Call fFreezeSheet(shtOutput)
    Call fSortDataInSheetSortSheetData(shtOutput, Array(3))

    Set shtOutput = Nothing
End Sub

Sub Sub_ToHomeSheet()
    shtMainMenu.Visible = xlSheetVisible
    shtMainMenu.Activate

'    If shtMainMenu.Visible = xlSheetVisible Then
'        shtMainMenu.Activate
'    Else
'        shtMainMenu.Visible = xlSheetVisible
'        ThisWorkbook.Worksheets(1).Activate
'    End If
End Sub

Sub sub_ResetOnError_Initialize()
    Err.Clear

    fGetProgressBar
    gProBar.ShowBar
    'On Error GoTo err_exit

    gsEnv = fGetEnvFromSysConf

    Call fEnableExcelOptionsAll

    gProBar.ChangeProcessBarValue 0.2
    Call sub_RemoveAllCommandBars

    gProBar.ChangeProcessBarValue 0.3
    Call fDeleteAllConditionFormatForAllSheets

   ' Call ThisWorkbook.fRefreshGetAllCommandbarsList

    gProBar.ChangeProcessBarValue 0.4
    Call ThisWorkbook.sub_WorkBookInitialization

    Call fSetIntialValueForShtMenuInitialize
    gProBar.ChangeProcessBarValue 1
err_exit:
    gProBar.DestroyBar
    Err.Clear
    ThisWorkbook.CheckCompatibility = False
    Set gProBar = Nothing
    'End
End Sub
Function fGetEnvFromSysConf() As String
    gsEnv = fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=DEVELOPMENT_OR_FORMAL_RELEASE", False)
    fGetEnvFromSysConf = gsEnv
End Function

Sub sub_SwitchDevProdMode()
    gsEnv = fGetEnvFromSysConf

    Call fEnableExcelOptionsAll

    If gsEnv = "DEV" Then
        gsEnv = "PROD"
    ElseIf gsEnv = "PROD" Then
        gsEnv = "DEV"
    End If

    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=DEVELOPMENT_OR_FORMAL_RELEASE" _
                                    , gsEnv, False)
    shtMenu.Activate
    Range("A1").Select
End Sub

Function fSetDEVUATPRODNotificationInSheetMenu()
    Const sDevNotifi = "This is DEV mode, please switch to PROD vresion by click the button above ""Switch Dev/Prod Mode"""

    Dim sNotifi As String
    Dim iColor As Long
    Dim iFontSize As Long
    Dim bBold As Boolean

    If gsEnv = "DEV" Then
        sNotifi = sDevNotifi

        iColor = RGB(0, 0, 255)
        iFontSize = 20
        bBold = True
    ElseIf gsEnv = "PROD" Then
        sNotifi = ""

        iColor = RGB(0, 0, 0)
        iFontSize = 10
        bBold = False
    Else
    End If

    shtMenu.Range("A1").Value = sNotifi
    shtMenu.Range("A1").Font.size = iFontSize
    shtMenu.Range("A1").Font.Color = iColor
    shtMenu.Range("A1").Font.Bold = bBold

    shtMenuCompInvt.Range("A1").Value = sNotifi
    shtMenuCompInvt.Range("A1").Font.size = iFontSize
    shtMenuCompInvt.Range("A1").Font.Color = iColor
    shtMenuCompInvt.Range("A1").Font.Bold = bBold
End Function

'*************************************************************************

Function fGetListAllModulesOfThisWorkbook() As Variant
    Dim arrOut()
    Dim iCnt As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent

    Set vbProj = ThisWorkbook.VBProject

    iCnt = vbProj.VBComponents.Count
    ReDim arrOut(1 To iCnt, 3)

    iCnt = 0
    For Each vbComp In vbProj.VBComponents
        iCnt = iCnt + 1
        arrOut(iCnt, 1) = "Modules"
        arrOut(iCnt, 2) = fVBEComponentTypeToString(vbComp.Type)
        arrOut(iCnt, 3) = vbComp.Name
    Next

    fGetListAllModulesOfThisWorkbook = arrOut
    Erase arrOut
End Function

Function fVBEComponentTypeToString(aType As VBIDE.vbext_ComponentType) As String
    Dim sOut As String

    Select Case aType
        Case VBIDE.vbext_ct_ActiveXDesigner
            sOut = "ActiveX Designer"
        Case VBIDE.vbext_ct_ClassModule
            sOut = "Class"
        Case VBIDE.vbext_ct_StdModule
            sOut = "Module"
        Case VBIDE.vbext_ct_Document
            sOut = "Document"
        Case VBIDE.vbext_ct_MSForm
            sOut = "User Form"
        Case Else
            sOut = "Unknown type: " & CStr(aType)
    End Select

    fVBEComponentTypeToString = sOut
End Function

Function fGetListAllSubFunctionsInThisWorkbook(arrModules()) As Variant
    Dim arrOut()
    Dim i As Long
    Dim iCnt As Long
    Dim sMod As String
    Dim lineNo As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Dim codeMod As VBIDE.CodeModule
    Dim procKind As VBIDE.vbext_ProcKind
    Dim funcName As String

    Set vbProj = ThisWorkbook.VBProject

    iCnt = 0
    ReDim arrOut(1 To 10000, 4)

    For i = LBound(arrModules, 1) To UBound(arrModules, 1)
        sMod = arrModules(i, 3)

        Set vbComp = vbProj.VBComponents(sMod)
        Set codeMod = vbComp.CodeModule

        lineNo = codeMod.CountOfDeclarationLines + 1

        Do Until lineNo >= codeMod.CountOfLines + 1
            funcName = codeMod.ProcOfLine(lineNo, procKind)

            If Not UCase(funcName) Like "CB*_CLICK" Then
                iCnt = iCnt + 1
                arrOut(iCnt, 1) = "Functions"
                arrOut(iCnt, 2) = sMod
                arrOut(iCnt, 3) = funcName
                arrOut(iCnt, 4) = ProcKindString(procKind)
            End If

            lineNo = codeMod.ProcStartLine(funcName, procKind) + codeMod.ProcCountLines(funcName, procKind) + 1
        Loop
    Next
    fGetListAllSubFunctionsInThisWorkbook = arrOut
    Erase arrOut
End Function

Function ProcKindString(procKind As VBIDE.vbext_ProcKind) As String
    Dim sOut As String

    Select Case procKind
        Case VBIDE.vbext_pk_Get
            sOut = "Property Get"
        Case VBIDE.vbext_pk_Let
            sOut = "Property Let"
        Case VBIDE.vbext_pk_Proc
            sOut = "Sub/Function"
        Case VBIDE.vbext_pk_Set
            sOut = "Property Set"
        Case Else
            sOut = "Unknown type: " & CStr(procKind)
    End Select
    ProcKindString = sOut
End Function

Function fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtTmp As Worksheet, Optional wb As Workbook) As Boolean
    Dim sTmp As String
    Dim response As VbMsgBoxResult

    If wb Is Nothing Then Set wb = ThisWorkbook

    sTmp = "tmpOutput"

    If fSheetExists(sTmp) Then
        wb.Worksheets(sTmp).Activate

        response = MsgBox("There is an existing sheet " & sTmp & ", to delete it, please press yes" _
                    & vbCr, vbCritical + vbYesNoCancel)
        If response = vbNo Then
            Set shtTmp = wb.Worksheets(sTmp)
        ElseIf response = vbYes Then    'vbYes
            Call fDeleteSheet(sTmp)
            Set shtTmp = fAddNewSheet(sTmp)
        Else
            fGetTmpSheetInWorkbookWhenNotExistsCreateIt = False
            Exit Function
        End If
    Else
        Set shtTmp = fAddNewSheet(sTmp)
    End If

    fGetTmpSheetInWorkbookWhenNotExistsCreateIt = True
End Function

Function fShtSysConf_SheetChange_DevProdChange(Target As Range)
    Dim rgAimed As Range
    Dim rgIntersect As Range

    Set rgAimed = fGetRangeFromExternalAddress(fGetSpecifiedConfigCellAddress(shtSysConf, "[Facility For Testing]", "Value" _
                        , "Setting Item ID=DEVELOPMENT_OR_FORMAL_RELEASE"))
    Set rgIntersect = Intersect(Target, rgAimed)

    If Not rgIntersect Is Nothing Then
        If rgIntersect.Areas.Count > 1 Then fErr "Please select only one cell."

        gsEnv = rgIntersect.Value

        Call fRemoveAllCommandbarsByConfig
        Call ThisWorkbook.sub_WorkBookInitialization
        Call fSetIntialValueForShtMenuInitialize
        Call fSetDEVUATPRODNotificationInSheetMenu

        fGetRibbonReference.Invalidate
    End If

    Set rgAimed = Nothing
    Set rgIntersect = Nothing
End Function

Sub sub_GenAlpabetList()
    Dim maxNum
    Dim lMax As Long
    Dim sMaxcol As String
    Dim arrList()

    If Not fPromptToOverWrite() Then Exit Sub

    maxNum = InputBox("How many letters to you want to generate? (either number or letter is ok, e.g., 20 or AF)", "Max Number letter")

    If fZero(maxNum) Then Exit Sub

    maxNum = Trim(maxNum)

    On Error Resume Next
    lMax = CLng(maxNum)
    sMaxcol = CStr(maxNum)
    Err.Clear

    If lMax > 0 Then
    ElseIf Len(sMaxcol) > 0 Then
        lMax = fLetter2Num(sMaxcol)
    End If

    If lMax <= 0 Or lMax > Columns.Count Then
        fMsgBox "the number you input is too small or too large, which should be with 1 - " & Columns.CountLarge
        Exit Sub
    End If

    Dim i As Long
    ReDim arrList(1 To lMax, 1)
    For i = 1 To lMax
        arrList(i, 1) = fNum2Letter(i)
    Next

    ActiveCell.Resize(UBound(arrList, 1), 1).Value = arrList
    Erase arrList
End Sub

Sub sub_GenNumberList()
    Dim maxNum
    Dim lMax As Long
    Dim sMaxcol As String
    Dim arrList()

    If Not fPromptToOverWrite() Then Exit Sub

    maxNum = InputBox("How many letters to you want to generate? ( e.g., 20 , 100)", "Max Number")
    If fZero(maxNum) Then Exit Sub

    maxNum = Trim(maxNum)

    On Error Resume Next
    lMax = CLng(maxNum)
    Err.Clear

    If lMax <= 0 Then
        fMsgBox "the number you input is too small or too large, which should be with 1 - " & Columns.CountLarge
        Exit Sub
    End If

    Dim i As Long
    ReDim arrList(1 To lMax, 1)
    For i = 1 To lMax
        arrList(i, 1) = i
    Next

    ActiveCell.Resize(UBound(arrList, 1), 1).Value = arrList
    Erase arrList

End Sub

Function fPromptToOverWrite() As Boolean
    fPromptToOverWrite = fPromptToConfirmToContinue("Data will be write to the current cell:" _
                & Replace(ActiveCell.Address, "$", "") & vbCr & "are you sure to continue?")
End Function
Function fPromptToConfirmToContinue(asAskMsg As String _
            , Optional aBBbMsgboxStyle As VbMsgBoxStyle = vbYesNoCancel + vbCritical + vbDefaultButton3 _
            , Optional bDoubleConfirm As Boolean = False) As Boolean
    fPromptToConfirmToContinue = False

    Dim response As VbMsgBoxResult
    response = MsgBox(Prompt:=asAskMsg, Buttons:=aBBbMsgboxStyle)

    If response <> vbYes Then Exit Function

    If bDoubleConfirm Then
        response = MsgBox(Prompt:="Are you sure to continue?", Buttons:=vbYesNoCancel + vbCritical + vbDefaultButton3)
        If response <> vbYes Then Exit Function
    End If

    fPromptToConfirmToContinue = True
End Function

'Sub AddFaceIDs()
'    Dim GName As String
'    Dim I As Integer, J As Single
'
'    For I = 6 To 1 Step -1 'Display from bottom to top
'        GName = "Group" & 600 * (I - 1) + 1 & "_" & 600 * I
'        On Error GoTo Endline
'        With Application.CommandBars.Add(GName)
'            .Visible = True
'            With .Controls
'                For J = 600 * (I - 1) + 1 To 600 * I
'                On Error Resume Next
'                With .Add(msoControlButton)
'                .FaceId = J
'                .Caption = J
'                End With
'                Next
'            End With
'        End With
'Endline:
'        With CommandBars(GName)
'            .Visible = True
'            .Width = 720 'contains 30?20 icons
'            .Left = 50 + (6 - I) * 20
'            .Top = 90 + (6 - I) * 20
'        End With
'    Next I
'End Sub
Sub Sub_FilterByActiveCell()
    Dim lMaxCol As Long
    lMaxCol = ActiveSheet.Cells(1, 1).End(xlToRight).Column
    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(ActiveSheet)

    If ActiveSheet.AutoFilterMode Then  'auto filter
        ActiveSheet.AutoFilter.ShowAllData
    Else
        fGetRangeByStartEndPos(ActiveSheet, 1, 1, 1, lMaxCol).AutoFilter
    End If

    Dim aActiveCellValue
    Dim lColToFilter As Long
    aActiveCellValue = ActiveCell.Value
    lColToFilter = ActiveCell.Column

    fGetRangeByStartEndPos(ActiveSheet, 1, 1, lMaxRow, lMaxCol).AutoFilter _
                Field:=lColToFilter _
                , Criteria1:="=*" & aActiveCellValue & "*" _
                , Operator:=xlAnd
End Sub

Sub Sub_FilterBySelectedCells()
    Dim rngSelected As Range

    Set rngSelected = Selection
    If fIfSelectedMoreThanOneRow(rngSelected) Then
        fMsgBox "不能选多行，只能选一行。"
        End
    End If

    'Call Sub_RemoveFilterForAcitveSheet("CLEAR_FILTER")
    Call Sub_RemoveFilterForAcitveSheet

    Dim lMaxRow As Long
    Dim lMaxCol As Long
    lMaxRow = fGetValidMaxRow(ActiveSheet)
    lMaxCol = fGetValidMaxCol(ActiveSheet)

    If ActiveSheet.AutoFilterMode Then ActiveSheet.AutoFilterMode = False
    fGetRangeByStartEndPos(ActiveSheet, 1, 1, 1, lMaxCol).AutoFilter

'    If ActiveSheet.AutoFilterMode Then  'auto filter
'        ActiveSheet.AutoFilter.ShowAllData
'    Else
'        fGetRangeByStartEndPos(ActiveSheet, 1, 1, 1, lMaxCol).AutoFilter
'    End If

    Dim eachCol As Integer
    Dim rgData As Range
    Set rgData = fGetRangeByStartEndPos(ActiveSheet, 1, 1, lMaxRow, lMaxCol)

    Dim rngEachArea As Range
    Dim eachCell As Range

    For Each rngEachArea In rngSelected.Areas
        For Each eachCell In rngEachArea
            If eachCell.Column > lMaxCol Then Exit For

            If IsNumeric(eachCell.Value) Then
                rgData.AutoFilter Field:=eachCell.Column _
                                , Criteria1:=eachCell.Value _
                                , Operator:=xlAnd
            Else
                rgData.AutoFilter Field:=eachCell.Column _
                                , Criteria1:="=*" & eachCell.Value & "*" _
                                , Operator:=xlAnd
            End If
        Next
    Next

    End
End Sub
Sub Sub_RemoveFilterForAcitveSheet(Optional ByVal asDegree As String = "SHOW_ALL_DATA")
    Call fRemoveFilterForSheet(ActiveSheet, asDegree)
End Sub

Sub sub_SortBySelectColumn()
    Dim sSelectContent As String
    Dim lSelectCol As Long
    sSelectContent = ActiveCell.Value
    lSelectCol = ActiveCell.Column

    Call Sub_RemoveFilterForAcitveSheet
    Call fSortDataInSheetSortSheetData(ActiveSheet, Array(ActiveCell.Column))

    Dim rgFound As Range
    Set rgFound = fFindInWorksheet(ActiveSheet.Columns(lSelectCol), sSelectContent, True, True)

    If Not rgFound Is Nothing Then rgFound.Select
    Set rgFound = Nothing
End Sub

Sub sub_SortBySelectedCells()
    Dim rngSelected As Range
    Dim sFirstValue
    Dim arrSortCol()

    Set rngSelected = Selection
'    If fIfSelectedMoreThanOneRow(rngSelected) Then
'        fMsgBox "不能选多行，只能选一行。"
'        End
'    End If

    sFirstValue = rngSelected.Cells(1, 1).Value

    Call Sub_RemoveFilterForAcitveSheet

    Dim lMaxRow As Long
    Dim lMaxCol As Long
    lMaxRow = fGetValidMaxRow(ActiveSheet)
    lMaxCol = fGetValidMaxCol(ActiveSheet)

    Dim eachCol As Integer
    Dim rgData As Range
    Set rgData = fGetRangeByStartEndPos(ActiveSheet, 1, 1, lMaxRow, lMaxCol)

    Dim rngEachArea As Range
    'Dim eachCell As Range
    Dim i As Integer
    Dim j As Integer

    i = 0
    For Each rngEachArea In rngSelected.Areas
        For j = rngEachArea.Column To rngEachArea.Column + rngEachArea.Columns.Count - 1
            If j > lMaxCol Then Exit For

            i = i + 1
            ReDim Preserve arrSortCol(i)
            arrSortCol(i) = j
        Next
'        For Each eachCell In rngEachArea
'            If eachCell.Column > lMaxCol Then Exit For
'
'            i = i + 1
'            ReDim Preserve arrSortCol(i)
'            arrSortCol(i) = eachCell.Column
'        Next
    Next

    If i > 0 Then Call fSortDataInSheetSortSheetData(ActiveSheet, arrSortCol)

    Dim rgFound As Range
    Set rgFound = fFindInWorksheet(ActiveSheet.Cells, CStr(sFirstValue), True, True)

    Debug.Print rngSelected.Cells(1, 1).Value
    If Not rgFound Is Nothing Then rgFound.Select
    Set rgFound = Nothing
    End
End Sub

Function fSetFilterForSheet(sht As Worksheet, aColToFilter, aFilterValue)
    If Not (IsArray(aColToFilter) And IsArray(aFilterValue) _
    Or Not IsArray(aColToFilter) And Not IsArray(aFilterValue)) Then
        fErr "param aColToFilter and aFilterValue must be array or non-array at the same time."
    End If

'    Dim myData As DataObject
'    Dim sOriginText As String
'
'    Set myData = New DataObject
'    myData.GetFromClipboard
'    On Error Resume Next
'    sOriginText = myData.GetText()
'    On Error GoTo 0

    fKeepCopyContent

    Dim lMaxRow As Long
    Dim lMaxCol As Long
    lMaxCol = sht.Cells(1, 1).End(xlToRight).Column
    lMaxRow = fGetValidMaxRow(sht)

    If sht.AutoFilterMode Then  'auto filter
        sht.AutoFilter.ShowAllData
    Else
        fGetRangeByStartEndPos(sht, 1, 1, 1, lMaxCol).AutoFilter
    End If

    Dim i As Integer
    If IsArray(aColToFilter) Then
        For i = LBound(aColToFilter) To UBound(aColToFilter)
            If Len(Trim(CStr(aFilterValue(i)))) > 0 Then _
            fGetRangeByStartEndPos(sht, 1, 1, lMaxRow, lMaxCol).AutoFilter _
                Field:=aColToFilter(i), Criteria1:="=*" & aFilterValue(i) & "*", Operator:=xlAnd
        Next
    Else
        fGetRangeByStartEndPos(sht, 1, 1, lMaxRow, lMaxCol).AutoFilter _
                Field:=aColToFilter, Criteria1:="=*" & aFilterValue & "*", Operator:=xlAnd
    End If

    'Call fGotoCell(sht.Range("A2"), True)
'    On Error Resume Next
'    myData.SetText sOriginText
'    'If fNzero(sOriginText) Then myData.SetText sOriginText
'    myData.PutInClipboard
'    On Error GoTo 0
'    Set myData = Nothing

    fCopyFromKept

'            If IsNumeric(eachCell.Value) Then
'                rgData.AutoFilter Field:=eachCell.Column _
'                                , Criteria1:=">=" & eachCell.Value _
'                                , Operator:=xlAnd
'            Else
'                rgData.AutoFilter Field:=eachCell.Column _
'                                , Criteria1:="=*" & eachCell.Value & "*" _
'                                , Operator:=xlAnd
'            End If
End Function

Function fCopyFilteredDataToRange(sht As Worksheet, colFiltered As Integer)
'    Dim myData As DataObject
'    Dim sOriginText As String
'    Set myData = New DataObject
'    myData.GetFromClipboard
'    On Error Resume Next
'    sOriginText = myData.GetText()
'    On Error GoTo 0

    fKeepCopyContent

    shtDataStage.Columns("A").ClearContents

    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow < 2 Then Exit Function

    fGetRangeByStartEndPos(sht, 2, CLng(colFiltered), lMaxRow, CLng(colFiltered)).Copy
    shtDataStage.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

'    On Error Resume Next
'    myData.SetText sOriginText
'    myData.PutInClipboard
'    On Error GoTo 0
'
'    Set myData = Nothing
    fCopyFromKept
End Function

Function fSheetIsNotVisible(sht As Worksheet) As Boolean
    fSheetIsNotVisible = (sht.Visible <> xlSheetVisible)
End Function

Function fSheetIsVisible(sht As Worksheet) As Boolean
    fSheetIsVisible = (sht.Visible = xlSheetVisible)
End Function

Sub subMain_ListAllSheets()

    Dim shtEach As Worksheet

    For Each shtEach In ThisWorkbook.Worksheets
        Debug.Print shtEach.CodeName & DELIMITER & shtEach.Name
    Next
End Sub

Function fDeleteRemoveDataFormatFromSheetLeaveHeader(ByRef shtParam As Worksheet)
    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(shtParam)

    If lMaxRow > 2 Then
        With fGetRangeByStartEndPos(shtParam, 2, 1, lMaxRow, fGetValidMaxCol(shtParam))
            .ClearContents
            '.ClearFormats
            .ClearComments
            .ClearNotes
            .ClearOutline
        End With
    End If
End Function

Function fOpenFileSelectDialogAndSetToSheetRange(rngAddrOrName As String _
                            , Optional asFileFilters As String = "" _
                            , Optional asTitle As String = "" _
                            , Optional shtParam As Worksheet)
    Dim sFile As String

    If shtParam Is Nothing Then Set shtParam = shtMenu

    sFile = fSelectFileDialog(Trim(shtParam.Range(rngAddrOrName).Value), , asTitle)
    If Len(sFile) > 0 Then shtParam.Range(rngAddrOrName).Value = sFile
End Function

Function fFindInWorksheet(rngToFindIn As Range, sWhatToFind As String _
                    , Optional abNotFoundThenError As Boolean = True _
                    , Optional abAllowMultiple As Boolean = False) As Range
    If Len(Trim(sWhatToFind)) <= 0 Then fErr "Wrong param sWhatToFind to fFindInWorksheet " & sWhatToFind

    Dim rngOut  As Range
    Dim rngFound As Range
    Dim lFoundCnt As Long
    Dim sFirstAddress As String

    Set rngFound = rngToFindIn.Find(What:=sWhatToFind _
                                    , after:=rngToFindIn.Cells(rngToFindIn.Rows.Count, rngToFindIn.Columns.Count) _
                                    , LookIn:=xlValues _
                                    , LookAt:=xlWhole _
                                    , SearchOrder:=xlByRows _
                                    , SearchDirection:=xlNext _
                                    , MatchCase:=False _
                                    , MatchByte:=False)
    Set rngOut = rngFound

    If rngFound Is Nothing Then
        If abNotFoundThenError Then
            fErr """" & sWhatToFind & """ cannot be found in sheet " & rngToFindIn.Parent.Name & "[" & rngToFindIn.Address & "], pls check your program."
        Else
            GoTo exit_function
        End If
    Else
        If Not abAllowMultiple Then
            sFirstAddress = rngFound.Address
            lFoundCnt = 1

            Do While True
                Set rngFound = rngToFindIn.Find(What:=sWhatToFind _
                                            , after:=rngFound _
                                            , LookIn:=xlValues _
                                            , LookAt:=xlWhole _
                                            , SearchOrder:=xlByRows _
                                            , SearchDirection:=xlNext _
                                            , MatchCase:=False _
                                            , MatchByte:=False)
                If rngFound Is Nothing Then Exit Do
                If rngFound.Address = sFirstAddress Then Exit Do

                lFoundCnt = lFoundCnt + 1
            Loop

            If lFoundCnt > 1 Then
                fErr lFoundCnt & " copies of """ & sWhatToFind & """ were found in sheet " & rngToFindIn.Parent.Name & ", pls check your program."
            End If
        End If
    End If
exit_function:
    Set fFindInWorksheet = rngOut
    Set rngOut = Nothing
    Set rngFound = Nothing
End Function

Function fGetRangeByStartEndPos(shtParam As Worksheet, alStartRow As Long, alStartCol As Long, alEndRow As Long, alEndCol As Long) As Range
    If alStartRow > alEndRow Then fErr "alStartRow > alEndRow in function fGetRangeByStartEndPos, please change your program to add the check logic before calling fGetRangeByStartEndPos"
    With shtParam
        Set fGetRangeByStartEndPos = .Range(.Cells(alStartRow, alStartCol), .Cells(alEndRow, alEndCol))
    End With
End Function

Function fReadRangeDatatoArrayByStartEndPos(shtParam As Worksheet, alStartRow As Long, alStartCol As Long, alEndRow As Long, alEndCol As Long) As Variant
    If alStartRow > alEndRow Then
        fReadRangeDatatoArrayByStartEndPos = Array()
    Else
        fReadRangeDatatoArrayByStartEndPos = fReadRangeDataToArray(fGetRangeByStartEndPos(shtParam, alStartRow, alStartCol, alEndRow, alEndCol))
    End If
End Function

Function fReadRangeDataToArray(rngParam As Range) As Variant
    Dim arrOut()

    If fRangeIsSingleCell(rngParam) Then
        ReDim arrOut(1 To 1, 1 To 1)
        arrOut(1, 1) = rngParam.Value
    Else
        arrOut = rngParam.Value
    End If

    fReadRangeDataToArray = arrOut
    Erase arrOut
End Function

Function fSetSpecifiedConfigCellValue(shtConfig As Worksheet, asTag As String, asRtnCol As String, asCriteria As String _
                                , sValue As String _
                                , Optional bDevUatProd As Boolean = False _
                                )
    Dim sAddr As String
    sAddr = fGetSpecifiedConfigCellAddress(shtConfig, asTag, asRtnCol, asCriteria, False, bDevUatProd)
    shtConfig.Range(sAddr).Value = sValue
End Function
Function fGetSpecifiedConfigCellValue(shtConfig As Worksheet, asTag As String, asRtnCol As String, asCriteria As String _
                                , Optional bDevUatProd As Boolean = False _
                                )
    Dim sAddr As String
    sAddr = fGetSpecifiedConfigCellAddress(shtConfig, asTag, asRtnCol, asCriteria, False, bDevUatProd)
    fGetSpecifiedConfigCellValue = shtConfig.Range(sAddr).Value
End Function
Function fGetSpecifiedConfigCellAddress(shtConfig As Worksheet, asTag As String, asRtnCol As String _
                                , asCriteria As String _
                                , Optional bAllowMultiple As Boolean = False _
                                , Optional bDevUatProd As Boolean = False _
                                )
    'asCriteria: colA=Value01, colB=Value02
    Dim arrColNames()
    Dim arrColValues()
    Dim iRtnColIndex As Integer
    Dim iEnvColIndex As Integer

    Call fSplitDataCriteria(asCriteria, arrColNames, arrColValues)

    iRtnColIndex = fEnlargeArayWithValue(arrColNames, asRtnCol)

    'DEV/UAT/PROD must be put at the end of arrColsName, since fFindMatchDataInArrayWithCriteria will read it.
    If bDevUatProd Then
        iEnvColIndex = fEnlargeArayWithValue(arrColNames, "DEV/UAT/PROD")
    End If

    Dim lConfigStartRow As Long _
        , lConfigStartCol As Long _
        , lConfigEndRow As Long _
        , lOutConfigHeaderAtRow As Long _
        , bNetValues As Boolean
    Dim arrConfigData()
    Dim arrColsIndex()

    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtConfig, arrColsName:=arrColNames _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Dim lMatchRow As Long
    Dim sErr As String
    lMatchRow = fFindMatchDataInArrayWithCriteria(arrConfigData, arrColsIndex, arrColValues, bAllowMultiple, sErr, bDevUatProd)

    If lMatchRow < 0 Then
        fErr sErr & " with criteria " & vbCr & asCriteria & vbCr & "gsEnv: " & gsEnv & vbCr & shtConfig.Name & vbCr & asTag
    End If

    fGetSpecifiedConfigCellAddress = shtConfig.Cells(lOutConfigHeaderAtRow + lMatchRow, lConfigStartCol + arrColsIndex(iRtnColIndex) - 1).Address(external:=True)
End Function

Function fFindMatchDataInArrayWithCriteria(arr(), arrColsIndex(), arrColValues() _
                                        , bAllowMultiple As Boolean _
                                        , ByRef asErrmsg As String _
                                , Optional bDevUatProd As Boolean = False _
                                        ) As Long
'-1:
' -2: more than 1 matched
' -3: no match
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Integer
    Dim bAllColAreSame As Boolean
    Dim lMatchCnt As Long
    Dim lOut As Long
    Dim sEachEnv As String

    If bDevUatProd And fZero(gsEnv) Then fErr "bDevUatProd is true but gsenv = blank"

    asErrmsg = ""
    lOut = -1
    lMatchCnt = 0
    For lEachRow = LBound(arr, 1) To UBound(arr, 1)
        If fArrayRowIsBlankHasNoData(arr, lEachRow) Then GoTo next_row

        If bDevUatProd Then
            sEachEnv = arr(lEachRow, arrColsIndex(UBound(arrColsIndex)))
            If Not (sEachEnv = gsEnv Or sEachEnv = "SHARED") Then GoTo next_row
        End If

        bAllColAreSame = True
        For i = LBound(arrColValues) To UBound(arrColValues)
            lEachCol = arrColsIndex(i)

            If Trim(CStr(arr(lEachRow, lEachCol))) <> arrColValues(i) Then
                bAllColAreSame = False
                GoTo next_row
            End If
        Next

        If bAllColAreSame Then
            lMatchCnt = lMatchCnt + 1
            lOut = lEachRow

            If bAllowMultiple Then GoTo exit_fun
        End If
next_row:
    Next

    If lMatchCnt > 1 Then
        If Not bAllowMultiple Then
            lOut = -2
            asErrmsg = lMatchCnt & " records were matched "
        End If
    ElseIf lMatchCnt <= 0 Then
        lOut = -3
        asErrmsg = "No record were matched "
    End If
exit_fun:
    fFindMatchDataInArrayWithCriteria = lOut
End Function

Function fSplitDataCriteria(asCriteria As String, ByRef arrColNames(), ByRef arrColValues())
    'asCriteria: colA=Value01, colB=Value02
    Dim arrCriteria
    Dim sCol As String
    Dim sValue As String
    Dim i As Integer
    Dim sEachCriteria As String

    arrCriteria = Split(asCriteria, ",")

    ReDim arrColNames(LBound(arrCriteria) To UBound(arrCriteria))
    ReDim arrColValues(LBound(arrCriteria) To UBound(arrCriteria))

    For i = LBound(arrCriteria) To UBound(arrCriteria)
        sEachCriteria = Trim(arrCriteria(i))    ' colA=Value01

        sCol = Trim(Split(sEachCriteria, "=")(0))
        sValue = Trim(Split(sEachCriteria, "=")(1))

        arrColNames(i) = sCol
        arrColValues(i) = sValue
    Next

    Erase arrCriteria
End Function

Function fWriteArray2Sheet(sht As Worksheet, arrData, Optional lStartRow As Long = 2, Optional lStartCol As Long = 1)
    If fArrayIsEmptyOrNoData(arrData) Then Exit Function

    If fGetArrayDimension(arrData) <> 2 Then
        fErr "Wrong array to paste to sheet: fGetArrayDimension(arrData) <> 2"
    End If

    sht.Cells(lStartRow, lStartCol).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData
End Function

Function fAppendArray2Sheet(sht As Worksheet, ByRef arrData, Optional lStartCol As Long = 1, Optional bEraseArray As Boolean = True)
    If fArrayIsEmptyOrNoData(arrData) Then Exit Function

'    If fGetArrayDimension(arrData) <> 2 Then
'        fErr "Wrong array to paste to sheet: fGetArrayDimension(arrData) <> 2"
'    End If

    Dim lFromRow As Long
    lFromRow = fGetValidMaxRow(sht) + 1

    sht.Cells(lFromRow, lStartCol).Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData
    If bEraseArray Then Erase arrData
End Function

Function fAutoFilterAutoFitSheet(sht As Worksheet, Optional alMaxCol As Long = 0 _
                                , Optional ColumnWidthAuto As Boolean = True)

    Dim lMaxCol As Long

    If alMaxCol > 0 Then
        lMaxCol = alMaxCol
    Else
        lMaxCol = fGetValidMaxCol(sht)
    End If

    If lMaxCol <= 0 Then Exit Function

    If sht.AutoFilterMode Then sht.AutoFilterMode = False

    fGetRangeByStartEndPos(sht, 1, 1, 1, lMaxCol).AutoFilter

    If ColumnWidthAuto Then sht.Cells.EntireColumn.AutoFit
    sht.Cells.EntireRow.AutoFit
End Function

Function fFreezeSheet(sht As Worksheet, Optional alSplitCol As Long = 0, Optional alSplitRow As Long = 1)
    sht.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = alSplitCol
    ActiveWindow.SplitRow = alSplitRow
    ActiveWindow.FreezePanes = True
End Function

Function fRemoveFilterForAllSheets(Optional wb As Workbook, Optional ByVal asDegree As String = "SHOW_ALL_DATA")
    If wb Is Nothing Then Set wb = ThisWorkbook

    On Error GoTo error_handling
    Dim sht As Worksheet
    For Each sht In wb.Worksheets
        Call fRemoveFilterForSheet(sht, asDegree)
    Next

error_handling:
    If Err.Number <> 0 Then
        MsgBox sht.Name & vbCr & Err.Description
    End If
    Set sht = Nothing
End Function
Function fDeleteBlankRowsFromAllSheets(Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim sht As Worksheet
    For Each sht In wb.Worksheets
        If sht.CodeName = shtSysConf.CodeName Then GoTo next_sht
        If sht.CodeName = shtMainMenu.CodeName Then GoTo next_sht
        If sht.CodeName = shtMenu.CodeName Then GoTo next_sht
        If sht.CodeName = shtMenuCompInvt.CodeName Then GoTo next_sht
        Call fDeleteBlankRowsFromSheet(sht)
next_sht:
    Next

    Set sht = Nothing
End Function
Function fDeleteBlankRowsFromSheet(sht As Worksheet)
    Dim lUsedRangMaxRow As Long
    Dim lValidMaxRow As Long

    lUsedRangMaxRow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
    lValidMaxRow = fGetValidMaxRow(sht)

    If lUsedRangMaxRow > lValidMaxRow Then
        sht.Rows((lValidMaxRow + 1) & ":" & lUsedRangMaxRow).Delete shift:=xlUp
    End If
End Function
Function fRemoveFilterForSheet(sht As Worksheet, Optional ByVal asDegree As String = "SHOW_ALL_DATA")
    asDegree = UCase(Trim(asDegree))

    Dim rgActiveCell As Range
    Set rgActiveCell = ActiveCell

    If sht.FilterMode Then  'advanced filter
        sht.ShowAllData
    End If

    If sht.AutoFilterMode Then  'auto filter
        If fZero(asDegree) Or asDegree = "SHOW_ALL_DATA" Then
            sht.AutoFilter.ShowAllData
        Else
            sht.AutoFilterMode = False
        End If
    End If

    rgActiveCell.Select
    Set rgActiveCell = Nothing

    'Call fGotoCell(sht.Range("A2"), True)
End Function

Function fActiveXControlExistsInSheet(sht As Worksheet, asControlName As String, ByRef objOut As Object) As Boolean
    Dim bOut As Boolean
    bOut = False

    On Error GoTo err_exit
    Set objOut = CallByName(sht, asControlName, VbGet)
    bOut = True

err_exit:
    Err.Clear
    fActiveXControlExistsInSheet = bOut
End Function
Function fActiveXControlExistsInSheet2(sht As Worksheet, asControlName As String, ByRef objOut As Object) As Boolean
    Dim obj As Object
    Dim bOut As Boolean
    bOut = False

    For Each obj In sht.DrawingObjects
        If obj.Name = asControlName Then
            Set objOut = obj
            bOut = True
            Exit For
        End If
    Next

err_exit:
    Set obj = Nothing
    fActiveXControlExistsInSheet2 = bOut
End Function

Function fSheetExists(asShtName As String, Optional ByRef shtOut As Worksheet, Optional wb As Workbook) As Boolean
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
    fSheetExists = bOut
End Function

Function fSheetExistsByCodeName(asShtCodeName As String, Optional ByRef shtOut As Worksheet, Optional wb As Workbook _
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
    fSheetExistsByCodeName = bOut

    If abPromptErrMsgIfNotFound Then
        If Not bOut Then fErr "The workbook " & wb.Name & " does not have a sheet whose code name is " & asShtCodeName
    End If
End Function

Function fGetSheetByCodeName(asShtCodeName As String, Optional wb As Workbook) As Worksheet
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

    Set fGetSheetByCodeName = shtOut
    Set shtOut = Nothing
    Set sht = Nothing
End Function
Function fIfExcelFileOpenedToCloseIt(sExcelFileFullPath As String)
    Dim wbTemp As Workbook

    If fExcelFileIsOpen(sExcelFileFullPath, wbTemp) Then
        fGetFSO
        'sExcelFileFullPath = gFSO.GetFile(sExcelFileFullPath).Path
        sExcelFileFullPath = fCheckPath(sExcelFileFullPath)

        If UCase(wbTemp.FullName) = UCase(sExcelFileFullPath) Then
            fErr "Excel File is open, pleae close it first." & vbCr & fGetFileBaseName(sExcelFileFullPath)
        Else
            fErr "Another file with the same name """ & fGetFileBaseName(sExcelFileFullPath) & """ is open, please close it first."
        End If
    End If

    Set wbTemp = Nothing
End Function

Function fSortDataInSheetSortSheetData(sht As Worksheet, arrSortByCols, Optional arrAscend)
'arrAscendDescend : array(true, false, true)
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    Dim rgToSort As Range
    Dim rgSortBy As Range
    Dim iSortOrder As XlSortOrder
    Dim i As Long
    Dim lSortCol As Long

    lMaxRow = fGetValidMaxRow(sht)
    lMaxCol = fGetValidMaxCol(sht)

    If lMaxRow <= 0 Or lMaxCol <= 0 Then Exit Function

    Set rgToSort = fGetRangeByStartEndPos(sht, 1, 1, lMaxRow, lMaxCol)

    sht.Sort.SortFields.Clear

    If IsArray(arrSortByCols) Then
        For i = LBound(arrSortByCols) To UBound(arrSortByCols)
            lSortCol = arrSortByCols(i)

            If Not IsMissing(arrAscend) Then
                If arrAscend(i) Then
                    iSortOrder = xlAscending
                Else
                    iSortOrder = xlDescending
                End If
            Else
                iSortOrder = xlAscending
            End If

            Set rgSortBy = fGetRangeByStartEndPos(sht, 2, lSortCol, lMaxRow, lSortCol)
            sht.Sort.SortFields.Add Key:=rgSortBy, SortOn:=xlSortOnValues _
                    , Order:=iSortOrder, DataOption:=xlSortNormal
        Next
    Else
        lSortCol = arrSortByCols

        If Not IsMissing(arrAscend) Then
            If arrAscend Then
                iSortOrder = xlAscending
            Else
                iSortOrder = xlDescending
            End If
        Else
            iSortOrder = xlAscending
        End If

        Set rgSortBy = fGetRangeByStartEndPos(sht, 2, lSortCol, lMaxRow, lSortCol)
        sht.Sort.SortFields.Add Key:=rgSortBy, SortOn:=xlSortOnValues _
                , Order:=iSortOrder, DataOption:=xlSortNormal
    End If

    sht.Sort.SetRange rgToSort
    sht.Sort.Header = xlYes
    sht.Sort.MatchCase = False
    sht.Sort.Orientation = xlTopToBottom
    sht.Sort.SortMethod = xlPinYin
    sht.Sort.Apply

    Set rgToSort = Nothing
    Set rgSortBy = Nothing
End Function

Function fProtectSheetAndAllowEdit(sht As Worksheet, Optional rngAllowEdit As Range _
                    , Optional lMaxRow As Long = 0, Optional lMaxCol As Long = 0 _
                    , Optional bLockColor As Boolean = True)

    If lMaxRow <= 0 Then lMaxRow = fGetValidMaxRow(sht)

    If lMaxRow < 2 Then Exit Function

    If lMaxCol <= 0 Then lMaxCol = fGetValidMaxCol(sht)

    Dim rgUsed As Range
    Set rgUsed = fGetRangeByStartEndPos(sht, 2, 1, lMaxRow, lMaxCol)

    If bLockColor Then
        rgUsed.Interior.Color = 13553360

        If Not rngAllowEdit Is Nothing Then
            'rngAllowEdit.Interior.Color = RGB(255, 255, 255)
            rngAllowEdit.Interior.Pattern = xlNone
            rngAllowEdit.Interior.TintAndShade = 0
            rngAllowEdit.Interior.PatternTintAndShade = 0
        End If
    End If

    sht.Cells.Locked = True

    If Not rngAllowEdit Is Nothing Then
        rngAllowEdit.Locked = False
        rngAllowEdit.FormulaHidden = False
    End If

    Set rgUsed = Nothing

    sht.Protect userinterfaceonly:=True, Password:=PW_PROTECT_SHEET _
        , DrawingObjects:=True _
        , Contents:=True _
        , Scenarios:=True _
        , AllowFormattingCells:=True _
        , AllowFormattingColumns:=True _
        , AllowFormattingRows:=True _
        , AllowInsertingColumns:=False _
        , AllowInsertingRows:=False _
        , AllowInsertingHyperlinks:=True _
        , AllowDeletingColumns:=False _
        , AllowDeletingRows:=False _
        , AllowSorting:=True _
        , AllowFiltering:=True _
        , AllowUsingPivotTables:=False

End Function

Function fUnProtectSheet(sht As Worksheet)
    sht.Unprotect Password:=PW_PROTECT_SHEET
End Function

Function fSetValidationListForRange(rngParam As Range, asValueListOrExternalAddr As String)
'asValueListOrExternalAddr
' 1) Formula1:="=$K$5:$K$21"    --> range().address( external:=true)
' 2) Formula1:="a,b,c,d,e,f,g"  --> "a, b, c, d, e, f"

    With rngParam.Validation
        .Delete
        .Add Type:=xlValidateList _
            , AlertStyle:=xlValidAlertStop _
            , Operator:=xlBetween _
            , Formula1:=asValueListOrExternalAddr
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Function

Function fSetValidationForDateRange(rngParam As Range)
    With rngParam.Validation
        .Delete
        .Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="1/1/2001", Formula2:="12/31/2099"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "输入错误！"
        .ErrorTitle = "输入错误！"
        .InputMessage = ""
        .ErrorMessage = "请输入正确的日期"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    Range("E6").Select
End Function
Function fSetValidationForNumberRange(rngParam As Range, aNumMin As Double, aNumMax As Double)
    With rngParam.Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween, Formula1:=aNumMin, Formula2:=aNumMax
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "输入错误！"
        .ErrorTitle = "输入错误！"
        .InputMessage = ""
        .ErrorMessage = "只允许输入" & aNumMin & "到" & aNumMax & "的数值。"
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Function

Function fModifyMoveActiveXButtonOnSheet(rngToPlaceTheButton As Range, sBtnTechName As String _
                                , Optional dblOffsetLeft As Double = 0, Optional dblOffsetTop As Double = 0 _
                                , Optional dblWidth As Double = 0, Optional dblHeight As Double = 0 _
                                , Optional lBackColor As Long = 0 _
                                , Optional lForeColor As Long = 0)
    Dim sht As Worksheet
    Set sht = rngToPlaceTheButton.Parent

    Dim obj As Object
    Dim eachObj As Object

    For Each eachObj In sht.OLEObjects
        If eachObj.Name = sBtnTechName Then
            Set obj = eachObj
            Exit For
        End If
    Next

    If obj Is Nothing Then fErr "Button " & sBtnTechName & " cannot be found in sheet " & sht.Name

    If dblWidth = 0 Then dblWidth = obj.Width
    If dblHeight = 0 Then dblHeight = obj.Height

    Dim offsetLeft As Double
    If dblOffsetLeft = 0 Then
        offsetLeft = (rngToPlaceTheButton.Width - dblWidth) / 2 - 2
    Else
        offsetLeft = dblOffsetLeft
    End If

    Dim offsetTop As Double
    If dblOffsetTop = 0 Then
        offsetTop = (rngToPlaceTheButton.Width - dblWidth) / 2 - 2
    Else
        offsetTop = dblOffsetTop
    End If

    obj.Left = rngToPlaceTheButton.Left + offsetLeft
    obj.Top = rngToPlaceTheButton.Top + offsetTop
    obj.Width = dblWidth
    obj.Height = dblHeight

    If lBackColor <> 0 Then obj.Object.BackColor = lBackColor
    If lForeColor <> 0 Then obj.Object.ForeColor = lForeColor

    Set obj = Nothing
    Set eachObj = Nothing
End Function

Function fSortDataInSheetSortSheetDataByFileSpec(asFileTag As String, arrSortByColNames, Optional arrAscend _
                                    , Optional shtData As Worksheet)
    Dim sFileSpecTag As String
    Dim shtToRead As Worksheet
    Dim dictColIndex As Dictionary
    Dim arrSortByCols()

    sFileSpecTag = fGetInputFileFileSpecTag(asFileTag)

    If shtData Is Nothing Then
        Set shtToRead = fGetInputFileSheetAfterLoadingToThisWorkBook(asFileTag)
    Else
        Set shtToRead = shtData
    End If

    Set dictColIndex = fReadInputFileSpecConfigItem(asFileTag, "LETTER_INDEX", shtToRead)

    ReDim arrSortByCols(LBound(arrSortByColNames) To UBound(arrSortByColNames))

    Dim i As Integer
    For i = LBound(arrSortByColNames) To UBound(arrSortByColNames)
        If Not dictColIndex.Exists(arrSortByColNames(i)) Then fErr "Column name does not exists " & arrSortByColNames(i)
        arrSortByCols(i) = dictColIndex(arrSortByColNames(i))
    Next

    Call fSortDataInSheetSortSheetData(shtToRead, arrSortByCols, arrAscend)

    Set dictColIndex = Nothing
    Set shtToRead = Nothing
End Function

Function fGetSelectedRowCount(rngParam As Range)
'    Dim eachArea As Range
'    Dim lRowCnt As Long
'
'    lRowCnt = 0
'    For Each eachArea In rngParam.Areas
'        lRowCnt = lRowCnt + eachArea.Rows.Count
'    Next
'    Set eachArea = Nothing
'
'    fGetSelectedRowCount = lRowCnt
End Function

Function fIfSelectedMoreThanOneRow(rngParam As Range) As Boolean
    Dim eachArea As Range
    Dim lRowNoSaved As Long
    Dim bOut As Boolean

    bOut = False
    lRowNoSaved = 0
    For Each eachArea In rngParam.Areas
        If eachArea.Rows.Count > 1 Then
            bOut = True
            Exit For
        Else
            If lRowNoSaved = 0 Then
                lRowNoSaved = eachArea.Row
            Else
                If eachArea.Row <> lRowNoSaved Then
                    bOut = True
                    Exit For
                End If
            End If
        End If
    Next
    Set eachArea = Nothing

    fIfSelectedMoreThanOneRow = bOut
End Function

Function fVeryHideSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetVeryHidden
End Function

Function fHideSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetHidden
End Function

Function fShowSheet(ByRef sht As Worksheet)
    sht.Visible = xlSheetVisible
    sht.Activate
End Function

Function fClearContentLeaveHeader(shtParam As Worksheet, Optional alHeaderByRow As Long = 1)
    fRemoveFilterForSheet shtParam

    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(shtParam)

    Dim lDataStartRow As Long
    lDataStartRow = alHeaderByRow + 1

    If lMaxRow > lDataStartRow Then
        On Error Resume Next
        With fGetRangeByStartEndPos(shtParam, lDataStartRow, 1, lMaxRow, fGetValidMaxCol(shtParam))
            .ClearContents
          '  .ClearFormats
            .ClearComments
            .ClearNotes
            .ClearOutline
        End With
        If Err.Number <> 0 Then Err.Clear
    End If
End Function

Function fCheckFileNameLength(ByRef asFileFullPath As String) As Boolean
    Dim bOut As Boolean

    asFileFullPath = Trim(asFileFullPath)

    bOut = CBool(Len(asFileFullPath) > 218)
    If bOut Then fErr "The file path is longer than 218, which is not allowed by Windows system."
End Function
Function fGetExcelFormatByFileExtension(ByVal asFileExtension As String) As XlFileFormat
    Dim iFileFormat As XlFileFormat

    Select Case "." & LCase(asFileExtension)
        Case ".csv"
            iFileFormat = xlCSV
        Case ".xls"
            iFileFormat = xlExcel8
        Case ".xlsx"
            iFileFormat = xlOpenXMLWorkbook
        Case ".xlsb"
            iFileFormat = xlOpenXMLWorkbookMacroEnabled
        Case ".txt"
            iFileFormat = xlCurrentPlatformText
        Case ".prn"
            iFileFormat = xlTextPrinter
        Case Else
            fErr "The file extension is not covered in function fGetExcelFormatByFileExtension, you need to revise this function."
    End Select
    fGetExcelFormatByFileExtension = iFileFormat
End Function
Function fGetExcelFormatByFileName(ByVal asFileFullPath As String) As XlFileFormat
    Dim sFileExt As String

    sFileExt = fGetFileExtension(asFileFullPath)

    fGetExcelFormatByFileName = fGetExcelFormatByFileExtension(sFileExt)
End Function
Function fCreateNewWorkbook(ByVal asFileFullPath As String, Optional ByVal asNewSheetName As String = "" _
                          , Optional ByVal asPassword As String = "") As Workbook
    Dim wbOut As Workbook

    fCheckFileNameLength asFileFullPath

    If Len(Trim(asNewSheetName)) <= 0 Then asNewSheetName = fGetFileNetName(asFileFullPath)

    Dim iOrig As Integer
    iOrig = Application.SheetsInNewWorkbook
    Application.SheetsInNewWorkbook = 1
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    wbOut.Worksheets(1).Name = asNewSheetName
    Application.SheetsInNewWorkbook = iOrig

    wbOut.BuiltinDocumentProperties("Keywords") = "RESTICTED"
    wbOut.BuiltinDocumentProperties("Comments") = "RESTICTED"

    ActiveWindow.DisplayGridlines = False

    Dim iFileFormat As XlFileFormat
    iFileFormat = fGetExcelFormatByFileName(asFileFullPath)
    wbOut.SaveAs Filename:=asFileFullPath, FileFormat:=iFileFormat, Password:=asPassword

    Set fCreateNewWorkbook = wbOut
    Set wbOut = Nothing
End Function

Function fCopySingleSheet2WorkBook(shtSource As Worksheet, wbCopyTo As Workbook, Optional ByVal asNewSheetName As String = "")
    Dim wbFrom As Workbook
    Set wbFrom = shtSource.Parent

    Dim sToShtName As String
    sToShtName = IIf(Len(Trim(asNewSheetName)) <= 0, shtSource.Name, asNewSheetName)

    If wbFrom.FullName = wbCopyTo.FullName Then
        If Len(Trim(asNewSheetName)) <= 0 Then
            fErr "Copying is withing the same workbook, you must specify a different sheet name by parameter asNewSheetName"
        End If
        If UCase(shtSource.Name) = UCase(asNewSheetName) Then
            fErr "Copying is withing the same workbook, parameter asNewSheetName cannot be same as the source sheet name " & "sheet name: " & asNewSheetName
        End If
    End If

    If fSheetExists(sToShtName, , wbCopyTo) Then
        fErr "There is already a sheet with the same name in workbook " & wbCopyTo.Name & vbCr & "sheet name: " & sToShtName
    End If

    Dim xlOrig As XlSheetVisibility
    xlOrig = shtSource.Visible
    shtSource.Visible = xlSheetVisible
    shtSource.Copy wbCopyTo.Worksheets(wbCopyTo.Worksheets.Count)
    shtSource.Visible = xlOrig

    wbCopyTo.ActiveSheet.Name = sToShtName
    ActiveWindow.DisplayGridlines = False

    Set wbFrom = Nothing
End Function

Function fCopySingleSheet2NewWorkbookFile(shtSource As Worksheet, asNewFileName As String, Optional ByVal asNewSheetName As String = "") As Workbook
    Dim sUniqueStr As String
    Dim wbOut As Workbook

    sUniqueStr = fGenRandomUniqueString()

    Call fIfExcelFileOpenedToCloseIt(asNewFileName)

    Call fDeleteFile(asNewFileName)

    Set wbOut = fCreateNewWorkbook(asNewFileName, sUniqueStr)

    Call fCopySingleSheet2WorkBook(shtSource, wbOut, asNewSheetName)
    Call fDeleteSheetIfExists(sUniqueStr, wbOut)

    Set fCopySingleSheet2NewWorkbookFile = wbOut
    Set wbOut = Nothing
End Function

Function fDeleteRowsFromSheetLeaveHeader(ByRef sht As Worksheet, Optional lHeaderByRow As Long = 1)
    Dim lMaxRow As Long
    Dim iOrigVisibility As XlSheetVisibility

    iOrigVisibility = sht.Visible
    sht.Visible = xlSheetVisible

    Call fRemoveFilterForSheet(sht)

    lMaxRow = sht.Range("A1").SpecialCells(xlCellTypeLastCell).Row

    If lMaxRow > lHeaderByRow Then
        sht.Rows(lHeaderByRow + 1 & ":" & lMaxRow).Delete shift:=xlUp
        Application.Goto sht.Cells(lHeaderByRow + 1, 1), True
    End If

    sht.Visible = iOrigVisibility
End Function

Function fGotoCell(rgGoTo As Range, Optional lScrollRow As Long = 0, Optional iScrollCol As Integer = 0)
'    Dim shtCurrActive As Worksheet
'
'    Set shtCurrActive = ActiveSheet

    Dim iOrigVisibility As XlSheetVisibility

    iOrigVisibility = rgGoTo.Parent.Visible
    rgGoTo.Parent.Visible = xlSheetVisible

    Application.Goto rgGoTo, False

    If lScrollRow > 0 Then ActiveWindow.ScrollRow = lScrollRow
    If iScrollCol > 0 Then ActiveWindow.ScrollColumn = iScrollCol

    rgGoTo.Parent.Visible = iOrigVisibility

    'shtCurrActive.Activate
    'Set shtCurrActive = Nothing
End Function

Function fActiveVisibleSwitchSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1" _
                            , Optional bHidePreviousActiveSheet As Boolean = False)
    Dim shtCurr As Worksheet
    Set shtCurr = ActiveSheet

    On Error Resume Next

    If shtToSwitch.Visible = xlSheetVisible Then
        If Not ActiveSheet Is shtToSwitch Then
            shtToSwitch.Visible = xlSheetVisible
            shtToSwitch.Activate
            Range(sRngAddrToSelect).Select
        Else
            shtToSwitch.Visible = xlSheetVeryHidden
        End If
    Else
        shtToSwitch.Visible = xlSheetVisible
        shtToSwitch.Activate
        Range(sRngAddrToSelect).Select
    End If

    If bHidePreviousActiveSheet Then
        If Not shtCurr Is shtToSwitch Then shtCurr.Visible = xlSheetVeryHidden
    End If

    Err.Clear
End Function

Function fShowActivateSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1" _
                            , Optional bHidePreviousActiveSheet As Boolean = False)
    Dim shtCurr As Worksheet
    Set shtCurr = ActiveSheet

    On Error Resume Next

    If shtToSwitch.Visible <> xlSheetVisible Then shtToSwitch.Visible = xlSheetVisible

    shtToSwitch.Activate
    Range(sRngAddrToSelect).Select

    If bHidePreviousActiveSheet Then
        If Not shtCurr Is shtToSwitch Then shtCurr.Visible = xlSheetVeryHidden
    End If

    Err.Clear
End Function
Function fShowAndActiveSheet(sht As Worksheet)
    sht.Visible = xlSheetVisible
    sht.Activate
End Function

Function fUpdateGDictInputFile_FileName(asFileTag As String, asFileName As String)
    Call fUpdateDictionaryItemValueForDelimitedElement(gDictInputFiles, asFileTag, InputFile.FilePath - InputFile.FileTag, asFileName)
End Function

Function fUpdateGDictInputFile_FileSpecTag(asFileTag As String, asFileSpecTag As String)
    Call fUpdateDictionaryItemValueForDelimitedElement(gDictInputFiles, asFileTag, InputFile.FileSpecTag - InputFile.FileTag, asFileSpecTag)
End Function

Function fSetValueBackToSysConf_InputFile_FileName(asFileTag As String, asFileName As String)
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Input Files]", "File Full Path", "Report ID=" & gsRptID & ",File Tag=" & asFileTag, asFileName)
End Function

Function fGetInputFileSheetAfterLoadingToThisWorkBook(asFileTag As String) As Worksheet
    Set fGetInputFileSheetAfterLoadingToThisWorkBook = ThisWorkbook.Worksheets(fGetInputFileSheetNameAfterLoadingToThisWorkBook(asFileTag))
End Function

Function fGetInputFileSheetNameAfterLoadingToThisWorkBook(asFileTag As String) As String
    Dim sOut As String
    Dim sSource As String
    sSource = fGetInputFileSourceType(asFileTag)

    Select Case sSource
        Case "PARSE_AS_TEXT"
            sOut = asFileTag
        Case "FILE_BINDED_IN_MACRO", "READ_FROM_DRIVE", "READ_PREV_STEP_OUTPUT_FILE"
            sOut = asFileTag
        Case "READ_PRE_EXISTING_SHEET", "READ_PREV_STEP_OUTPUT_SHEET"
            sOut = fGetInputFileFileName(asFileTag)
        Case "READ_SHEET_BINDED_IN_MACRO"
            sOut = ""
        Case Else
            fErr "wrong sSource" & sSource
    End Select

    fGetInputFileSheetNameAfterLoadingToThisWorkBook = sOut
End Function

Function fConvertFomulaToValueForSheetIfAny(sht As Worksheet)
    Dim rng As Range

    On Error Resume Next
    Set rng = sht.Cells.SpecialCells(xlCellTypeFormulas)
    Err.Clear

    If rng Is Nothing Then Exit Function

'    Dim eachRng
'    For Each eachRng In rng.Areas
'        eachRng.Value = eachRng.Value
'    Next

    rng.Parent.UsedRange.Value = rng.Parent.UsedRange.Value
End Function

Function fSaveAndCloseWorkBook(wb As Workbook)
    If wb Is Nothing Then fErr "workbook passed to fSaveAndCloseWorkBook is nothing, please check your program."

    wb.Saved = True
    wb.CheckCompatibility = False
    wb.Save
    wb.CheckCompatibility = True
    wb.Close
    Set wb = Nothing
End Function
Function fCloseWorkBookWithoutSave(wb As Workbook)
    If wb Is Nothing Then Exit Function     'fErr "workbook passed to fCloseWorkBookWithoutSave is nothing, please check your program."

    wb.Saved = True
    wb.Close savechanges:=False
    Set wb = Nothing
End Function
Function fImportSingleSheetExcelFileToThisWorkbook(sExcelFileFullPath As String, sNewSheet As String _
                        , Optional asShtToImport As String = "", Optional wb As Workbook)
    Call fIfExcelFileOpenedToCloseIt(sExcelFileFullPath)

    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim wbSource As Workbook
    Set wbSource = Workbooks.Open(Filename:=sExcelFileFullPath, ReadOnly:=True)

    asShtToImport = Trim(asShtToImport)

    If Len(asShtToImport) <= 0 Then
        wbSource.Worksheets(1).Copy after:=wb.Worksheets(wb.Worksheets.Count)
    Else
        If Not fSheetExists(asShtToImport, , wbSource) Then
            fErr "There is no sheet named """ & asShtToImport & """ in workbook " & sExcelFileFullPath
        End If

        wbSource.Worksheets(asShtToImport).Copy after:=wb.Worksheets(wb.Worksheets.Count)
    End If

    wb.ActiveSheet.Name = sNewSheet
    ActiveWindow.DisplayGridlines = False

    Call fConvertFomulaToValueForSheetIfAny(wb.Worksheets(sNewSheet))

    Call fCloseWorkBookWithoutSave(wbSource)
End Function

Function fLoadFileByFileTag(asFileTag As String)
    Dim sFileFullPath As String
    Dim sSource As String
    Dim sReloadOrNot As String
    Dim sShtToImport As String
    Dim sShtToBeAdded As String

    sSource = fGetInputFileSourceType(asFileTag)
    If sSource = "READ_SHEET_BINDED_IN_MACRO" _
    Or sSource = "READ_PRE_EXISTING_SHEET" _
    Or sSource = "READ_PREV_STEP_OUTPUT_SHEET" Then Exit Function
'    sSource = "FILE_BINDED_IN_MACRO" _
'    Or sSource = "READ_PREV_STEP_OUTPUT_FILE"

    sFileFullPath = fGetInputFileFileName(asFileTag)
    sReloadOrNot = fGetInputFileReloadOrNot(asFileTag)
    sShtToImport = fGetInputFileSheetToImport(asFileTag)

    sShtToBeAdded = fGetInputFileSheetNameAfterLoadingToThisWorkBook(asFileTag)

    If fSheetExists(sShtToBeAdded) Then
        If sReloadOrNot = "RELOAD" Or fZero(sReloadOrNot) Then
            Call fDeleteSheet(sShtToBeAdded)
        Else
            Exit Function
        End If
    End If

    Select Case sSource
        Case "PARSE_AS_TEXT"
            Call fReadTxtFile2NewSheet(sFileFullPath, sShtToBeAdded, asFileTag)
        Case "FILE_BINDED_IN_MACRO", "READ_FROM_DRIVE", "READ_PREV_STEP_OUTPUT_FILE"
            Call fImportSingleSheetExcelFileToThisWorkbook(sFileFullPath, sShtToBeAdded)
            Call fRemoveFilterForSheet(ThisWorkbook.Worksheets(sShtToBeAdded))
        Case "READ_PRE_EXISTING_SHEET", "READ_PREV_STEP_OUTPUT_SHEET"
        Case "READ_SHEET_BINDED_IN_MACRO"
            Exit Function
        Case Else
            fErr "wrong sSource" & sSource
    End Select

End Function

Function fReadTxtColSpec(asFileTag As String) As Variant
    Dim arrOut()

    Dim sFileSpecTag As String
    sFileSpecTag = fGetInputFileFileSpecTag(asFileTag)

    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = sFileSpecTag
    ReDim arrColsName(1 To 2)
    arrColsName(1) = "Column Index"
    arrColsName(2) = "TXT Format Only For Text File"

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtFileSpec _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, 1, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
'    Call fValidateBlankInArray(arrConfigData, Company.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID")

    Dim dict As Dictionary
    Set dict = New Dictionary

    Dim lEachRow As Long
    Dim sLetterIndex As String
    Dim sTxtFormat As String
    Dim lTxtFormat As Long

    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row

        sLetterIndex = Trim(arrConfigData(lEachRow, 1))
        sTxtFormat = Trim(arrConfigData(lEachRow, 2))

        If Len(sTxtFormat) <= 0 Then
            lTxtFormat = 1
        Else
            lTxtFormat = fConvertTxtImportFormatToNum(sTxtFormat)
        End If

        dict.Add fLetter2Num(sLetterIndex), lTxtFormat
next_row:
    Next

    Erase arrColsName
    Erase arrConfigData

    Dim lColindex As Long
    Dim lMaxCol As Long

    lMaxCol = WorksheetFunction.Max(dict.Keys)

    ReDim arrOut(1 To lMaxCol)
    For lEachRow = 0 To dict.Count - 1
        lColindex = dict.Keys(lEachRow)
        arrOut(lColindex) = dict.Items(lEachRow)
    Next

    For lEachRow = LBound(arrOut) To UBound(arrOut)
        If Len(CStr(arrOut(lEachRow))) <= 0 Then
            arrOut(lEachRow) = 9    'xlSkipColumn
        End If
    Next

    fReadTxtColSpec = arrOut()
    Erase arrOut
    Set dict = Nothing
End Function
Function fImportTxtFile(sFileFullPath, arrColFormat, asDelmiter As String _
                        , alTextFilePlatForm As Long, ByRef shtTo As Worksheet) As Worksheet
    If fArrayIsEmptyOrNoData(arrColFormat) Then arrColFormat = Array(1)

    shtTo.Cells.ClearContents

    With shtTo.QueryTables.Add(Connection:="TEXT;" & sFileFullPath _
        , destination:=shtTo.Range("$A$1"))
        '.CommandType = 0
        .Name = shtTo.Name
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .TextFilePromptOnRefresh = False
        .TextFilePlatform = alTextFilePlatForm
        .TextFileStartRow = 1
        .TextFileParseType = xlDelimited
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileConsecutiveDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSemicolonDelimiter = False
        .TextFileCommaDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileOtherDelimiter = asDelmiter
        .TextFileColumnDataTypes = arrColFormat
        .TextFileTrailingMinusNumbers = True
        .Refresh BackgroundQuery:=False
    End With

End Function

Function fReadTxtFile2NewSheet(sFileFullPath As String, sShtToBeAdded As String, asFileTag As String)
    Dim shtToAdd As Worksheet
    Set shtToAdd = fAddNewSheet(sShtToBeAdded)

    Dim arrColFormat()
    arrColFormat = fReadTxtColSpec(asFileTag)

    Dim i As Long
    Dim iConfiMaxCol As Long

    iConfiMaxCol = UBound(arrColFormat)
    ReDim Preserve arrColFormat(LBound(arrColFormat) To iConfiMaxCol + 200)
    For i = iConfiMaxCol + 1 To iConfiMaxCol + 200
        arrColFormat(i) = 9         'xlSkipColumn
    Next

    Dim sColDelimiter As String
    Dim lPlatForm As Long

    sColDelimiter = Trim(Split(gDictTxtFileSpec(asFileTag), DELIMITER)(0))
    lPlatForm = CLng(Split(gDictTxtFileSpec(asFileTag), DELIMITER)(1))

    Call fImportTxtFile(sFileFullPath, arrColFormat, sColDelimiter, lPlatForm, shtToAdd)

    fDeleteRemoveConnections shtToAdd.Parent
End Function

Function fDeleteRemoveConnections(wb As Workbook)
    Dim i As Long

    For i = wb.Connections.Count To 1 Step -1
        wb.Connections(i).Delete
    Next
End Function

Function fCheckIfSheetHasNodata_RaiseErrToStop(arr, sht As Worksheet)
    gbNoData = fArrayIsEmptyOrNoData(arr)

    Dim sSheet As String
    If Not sht Is Nothing Then sSheet = sht.Name
    If gbNoData Then fErr "Input File " & sSheet & " has no qualified data!"
End Function

Function fFindHeaderAtLineInFileSpec(rngConfigBlock As Range, arrColsName) As Long
    Dim lColAtRow As Long
    Dim lEachCol As Long
    Dim sEachColName As String
    Dim rngFound As Range

    lColAtRow = 0
    For lEachCol = LBound(arrColsName) To UBound(arrColsName)
        sEachColName = Trim(arrColsName(lEachCol))
        sEachColName = Replace(sEachColName, "*", "~*")

        Set rngFound = fFindInWorksheet(rngConfigBlock, sEachColName)

        If lColAtRow <> 0 Then
            If lColAtRow <> rngFound.Row Then
                fErr "Columns are not at the same row."
            End If
        Else
            lColAtRow = rngFound.Row
        End If
    Next

    Set rngFound = Nothing

    fFindHeaderAtLineInFileSpec = lColAtRow
End Function

Function fGetTableLevelConfig(rngConfigBlock As Range, asTableLevelConf As String) As String
    Dim shtParent As Worksheet
    Dim rgFound As Range
    Dim rgTarget As Range
    Dim lValueColSameAsDisplayName As Long

    Set shtParent = rngConfigBlock.Parent

    Set rgFound = fFindInWorksheet(rngConfigBlock, "Column Display Name")
    lValueColSameAsDisplayName = rgFound.Column

    Set rgFound = fFindInWorksheet(rngConfigBlock, asTableLevelConf)

    Set rgTarget = rgFound.Offset(0, lValueColSameAsDisplayName - rgFound.Column)

    If fZero(rgTarget.Value) Then
        fErr "asTableLevelConf cannot be blank in " & shtParent.Name & vbCr & "range: " & rngConfigBlock.Address
    End If

    fGetTableLevelConfig = Trim(rgTarget.Value)
    Set rgTarget = Nothing
    Set rgFound = Nothing
    Set shtParent = Nothing
End Function

Function fReadInputFileSpecConfig(sFileSpecTag As String, ByRef dictLetterIndex As Dictionary _
                                , Optional ByRef dictArrayIndex As Dictionary _
                                , Optional ByRef dictDisplayName As Dictionary _
                                , Optional ByRef dictRawType As Dictionary _
                                , Optional ByRef dictDataFormat As Dictionary _
                                , Optional ByRef bReadWholeSheetData As Boolean _
                                , Optional shtData As Worksheet _
                                , Optional alHeaderAtRow As Long = 1)
    'Dim asTag As String
    Dim arrColsName()
    Dim arrColsIndex()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    Const TECH_NAME = 1
    Const DISPLAY_NAME = 2
    Const LETTER_INDEX = 3
    Const ARRAY_INDEX = 4
    Const RAW_DATA_TYPE = 5
    Const DATA_FORMAT = 6

    ReDim arrColsName(TECH_NAME To DATA_FORMAT)

    arrColsName(TECH_NAME) = "Column Tech Name"
    arrColsName(DISPLAY_NAME) = "Column Display Name"
    arrColsName(LETTER_INDEX) = "Column Index"
    arrColsName(ARRAY_INDEX) = "Array Index"
    arrColsName(RAW_DATA_TYPE) = "Raw Data Type"
    arrColsName(DATA_FORMAT) = "Data Format"

    Dim rngConfigBlock As Range
    Set rngConfigBlock = fFindRageOfFileSpecConfigBlock(sFileSpecTag)

    Dim iCol_TxtFormat As Long
    Dim iCol_OutputAsInput As Long
    Dim bTxtTemplate As Boolean
    Dim bOutputAsInput As Boolean
    Dim sGetColIndexBy As String
    Dim sReadSheetDataBy As String
    Dim bDynamic As Boolean

    bTxtTemplate = fFileSpecTemplateHasAdditionalHeader(rngConfigBlock, "TXT Format Only For Text File")
    If bTxtTemplate Then iCol_TxtFormat = fEnlargeArayWithValue(arrColsName, "TXT Format Only For Text File")

    bOutputAsInput = fFileSpecTemplateHasAdditionalHeader(rngConfigBlock, Array("Column Attr", "Column Width"))
    If bOutputAsInput Then iCol_OutputAsInput = fEnlargeArayWithValue(arrColsName, "Column Attr")

    sGetColIndexBy = UCase(fGetTableLevelConfig(rngConfigBlock, "Get Column Index By:"))
    sReadSheetDataBy = UCase(fGetTableLevelConfig(rngConfigBlock, "Read Sheet's Data By:"))

    bDynamic = (sGetColIndexBy <> "FIXED_LETTERS")
    bReadWholeSheetData = (sReadSheetDataBy = "READ_WHOLE_SHEET")

    Dim sErrPos As String
    sErrPos = vbCr & vbCr & "Sheet Name: " & shtFileSpec.Name & vbCr & "Range:" & rngConfigBlock.Address

    If bDynamic Then
        If iCol_TxtFormat > 0 Then fErr "dynamic (COLUMNS_NAME) cannot be specified for Txt Template" & sErrPos
        If bOutputAsInput Then fErr "dynamic (COLUMNS_NAME) cannot be specified for OutputAsInput Template" & sErrPos
        If shtData Is Nothing Then fErr "dynamic (COLUMNS_NAME), but shtData is not provided(nothing)." & sErrPos
    End If

    Call fReadConfigBlockToArray(asTag:=sFileSpecTag, shtParam:=shtFileSpec _
                                , arrConfigData:=arrConfigData _
                                , arrColsName:=arrColsName _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True)
    If lConfigHeaderAtRow >= lConfigEndRow Then fErr "No data is configured  " & sErrPos
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(TECH_NAME), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))

    If bDynamic Then ' by col dislay name    '"Column Display Name"
        Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(DISPLAY_NAME), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
        'Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(LETTER_INDEX), True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
    Else   'by speicified letter  '"Column Index"     Txt Template
        If Not bOutputAsInput Then
            Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(LETTER_INDEX), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
        End If
    End If

    If Not bReadWholeSheetData Then
        Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(ARRAY_INDEX), True, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
    End If
'    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
'    Call fValidateBlankInArray(arrConfigData, 2, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")

    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sColTechName As String
    Dim sDisplayName As String
    Dim sLetterIndex As String
    Dim lColLetter2Num As Long
    Dim sArrayIndex As String
    Dim lColArray2Num As Long
    Dim arrTxtNonImportCol()
    Dim dictActualRow As New Dictionary

    sErrPos = sErrPos & vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "

    Set dictLetterIndex = New Dictionary
    Set dictArrayIndex = New Dictionary
    Set dictDisplayName = New Dictionary
    Set dictRawType = New Dictionary
    Set dictDataFormat = New Dictionary

    'Dim dictTmpArrayInd As New Dictionary
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row

        lActualRow = lConfigHeaderAtRow + lEachRow
        sColTechName = Trim(arrConfigData(lEachRow, arrColsIndex(TECH_NAME)))
        sDisplayName = Trim(arrConfigData(lEachRow, arrColsIndex(DISPLAY_NAME)))
        sArrayIndex = Trim(arrConfigData(lEachRow, arrColsIndex(ARRAY_INDEX)))
        sLetterIndex = Trim(arrConfigData(lEachRow, arrColsIndex(LETTER_INDEX)))

        If bOutputAsInput Then
            If Trim(arrConfigData(lEachRow, arrColsIndex(iCol_OutputAsInput))) = "NOT_SHOW_UP" Then
                If Len(sLetterIndex) > 0 Then
                    fErr "Col Letter Index should be blank for NOT_SHOW_UP: " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(iCol_OutputAsInput)
                End If
                GoTo next_row
            End If
        End If

        If bTxtTemplate Then
            If Trim(arrConfigData(lEachRow, arrColsIndex(iCol_TxtFormat))) = "xlSkipColumn" Then
                If Len(sArrayIndex) > 0 Then fErr "ArrayIndex should be blank when xlSkipColumn is specified " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(iCol_TxtFormat)

                Call fEnlargeArayWithValue(arrTxtNonImportCol, fLetter2Num(sLetterIndex))
                GoTo next_row
            End If
        End If

        If Not bDynamic Then
            'If Len(sLetterIndex) > 0 Then
                lColLetter2Num = fLetter2Num(sLetterIndex)

                If lColLetter2Num <= 0 Or lColLetter2Num > Columns.Count Then
                    fErr "Col Letter Index is invalid,should be A - XFD: " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(LETTER_INDEX)
                End If
                dictLetterIndex.Add sColTechName, lColLetter2Num
            'End If
        End If
        If Not bReadWholeSheetData Then
            If Len(sArrayIndex) > 0 Then
                If Not IsNumeric(sArrayIndex) Then
                    fErr "Col Array Index is invalid,should be  1, 2, 3, ...: " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(LETTER_INDEX)
                End If
                lColArray2Num = CLng(sArrayIndex)

                If lColArray2Num <= 0 Or lColArray2Num > Columns.Count Then
                    fErr "Col Array Index is invalid,should be 1 - " & Columns.Count & ": " & Replace(sErrPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(ARRAY_INDEX)
                End If
                dictArrayIndex.Add sColTechName, lColArray2Num
            End If
        End If

        dictDisplayName.Add sColTechName, sDisplayName
        dictRawType.Add sColTechName, UCase(Trim(arrConfigData(lEachRow, arrColsIndex(RAW_DATA_TYPE))))
        dictDataFormat.Add sColTechName, Trim(arrConfigData(lEachRow, arrColsIndex(DATA_FORMAT)))

        dictActualRow.Add sColTechName, lActualRow
next_row:
    Next

    If dictActualRow.Count <= 0 Then fErr "Cxxxxxxxxxxx"

    Dim lTxtMaxCol As Long
    Dim iTxt As Long
    Dim arrTxt
    If bTxtTemplate Then
        lTxtMaxCol = WorksheetFunction.Max(dictLetterIndex.Items)
        arrTxt = dictLetterIndex.Items
        For iTxt = 1 To lTxtMaxCol
            If InArray(arrTxt, iTxt) < 0 Then
                Call fEnlargeArayWithValue(arrTxtNonImportCol, iTxt)
            End If
        Next

        Erase arrTxt
    End If

    If Not bReadWholeSheetData Then
        If dictArrayIndex.Count <= 0 Then
            fErr "READ_SPECIFIED_COLUMNS is specified, but no array index is not specified"
        End If
    End If

    If bTxtTemplate Then
        fGetRangeByStartEndPos(shtFileSpec, lConfigHeaderAtRow + 1, lConfigStartCol + arrColsIndex(LETTER_INDEX), lConfigEndRow, lConfigStartCol + arrColsIndex(LETTER_INDEX)).ClearContents
        If fRecalculateColumnIndexByRemoveNonImportTxtCol(dictLetterIndex, arrTxtNonImportCol) Then
            For lEachRow = 0 To dictActualRow.Count - 1
                shtFileSpec.Cells(dictActualRow.Items(lEachRow), lConfigStartCol + arrColsIndex(LETTER_INDEX)) = _
                    fNum2Letter(dictLetterIndex.Items(lEachRow)) 'this is for reference
            Next
        End If
    End If

    Dim arrDisplayNames()
    Dim arrDynamicColIndex()
    If bDynamic Then
        arrDisplayNames = dictDisplayName.Items
        Call fFindAllColumnsIndexByColNames(shtData.Rows(alHeaderAtRow), arrDisplayNames, arrDynamicColIndex)

        If Not Base0(arrDisplayNames) Then fErr "arrDisplayNames is not based from 0"
        For lEachRow = LBound(arrDynamicColIndex) To UBound(arrDynamicColIndex)
            dictLetterIndex.Add dictDisplayName.Keys(lEachRow), arrDynamicColIndex(lEachRow)
        Next
        For lEachRow = 0 To dictActualRow.Count - 1
            shtFileSpec.Cells(dictActualRow.Items(lEachRow), lConfigStartCol + arrColsIndex(LETTER_INDEX) - 1) = _
                    fNum2Letter(dictLetterIndex.Items(lEachRow))
        Next
    End If

    Erase arrDisplayNames
    Erase arrDynamicColIndex
    Erase arrDisplayNames
    Erase arrColsName
    Erase arrColsIndex
    Erase arrConfigData
    Set dictActualRow = Nothing
    Set rngToFindIn = Nothing
End Function

Function fRecalculateColumnIndexByRemoveNonImportTxtCol(ByRef dictLetterIndex As Dictionary, arrTxtNonImportCol()) As Boolean
    Dim bOut As Boolean
    bOut = False

    If fArrayIsEmptyOrNoData(arrTxtNonImportCol) Then GoTo exit_fun

    Call fSortArayDesc(arrTxtNonImportCol)

    Dim iArrayIndex As Long
    Dim iDictIndex As Long
    Dim iColIndex As Long

    For iArrayIndex = LBound(arrTxtNonImportCol) To UBound(arrTxtNonImportCol)
        iColIndex = arrTxtNonImportCol(iArrayIndex)

        For iDictIndex = 0 To dictLetterIndex.Count - 1
            If dictLetterIndex.Items(iDictIndex) > iColIndex Then
                dictLetterIndex(dictLetterIndex.Keys(iDictIndex)) = dictLetterIndex.Items(iDictIndex) - 1
            ElseIf dictLetterIndex.Items(iDictIndex) = iColIndex Then
                fErr "abnormal in fRecalculateColumnIndexByRemoveNonImportTxtCol"
            Else
                'Debug.Print dictLetterIndex.Items(iDictIndex) & " - " & iColIndex
            End If
        Next iDictIndex
    Next iArrayIndex

    bOut = True
exit_fun:
    fRecalculateColumnIndexByRemoveNonImportTxtCol = bOut
End Function

Function fConvertTxtImportFormatToNum(sDesc As String) As Integer
    Dim iOut As Integer

    Select Case sDesc
        Case "xlGeneralFormat"
            iOut = 1
        Case "xlTextFormat"
            iOut = 2
        Case "xlMDYFormat"
            iOut = 3
        Case "xlDMYFormat"
            iOut = 4
        Case "xlYMDFormat"
            iOut = 5
        Case "xlMYDFormat"
            iOut = 6
        Case "xlDYMFormat"
            iOut = 7
        Case "xlYDMFormat"
            iOut = 8
        Case "xlSkipColumn"
            iOut = 9
        Case "xlEMDFormat"
            iOut = 10
        Case Else
    End Select

    fConvertTxtImportFormatToNum = iOut
End Function

Function fReadSheetDataByConfig(asFileTag As String, ByRef dictColIndex As Dictionary, ByRef arrDataOut() _
                                , Optional ByRef dictColFormat As Dictionary _
                                , Optional ByRef dictRawType As Dictionary _
                                , Optional ByRef dictDisplayName As Dictionary _
                                , Optional alDataFromRow As Long = 2 _
                                , Optional ByRef shtData As Worksheet)
    Dim sFileSpecTag As String
    Dim shtToRead As Worksheet

    sFileSpecTag = fGetInputFileFileSpecTag(asFileTag)

    If shtData Is Nothing Then
        Set shtToRead = fGetInputFileSheetAfterLoadingToThisWorkBook(asFileTag)
        Set shtData = shtToRead
    Else
        Set shtToRead = shtData
    End If

    Dim bReadWholeSheetData As Boolean
    Dim dictLetterIndex As Dictionary
    Dim dictArrayIndex As Dictionary
    Call fReadInputFileSpecConfig(sFileSpecTag:=sFileSpecTag _
                                , dictLetterIndex:=dictLetterIndex _
                                , dictArrayIndex:=dictArrayIndex _
                                , dictDisplayName:=dictDisplayName _
                                , dictRawType:=dictRawType _
                                , dictDataFormat:=dictColFormat _
                                , bReadWholeSheetData:=bReadWholeSheetData _
                                , shtData:=shtToRead _
                                , alHeaderAtRow:=alDataFromRow - 1)
    If bReadWholeSheetData Then
        Call fCopyReadWholeSheetData2Array(shtToRead, arrDataOut, dictLetterIndex, alDataFromRow)
        Call fConvertDataToTheirRawDataType(arrDataOut, dictLetterIndex, dictRawType, dictColFormat)
    Else
        Call fReadSpecifiedColsToArrayByConfig(shtData:=shtToRead, dictLetterIndex:=dictLetterIndex, dictArrayIndex:=dictArrayIndex _
                    , dictRawType:=dictRawType, dictColFormat:=dictColFormat _
                     , arrDataOut:=arrDataOut, alDataFromRow:=alDataFromRow)
    End If

    If bReadWholeSheetData Then
        Set dictColIndex = dictLetterIndex
    Else
        Set dictColIndex = dictArrayIndex
    End If

    Set shtToRead = Nothing
    Set dictLetterIndex = Nothing
    Set dictArrayIndex = Nothing
End Function

'Function fGetReadInputFileSpecConfigItem(sFileSpecTag As String, asItem As String) As Variant
'    Dim dictColFormat As Dictionary _
'                                , dictRawType As Dictionary _
'                                , dictDisplayName As Dictionary
''                                , alDataFromRow As Long _
''                                , shtData As Worksheet
'    Dim dictLetterIndex As Dictionary
'    Dim dictArrayIndex As Dictionary
'
'    Call fReadInputFileSpecConfig(sFileSpecTag:=sFileSpecTag _
'                                , dictLetterIndex:=dictLetterIndex _
'                                , dictArrayIndex:=dictArrayIndex _
'                                , dictDisplayName:=dictDisplayName _
'                                , dictRawType:=dictRawType _
'                                , dictDataFormat:=dictColFormat _
'                                , bReadWholeSheetData:=bReadWholeSheetData _
'                                , shtData:=shtToRead _
'                                , alHeaderAtRow:=alDataFromRow - 1)
'    Select Case asItem
'        Case "TXT_COL_FORMAT"
'            set fGetReadInputFileSpecConfigItem =
'        Case Else
'
'    End Select
'
'    Set dictColFormat = Nothing
'    Set dictRawType = Nothing
'    Set dictDisplayName = Nothing
'    Set dictLetterIndex = Nothing
'    Set dictArrayIndex = Nothing
'End Function

Function fReadSpecifiedColsToArrayByConfig(shtData As Worksheet, dictLetterIndex As Dictionary, dictArrayIndex As Dictionary _
                    , dictRawType As Dictionary, dictColFormat As Dictionary _
                     , arrDataOut(), Optional alDataFromRow As Long = 2)

    Dim lColCopyFrom As Long
    Dim lColCopyTo As Long
    Dim lArrayMaxCol As Long
    Dim lShtMaxRow As Long

    lArrayMaxCol = WorksheetFunction.Max(dictArrayIndex.Items)
    lShtMaxRow = fGetValidMaxRow(shtData)

    If lShtMaxRow < alDataFromRow Then arrDataOut = Array(): Exit Function

    ReDim arrDataOut(1 To lShtMaxRow - alDataFromRow + 1, 1 To lArrayMaxCol)

    Dim i As Long
    Dim sTechName As String
    Dim sColType As String
    Dim arrEachCol()
    Dim lEachRow As Long
    For i = 0 To dictArrayIndex.Count - 1
        sTechName = dictArrayIndex.Keys(i)

        lColCopyFrom = dictLetterIndex(sTechName)
        lColCopyTo = dictArrayIndex(sTechName)

        sColType = UCase(dictRawType(sTechName))
        arrEachCol = fReadRangeDatatoArrayByStartEndPos(shtData, alDataFromRow, lColCopyFrom, lShtMaxRow, lColCopyFrom)

        If sColType = "DATE" Or sColType = "STRING_PERCENTAGE" Or sColType = "RMB_CURRENCY" Then
            For lEachRow = LBound(arrEachCol, 1) To UBound(arrEachCol, 1)
                arrDataOut(lEachRow, lColCopyTo) = fCType(arrEachCol(lEachRow, 1), sColType, dictColFormat(sTechName))
            Next
        Else
            For lEachRow = LBound(arrEachCol, 1) To UBound(arrEachCol, 1)
                arrDataOut(lEachRow, lColCopyTo) = arrEachCol(lEachRow, 1)
            Next
        End If
    Next

End Function

Function fConvertDataToTheirRawDataType(ByRef arrData(), dictLetterIndex As Dictionary, dictRawType As Dictionary, dictColFormat As Dictionary)
    Dim i As Long
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim sTechName As String
    Dim sColType As String

    For i = 0 To dictLetterIndex.Count - 1
        sTechName = dictLetterIndex.Keys(i)
        sColType = UCase(dictRawType(sTechName))

        If sColType = "DATE" Or sColType = "STRING_PERCENTAGE" Or sColType = "RMB_CURRENCY" Then
            lEachCol = dictLetterIndex.Items(i)

            For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
                arrData(lEachRow, lEachCol) = fCType(arrData(lEachRow, lEachCol), sColType, dictColFormat(sTechName))
            Next
        End If
    Next
End Function

Function fCType(aValue, asToType As String, asFormat As String) As Variant
    Dim aOut As Variant
    Dim sDataType As String
    Dim bOrigToAreSame As Boolean

    If IsEmpty(aValue) Then fCType = aValue: Exit Function

    asToType = UCase(asToType)
    sDataType = UCase(TypeName(aValue))

    bOrigToAreSame = False
    Select Case asToType
        Case "STRING", "TEXT"
            If sDataType = "STRING" Then bOrigToAreSame = True
        Case "DATE"
            If aValue = 0 Then fCType = 0: Exit Function
            If sDataType = "DATE" Then bOrigToAreSame = True
        Case "DECIMAL"
            If sDataType = "DECIMAL" Or sDataType = "DOUBLE" Or sDataType = "SINGLE" Or sDataType = "CURRENCY" Then
                bOrigToAreSame = True
            End If
        Case "NUMBER"
            If sDataType = "BYTE" Or sDataType = "INTEGER" Or sDataType = "LONG" Or sDataType = "LONGLONG" Or sDataType = "LONGPRT" Then
                bOrigToAreSame = True
            End If
        Case "STRING_PERCENTAGE"
        Case "RMB_CURRENCY"

        Case Else
            fErr "wrong param asToType"
    End Select

    If bOrigToAreSame Then fCType = aValue: Exit Function

    Select Case asToType
        Case "STRING", "TEXT"
            fCType = CStr(aValue)
        Case "DATE"
            Dim dtTmp As Date
            dtTmp = fCdateStr(CStr(aValue), asFormat)

            If dtTmp <= 0 Then
                fErr "Wrong date value: " & aValue & ", please check your data, or contact with IT support."
            End If
            fCType = dtTmp
        Case "DECIMAL"
            fCType = CDbl(aValue)
        Case "NUMBER"
            fCType = CLng(aValue)
        Case "STRING_PERCENTAGE"
            fCType = fCPercentage2Dbl(aValue)
        Case "RMB_CURRENCY"
            fCType = fCRMBCurrency2Dbl(aValue)
        Case Else
            fErr "wrong param asToType"
    End Select
End Function

Function fCdateStr(sDate As String, Optional sFormat As String = "") As Date
    Dim sYear As String
    Dim sMonth As String
    Dim sDay As String

    sDate = Trim(sDate)
    If Len(sDate) <= 0 Then Exit Function

    Dim bSplit As Boolean
    Dim sDelimiter As String
    Const DATE_DELIMITERS = "-/._."

    bSplit = False

    Dim i As Integer
    For i = 1 To Len(DATE_DELIMITERS)
        If InStr(sDate, Mid(DATE_DELIMITERS, i, 1)) > 0 Then
            sDelimiter = Mid(DATE_DELIMITERS, i, 1)
            bSplit = True
            Exit For
        End If
    Next

    sFormat = Replace(sFormat, ">", "")
    sFormat = Replace(sFormat, "<", "")

    If bSplit Then sFormat = Replace(sFormat, sDelimiter, "/")

    If bSplit And Len(sFormat) <= 0 Then fErr "The date has delimiter, but you did not specify the format:" & vbCr & "Date:" & sDate & vbCr & "Format:" & sFormat

    Select Case UCase(sFormat)
        Case "DDMMMYY", "DDMMMYYYY"
            sYear = Mid(sDate, 6)
            sMonth = fConvertMMM2Num(Mid(sDate, 3, 3))
            sDay = Left(sDate, 2)
        Case "MMDDYY", "MMDDYYYY"
            sYear = Mid(sDate, 5)
            sMonth = Left(sDate, 2)
            sDay = Mid(sDate, 3, 2)
        Case "DDMMYY", "DDMMYYYY"
            sYear = Mid(sDate, 5)
            sMonth = Mid(sDate, 3, 2)
            sDay = Left(sDate, 2)
        Case "YYMMDD"
            sYear = Left(sDate, 2)
            sMonth = Mid(sDate, 3, 2)
            sDay = Mid(sDate, 5, 2)
        Case "YYYYMMDD"
            sYear = Left(sDate, 4)
            sMonth = Mid(sDate, 5, 2)
            sDay = Mid(sDate, 7, 2)
        Case "YY/MM/DD", "YYYY/MM/DD"
            sYear = Split(sDate, sDelimiter)(0)
            sMonth = Split(sDate, sDelimiter)(1)
            sDay = Split(sDate, sDelimiter)(2)
        Case Else
            fErr "sFormat is not covered in fCdateStr, please change this function." & vbCr _
             & "sFormat: " & sFormat & vbCr _
             & "sDelimiter: " & sDelimiter & vbCr _
             & "sDate: " & sDate
    End Select

    If val(sYear) <= 0 Then fErr "Year is in correct in date:" & vbCr & sDate
    If val(sMonth) <= 0 Then fErr "Month is in correct in date:" & vbCr & sDate
    If val(sDay) <= 0 Then fErr "Day is in correct in date:" & vbCr & sDate

    fCdateStr = DateSerial(CLng(sYear), CLng(sMonth), CLng(sDay))
End Function

Function fCPercentage2Dbl(ByVal aValue As String) As Double
    aValue = Trim(aValue)
    aValue = Left(aValue, Len(aValue) - 1)
    fCPercentage2Dbl = val(aValue) / 100
End Function
Function fCRMBCurrency2Dbl(ByVal aValue As String) As Double
    aValue = Trim(aValue)

    If Left(aValue, 1) = "￥" Then
        aValue = Right(aValue, Len(aValue) - 1)
        fCRMBCurrency2Dbl = val(aValue)
    Else
        fCRMBCurrency2Dbl = val(aValue)
    End If
End Function

Function fCopyReadWholeSheetData2Array(shtToRead As Worksheet, ByRef arrDataOut() _
            , Optional dictLetterIndex As Dictionary, Optional alDataFromRow As Long = 2, Optional alMaxCol As Long = 0)
    Dim lMaxRow As Long
    Dim lMaxCol As Long

    lMaxRow = fGetValidMaxRow(shtToRead)
    If lMaxRow < alDataFromRow Then arrDataOut = Array(): Exit Function

    If alMaxCol > 0 Then
        lMaxCol = alMaxCol
    Else
        If dictLetterIndex Is Nothing Then
            lMaxCol = fGetValidMaxCol(shtToRead)
        Else
            lMaxCol = WorksheetFunction.Max(dictLetterIndex.Items)
        End If
    End If

    arrDataOut = fReadRangeDatatoArrayByStartEndPos(shtToRead, alDataFromRow, 1, lMaxRow, lMaxCol)
End Function

Function fReadMasterSheetData(asFileTag As String, Optional shtData As Worksheet, Optional asDataFromRow As Long = 2 _
        , Optional bNoDataError As Boolean = False)
    Call fReadSheetDataByConfig(asFileTag:=asFileTag, dictColIndex:=dictMstColIndex, arrDataOut:=arrMaster _
                                , dictColFormat:=dictMstCellFormat _
                                , dictRawType:=dictMstRawType _
                                , dictDisplayName:=dictMstDisplayName _
                                , alDataFromRow:=asDataFromRow _
                                , shtData:=shtData)
    If bNoDataError Then Call fCheckIfSheetHasNodata_RaiseErrToStop(arrMaster, shtData)
End Function

Function fPrepareOutputSheetHeaderAndTextColumns(shtOutput As Worksheet)
    Dim i As Long
    Dim arrHeader()
    Dim lMaxCol As Long

    lMaxCol = WorksheetFunction.Max(dictRptColIndex.Items)

    ReDim arrHeader(1 To 1, 1 To lMaxCol)

    Dim lEachCol As Long
    For lEachCol = 1 To lMaxCol
        For i = 0 To dictRptColIndex.Count - 1
            If dictRptColIndex.Items(i) = lEachCol Then
                arrHeader(1, lEachCol) = dictRptDisplayName(dictRptColIndex.Keys(i))
                Exit For
            End If
        Next
    Next
    shtOutput.Range("A1").Resize(1, lMaxCol).Value = arrHeader
    Erase arrHeader

    Dim rgHeader As Range
    Set rgHeader = fGetRangeByStartEndPos(shtOutput, 1, 1, 1, lMaxCol)

    rgHeader.HorizontalAlignment = xlCenter
    rgHeader.VerticalAlignment = xlCenter
    rgHeader.WrapText = True
    rgHeader.Orientation = 0
    rgHeader.AddIndent = False
    rgHeader.IndentLevel = 0
    rgHeader.ShrinkToFit = False
    rgHeader.ReadingOrder = xlContext
    rgHeader.MergeCells = False
    Set rgHeader = Nothing

    Call fPresetColsNumberFormat2TextForOuputSheet(shtOutput)
End Function

Function fPresetColsNumberFormat2TextForOuputSheet(shtOutput As Worksheet, Optional lMaxRow As Long = 0)
    If lMaxRow = 0 Then lMaxRow = Rows.Count - 1

    Dim i As Long
    Dim iColIndex As Long
    Dim sColTech  As String
    Dim sColType As String

    For i = 0 To dictRptRawType.Count - 1
        sColTech = dictRptRawType.Keys(i)

        If dictRptColAttr(sColTech) = "NOT_SHOW_UP" Then GoTo next_col

        'sFormatStr = dictRptDataFormat(sColTech)
        sColType = UCase(dictRptRawType(sColTech))
        If sColType = "STRING" Or sColType = "TEXT" Then
            iColIndex = dictRptColIndex(sColTech)
            Call fSetNumberFormatForRange(fGetRangeByStartEndPos(shtOutput, 2, iColIndex, lMaxRow, iColIndex), "@")
        End If
next_col:
    Next
End Function

Function fSetNumberFormatForRange(rng As Range, Optional sFormat As String = "General")
    rng.NumberFormat = sFormat
    'rng.Value = rng.Value
End Function

Function fPostProcess(ByRef shtOutput As Worksheet)
    Call fDeleteNotShowUpColumns(shtOutput)
End Function

Function fDeleteNotShowUpColumns(ByRef shtOutput As Worksheet)
    Dim i As Long
    Dim iColIndex As Long
    Dim sColTech  As String
    Dim sColType As String

    If dictRptColIndex Is Nothing Then fErr "dictRptColIndex is even empty , pls call fReadSysConfig_Output first"

    For i = dictRptColIndex.Count - 1 To 0 Step -1
        sColTech = dictRptColIndex.Keys(i)
        sColType = dictRptColAttr(sColTech)

        If sColType = "NOT_SHOW_UP" Then
            shtOutput.Columns(dictRptColIndex.Items(i)).Delete shift:=xlToLeft
        End If
next_col:
    Next
End Function

Function fSaveWorkBookNotClose(wb As Workbook)
    wb.Saved = True
    wb.Close savechanges:=False
    Set wb = Nothing
End Function
'Function fCloseWorkBookWithoutSave(wb As Workbook)
'    wb.CheckCompatibility = False
'    wb.Save
'    If gbCheckCompatibility Then wb.CheckCompatibility = True
'End Function

Function fCleanSheetOutputResetSheetOutput(ByRef shtOutput As Worksheet)
    Call fRemoveFilterForSheet(shtOutput)
    shtOutput.Cells.ClearContents
    shtOutput.Cells.ClearFormats
    shtOutput.Cells.ClearContents
    shtOutput.Cells.ClearContents
    shtOutput.UsedRange.Delete shift:=xlUp
End Function

Function fClearDataFromSheetLeaveHeader(ByRef shtOutput As Worksheet)
    Call fRemoveFilterForSheet(shtOutput)

    Dim lMaxRow As Long
    lMaxRow = fGetValidMaxRow(shtOutput)

    If lMaxRow > 2 Then
        With fGetRangeByStartEndPos(shtOutput, 2, 1, lMaxRow, fGetValidMaxCol(shtOutput))
            .ClearContents
            '.ClearFormats
            .ClearComments
            .ClearNotes
            .ClearOutline
        End With
    End If
End Function

Function fRedimArrOutputBaseArrMaster()
    Dim lMaxCol As Long
    lMaxCol = fGetReportMaxColumn()
    ReDim arrOutput(1 To UBound(arrMaster, 1), 1 To lMaxCol)
End Function

Function fGetReportMaxColumn() As Long
    fGetReportMaxColumn = WorksheetFunction.Max(dictRptColIndex.Items)
End Function

Function fFormatOutputSheet(ByRef shtOutput As Worksheet, Optional lRowFrom As Long = 2)
    Dim lMaxRow As Long
    Dim lMaxCol As Long
    lMaxRow = fGetValidMaxRow(shtOutput)
    lMaxCol = fGetValidMaxCol(shtOutput)

    If lMaxRow < 2 Then Exit Function

    Call fSetFormatBoldOrangeBorderForRangeEspeciallyForHeader(fGetRangeByStartEndPos(shtOutput, 1, 1, 1, lMaxCol))

    Call fBasicCosmeticFormatSheet(shtOutput, lMaxCol)
    Call fFormatReportByConfigByCopyFormat(shtOutput, lMaxCol, lRowFrom, lMaxRow)

    If lMaxRow < 5000 Then
        Call fSetFormatForOddEvenLineByFixColor(shtOutput, lMaxCol, lRowFrom, lMaxRow)
        'Call fSetConditionFormatForOddEvenLine(shtOutput, lMaxCol, lRowFrom, lMaxRow)
    Else
        Call fSetConditionFormatForOddEvenLine(shtOutput, lMaxCol, lRowFrom, lMaxRow)
    End If

    Call fSetBorderLineForSheet(shtOutput, lMaxCol, lRowFrom, lMaxRow)
    Call fSetNumberFormatForOutputSheetByConfigExceptTextCol(shtOutput, lMaxCol, lRowFrom, lMaxRow)
    Call fSetBackNumberFormat2TextForCols(shtOutput)
    Call fSetColumnWidthForOutputSheetByConfig(shtOutput)
End Function

Function fSetColumnWidthForOutputSheetByConfig(ByRef shtOutput As Worksheet)
    Dim i As Long

    For i = 0 To dictRptColWidth.Count - 1
        If dictRptColWidth.Items(i) <> 0 Then
            shtOutput.Columns(dictRptColIndex(dictRptColWidth.Keys(i))).ColumnWidth = CDbl(dictRptColWidth.Items(i))
        End If
    Next
End Function
Function fBasicCosmeticFormatSheet(ByRef sht As Worksheet, Optional lMaxCol As Long = 0)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(sht)

    sht.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.SplitColumn = 0
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    ActiveWindow.DisplayGridlines = False
    'Call fFreezeSheet(sht, alSplitCol, alSplitRow)

    If sht.AutoFilterMode Then sht.AutoFilterMode = False
    fGetRangeByStartEndPos(sht, 1, 1, 1, lMaxCol).AutoFilter
    'sht.Cells.EntireColumn.AutoFit
    sht.Cells.EntireRow.AutoFit
    'sht.Range("A2").Select
End Function

Function fFormatReportByConfigByCopyFormat(ByRef shtOutput As Worksheet, Optional lMaxCol As Long = 0 _
                , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0 _
                , Optional dictColIndex As Dictionary _
                , Optional dictColCellFormat As Dictionary _
                , Optional bOddEvenColor As Boolean = True)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtOutput)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtOutput)

'    Call fSetFormatBoldOrangeBorderForRangeEspeciallyForHeader(fGetRangeByStartEndPos(shtOutput, 1, 1, 1, lMaxCol))

    If lRowTo < lRowFrom Then Exit Function

    Dim i As Long
    Dim sColTech As String
    Dim lEachCol As Long
    Dim rgFrom As Range
    Dim rgTo As Range
    Dim oDisFormat As DisplayFormat

    If dictColIndex Is Nothing Then Set dictColIndex = dictRptColIndex
    If dictColCellFormat Is Nothing Then Set dictColCellFormat = dictRptCellFormat

    For i = 0 To dictColIndex.Count - 1
        sColTech = dictColIndex.Keys(i)
        lEachCol = CLng(dictColIndex.Items(i))

        Set rgFrom = shtFileSpec.Range(dictColCellFormat(sColTech))
        Set rgTo = fGetRangeByStartEndPos(shtOutput, lRowFrom, lEachCol, lRowTo, lEachCol)

        If rgFrom.Interior.Color <> RGB(255, 255, 255) Or Not bOddEvenColor Then
            rgFrom.Copy
            rgTo.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Application.CutCopyMode = False
        Else
            Set oDisFormat = rgFrom.DisplayFormat
            rgTo.HorizontalAlignment = oDisFormat.HorizontalAlignment
            rgTo.VerticalAlignment = oDisFormat.VerticalAlignment
            rgTo.Font.Bold = oDisFormat.Font.Bold
            rgTo.Font.Italic = oDisFormat.Font.Italic
            rgTo.Font.FontStyle = oDisFormat.Font.FontStyle
            rgTo.Font.Strikethrough = oDisFormat.Font.Strikethrough
            rgTo.Font.Underline = oDisFormat.Font.Underline
            rgTo.Font.ThemeFont = oDisFormat.Font.ThemeFont
            rgTo.Font.Color = oDisFormat.Font.Color
        End If
    Next
    Set rgFrom = Nothing
    Set rgTo = Nothing
End Function

Function fSetBorderLineForSheet(ByRef shtOutput As Worksheet, Optional lMaxCol As Long = 0 _
                , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtOutput)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtOutput)

    Call fSetBorderLineForRange(fGetRangeByStartEndPos(shtOutput, lRowFrom, 1, lRowTo, lMaxCol))
End Function

Function fSetBorderLineForRange(ByRef rng As Range)
    With rng
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
'            .LineStyle = xlContinuous
'            .ColorIndex = xlAutomatic
'            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeTop)
'            .LineStyle = xlContinuous
'            .ColorIndex = xlAutomatic
'            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeBottom)
'            .LineStyle = xlContinuous
'            .ColorIndex = xlAutomatic
'            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlEdgeRight)
'            .LineStyle = xlContinuous
'            .ColorIndex = xlAutomatic
'            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlInsideVertical)
'            .LineStyle = xlContinuous
'            .ColorIndex = xlAutomatic
'            .TintAndShade = 0
            .Weight = xlHairline
        End With
        With .Borders(xlInsideHorizontal)
'            .LineStyle = xlContinuous
'            .ColorIndex = xlAutomatic
'            .TintAndShade = 0
            .Weight = xlHairline
        End With
    End With
End Function

Function fSetNumberFormatForOutputSheetByConfigExceptTextCol(ByRef shtOutput As Worksheet, Optional lMaxCol As Long = 0 _
                                                            , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0 _
                                                            )
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtOutput)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtOutput)

    If lRowTo < lRowFrom Then Exit Function

    Dim i As Long
    Dim sColTech As String
    Dim lEachCol As Long
    Dim sFormat As String
    Dim sColType As String

    For i = 0 To dictRptRawType.Count - 1
        sColTech = dictRptRawType.Keys(i)

        If dictRptColAttr(sColTech) = "NOT_SHOW_UP" Then GoTo next_col

        sColType = UCase(dictRptRawType(sColTech))
        sFormat = dictRptDataFormat(sColTech)
        lEachCol = dictRptColIndex(sColTech)

        Select Case sColType
            Case "NUMBER"
                If Len(sFormat) <= 0 Then
                    Call fSetNumberFormatForRange(fGetRangeByStartEndPos(shtOutput, lRowFrom, lEachCol, lRowTo, lEachCol), "0_")
                Else
                    Call fSetNumberFormatForRange(fGetRangeByStartEndPos(shtOutput, lRowFrom, lEachCol, lRowTo, lEachCol), sFormat)
                End If
            Case Else
                If Len(sFormat) > 0 Then Call fSetNumberFormatForRange(fGetRangeByStartEndPos(shtOutput, lRowFrom, lEachCol, lRowTo, lEachCol), sFormat)
        End Select
next_col:
    Next
End Function

Function fSetBackNumberFormat2TextForCols(ByRef shtOutput As Worksheet _
                                                            , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0 _
                                                            , Optional lMaxCol As Long = 0)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtOutput)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtOutput)

    If lRowTo < lRowFrom Then Exit Function

    Dim i As Long
    Dim lEachRow As Long
    Dim sColTech As String
    Dim lEachCol As Long
    Dim sFormat As String
    Dim sColType As String
    Dim arrData()

    For i = 0 To dictRptRawType.Count - 1
        sColTech = dictRptRawType.Keys(i)

        If dictRptColAttr(sColTech) = "NOT_SHOW_UP" Then GoTo next_col

        sColType = UCase(dictRptRawType(sColTech))
        sFormat = dictRptDataFormat(sColTech)
        lEachCol = dictRptColIndex(sColTech)

        If (sColType = "STRING" Or sColType = "TEXT") And Len(sFormat) > 0 Then
            Call fSetNumberFormatForRange(fGetRangeByStartEndPos(shtOutput, lRowFrom, lEachCol, lRowTo, lEachCol), "@")

            If Len(sFormat) > 0 Then
                arrData = fReadRangeDatatoArrayByStartEndPos(shtOutput, lRowFrom, lEachCol, lRowTo, lRowTo)

                For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
                    arrData(lEachRow, 1) = format(arrData(lEachRow, 1), sFormat)
                Next

                shtOutput.Cells(lRowFrom, lEachCol).Resize(UBound(arrData, 1), 1).Value = arrData
            End If
        End If
next_col:
    Next

End Function

Function fSetFormatForOddEvenLineByFixColor(ByRef shtOutput As Worksheet, Optional lMaxCol As Long = 0 _
                                            , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtOutput)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtOutput)

    If lRowTo < lRowFrom Then Exit Function

    Dim rgOddLInes As Range
    Dim rgEvenLInes As Range
    Dim lEachRow As Long

    For lEachRow = lRowFrom To lRowTo
        If (lEachRow Mod 2) = 0 Then
            If rgEvenLInes Is Nothing Then
                Set rgEvenLInes = fGetRangeByStartEndPos(shtOutput, lEachRow, 1, lEachRow, lMaxCol)
            Else
                Set rgEvenLInes = Union(rgEvenLInes, fGetRangeByStartEndPos(shtOutput, lEachRow, 1, lEachRow, lMaxCol))
            End If
        Else
            If rgOddLInes Is Nothing Then
                Set rgOddLInes = fGetRangeByStartEndPos(shtOutput, lEachRow, 1, lEachRow, lMaxCol)
            Else
                Set rgOddLInes = Union(rgOddLInes, fGetRangeByStartEndPos(shtOutput, lEachRow, 1, lEachRow, lMaxCol))
            End If
        End If
    Next

    Dim sAddr As String
    If Not rgEvenLInes Is Nothing Then
        'sAddr = fGetSpecifiedConfigCellAddress(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=REPORT_EVEN_LINE_COLOR")
        sAddr = fGetSysMiscConfig("REPORT_EVEN_LINE_COLOR")
        rgEvenLInes.Interior.Color = fGetRangeFromExternalAddress(sAddr).Interior.Color
    End If
    If Not rgOddLInes Is Nothing Then
       ' sAddr = fGetSpecifiedConfigCellAddress(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=REPORT_ODD_LINE_COLOR")
        sAddr = fGetSysMiscConfig("REPORT_ODD_LINE_COLOR")
        rgOddLInes.Interior.Color = fGetRangeFromExternalAddress(sAddr).Interior.Color
    End If
    Set rgEvenLInes = Nothing
    Set rgOddLInes = Nothing
End Function

Function fDeleteAllConditionFormatFromSheet(ByRef shtParam As Worksheet)
    shtParam.Cells.FormatConditions.Delete
End Function
Function fDeleteAllConditionFormatForAllSheets(Optional wb As Workbook)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim sht As Worksheet
    For Each sht In wb.Worksheets
        Call fDeleteAllConditionFormatFromSheet(sht)
    Next

    Set sht = Nothing
End Function
Function fSetConditionFormatForOddEvenLine(ByRef shtParam As Worksheet, Optional lMaxCol As Long = 0 _
                                            , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0 _
                                            , Optional arrKeyColsNotBlank _
                                            , Optional bExtendToMore10ThousRows As Boolean = False)
'arrKeyColsNotBlank
'    1. singlecol: 1
'    1. array(1,2,3)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtParam)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtParam)

    If lMaxCol <= 0 Then Exit Function

    If bExtendToMore10ThousRows Then lRowTo = lRowTo + 100000

    If lRowTo < lRowFrom Then Exit Function

    Dim rngCondFormat As Range
    Set rngCondFormat = fGetRangeByStartEndPos(shtParam, lRowFrom, 1, lRowTo, lMaxCol)

    Dim sAddr As String
    Dim sKeyColsFormula As String
    Dim sFormula As String
    Dim lColor As Long
    Dim i As Integer
    Dim sColLetter As String
    Dim aFormatCondition As FormatCondition

    If Not IsMissing(arrKeyColsNotBlank) Then
        If IsArray(arrKeyColsNotBlank) Then
            For i = LBound(arrKeyColsNotBlank) To UBound(arrKeyColsNotBlank)
                sColLetter = fNum2Letter(arrKeyColsNotBlank(i))
                sKeyColsFormula = sKeyColsFormula & "," & "len(trim($" & sColLetter & lRowFrom & ")) > 0"
            Next
            If Len(sKeyColsFormula) > 0 Then sKeyColsFormula = Right(sKeyColsFormula, Len(sKeyColsFormula) - 1)
            sKeyColsFormula = sKeyColsFormula & ","
        Else
            sColLetter = fNum2Letter(arrKeyColsNotBlank)
            sKeyColsFormula = "len(trim($" & sColLetter & lRowFrom & ")) > 0"
            sKeyColsFormula = sKeyColsFormula
        End If
    Else
        sKeyColsFormula = ""
    End If

    sFormula = "=And( " & sKeyColsFormula & "mod(row(),2)=0)"

    Set aFormatCondition = rngCondFormat.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
    aFormatCondition.SetFirstPriority
    aFormatCondition.StopIfTrue = False

    'sAddr = fGetSpecifiedConfigCellAddress(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=REPORT_EVEN_LINE_COLOR")
    sAddr = fGetSysMiscConfig("REPORT_EVEN_LINE_COLOR")
    lColor = fGetRangeFromExternalAddress(sAddr).Interior.Color
    aFormatCondition.Interior.Color = lColor

    sFormula = "=And( " & sKeyColsFormula & "mod(row(),2)<>0)"
    Set aFormatCondition = rngCondFormat.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
    aFormatCondition.SetFirstPriority
    aFormatCondition.StopIfTrue = False

    'sAddr = fGetSpecifiedConfigCellAddress(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=REPORT_ODD_LINE_COLOR")
    sAddr = fGetSysMiscConfig("REPORT_ODD_LINE_COLOR")
    lColor = fGetRangeFromExternalAddress(sAddr).Interior.Color
    aFormatCondition.Interior.Color = lColor

    Set aFormatCondition = Nothing
End Function

Function fSetConditionFormatForBorders(ByRef shtParam As Worksheet, Optional lMaxCol As Long = 0 _
                                            , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0 _
                                            , Optional arrKeyColsNotBlank _
                                            , Optional bExtendToMore10ThousRows As Boolean = False)
'arrKeyColsNotBlank
'    1. singlecol: 1
'    1. array(1,2,3)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtParam)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtParam)

    If lMaxCol <= 0 Then Exit Function
    If bExtendToMore10ThousRows Then lRowTo = lRowTo + 100000

    If lRowTo < lRowFrom Then Exit Function

    Dim rngCondFormat As Range
    Set rngCondFormat = fGetRangeByStartEndPos(shtParam, lRowFrom, 1, lRowTo, lMaxCol)

    Dim sAddr As String
    Dim sKeyColsFormula As String
    Dim sFormula As String
    Dim lColor As Long
    Dim i As Integer
    Dim sColLetter As String
    Dim aFormatCondition As FormatCondition

    If Not IsMissing(arrKeyColsNotBlank) Then
        If IsArray(arrKeyColsNotBlank) Then
            For i = LBound(arrKeyColsNotBlank) To UBound(arrKeyColsNotBlank)
                sColLetter = fNum2Letter(arrKeyColsNotBlank(i))
                sKeyColsFormula = sKeyColsFormula & "," & "len(trim($" & sColLetter & lRowFrom & ")) > 0"
            Next
            If Len(sKeyColsFormula) > 0 Then sKeyColsFormula = Right(sKeyColsFormula, Len(sKeyColsFormula) - 1)
        Else
            sColLetter = fNum2Letter(arrKeyColsNotBlank)
            sKeyColsFormula = "len(trim($" & sColLetter & lRowFrom & ")) > 0"
            sKeyColsFormula = sKeyColsFormula
        End If
    Else
        sKeyColsFormula = ""
    End If

    sFormula = "=And( " & sKeyColsFormula & ")"

    Set aFormatCondition = rngCondFormat.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
    aFormatCondition.SetFirstPriority
    aFormatCondition.StopIfTrue = False

    aFormatCondition.Borders(xlLeft).Weight = xlHairline
    aFormatCondition.Borders(xlRight).Weight = xlHairline
    aFormatCondition.Borders(xlTop).Weight = xlHairline
    aFormatCondition.Borders(xlBottom).Weight = xlHairline

    Set aFormatCondition = Nothing
End Function
Function fSetFormatBoldOrangeBorderForHeader(ByRef sht As Worksheet, Optional lMaxCol As Long = 0 _
                                            , Optional lHeaderRowFrom As Long = 1, Optional lHeaderRowTo As Long = 1 _
                                            )
    If lMaxCol <= 0 Then lMaxCol = fGetValidMaxCol(sht)

    fSetFormatBoldOrangeBorderForRangeEspeciallyForHeader fGetRangeByStartEndPos(sht, lHeaderRowFrom, 1, lHeaderRowTo, lMaxCol)
End Function
Function fSetFormatBoldOrangeBorderForRangeEspeciallyForHeader(ByRef rgTarget As Range)
    Dim lColor As Long
    Dim sAddr As String

    'sAddr = fGetSpecifiedConfigCellAddress(shtSysConf, "[System Misc Settings]", "Value", "Setting Item ID=REPORT_HEADER_LINE_COLOR")
    sAddr = fGetSysMiscConfig("REPORT_HEADER_LINE_COLOR")
    lColor = fGetRangeFromExternalAddress(sAddr).Interior.Color

    With rgTarget
        .Font.Bold = True
        .Interior.Color = lColor

        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
        End With
    End With
End Function

Function fCheckIfGotBusinessError(Optional bMsgbox As Boolean = True) As Boolean
    If fNzero(gsBusinessErrorMsg) Then
        If bMsgbox Then fMsgBox gsBusinessErrorMsg
        fCheckIfGotBusinessError = True
        Exit Function
    End If

    If Err.Number <> 0 Or gErrNum <> 0 Then
        If Err.Number = vbObjectError + BUSINESS_ERROR_NUMBER Or gErrNum = vbObjectError + BUSINESS_ERROR_NUMBER Then
            fCheckIfGotBusinessError = True
            Exit Function
        End If
    End If

    If gErrNum <> 0 Then
        fCheckIfGotBusinessError = True:            Exit Function
    Else
        'fCheckIfGotBusinessError = False:            Exit Function
    End If

    fCheckIfGotBusinessError = False
End Function
Function fCheckIfUnCapturedExceptionAbnormalError() As Boolean
    If Err.Number <> 0 And Err.Number <> vbObjectError + BUSINESS_ERROR_NUMBER Then
        fCheckIfUnCapturedExceptionAbnormalError = True
'
'        If Err.Number = vbObjectError + CONFIG_ERROR_NUMBER Then
'        Else
            fMsgBox "Error has occurred:" _
                    & vbCr & vbCr _
                    & "Error Number: " & Err.Number & vbCr _
                    & "Error Description:" & Err.Description
'        End If
        Exit Function
    End If

    fCheckIfUnCapturedExceptionAbnormalError = False
End Function

Function fPrepareHeaderToSheet(shtParam As Worksheet, arrHeaders, Optional alHeaderAtRow As Long = 1)
    Dim i As Integer
    Dim iV As Integer
    Dim arrHeaderHorizontal()

    ReDim arrHeaderHorizontal(1 To 1, 1 To UBound(arrHeaders) - LBound(arrHeaders) + 1)
    iV = 0
    For i = LBound(arrHeaders) To UBound(arrHeaders)
        iV = iV + 1
        arrHeaderHorizontal(1, iV) = arrHeaders(i)
    Next

    shtParam.Cells(alHeaderAtRow, 1).Resize(1, iV).Value = arrHeaderHorizontal
    Erase arrHeaderHorizontal
    Erase arrHeaders
End Function

Function fReadInputFileSpecConfigItem(asFileTag As String, asWhatToReturn As String _
                            , Optional shtData As Worksheet, Optional alDataFromRow As Long = 2)
    Dim bReadWholeSheetData As Boolean
    Dim dictLetterIndex As Dictionary
    Dim dictArrayIndex As Dictionary
    Dim dictColFormat As Dictionary
    Dim dictRawType As Dictionary
    Dim dictDisplayName As Dictionary

    Dim sFileSpecTag As String
    Dim shtToRead As Worksheet

    sFileSpecTag = fGetInputFileFileSpecTag(asFileTag)

    If shtData Is Nothing Then
        Set shtToRead = fGetInputFileSheetAfterLoadingToThisWorkBook(asFileTag)
    Else
        Set shtToRead = shtData
    End If

    Call fReadInputFileSpecConfig(sFileSpecTag:=sFileSpecTag _
                                , dictLetterIndex:=dictLetterIndex _
                                , dictArrayIndex:=dictArrayIndex _
                                , dictDisplayName:=dictDisplayName _
                                , dictRawType:=dictRawType _
                                , dictDataFormat:=dictColFormat _
                                , bReadWholeSheetData:=bReadWholeSheetData _
                                , shtData:=shtToRead _
                                , alHeaderAtRow:=alDataFromRow - 1)
    Set shtToRead = Nothing

    If asWhatToReturn = "LETTER_INDEX" Then
        Set fReadInputFileSpecConfigItem = dictLetterIndex
    Else
        fErr "wrong param: " & asWhatToReturn
    End If

    Set dictLetterIndex = Nothing
    Set dictArrayIndex = Nothing
    Set dictColFormat = Nothing
    Set dictRawType = Nothing
End Function

Function fSetFormatForExceptionCells(shtOutput As Worksheet, dErrOrWarningRows As Dictionary, asColorTag As String)
    Dim sColorAddr As String
    Dim lColor  As Long

    If dErrOrWarningRows Is Nothing Then Exit Function
    If dErrOrWarningRows.Count <= 0 Then Exit Function

    sColorAddr = fGetSysMiscConfig(asColorTag)
    lColor = fGetRangeFromExternalAddress(sColorAddr).Interior.Color

    Dim rgTarget As Range

    Dim i As Long
    Dim j As Integer
    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim arrExceptionCols

    For i = 0 To dErrOrWarningRows.Count - 1
        lEachRow = dErrOrWarningRows.Keys(i)

        arrExceptionCols = Split(dErrOrWarningRows.Items(i), DELIMITER)
        For j = LBound(arrExceptionCols) To UBound(arrExceptionCols)
            lEachCol = arrExceptionCols(j)

            If rgTarget Is Nothing Then
                Set rgTarget = shtOutput.Cells(lEachRow, lEachCol)
            Else
                Set rgTarget = Union(rgTarget, shtOutput.Cells(lEachRow, lEachCol))
            End If
        Next
    Next

    If Not rgTarget Is Nothing Then rgTarget.Interior.Color = lColor

    Set rgTarget = Nothing
End Function

Function fGetDictionayDelimiteredItemsCount(ByRef dict As Dictionary, Optional sDelimiter As String = ",") As Long
    Dim lCount  As Long
    Dim i As Long

    lCount = 0
    For i = 0 To dict.Count - 1
        lCount = lCount + UBound(Split(dict.Items(i), sDelimiter)) + 1
    Next

    fGetDictionayDelimiteredItemsCount = lCount
End Function

Function fWorkbookHasMoreThanOneSheetVisible(Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook

    Dim sht As Worksheet
    Dim iCnt As Integer

    iCnt = 0
    For Each sht In wb.Worksheets
        If sht.Visible = xlSheetVisible Then
            iCnt = iCnt + 1
        End If

        If iCnt > 1 Then Exit For
    Next

    Set sht = Nothing

    fWorkbookHasMoreThanOneSheetVisible = (iCnt > 1)
End Function

Enum InputFile
    [_first] = 1
    ReportID = 1
    FileTag = 2
    FilePath = 3
    source = 4
    ReLoadOrNot = 5
    FileSpecTag = 6
    Env = 7
    DefaultSheet = 8
    PivotTableTag = 9
    RowNo = 10
    [_last] = 10
End Enum

Function fFindAllColumnsIndexByColNames(rngToFindIn As Range, arrColsName, ByRef arrColsIndex() _
                                , Optional ByRef alHeaderAtRow As Long, Optional bReturnLetter As Boolean = False)
    If fArrayIsEmptyOrNoData(arrColsName) Then fErr "arrColsName is empty."
    If fArrayHasBlankValue(arrColsName) Then fErr "arrColsName has blank element." & vbCr & Join(arrColsName, vbCr)
    If fArrayHasDuplicateElement(arrColsName) Then fErr "arrColsName has duplicate element."

    ReDim arrColsIndex(LBound(arrColsName) To UBound(arrColsName))

    Dim lColAtRow As Long
    Dim lEachCol As Long
    Dim sEachColName As String
    Dim rngFound As Range

    lColAtRow = 0
    For lEachCol = LBound(arrColsName) To UBound(arrColsName)
        sEachColName = Trim(arrColsName(lEachCol))
        sEachColName = Replace(sEachColName, "*", "~*")

        Set rngFound = fFindInWorksheet(rngToFindIn, sEachColName)

        If lColAtRow <> 0 Then
            If lColAtRow <> rngFound.Row Then
                fErr "Columns are not at the same row."
            End If
        Else
            lColAtRow = rngFound.Row
        End If

        If bReturnLetter Then
            arrColsIndex(lEachCol) = fNum2Letter(rngFound.Column)
        Else
            arrColsIndex(lEachCol) = rngFound.Column
        End If
    Next

    alHeaderAtRow = lColAtRow
    Set rngFound = Nothing
End Function
'
'Function fValidateDuplicateKeys(arrConfigData(), arrColsIndex(), arrKeyCols, lHeaderAtRow As Long, lStartCol As Long)
'    If fArrayIsEmptyOrNoData(arrKeyCols) Then Exit Function
'
'    Dim lEachRow As Long
'    Dim lEachCol As Long
'    Dim i As Long
'    Dim sKeyStr As String
'    Dim dict As New Dictionary
'
'    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
'        sKeyStr = ""
'
'        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
'            'lEachCol = arrColsIndex(arrKeyCols(i) - 1)
'            lEachCol = arrColsIndex(arrKeyCols(i))
'            sKeyStr = sKeyStr & Trim(CStr(arrConfigData(lEachRow, lEachCol)))
'        Next
'
'        If dict.Exists(sKeyStr) Then
'            fErr "Duplicate key " & sKeyStr & " was found " & vbCr & "at row: " & (lHeaderAtRow + lEachRow) _
'                     & ", column: " & fNum2Letter((lStartCol + lEachCol))
'        Else
'            dict.Add sKeyStr, 0
'        End If
'    Next
'
'    Set dict = Nothing
'End Function

Function fValidateDuplicateKeysForConfigBlock(arrConfigData(), arrColsIndex(), arrKeyCols, lHeaderAtRow As Long, lStartCol As Long)
    If fArrayIsEmptyOrNoData(arrKeyCols) Then Exit Function

    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long
    Dim sKeyStr As String
    Dim dict As New Dictionary

    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        sKeyStr = ""

        For i = LBound(arrKeyCols) To UBound(arrKeyCols)
            'lEachCol = arrColsIndex(arrKeyCols(i) - 1)
            lEachCol = arrColsIndex(arrKeyCols(i))
            sKeyStr = sKeyStr & Trim(CStr(arrConfigData(lEachRow, lEachCol)))
        Next

        If dict.Exists(sKeyStr) Then
            fErr "Duplicate key " & sKeyStr & " was found " & vbCr & "at row: " & (lHeaderAtRow + lEachRow) _
                     & ", column: " & fNum2Letter((lStartCol + lEachCol))
        Else
            dict.Add sKeyStr, 0
        End If
    Next

    Set dict = Nothing
End Function

'Function fReadConfigBlockToArrayValidated(asTag As String, rngToFindIn As Range, arrColsName
Function fReadConfigBlockToArrayValidated(asTag As String, shtParam As Worksheet, arrColsName _
                                , Optional arrKeyCols _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                , Optional bNetValues As Boolean = True) As Variant
    'arrKeyCols:  array(1, 2, 3, 5), or unnecessary: array()
    Dim arrConfigData()
    Dim arrColsIndex()
    Dim arrOut()

    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtParam, arrColsName:=arrColsName _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=abNoDataConfigThenError _
                                )

    If fArrayIsEmptyOrNoData(arrConfigData) Then GoTo exit_fun

    'Call fValidateDuplicateKeys(arrConfigData, arrColsIndex, arrKeyCols, lOutConfigHeaderAtRow, lConfigStartCol)

    If bNetValues Then
        ReDim arrOut(LBound(arrConfigData, 1) To UBound(arrConfigData, 1), 1 To UBound(arrColsIndex) - LBound(arrColsIndex) + 1)

        Dim lEachRow As Long
        Dim lEachCol As Long
        Dim i As Long

        For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
            'i = LBound(arrColsIndex) + 1
            i = LBound(arrColsIndex)
            For lEachCol = LBound(arrColsIndex) To UBound(arrColsIndex)
                arrOut(lEachRow, i) = arrConfigData(lEachRow, arrColsIndex(lEachCol))
                i = i + 1
            Next
        Next
    End If
exit_fun:
    Erase arrColsIndex

    If bNetValues Then
        fReadConfigBlockToArrayValidated = arrOut
    Else
        fReadConfigBlockToArrayValidated = arrConfigData
    End If

    Erase arrConfigData
    Erase arrOut
End Function
'Function fReadConfigBlockToArrayNet(asTag As String, rngToFindIn As Range, arrColsName()
Function fReadConfigBlockToArrayNet(asTag As String, shtParam As Worksheet, arrColsName _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                ) As Variant
    Dim arrOut()
    Dim arrColsIndex()
    Dim arrConfigData()

    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtParam, arrColsName:=arrColsName _
                                , arrConfigData:=arrConfigData _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lOutConfigHeaderAtRow _
                                , abNoDataConfigThenError:=abNoDataConfigThenError _
                                )
    If fArrayIsEmptyOrNoData(arrConfigData) Then GoTo exit_fun

    ReDim arrOut(LBound(arrConfigData, 1) To UBound(arrConfigData, 1), LBound(arrColsIndex) To UBound(arrColsIndex))

    Dim lEachRow As Long
    Dim lEachCol As Long
    Dim i As Long

    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        i = LBound(arrColsIndex)
        For lEachCol = LBound(arrColsIndex) To UBound(arrColsIndex)
            arrOut(lEachRow, i) = arrConfigData(lEachRow, arrColsIndex(lEachCol))
            i = i + 1
        Next
    Next
exit_fun:
    Erase arrColsIndex
    Erase arrConfigData
    fReadConfigBlockToArrayNet = arrOut
    Erase arrOut
End Function
'Function fReadConfigBlockToArray(asTag As String, rngToFindIn As Range, arrColsName
Function fReadConfigBlockToArray(asTag As String, shtParam As Worksheet, arrColsName _
                                , ByRef arrConfigData() _
                                , ByRef arrColsIndex() _
                                , Optional ByRef lConfigStartRow As Long _
                                , Optional ByRef lConfigStartCol As Long _
                                , Optional ByRef lConfigEndRow As Long _
                                , Optional ByRef lOutConfigHeaderAtRow As Long _
                                , Optional abNoDataConfigThenError As Boolean = False _
                                )
    arrConfigData = Array()

    'Dim shtConfig As Worksheet
    'Set shtConfig = rngToFindIn.Parent
    Dim rngToFindIn As Range

    Call fReadConfigBlockStartEnd(asTag, shtParam, lConfigStartRow, lConfigStartCol, lConfigEndRow)

    If lConfigEndRow < lConfigStartRow + 1 Then
        If abNoDataConfigThenError Then
            fErr "No data is configured under tag " & asTag & " in sheet " & shtParam.Name & vbCr _
                    & "You must leave at least one blank line after the tag."
        End If
    End If

    Set rngToFindIn = fGetRangeByStartEndPos(shtParam, lConfigStartRow, lConfigStartCol, lConfigEndRow, Columns.Count)
    Call fFindAllColumnsIndexByColNames(rngToFindIn, arrColsName, arrColsIndex, lOutConfigHeaderAtRow)

    Dim lColsMinCol As Long
    Dim lColsMaxCol As Long

    lColsMinCol = Application.WorksheetFunction.Min(arrColsIndex)
    lColsMaxCol = Application.WorksheetFunction.Max(arrColsIndex)

    lConfigEndRow = fGetValidMaxRowOfRange(fGetRangeByStartEndPos(shtParam, lConfigStartRow, lConfigStartCol, lConfigEndRow, lColsMaxCol))

    If lConfigEndRow > lOutConfigHeaderAtRow Then
        arrConfigData = fReadRangeDatatoArrayByStartEndPos(shtParam, lOutConfigHeaderAtRow + 1, lColsMinCol, lConfigEndRow, lColsMaxCol)
    End If

    lConfigStartCol = lColsMinCol

    Dim lEachCol As Long
    'change 10, 15, 20, to 1, 6, 11
    For lEachCol = UBound(arrColsIndex) To LBound(arrColsIndex) Step -1
        arrColsIndex(lEachCol) = arrColsIndex(lEachCol) - lColsMinCol + 1
    Next

    'Set shtConfig = Nothing
End Function

Function fFindRageOfFileSpecConfigBlock(sFileSpecTag As String) As Range
    Dim rngConfigBlock As Range
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long

    Call fReadConfigBlockStartEnd(sFileSpecTag, shtFileSpec, lConfigStartRow, lConfigStartCol, lConfigEndRow)
    Set fFindRageOfFileSpecConfigBlock = fGetRangeByStartEndPos(shtFileSpec, lConfigStartRow, lConfigStartCol, lConfigEndRow, Columns.Count)
End Function
'Function fReadConfigBlockStartEnd(sFileSpecTag As String, rngToFindIn As Range
Function fReadConfigBlockStartEnd(sFileSpecTag As String, shtParam As Worksheet _
                                , ByRef lOutBlockStartRow As Long _
                                , ByRef lOutBlockStartCol As Long _
                                , ByRef lOutBlockEndRow As Long)

    'Dim shtSource As Worksheet
    Dim lMaxRow As Long
    Dim rngTagFound As Range
    Dim lTagRow As Long
    Dim lTagCol As Long

    'Set shtSource = rngToFindIn.Parent
    lMaxRow = fGetValidMaxRow(shtParam)

    Set rngTagFound = fFindInWorksheet(shtParam.Cells, sFileSpecTag)
    lTagRow = rngTagFound.Row
    lTagCol = rngTagFound.Column

    Set rngTagFound = fFindInWorksheet(fGetRangeByStartEndPos(shtParam, lTagRow + 1, lTagCol, lMaxRow, lTagCol) _
                                    , "[*]", False, True)
    If rngTagFound Is Nothing Then
        lOutBlockEndRow = lMaxRow
    Else
        lOutBlockEndRow = rngTagFound.Row - 1
    End If

    lOutBlockStartRow = lTagRow + 1
    lOutBlockStartCol = lTagCol

    'Set shtSource = Nothing
    Set rngTagFound = Nothing
End Function

Function fReadConfigInputFiles(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID

    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Input Files]"
    ReDim arrColsName(InputFile.ReportID To InputFile.PivotTableTag)

    arrColsName(InputFile.ReportID) = "Report ID"
    arrColsName(InputFile.FileTag) = "File Tag"
    arrColsName(InputFile.FilePath) = "File Full Path"
    arrColsName(InputFile.source) = "Source"
    arrColsName(InputFile.ReLoadOrNot) = "When Data Already Loaded To Sheet"
    arrColsName(InputFile.FileSpecTag) = "File Spec Tag"
    arrColsName(InputFile.Env) = "DEV/UAT/PROD"
    arrColsName(InputFile.DefaultSheet) = "Which Sheet To Import"
    arrColsName(InputFile.PivotTableTag) = "Pivot Table Tag To Be Created From This Data Source"

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )

    Call fValidateDuplicateInArray(arrConfigData, Array(InputFile.ReportID, InputFile.FileTag, InputFile.Env), False _
        , shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID + File Tag")
    Call fValidateBlankInArray(arrConfigData, InputFile.ReportID, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID")
    Call fValidateBlankInArray(arrConfigData, InputFile.FileTag, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "File Tag")
    Call fValidateBlankInArray(arrConfigData, InputFile.source, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Source")
    Call fValidateBlankInArray(arrConfigData, InputFile.Env, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "DEV/UAT/PROD")

    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sRptNameStr As String
    Dim sValueStr As String
    Dim sFileTag As String
    Dim sSource As String
    Dim sEnv As String
    Dim sShtToImport  As String
    Dim sFileName  As String

    Dim sPos As String
    sPos = vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "

    Set gDictInputFiles = New Dictionary

    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row

        sRptNameStr = DELIMITER & Trim(arrConfigData(lEachRow, InputFile.ReportID)) & DELIMITER
        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row

        sEnv = Trim(arrConfigData(lEachRow, InputFile.Env))

        If Not (sEnv = gsEnv Or sEnv = "SHARED") Then GoTo next_row

        lActualRow = lConfigHeaderAtRow + lEachRow

        sSource = Trim(arrConfigData(lEachRow, InputFile.source))
        'sShtToImport = arrConfigData(lEachRow, InputFile.DefaultSheet)
        sFileName = Trim(arrConfigData(lEachRow, InputFile.FilePath))

        If sSource = "READ_PRE_EXISTING_SHEET" Then
            If Not fSheetExists(sFileName) Then
                fErr "READ_PRE_EXISTING_SHEET, but sheet specified does not exists, you may have to run the previous steps first:" _
                & sFileName & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(InputFile.DefaultSheet)
            End If
        ElseIf sSource = "READ_PRE_EXISTING_SHEET" Then
        End If

        sFileTag = Trim(arrConfigData(lEachRow, InputFile.FileTag))
        sValueStr = fComposeStrForInputFile(arrConfigData, lEachRow)

        gDictInputFiles.Add sFileTag, sValueStr
        Call fUpdateDictionaryItemValueForDelimitedElement(gDictInputFiles, sFileTag, InputFile.RowNo - InputFile.FileTag, lActualRow)
next_row:
    Next

    Erase arrConfigData
    Erase arrColsName
End Function

Function fComposeStrForInputFile(arrConfigData, lEachRow As Long) As String
    Dim sOut As String
    Dim i As Integer

    For i = InputFile.FilePath To InputFile.PivotTableTag
        sOut = sOut & DELIMITER & Trim(arrConfigData(lEachRow, i))
    Next

    fComposeStrForInputFile = Right(sOut, Len(sOut) - 1)
End Function

Function fReadConfigWholeColsToDictionary(shtConfig As Worksheet, asTag As String, asKeyNotNullCol As String, asRtnCol As String) As Dictionary
    If fZero(asTag) Or fZero(asKeyNotNullCol) Or fZero(asRtnCol) Then fErr "Wrong param"

    Dim bRtnColIsKeyCol As Boolean
    bRtnColIsKeyCol = (Trim(asKeyNotNullCol) = Trim(asRtnCol))

    Dim arrColNames()
    ReDim arrColNames(0 To 1)
    arrColNames(0) = asKeyNotNullCol
    arrColNames(1) = asRtnCol

    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    arrKeyColsForValidation = Array(1, 2)

'    arrConfigData = fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=shtConfig.Cells _
'                                , arrColsName:=arrColsName _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
End Function

'
Function fReadConfigWholeMultipleColsToArray(shtConfig As Worksheet, asTag As String, arrColsName) As Variant
'arrKeyColsForValidation : Array(1, 2, 5)
    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Erase arrColsName
    fReadConfigWholeMultipleColsToArray = arrConfigData
    Erase arrConfigData
End Function

'
Function fGetReadConfigWholePairColsValueAsArray(shtConfig As Worksheet, asTag As String, arrFetchCols, Optional arrKeyColsForValidation) As Variant
    Dim dict As Dictionary
    Dim arrOut()

    Set dict = fGetReadConfigWholePairColsValueAsDictionary()

    Call fCopyDictionaryItems2Array(dict, arrOut)

    Set dict = Nothing

    fGetReadConfigWholePairColsValueAsArray = arrOut
    Erase arrOut
End Function

Function fGetReadConfigWholePairColsValueAsDictionary(shtConfig As Worksheet, asTag As String _
                    , asKeyNotNullCol As String, asRtnCol As String) As Dictionary
    If fZero(asKeyNotNullCol) Or fZero(asRtnCol) Then fErr "Wrong param"

    Dim bRtnColIsKeyCol As Boolean
    bRtnColIsKeyCol = (Trim(asKeyNotNullCol) = Trim(asRtnCol))

    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    ReDim arrColsName(1 To 2)
    arrColsName(1) = Trim(asKeyNotNullCol)
    arrColsName(2) = Trim(asRtnCol)

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Call fValidateDuplicateInArray(arrConfigData, 1, False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, asKeyNotNullCol)
    If Not bRtnColIsKeyCol Then
        Call fValidateDuplicateInArray(arrConfigData, 2, False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, asRtnCol)
    End If

'    Call fValidateBlankInArray(arrConfigData, Company.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID")

    Set fReadConfigCompanyList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, Company.REPORT_ID _
            , Array(Company.id, Company.Name, Company.Commission, Company.Selected), DELIMITER)
    Erase arrColsName
    Erase arrConfigData

End Function

'Function fGetReadConfigWholeSingleColValueAsArray(shtConfig As Worksheet, asTag As String, arrFetchCols, Optional arrKeyColsForValidation) As Variant
'    Dim dict As Dictionary
'    Dim arrOut()
'
'    Set dict = fGetReadConfigWholeSingleColValueAsDictionary()
'
'    Call fCopyDictionaryItems2Array(dict, arrOut)
'
'    Set dict = Nothing
'
'    fGetReadConfigWholePairColsValueAsArray = arrOut
'    Erase arrOut
'End Function
Function fGetReadConfigWholeSingleColValueAsArray(shtConfig As Worksheet, asTag As String _
                    , asColName As String _
                    , Optional IgnoreBlankKeys As Boolean = False _
                    , Optional WhenKeyIsDuplicateError As Boolean = True) As Variant
    If fZero(asColName) Then fErr "Wrong param"

    'Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    ReDim arrColsName(1 To 1)
    arrColsName(1) = Trim(asColName)

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtConfig _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )
    Erase arrColsName

    If WhenKeyIsDuplicateError Then
        Call fValidateDuplicateInArray(arrConfigData, 1, True, shtConfig, lConfigHeaderAtRow, lConfigStartCol, asColName)
    End If

    If Not IgnoreBlankKeys Then
        Call fValidateBlankInArray(arrConfigData, 1, shtConfig, lConfigHeaderAtRow, lConfigStartCol, asColName)
    End If

    Dim dict As Dictionary
    Set dict = fReadArray2DictionaryOnlyKeys(arrConfigData, 1, IgnoreBlankKeys, WhenKeyIsDuplicateError)

    Dim arrOut()
    Call fCopyDictionaryKeys2Array(dict, arrOut)

    Erase arrConfigData
    Set dict = Nothing

    fGetReadConfigWholeSingleColValueAsArray = arrOut
End Function

Function fReadSysConfig_InputTxtSheetFile(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID

    Call fReadConfigInputFiles(asReportID)
    Call fReadTxtFileImportConfig

    Dim i As Long
    Dim sFileTag As String
    Dim sSource As String
    Dim sDependantFileTag As String
    Dim sDependantFileOrSheet As String
    Dim sFileSpec As String

    For i = 0 To gDictInputFiles.Count - 1
        sFileTag = gDictInputFiles.Keys(i)
        sSource = fGetInputFileSourceType(sFileTag)

        If sSource = "READ_PREV_STEP_OUTPUT_SHEET" Or sSource = "READ_PREV_STEP_OUTPUT_FILE" Then
            sDependantFileTag = fGetInputFileFileName(sFileTag)

            sDependantFileOrSheet = fGetReneratedReport(sDependantFileTag)

            If sSource = "READ_PREV_STEP_OUTPUT_SHEET" Then
                If Not fSheetExists(sDependantFileOrSheet) Then
                    fErr "Dependant ourput sheet as below does not exist, you may have to run the previous steps first." _
                        & vbCr & "File Tag: " & sFileTag _
                        & vbCr & "sDependantFileTag Tag: " & sDependantFileTag _
                        & vbCr & "sDependantFileOrSheet: " & sDependantFileOrSheet
                End If
            Else
                If Not fFileExists(sDependantFileOrSheet) Then
                    fErr "Dependant ourput FILE as below does not exist, you may have to run the previous steps first." _
                        & vbCr & "File Tag: " & sFileTag _
                        & vbCr & "sDependantFileTag Tag: " & sDependantFileTag _
                        & vbCr & "sDependantFileOrSheet: " & sDependantFileOrSheet
                End If
            End If

            Call fUpdateGDictInputFile_FileName(sFileTag, sDependantFileOrSheet)

            sFileSpec = fGetOutputReportItem(sDependantFileTag, "FILE_SPEC")
            Call fUpdateGDictInputFile_FileSpecTag(sFileTag, sFileSpec)
        End If
    Next

    Call fCrossValidateInputFileTxtSheetFile
End Function

Function fReadTxtFileImportConfig(Optional asReportID As String = "")
    If asReportID = "" Then asReportID = gsRptID

    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Input Txt Files Import Config]"

    Const REPORT_ID = 1
    Const FILE_TAG = 2
    Const COL_DELIMITER = 3
    Const PLATFORM = 4
    ReDim arrColsName(REPORT_ID To PLATFORM)

    arrColsName(REPORT_ID) = "Report ID"
    arrColsName(FILE_TAG) = "File Tag"
    arrColsName(COL_DELIMITER) = "Column Delimiter"
    arrColsName(PLATFORM) = "TextFilePlatForm"

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True _
                                )

    Call fValidateDuplicateInArray(arrConfigData, Array(REPORT_ID, FILE_TAG), False _
        , shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID + File Tag")
    Call fValidateBlankInArray(arrConfigData, REPORT_ID, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID")
    Call fValidateBlankInArray(arrConfigData, FILE_TAG, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "File Tag")
    Call fValidateBlankInArray(arrConfigData, COL_DELIMITER, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Column Delimiter")
    Call fValidateBlankInArray(arrConfigData, PLATFORM, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "TextFilePlatForm")

    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sRptNameStr As String
    Dim sValueStr As String
    Dim sFileTag As String

    Dim sPos As String
    sPos = vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "

    Set gDictTxtFileSpec = New Dictionary

    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row

        sRptNameStr = DELIMITER & Trim(arrConfigData(lEachRow, REPORT_ID)) & DELIMITER
        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row

        lActualRow = lConfigHeaderAtRow + lEachRow

        sFileTag = Trim(arrConfigData(lEachRow, FILE_TAG))

        sValueStr = Trim(arrConfigData(lEachRow, COL_DELIMITER)) & DELIMITER _
                & Trim(arrConfigData(lEachRow, PLATFORM))

        gDictTxtFileSpec.Add sFileTag, sValueStr
next_row:
    Next

    Erase arrConfigData
    Erase arrColsName
End Function

Function fReadSysConfig_Output(Optional asReportID As String, Optional asRptType As String) As String
    Dim sRptFileName As String
    Dim asFileSpecTag As String

    If asReportID = "" Then asReportID = gsRptID
    sRptFileName = fReadConfigOutputFiles(asReportID, asRptType, asFileSpecTag)

    If fNzero(asFileSpecTag) Then
        Call fGetOutputReportColsConfig(asFileSpecTag)
    End If

    fReadSysConfig_Output = sRptFileName
End Function

Function fReadConfigOutputFiles(Optional asReportID As String = "" _
            , Optional asRptType As String, Optional asFileSpecTag As String) As String
    If asReportID = "" Then asReportID = gsRptID

    Dim asTag As String
    Dim arrColsName()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    asTag = "[Output Files]"

    ReDim arrColsName(1 To 5)
    arrColsName(1) = "Report ID"
    arrColsName(2) = "Output Type"
    arrColsName(3) = "File Full Path or Sheet Name"
    arrColsName(4) = "File Spec Tag"
    arrColsName(5) = "DEV/UAT/PROD"

    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtSysConf _
                                , arrColsName:=arrColsName _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True)

    Call fValidateDuplicateInArray(arrConfigData, Array(1, 5), False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Report ID")
    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
    Call fValidateBlankInArray(arrConfigData, 2, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")
    Call fValidateBlankInArray(arrConfigData, 5, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")

    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sRptNameStr As String
    Dim sEnv As String
    Dim sRptFileName As String

    Dim bFound As Boolean
    Dim sPos As String
    sPos = vbCr & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "

    bFound = False
    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row

        sRptNameStr = DELIMITER & Trim(arrConfigData(lEachRow, 1)) & DELIMITER
        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row

        sEnv = Trim(arrConfigData(lEachRow, 5))

        If Not (sEnv = gsEnv Or sEnv = "SHARED") Then GoTo next_row

        lActualRow = lConfigHeaderAtRow + lEachRow
        asRptType = Trim(arrConfigData(lEachRow, 2))
        sRptFileName = Trim(arrConfigData(lEachRow, 3))
        asFileSpecTag = Trim(arrConfigData(lEachRow, 4))

        If asRptType = "SHEET" Then
            If fZero(asFileSpecTag) Then
                fErr "Sheet name cannot be blnak when SHEET specified  " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If
        End If

        bFound = True
next_row:
    Next

    Erase arrConfigData
    Erase arrColsName

    If Not bFound Then
        fErr "[" & asReportID & "] was not configured under : " & asTag & vbCr & shtSysConf.Name
    End If

    fReadConfigOutputFiles = sRptFileName
End Function

Function fGetOutputReportColsConfig(asFileSpecTag As String)
    asFileSpecTag = Trim(asFileSpecTag)
    If Not (Left(asFileSpecTag, 1) = "[" And Right(asFileSpecTag, 1) = "]") Then
        fErr "File Spec Tag is incorrect, which should be like [Output Format - xxx]"
    End If

    Call fReadOutputFileSpecConfig(asTag:=asFileSpecTag _
                                    , dictColsIndex:=dictRptColIndex _
                                    , dictColsName:=dictRptDisplayName _
                                    , dictRawType:=dictRptRawType _
                                    , dictCellFormat:=dictRptCellFormat _
                                    , dictDataFormat:=dictRptDataFormat _
                                    , dictColWidth:=dictRptColWidth _
                                    , dictColAttr:=dictRptColAttr _
                                    )
End Function

Function fReadOutputFileSpecConfig(asTag As String _
                                    , ByRef dictColsIndex As Dictionary _
                                    , ByRef dictColsName As Dictionary _
                                    , ByRef dictRawType As Dictionary _
                                    , ByRef dictCellFormat As Dictionary _
                                    , ByRef dictDataFormat As Dictionary _
                                    , ByRef dictColWidth As Dictionary _
                                    , ByRef dictColAttr As Dictionary _
                                    )

    'Dim asTag As String
    Dim arrColsName()
    Dim arrColsIndex()
    Dim rngToFindIn As Range
    Dim arrConfigData()
    Dim lConfigStartRow As Long
    Dim lConfigStartCol As Long
    Dim lConfigEndRow As Long
    Dim lConfigHeaderAtRow As Long

    ReDim arrColsName(1 To 9)
    arrColsName(1) = "Column Tech Name"
    arrColsName(2) = "Column Display Name"
    arrColsName(3) = "Column Index"
    arrColsName(4) = "Array Index"
    arrColsName(5) = "Raw Data Type"
    arrColsName(6) = "Cell Format"
    arrColsName(7) = "Data Format"
    arrColsName(8) = "Column Width"
    arrColsName(9) = "Column Attr"

    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtFileSpec _
                                , arrConfigData:=arrConfigData _
                                , arrColsName:=arrColsName _
                                , arrColsIndex:=arrColsIndex _
                                , lConfigStartRow:=lConfigStartRow _
                                , lConfigStartCol:=lConfigStartCol _
                                , lConfigEndRow:=lConfigEndRow _
                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
                                , abNoDataConfigThenError:=True)
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(1), False, shtFileSpec, lConfigHeaderAtRow, lConfigStartCol, arrColsName(1))
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(2), False, shtFileSpec, lConfigHeaderAtRow, lConfigStartCol, arrColsName(2))
    Call fValidateDuplicateInArray(arrConfigData, arrColsIndex(3), True, shtFileSpec, lConfigHeaderAtRow, lConfigStartCol, arrColsName(3))
'    Call fValidateBlankInArray(arrConfigData, 1, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Output Type")

    Dim lEachRow As Long
    Dim lActualRow As Long
    Dim sColTechName As String
    Dim sRptFileName As String
    Dim sColWidth As String
    Dim sColLetterIndex As String
    Dim lColLetter2Num As Long
    Dim sColAttr As String
    Dim lColWidth As Double

    Dim sPos As String
    sPos = vbCr & vbCr & "Sheet: " & shtFileSpec.Name & vbCr & "Row : $ACTUAL_ROW$" & vbCr & "Column: "

    Set dictColsIndex = New Dictionary
    Set dictColsName = New Dictionary
    Set dictRawType = New Dictionary
    Set dictCellFormat = New Dictionary
    Set dictDataFormat = New Dictionary
    Set dictColWidth = New Dictionary
    Set dictColAttr = New Dictionary

    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row

        lActualRow = lConfigHeaderAtRow + lEachRow

        sColTechName = Trim(arrConfigData(lEachRow, arrColsIndex(1)))
        sColLetterIndex = Trim(arrConfigData(lEachRow, arrColsIndex(3)))
        sColAttr = Trim(arrConfigData(lEachRow, arrColsIndex(9)))

        dictColsName.Add sColTechName, Trim(arrConfigData(lEachRow, arrColsIndex(2)))
        dictRawType.Add sColTechName, UCase(Trim(arrConfigData(lEachRow, arrColsIndex(5))))
        dictCellFormat.Add sColTechName, shtFileSpec.Cells(lActualRow, lConfigStartCol + arrColsIndex(6) - 1).Address
        dictDataFormat.Add sColTechName, Trim(arrConfigData(lEachRow, arrColsIndex(7)))

        If sColAttr = "NOT_SHOW_UP" Then
            If Len(sColLetterIndex) > 0 Then
                fErr "Col Letter Index should be blank for NOT_SHOW_UP: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If

            dictColsIndex.Add sColTechName, -1
        Else
            If Len(sColLetterIndex) <= 0 Then
                fErr "Col Letter Index should NOT be blank: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If

            lColLetter2Num = fLetter2Num(sColLetterIndex)

            If lColLetter2Num <= 0 Or lColLetter2Num > Columns.Count Then
                fErr "Col Letter Index is invalid,should be A - XFD: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(3)
            End If

            dictColsIndex.Add sColTechName, lColLetter2Num
        End If

        sColWidth = Trim(arrConfigData(lEachRow, arrColsIndex(8)))
        If Len(sColWidth) > 0 Then
            If Not IsNumeric(sColWidth) Then
                fErr "Col Width should be numeric: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(8)

                If CDbl(sColWidth) <= 0 Or CDbl(sColWidth) > 255 Then
                    fErr "Col Width should be > 0 and <= 255: " & Replace(sPos, "$ACTUAL_ROW$", lActualRow) & arrColsName(8)
                End If
            End If

            lColWidth = CDbl(sColWidth)
        Else
            lColWidth = 0
        End If

        dictColWidth.Add sColTechName, lColWidth
        dictColAttr.Add sColTechName, sColAttr
next_row:
    Next

    Dim lNextCol As Long
    lNextCol = WorksheetFunction.Max(dictColsIndex.Items)

    For lEachRow = 0 To dictColAttr.Count - 1
        If dictColAttr.Items(lEachRow) = "NOT_SHOW_UP" Then
            lNextCol = lNextCol + 1
            dictColsIndex(dictColAttr.Keys(lEachRow)) = lNextCol
        End If
    Next

    Erase arrConfigData
    Erase arrColsName
    Erase arrColsIndex
End Function

Function fCrossValidateInputFileTxtSheetFile()
    Dim i As Long
    Dim sFileTag As String
    Dim sSource As String
    Dim sDependantFileTag As String
    Dim sDependantSheet As String
    Dim lActualRow As Long

    For i = 0 To gDictInputFiles.Count - 1
        sFileTag = gDictInputFiles.Keys(i)
        sSource = fGetInputFileSourceType(sFileTag)

        If sSource = "FILE_BINDED_IN_MACRO" Or sSource = "READ_FROM_DRIVE" Then
            fIfExcelFileOpenedToCloseIt fGetInputFileFileName(sFileTag)
        ElseIf sSource = "PARSE_AS_TEXT" Then
'            If Not gDictTxtFileSpec Is Nothing Then
'                If Not gDictTxtFileSpec.Exists(sFileTag) Then
'                    lActualRow = fGetInputFileRowNo(sFileTag)
'                    fErr "The file Tag below is configured as " & sSource & ", but [TXT File Importing Specification] does not have one, pls check." _
'                        & vbCr & "File Tag: " & sFileTag _
'                        & vbCr & "Row: " & lActualRow
'                End If
'            End If
        End If
    Next

    If Not gDictTxtFileSpec Is Nothing Then
        For i = 0 To gDictTxtFileSpec.Count - 1
            sFileTag = gDictTxtFileSpec.Keys(i)

            If Not gDictInputFiles.Exists(sFileTag) Then
                fErr sFileTag & " is coonfigured under [TXT File Importing Specification], but does not exist in [Input Files], pls check." _
                    & vbCr & "File Tag: " & sFileTag
            End If
        Next
    End If
End Function
Function fGetInputFileFileName(asFileTag As String) As String
    fGetInputFileFileName = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.FilePath - InputFile.ReportID - 2)
End Function
Function fGetInputFileSourceType(asFileTag As String) As String
    fGetInputFileSourceType = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.source - InputFile.ReportID - 2)
End Function
Function fGetInputFileFileSpecTag(asFileTag As String) As String
    If Not gDictInputFiles.Exists(asFileTag) Then fErr asFileTag & " does not exist in " & Join(gDictInputFiles.Keys, vbCr)
    fGetInputFileFileSpecTag = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.FileSpecTag - InputFile.ReportID - 2)
End Function
Function fGetInputFileRowNo(asFileTag As String) As String
    fGetInputFileRowNo = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.RowNo - InputFile.ReportID - 2)
End Function
Function fGetInputFileReloadOrNot(asFileTag As String) As String
    fGetInputFileReloadOrNot = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.ReLoadOrNot - InputFile.ReportID - 2)
End Function
Function fGetInputFileSheetToImport(asFileTag As String) As String
    fGetInputFileSheetToImport = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.DefaultSheet - InputFile.ReportID - 2)
End Function
Function fGetInputFilePivotTableTag(asFileTag As String) As String
    fGetInputFilePivotTableTag = Split(gDictInputFiles(asFileTag), DELIMITER)(InputFile.PivotTableTag - InputFile.ReportID - 2)
End Function

Function fGetReneratedReport(Optional asReportID As String = "") As String
    If fZero(asReportID) Then asReportID = gsRptID

    fGetReneratedReport = fGetSpecifiedConfigCellValue(shtSysConf, "[Generate Report File / Sheet]" _
                            , "Genearted File Or Sheet", "Report ID=" & asReportID)
End Function

Function fGetOutputReportItem(Optional asReportID As String = "", Optional sType As String) As String
    If fZero(asReportID) Then asReportID = gsRptID

    If sType = "FILE_SPEC" Then
        fGetOutputReportItem = fGetSpecifiedConfigCellValue(shtSysConf, "[Output Files]" _
                            , "File Spec Tag", "Report ID=" & asReportID)
    Else
        fErr "wrong param fGetOutputReportItem"
    End If
End Function

Function fSetReneratedReport(Optional asReportID As String = "", Optional sRptFile As String = "") As String
    If fZero(asReportID) Then asReportID = gsRptID
    If fZero(sRptFile) Then sRptFile = gsRptFilePath

    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Generate Report File / Sheet]" _
                            , "Genearted File Or Sheet", "Report ID=" & asReportID, sRptFile)
End Function

Public Const BUSINESS_ERROR_NUMBER = 10000
'Public Const CONFIG_ERROR_NUMBER = 20000
Public Const DELIMITER = "|"

Public gErrNum As Long
Public gErrMsg As String

'=======================================
Public dictMiscConfig As Dictionary
'=======================================
Public arrMaster()
'=======================================
Public arrOutput()
'=======================================
Public gFSO As FileSystemObject
Public gRegExp As VBScript_RegExp_55.RegExp
Public Const PW_PROTECT_SHEET = "abcd1234"

Function fSetBackToConfigSheetAndUpdategDict_UserTicket()
    Dim ckb As Object

    'Dim eachObj As Object

    'for each eachobj in shtmenu.
    Dim i As Long
    Dim sCompanyID As String
    Dim sTickValue As String

    For i = 0 To dictCompList.Count - 1
        sCompanyID = dictCompList.Keys(i)

        If Not fActiveXControlExistsInSheet(shtCurrMenu, fGetCompany_CheckBoxName(sCompanyID), ckb) Then GoTo next_company

        sTickValue = IIf(ckb.Value, "Y", "N")

        Call fSetSpecifiedConfigCellValue(shtStaticData, "[Sales Company List]", "User Ticked", "Company ID=" & sCompanyID, sTickValue)
        Call fUpdateDictionaryItemValueForDelimitedElement(dictCompList, sCompanyID, Company.Selected - Company.REPORT_ID, sTickValue)
next_company:
    Next
End Function

Function fSetBackToConfigSheetAndUpdategDict_InputFiles()
    Dim i As Integer
    Dim sEachCompanyID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String

    For i = 0 To dictCompList.Count - 1
        sEachCompanyID = dictCompList.Keys(i)
        'sFilePathRange = "rngSalesFilePath_" & sEachCompanyID

        If fGetCompany_UserTicked(sEachCompanyID) = "Y" Then
            sFilePathRange = fGetCompany_InputFileTextBoxName(sEachCompanyID)
            sEachFilePath = Trim(shtCurrMenu.Range(sFilePathRange).Value)
        Else
            sEachFilePath = "User not selected."
        End If

        Call fSetValueBackToSysConf_InputFile_FileName(sEachCompanyID, sEachFilePath)
        Call fUpdateGDictInputFile_FileName(sEachCompanyID, sEachFilePath)

        'Call fSetSalesInfoFileToMainConfig(sEachCompanyID, sEachFilePath)
    Next

'    sFile = Trim(shtMenu.Range("rngSalesFilePath_GY").Value)
'
'    Call fSetValueBackToSysConf_InputFile_FileName("GY", sFile)
'    Call fUpdateGDictInputFile_FileName("GY", sFile)

End Function

Function fSetIntialValueForShtMenuInitialize()

End Function

Function fSetConditionFormatForFundamentalSheets()
'    Call fClearConditionFormatAndAdd(shtCompanyNameReplace, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtHospital, Array(1), True)
'    Call fClearConditionFormatAndAdd(shtHospitalReplace, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtProductMaster, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtProductNameReplace, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtProductProducerReplace, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtProductSeriesReplace, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtProductUnitRatio, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtProductProducerMaster, Array(1), True)
'    Call fClearConditionFormatAndAdd(shtProductNameMaster, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtSalesManMaster, Array(1), True)
'    Call fClearConditionFormatAndAdd(shtFirstLevelCommission, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSecondLevelCommission, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSelfSalesOrder, Array(1, 2, 3), True)         'to-do
'    Call fClearConditionFormatAndAdd(shtSelfSalesPreDeduct, Array(1, 2, 3, 4), True)       'to-do
'    Call fClearConditionFormatAndAdd(shtSalesManCommConfig, Array(1, 2, 3, 4, 5, 6), True)
'    Call fClearConditionFormatAndAdd(shtException, Array(1), True)
'
'    Call fClearConditionFormatAndAdd(shtSelfPurchaseOrder, Array(1, 2, 3, 4, 5), True)       'to-do
'    Call fClearConditionFormatAndAdd(shtSelfInventory, Array(1, 2, 3, 5), True)       'to-do
'
'    Call fClearConditionFormatAndAdd(shtNewRuleProducts, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtPromotionProduct, Array(2, 3, 4, 5), True)
'    Call fClearConditionFormatAndAdd(shtCZLInventory, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtSelfInventory, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLPurchaseOrder, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLInventory, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLInvDiff, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLRolloverInv, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompInvCalcd, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompInvUnified, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompRolloverInv, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompInvDiff, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtProductTaxRate, Array(ProdTaxRate.ProductProducer, ProdTaxRate.ProductName, ProdTaxRate.ProductSeries, ProdTaxRate.TaxRate), True)
'
'    Call fClearConditionFormatAndAdd(shtRefund, Array(1, 2, 3, 4, 5), True)
End Function

Function fClearConditionFormatAndAdd(sht As Worksheet, arrKeysCols, Optional bExtendToMore10ThousRows As Boolean = True)
    Call fDeleteAllConditionFormatFromSheet(sht)
    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
    Call fSetConditionFormatForBorders(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
End Function

Private Sub Workbook_Open()

    Call sub_WorkBookInitialization

End Sub

Sub sub_WorkBookInitialization()
    'shtBillIn
    '=========================================================
    Call fDeleteAllConditionFormatFromSheet(shtBillIn)
    Call fSetConditionFormatForOddEvenLine(shtBillIn, , , , Array(BillIn.FromCompany), True)
    Call fSetConditionFormatForBorders(shtBillIn, , , , Array(BillIn.FromCompany), True)

    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBillIn, 2, BillIn.Amount, Rows.Count, BillIn.Amount), 0, 999999999)

    'shtBillOut
    '=========================================================
    Call fDeleteAllConditionFormatFromSheet(shtBillOut)
    Call fSetConditionFormatForOddEvenLine(shtBillOut, , , , Array(BillOut.toCompany), True)
    Call fSetConditionFormatForBorders(shtBillOut, , , , Array(BillOut.toCompany), True)

    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBillOut, 2, BillOut.Amount, Rows.Count, BillOut.Amount), 0, 999999999)

    'shtBusinessDetails
    '=========================================================
    Call fDeleteAllConditionFormatFromSheet(shtBusinessDetails)
   ' Call fSetConditionFormatForOddEvenLine(shtBusinessDetails, , shtBusinessDetails.DataStartRow, , Array(BuzDetail.VendorName, BuzDetail.PLATFORM), True)
    Call fSetConditionFormatForBorders(shtBusinessDetails, , shtBusinessDetails.DataStartRow, , Array(BuzDetail.VendorName, BuzDetail.PLATFORM), True)

    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_Qty, Rows.Count, BuzDetail.Point_Qty), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_Price, Rows.Count, BuzDetail.Point_Price), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_CurrDayPrice, Rows.Count, BuzDetail.Point_CurrDayPrice), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_DaysNum, Rows.Count, BuzDetail.Point_DaysNum), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_Amt, Rows.Count, BuzDetail.Point_Amt), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.DownLoad_Qty, Rows.Count, BuzDetail.DownLoad_Qty), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.DownLoad_Price, Rows.Count, BuzDetail.DownLoad_Price), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.DownLoad_Amt, Rows.Count, BuzDetail.DownLoad_Amt), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Credit_Qty, Rows.Count, BuzDetail.Credit_Qty), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Credit_Price, Rows.Count, BuzDetail.Credit_Price), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Credit_Amt, Rows.Count, BuzDetail.Credit_Amt), 0, 999999999)
End Sub

Function fFileSpecTemplateHasAdditionalHeader(rngConfigBlock As Range, arrHeadersToFind) As Boolean
    Dim arrColsName(1 To 6)
    arrColsName(1) = "Column Tech Name"
    arrColsName(2) = "Column Display Name"
    arrColsName(3) = "Column Index"
    arrColsName(4) = "Array Index"
    arrColsName(5) = "Raw Data Type"
    arrColsName(6) = "Data Format"

    Dim rngHeader As Range
    Dim lHeaderAtLine As Long
    Dim rngFound As Range
    Dim i As Integer

    lHeaderAtLine = fFindHeaderAtLineInFileSpec(rngConfigBlock, arrColsName)
    Set rngHeader = fGetRangeByStartEndPos(shtFileSpec, lHeaderAtLine, rngConfigBlock.Column, lHeaderAtLine, Columns.Count)

    If IsArray(arrHeadersToFind) Then
        For i = LBound(arrHeadersToFind) To UBound(arrHeadersToFind)
            Set rngFound = fFindInWorksheet(rngHeader, CStr(arrHeadersToFind(i)), False)
            If rngFound Is Nothing Then
                Exit For
            End If
        Next
    Else
        Set rngFound = fFindInWorksheet(rngHeader, CStr(arrHeadersToFind), False)
    End If

    fFileSpecTemplateHasAdditionalHeader = (Not rngFound Is Nothing)
    Set rngHeader = Nothing
    Set rngFound = Nothing
End Function

Sub subMainValidateSourceCodeFile()
    On Error GoTo error_handling

    Call fInitialization

    Call fBackUpTextFileWithDefaultFileName(SOURCE_CODE_LIBRARY_FILE)

    Call fAppendBlankLineToTheEndOfTextFile(SOURCE_CODE_LIBRARY_FILE)
    Call fTrimTrailingBlanksForTextFile(SOURCE_CODE_LIBRARY_FILE)
    Call fDeleteMultipleBlankLinesFromTextFile(SOURCE_CODE_LIBRARY_FILE)

    Dim dictFuns As Dictionary
    Dim dictToDeleteRows As Dictionary
    Dim arrFileLines
    Dim iCnt As Long

    arrFileLines = fReadTextFileAllLinesToArray(SOURCE_CODE_LIBRARY_FILE)

    Set dictFuns = fGetAllFunctionListFromFile(arrFileLines, SOURCE_CODE_LIBRARY_FILE)
    iCnt = fScanDuplicateFunctionsWriteToTwoTextFileForCompare(arrFileLines, dictFuns, COMPARE_TMP_FILE_LEFT, COMPARE_TMP_FILE_RIGHT, dictToDeleteRows)

    Call fDeleteLinesFromTextFileBySortedDictionary(SOURCE_CODE_LIBRARY_FILE, dictToDeleteRows)
    Call fDeleteMultipleBlankLinesFromTextFile(SOURCE_CODE_LIBRARY_FILE)

    Set dictFuns = Nothing
    Set dictToDeleteRows = Nothing
    Erase arrFileLines

    If iCnt > 0 Then
        Shell """" & BEYOND_COMPARE_EXE & """ """ & COMPARE_TMP_FILE_LEFT & """  """ & COMPARE_TMP_FILE_RIGHT & """", vbMaximizedFocus
        MsgBox iCnt & " functions are found with multiple copies, please check the [beyond compare] screen.", vbExclamation
    Else
        MsgBox "all are ok", vbInformation
    End If

error_handling:
    If gErrNum <> 0 Then GoTo reset_excel_options

    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options

reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub

Function fClearGlobalVarialesResetOption()
    Set gFSO = Nothing
    Set gRegExp = Nothing

    Set gProBar = Nothing
    Set dictNavigate = Nothing
    Set dictWbListCurrPos = Nothing

    fEnableExcelOptionsAll
End Function

Function fCreateTextFileInUnicode(sFileFullPath As String)
    fGetFSO
    With gFSO.CreateTextFile(sFileFullPath, True, True)
        .Close
    End With
End Function

Function fDeleteLinesFromTextFileBySortedDictionary(sFileFullPath As String, dictToDeleteRows As Dictionary, Optional sLineBreak As String = vbCrLf) ', Optional bUnicode As Boolean = False,
    'keys: starting from 0
    ' 10-15, 25
    Dim i As Long
    Dim sKey As String
    Dim lStartLine As Long
    Dim lEndLine As Long
    Dim arrFileLines
    Dim sTmpFile As String
    Dim iFileNum As Integer
    Dim lLastLine As Long
    Dim lEachRow As Long

    sTmpFile = fGetFileNetName(sFileFullPath, True) & "." & format(Now(), "yyyymmddHHMMSS")

    arrFileLines = fReadTextFileAllLinesToArray(sFileFullPath, sLineBreak)

    'If bUnicode Then fCreateTextFileInUnicode (sTmpFile)

    iFileNum = FreeFile
    Open sTmpFile For Output As #iFileNum

    lLastLine = -1
    For i = 0 To dictToDeleteRows.Count - 1
        sKey = dictToDeleteRows.Keys(i)

        If InStr(sKey, "-") Then
            lStartLine = Split(sKey, "-")(0)
            lEndLine = Split(sKey, "-")(1)
        Else
            lStartLine = sKey
            lEndLine = sKey
        End If

        If lStartLine > lLastLine Then
            For lEachRow = lLastLine + 1 To lStartLine - 1
                Print #iFileNum, arrFileLines(lEachRow)
            Next

            lLastLine = lEndLine
        End If
    Next

    If lLastLine < UBound(arrFileLines) Then
        For lEachRow = lLastLine + 1 To UBound(arrFileLines)
            Print #iFileNum, arrFileLines(lEachRow)
        Next
    End If

    Erase arrFileLines
    Close #iFileNum

    Kill sFileFullPath
    Name sTmpFile As sFileFullPath
End Function

Function fScanDuplicateFunctionsWriteToTwoTextFileForCompare(arrFileLines, dictFuns As Dictionary _
                    , asLeftFile As String, asRightFile As String, dictToDeleteRows As Dictionary) As Long
    Dim iCnt As Long
    Dim i As Long
    Dim sFunLines As String
    Dim sFunName As String
    Dim iLeftFileNum As Integer
    Dim iRightFileNum As Integer
    Dim arrInstance
    Dim lStartLine As Long
    Dim lEndLine As Long
    Dim lFirstStartLine As Long
    Dim lFirstEndLine As Long
    Dim j As Long
    Dim k As Long

    Set dictToDeleteRows = New Dictionary

    iCnt = 0
    For i = 0 To dictFuns.Count - 1
        sFunLines = dictFuns.Items(i)

        If InStr(sFunLines, "|") > 0 Then
            sFunName = dictFuns.Keys(i)

            If iCnt <= 0 Then
                iLeftFileNum = FreeFile
                Open asLeftFile For Output As #iLeftFileNum

                iRightFileNum = FreeFile
                Open asRightFile For Output As iRightFileNum
            End If

            arrInstance = Split(sFunLines, "|")

            For j = LBound(arrInstance) To UBound(arrInstance)
                lStartLine = Split(arrInstance(j), ",")(0)
                lEndLine = Split(arrInstance(j), ",")(1)

                If ((j + 1) Mod 2) = 1 Then
                    Print #iLeftFileNum, "'============================== line " & lStartLine + 1 & " - " & lEndLine + 1 & " ========================="

                    For k = lStartLine To lEndLine
                        Print #iLeftFileNum, arrFileLines(k)
                    Next
                Else
                    Print #iRightFileNum, "'============================== line " & lStartLine + 1 & " - " & lEndLine + 1 & " ========================="

                    For k = lStartLine To lEndLine
                        Print #iRightFileNum, arrFileLines(k)
                    Next
                End If

                If j = LBound(arrInstance) Then
                    lFirstStartLine = lStartLine
                    lFirstEndLine = lEndLine
                ElseIf j = LBound(arrInstance) + 1 Then
                    If fTwoSectionAreSameInTextFile(arrFileLines, lFirstStartLine, lFirstEndLine, lStartLine, lEndLine) Then
                        dictToDeleteRows.Add CStr(lFirstStartLine) & "-" & CStr(lFirstEndLine), ""
                    End If
                End If
            Next

            If ((j + 1) Mod 2) = 0 Then
                For k = lStartLine To lEndLine
                    Print #iRightFileNum, j & " copies " & sFunName & " !!!"
                Next
            End If

            iCnt = iCnt + 1
            Erase arrInstance
        End If
    Next
    If iCnt > 0 Then
        Close #iLeftFileNum
        Close #iRightFileNum
    End If

    fScanDuplicateFunctionsWriteToTwoTextFileForCompare = iCnt
End Function

Function fTwoSectionAreSameInTextFile(arrFileLines, lFirstStartLine, lFirstEndLine, l2ndStartLine, l2ndEndLine)
    Dim lEachRow As Long
    Dim bOut As Boolean

    If lFirstEndLine - lFirstStartLine <> l2ndEndLine - l2ndStartLine Then fTwoSectionAreSameInTextFile = False: Exit Function

    bOut = True
    For lEachRow = 0 To lFirstEndLine - lFirstStartLine
        If UCase(Replace(arrFileLines(lFirstStartLine + lEachRow), " ", "")) <> UCase(Replace(arrFileLines(l2ndStartLine + lEachRow), " ", "")) Then
            bOut = False
            Exit For
        End If
    Next

    fTwoSectionAreSameInTextFile = bOut
End Function

Function fGetAllFunctionListFromFile(arrFileLines, sSourceCodeFile As String) As Dictionary
    Dim dictFuns As Dictionary
    Dim lEachRow As Long
    Dim sEachLine As String
    Dim sLineContent As String
    Dim sScopeTag 'As String
    Dim sFunOrSub 'As String
    Dim bNewFunStart As Boolean
    Dim lFunStartLine As Long
    Dim lFunEndLine As Long
    Dim lFunNameEndPos As Long
    Dim sFunName As String

    Set dictFuns = New Dictionary

    For lEachRow = LBound(arrFileLines) To UBound(arrFileLines)
        sEachLine = Trim(arrFileLines(lEachRow))

        sLineContent = Replace(sEachLine, vbTab, " ")
        sLineContent = WorksheetFunction.Trim(sLineContent)

        If Left(sLineContent, 1) = "'" Then GoTo next_row

        For Each sScopeTag In Array("PUBLIC ", "PRIVATE ", "FRIEND ")
            If UCase(Left(sLineContent, Len(sScopeTag))) = sScopeTag Then
                sLineContent = Right(sLineContent, Len(sLineContent) - Len(sScopeTag))
                Exit For
            End If
        Next

        For Each sFunOrSub In Array("FUNCTION ", "SUB ")
            If UCase(Left(sLineContent, Len(sFunOrSub))) = sFunOrSub Then
                sLineContent = Right(sLineContent, Len(sLineContent) - Len(sFunOrSub))
                bNewFunStart = True
                Exit For
            End If

            bNewFunStart = False
        Next

        If bNewFunStart Then
            If lFunStartLine > 0 And lFunEndLine <= 0 Then
                fErr "End function/sub for " & sFunName & " has not been found, but a new fun/sub is detected."
            End If
        End If

        If bNewFunStart Then
            lFunNameEndPos = InStr(sLineContent, "(")
            If lFunNameEndPos <= 0 Then
                lFunNameEndPos = InStr(sLineContent, " _")
            End If

            If lFunNameEndPos <= 0 Then fErr "function/sub declaration is incorrect at line " & lEachRow + 1

            sFunName = Trim(Left(sLineContent, lFunNameEndPos - 1))

            lFunStartLine = lEachRow
            lFunEndLine = 0
        Else
            If UCase(sLineContent) Like "END FUNCTION*" Then
                If lFunStartLine <= 0 Then fErr "FUNCTION declaration is not found, but a end function is detected at line " & lEachRow + 1
                lFunEndLine = lEachRow
            ElseIf UCase(sLineContent) Like "END SUB*" Then
                If lFunStartLine <= 0 Then fErr "sub declaration is not found, but a end sub is detected at line " & lEachRow + 1
                lFunEndLine = lEachRow
            Else
            End If

            If lFunEndLine > 0 Then
                If Not dictFuns.Exists(sFunName) Then
                    dictFuns.Add sFunName, lFunStartLine & "," & lFunEndLine
                Else
                    dictFuns(sFunName) = dictFuns(sFunName) & "|" & lFunStartLine & "," & lFunEndLine
                End If

                lFunStartLine = 0
                lFunEndLine = 0
            End If
        End If
next_row:
    Next

    Set fGetAllFunctionListFromFile = dictFuns
    Set dictFuns = Nothing
End Function

Function fDeleteMultipleBlankLinesFromTextFile(sFileFullPath As String, Optional sLineBreak As String = vbCrLf)
    Dim arrFileLines
    Dim sTmpFile As String
    Dim lEachRow As Long
    Dim sEachLine As String
    Dim bNewBlankStart As Boolean
    Dim lLastValidRow As Long

    sTmpFile = fGetFileNetName(sFileFullPath, True) & "." & format(Now(), "yyyymmddHHMMSS")

    arrFileLines = fReadTextFileAllLinesToArray(sFileFullPath, sLineBreak)

    For lEachRow = UBound(arrFileLines) To LBound(arrFileLines) Step -1
        sEachLine = arrFileLines(lEachRow)

        If Len(Trim(sEachLine)) > 0 Then
            lLastValidRow = lEachRow
            Exit For
        End If
    Next

    Dim iFileNum As Long

    iFileNum = FreeFile
    Open sTmpFile For Output As #iFileNum

    bNewBlankStart = True
    For lEachRow = LBound(arrFileLines) To lLastValidRow    'UBound(arrFileLines)
        sEachLine = arrFileLines(lEachRow)

        If Len(Trim(sEachLine)) <= 0 Then
            If bNewBlankStart Then
                Print #iFileNum, ""
                bNewBlankStart = False
            Else
            End If
        Else
            Print #iFileNum, sEachLine
            bNewBlankStart = True
        End If
    Next

    Erase arrFileLines
    Close #iFileNum

    Kill sFileFullPath
    Name sTmpFile As sFileFullPath
End Function

Function fTrimTrailingBlanksForTextFile(sFileFullPath As String, Optional sLineBreak As String = vbCrLf)
    Dim arrFileLines
    Dim sTmpFile As String
    Dim lEachRow As Long

    sTmpFile = fGetFileNetName(sFileFullPath, True) & "." & format(Now(), "yyyymmddHHMMSS")

    arrFileLines = fReadTextFileAllLinesToArray(sFileFullPath, sLineBreak)

    Dim iFileNum As Long

    iFileNum = FreeFile
    Open sTmpFile For Output As #iFileNum

    For lEachRow = LBound(arrFileLines) To UBound(arrFileLines)
        Print #iFileNum, RTrim(arrFileLines(lEachRow))
    Next

    Erase arrFileLines

    Close #iFileNum

    Kill sFileFullPath
    Name sTmpFile As sFileFullPath
End Function

Function fReadTextFileAllLinesToArray(sFileFullPath As String, Optional sLineBreak As String = vbCrLf)
    Dim sContent
    Dim iFileNum As Long

    On Error GoTo exit_fun

    iFileNum = FreeFile
    Open sFileFullPath For Input As #iFileNum

    'sContent = Input(LOF(iFileNum), #iFileNum)
    sContent = StrConv(InputB(LOF(iFileNum), #iFileNum), vbUnicode)

    Close #iFileNum

    Dim arrFileLines
    Dim sLastLine As String

    fReadTextFileAllLinesToArray = Split(sContent, sLineBreak)
    sContent = ""
exit_fun:
    Close #iFileNum
    If Err.Number <> 0 Then fErr Err.Description
End Function

Function fAppendBlankLineToTheEndOfTextFile(sFileFullPath As String, Optional sLineBreak As String = vbCrLf)
    Dim arrFileLines
    Dim sLastLine As String

    arrFileLines = fReadTextFileAllLinesToArray(sFileFullPath, sLineBreak)

    sLastLine = Trim(arrFileLines(UBound(arrFileLines)))
    Erase arrFileLines

    If Len(sLastLine) > 0 Then
        Call fAppendToTextFile(sFileFullPath, "")
    End If
End Function

Function fAppendToTextFile(sFileFullPath As String, sWhatToAppend As String, Optional AppendNewLineFeed As Boolean = True) ', Optional bUnicode As Boolean = False)
    Dim iFileNum As Long

    'If bUnicode Then sWhatToAppend = StrConv(sWhatToAppend, vbUnicode)

    iFileNum = FreeFile

    Open sFileFullPath For Append As #iFileNum

    If AppendNewLineFeed Then
        Print #iFileNum, sWhatToAppend
    Else
        Print #iFileNum, sWhatToAppend;
    End If

    Close #iFileNum
End Function

Function fBackUpTextFileWithDefaultFileName(sFileFullPath As String)
    Dim sTmpFile As String
    Dim lEachRow As Long

    sTmpFile = fGetFileNetName(sFileFullPath, True) & "." & Timer() * 100 & ".bak"

    fGetFSO
    Call gFSO.CopyFile(sFileFullPath, sTmpFile, True)

End Function

Function fInitialization()
    Err.Clear

    gErrNum = 0
    gErrMsg = ""

    Call fDisableExcelOptionsAll

    Application.ScreenUpdating = True   ' for testing

    Call fRemoveFilterForAllSheets
End Function

Function fSetFocus(controlOnForm)
    controlOnForm.SelStart = 0
    controlOnForm.SelLength = Len(controlOnForm.Value)
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
        Kill sFolder & "*"
    End If
End Function

Function fOpenWorkbook(ByVal asFileFullPath As String _
                , Optional ByRef bAlreadyOpened As Boolean, Optional ByVal bWhenOpenedToCloseFirst As Boolean = False _
                , Optional asSheetName As String = "", Optional ByRef shtOut As Worksheet _
                , Optional bReadOnly As Boolean = True) As Workbook
    Dim wbOut As Workbook
    Dim sTmp As String

    bAlreadyOpened = False
    If fExcelFileIsOpen(asFileFullPath, wbOut) Then
        If UCase(wbOut.FullName) <> UCase(Trim(asFileFullPath)) Then
            sTmp = wbOut.FullName
            Set wbOut = Nothing
            fErr "Another excel file with the same file name has already been open, please close it first" & vbCr _
                 & "File Name: " & fGetFileBaseName(asFileFullPath) & vbCr & vbCr _
                 & "File fullpath: " & sTmp
        Else
            bAlreadyOpened = True
            If bWhenOpenedToCloseFirst Then
                fErr "The file is already opened, please close it first. Or you can change the parameter(bWhenOpenedToCloseFirst) to use the open file directly."
            End If
        End If
    End If

    If wbOut Is Nothing Then
        'Application.AutomationSecurity = msoAutomationSecurityForceDisable
        Set wbOut = Workbooks.Open(Filename:=asFileFullPath, ReadOnly:=bReadOnly)
        wbOut.Saved = True
       ' Application.AutomationSecurity = msoAutomationSecurityByUI
    End If

    'Call fExcelFileOpenedToCloseIt(asFileFullPath)

    If Len(Trim(asSheetName)) > 0 Then
        If Not fSheetExists(asSheetName, , wbOut) Then
            fErr "The workbook does not have a sheet named as [" & asSheetName & "], please chedk."
        End If
    End If
    Set fOpenWorkbook = wbOut
    Set wbOut = Nothing
End Function
Function fFolderExists(sFolder As String) As Boolean
    fGetFSO
    fFolderExists = gFSO.FolderExists(sFolder)
End Function

Function fSetRowHeightForExceedingThreshold(sht As Worksheet, lStartRow As Long, lEndRow As Long _
                                        , Optional dblRowHeightThreshold As Double = 16)
    Dim lEachRow As Long
    Dim rgAll As Range
    Dim rgTarget As Range

    Set rgAll = sht.Rows(lStartRow & ":" & lEndRow)

    For lEachRow = 1 To lEndRow - lStartRow + 1

        If rgAll.Rows(lEachRow).RowHeight < dblRowHeightThreshold Then
            If rgTarget Is Nothing Then
                Set rgTarget = rgAll.Rows(lEachRow)
            Else
                Set rgTarget = Union(rgTarget, rgAll.Rows(lEachRow))
            End If
        End If
    Next

    If Not rgTarget Is Nothing Then
        rgTarget.RowHeight = dblRowHeightThreshold
    End If

    Set rgAll = Nothing
    Set rgTarget = Nothing
End Function

Function fExcelFileIsOpen(sExcelFileName As String, Optional wbOut As Workbook) As Boolean
    On Error Resume Next
    Set wbOut = Workbooks(fGetFileBaseName(sExcelFileName))
    Err.Clear

    fExcelFileIsOpen = (Not wbOut Is Nothing)
End Function
Function fExactExcelFileIsopen(sExcelFileName As String, Optional wbOut As Workbook) As Boolean
    Dim bOut As Boolean

    bOut = False

    On Error Resume Next
    Set wbOut = Workbooks(fGetFileBaseName(sExcelFileName))
    Err.Clear

    If wbOut Is Nothing Then GoTo exit_fun

    If UCase(wbOut.FullName) <> UCase(Trim(sExcelFileName)) Then
        Set wbOut = Nothing
        GoTo exit_fun
    Else
        bOut = True: GoTo exit_fun
    End If

exit_fun:
    fExactExcelFileIsopen = bOut
End Function
Function fExcelFileOpenedToCloseIt(sExcelFileFullPath As String, Optional wbOut As Workbook _
                            , Optional bRaiseErrIfOpened As Boolean = True _
                            , Optional bActiveItIfAlreadyOpened As Boolean = True) As Boolean
    Dim bIsOpenAlready As Boolean
    Dim sFileBaseName As String

    bIsOpenAlready = fExcelFileIsOpen(sExcelFileFullPath, wbOut)

    If bIsOpenAlready Then
        If bActiveItIfAlreadyOpened Then wbOut.Activate
        'fGetFSO
        'sExcelFileFullPath = gFSO.GetFile(sExcelFileFullPath).Path
        sExcelFileFullPath = fCheckPath(sExcelFileFullPath)
        sFileBaseName = fGetFileBaseName(sExcelFileFullPath)

        If UCase(wbOut.FullName) = UCase(sExcelFileFullPath) Then
            If bRaiseErrIfOpened Then
                fErr "Excel File is open, pleae close it first." & vbCr & sFileBaseName
            Else
                MsgBox "Excel File is open, pleae close it first." & vbCr & sFileBaseName, vbExclamation
            End If
        Else
            Set wbOut = Nothing

            If bRaiseErrIfOpened Then
                fErr "Another file with the same name """ & sFileBaseName & """ is open, please close it first."
            Else
                MsgBox "Another file with the same name """ & sFileBaseName & """ is open, please close it first.", vbExclamation
            End If
        End If
    End If

    fExcelFileOpenedToCloseIt = bIsOpenAlready
End Function

Function fErr(Optional sMsg As String = "") As VbMsgBoxResult
    gErrNum = vbObjectError + CONFIG_ERROR_NUMBER
    'gbBusinessError = True
    gErrMsg = sMsg
    'If fNzero(sMsg) Then fMsgBox "Error: " & vbCr & vbCr & sMsg, vbCritical
    If fNzero(sMsg) Then fMsgBox sMsg, vbCritical

    Err.Raise gErrNum, "", "Program is to be terminated."
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
