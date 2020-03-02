Attribute VB_Name = "MB_0_RibbonButton"
Option Explicit
Option Base 1
Private Const POTENTIOAL_UNSOLVED_EXTERNAL_LINES = "Potential_Links"

Sub subMain_OpenAcitveWorkbookLocation()
    If Workbooks.Count <= 0 Then Exit Sub
    
    If Len(ActiveWorkbook.Path) <= 0 Then Exit Sub
    
    Call fOpenAcitveWorkbookLocation
End Sub
Sub subMain_DisplayWorkbookFullPath()
    If Workbooks.Count <= 0 Then Exit Sub
    
    'MsgBox ActiveWorkbook.FullName
    
    Dim sFullPath As String
    sFullPath = ActiveWorkbook.FullName
    
    Dim myData As DataObject

    Set myData = New DataObject
    
    myData.SetText sFullPath
    myData.PutInClipboard
        
    Set myData = Nothing
End Sub
Sub subMain_BackupActiveWorkbook()
    If Workbooks.Count <= 0 Then Exit Sub
    
    'MsgBox ActiveWorkbook.FullName
    
    If Len(ActiveWorkbook.Path) <= 0 Then Exit Sub
    Call fBackupActiveWorkbook(ActiveWorkbook)
End Sub
Sub btnExportSourceCode_onAction()
    If Workbooks.Count <= 0 Then Exit Sub
    
    Call sub_ExportSourceCodeToFolder
End Sub

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
'Function fActiveVisibleSwitchSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1")
'    Dim shtCurr As Worksheet
'    Set shtCurr = ActiveSheet
'
'    On Error Resume Next
'
'    If shtToSwitch.Visible = xlSheetVisible Then
'        If ActiveSheet Is shtToSwitch Then
'            shtToSwitch.Visible = xlSheetVisible
'            shtToSwitch.Activate
'            Range(sRngAddrToSelect).Select
'        Else
'            shtToSwitch.Visible = xlSheetVeryHidden
'        End If
'    Else
'        shtToSwitch.Visible = xlSheetVisible
'        shtToSwitch.Activate
'        Range(sRngAddrToSelect).Select
'    End If
'
'    If bHidePreviousActiveSheet Then
'        If Not shtCurr Is shtToSwitch Then shtCurr.Visible = xlSheetVeryHidden
'    End If
'
'    err.Clear
'End Function
Function fHideAllSheetExcept(ParamArray arr())
    Dim sht 'As Worksheet
    Dim shtConvt 'As Worksheet
    Dim wbSht 'As Worksheet
    
    On Error Resume Next
    
    For Each wbSht In ThisWorkbook.Worksheets
        For Each sht In arr
            Set shtConvt = sht
            If wbSht Is shtConvt Then
                'sht.Visible = xlSheetVisible
                GoTo next_wbsheet
            End If
        Next
        
        wbSht.Visible = xlSheetVeryHidden
next_wbsheet:
    Next
    
    Set shtConvt = Nothing
    Err.Clear
End Function

Sub subMain_ValidateAllSheetsData()
    On Error GoTo exit_sub
    
    fGetProgressBar
    gProBar.ShowBar
    gProBar.ChangeProcessBarValue 0.1
    If Not shtCompanyNameReplace.fValidateSheet(False) Then GoTo exit_sub
    If Not shtHospital.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductMaster.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.2
    If Not shtProductNameMaster.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductProducerMaster.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.3
    If Not shtSalesManMaster.fValidateSheet(False) Then GoTo exit_sub
    If Not shtSalesManCommConfig.fValidateSheet(False) Then GoTo exit_sub
    
    gProBar.ChangeProcessBarValue 0.4
    If Not shtNewRuleProducts.fValidateSheet(False) Then GoTo exit_sub
    
    If Not shtHospitalReplace.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.5
    If Not shtProductProducerReplace.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductNameReplace.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.7
    If Not shtProductSeriesReplace.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductUnitRatio.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 0.8
    If Not shtFirstLevelCommission.fValidateSheet(False) Then GoTo exit_sub
    If Not shtSecondLevelCommission.fValidateSheet(False) Then GoTo exit_sub
    gProBar.ChangeProcessBarValue 1
    If Not shtSelfPurchaseOrder.fValidateSheet(False) Then GoTo exit_sub
    If Not shtSelfSalesOrder.fValidateSheet(False) Then GoTo exit_sub
    If Not shtPromotionProduct.fValidateSheet(False) Then GoTo exit_sub
    If Not shtProductTaxRate.fValidateSheet(False) Then GoTo exit_sub
    
    gProBar.DestroyBar
    fMsgBox "没有发现错误！", vbInformation
exit_sub:
    'If Err.Number <> 0 Then fMsgBox Err.Number
    gProBar.DestroyBar
End Sub

'Sub subMain_BackToLastPosition()
'    Dim sLastSheetName As String
'    Dim lLastMaxRow As Long
'    Dim lPrevMaxRow As Long
'    Dim bFound As Boolean
'
'    Const LAST_COL = 2
'    Const PREV_COL = 3
'
'    bFound = False
'    On Error GoTo exit_sub
'
'    Dim shtLast As Worksheet
'    Dim lEachRow As Long
'
'    lLastMaxRow = shtDataStage.Cells(Rows.Count, LAST_COL).End(xlUp).Row
'
'    For lEachRow = lLastMaxRow To 1 Step -1
'        sLastSheetName = Trim(shtDataStage.Cells(lEachRow, LAST_COL).Value)
'        shtDataStage.Cells(lEachRow, LAST_COL).ClearContents
'
'        If fZero(sLastSheetName) Then GoTo previous_row
'
'        If fSheetExists(sLastSheetName) Then
'            Set shtLast = ThisWorkbook.Worksheets(sLastSheetName)
'
'            If UCase(shtLast.Name) = UCase(ActiveSheet.Name) Then
'                Call fAppendDataToLastCellOfColumn(shtDataStage, PREV_COL, sLastSheetName)
'            Else
'                If fSheetIsVisible(shtLast) Then
'                    'Application.EnableEvents = False
'                    shtLast.Activate
'                    'Application.EnableEvents = True
'                    bFound = True
'                    Exit For
'                End If
'            End If
'        End If
'
'previous_row:
'    Next
'
'    If bFound Then
'        Call fAppendDataToLastCellOfColumn(shtDataStage, PREV_COL, sLastSheetName)
'    End If
'
'exit_sub:
'    Set shtLast = Nothing
'    'Application.EnableEvents = True
'End Sub

'Sub subMain_BackToPreviousPosition()
'    Dim sPrevSheetName As String
'    Dim lPrevMaxRow As Long
'    Dim lLastMaxRow As Long
'    Dim bFound As Boolean
'
'    Const LAST_COL = 2
'    Const PREV_COL = 3
'
'    bFound = False
'    On Error GoTo exit_sub
'
'    Dim shtPrev As Worksheet
'    Dim lEachRow As Long
'
'    lPrevMaxRow = shtDataStage.Cells(Rows.Count, PREV_COL).End(xlUp).Row
'
'    For lEachRow = lPrevMaxRow To 1 Step -1
'        sPrevSheetName = Trim(shtDataStage.Cells(lEachRow, PREV_COL).Value)
'        shtDataStage.Cells(lEachRow, PREV_COL).ClearContents
'
'        If fZero(sPrevSheetName) Then GoTo previous_row
'
'        If fSheetExists(sPrevSheetName) Then
'            Set shtPrev = ThisWorkbook.Worksheets(sPrevSheetName)
'
'            If UCase(shtPrev.Name) = UCase(ActiveSheet.Name) Then
'                Call fAppendDataToLastCellOfColumn(shtDataStage, LAST_COL, sPrevSheetName)
'            Else
'                If fSheetIsVisible(shtPrev) Then
'                    'Application.EnableEvents = False
'                    shtPrev.Activate
'                    'Application.EnableEvents = True
'                    bFound = True
'                    Exit For
'                End If
'            End If
'        End If
'
'previous_row:
'    Next
'
'    If bFound Then
'        Call fAppendDataToLastCellOfColumn(shtDataStage, LAST_COL, sPrevSheetName)
'    End If
'
'exit_sub:
'    Set shtPrev = Nothing
'    'Application.EnableEvents = True
'End Sub

Function fAppendDataToLastCellOfColumn(ByRef sht As Worksheet, alCol As Long, aValue)
    Dim lMaxRow As Long
    lMaxRow = sht.Cells(Rows.Count, alCol).End(xlUp).Row
    
    If lMaxRow <= 1 Then
        If fZero(sht.Cells(lMaxRow, alCol).value) Then
            sht.Cells(lMaxRow, alCol).value = aValue
        Else
            sht.Cells(lMaxRow + 1, alCol).value = aValue
        End If
    Else
        sht.Cells(lMaxRow + 1, alCol).value = aValue
    End If
End Function

Sub Sub_DataMigration()
    On Error GoTo error_handling
    
    fInitialization

    Dim arrSource()
    Dim sOldFile As String
    Dim arrSheetsToMigr
    
    'to-do
    arrSheetsToMigr = Array(shtHospital _
                            , shtProductProducerMaster _
                            , shtProductNameMaster _
                            , shtProductMaster _
                            , shtSalesManMaster _
                            , shtHospitalReplace _
                            , shtProductProducerReplace _
                            , shtProductNameReplace _
                            , shtProductSeriesReplace _
                            , shtProductUnitRatio _
                            , shtSalesManCommConfig _
                            , shtSelfPurchaseOrder _
                            , shtSelfSalesOrder _
                            , shtFirstLevelCommission _
                            , shtSecondLevelCommission _
                            , shtNewRuleProducts _
                            , shtCompanyNameReplace _
                            , shtCZLRolloverInv _
                            , shtSalesCompRolloverInv _
                            , shtProductTaxRate _
                            , shtPromotionProduct _
                              )

    sOldFile = fSelectFileDialog(, "Macro File=*.xlsm", "Old Version With Latest User Data")
    If fZero(sOldFile) Then Exit Sub
    
    Call fExcelFileOpenedToCloseIt(sOldFile)
    
    Dim wbSource As Workbook
    Dim shtSource As Worksheet
    Dim eachSheet
    Dim shtTargetEach As Worksheet
    
    Application.AutomationSecurity = msoAutomationSecurityForceDisable

    Set wbSource = Workbooks.Open(Filename:=sOldFile, ReadOnly:=True)
    
    For Each eachSheet In arrSheetsToMigr
        Set shtTargetEach = eachSheet
        
        Set shtSource = fFindSheetBySheetCodeName(wbSource, shtTargetEach)
        Call fRemoveFilterForSheet(shtSource)
        
        Call fConvertFomulaToValueForSheetIfAny(shtSource)
        Call fCopyReadWholeSheetData2Array(shtSource, arrSource)
        'arrSource = wbSource.shtProductMaster.UsedRange.Value2
        Call fDeleteRemoveDataFormatFromSheetLeaveHeader(shtTargetEach)
        
        Call fWriteArray2Sheet(shtTargetEach, arrSource)
        
        If UBound(arrSource, 1) - LBound(arrSource, 1) + 2 <> fGetValidMaxRow(shtTargetEach) Then
            fErr "UBound(arrSource, 1) - LBound(arrSource, 1) + 2 <> fGetValidMaxRow(shtTargetEach)"
        End If
        
        Erase arrSource
    Next
    
    Call fCloseWorkBookWithoutSave(wbSource)
error_handling:
    If Err.Number <> 0 Then MsgBox Err.Description
    
    Erase arrSource
    If Not wbSource Is Nothing Then Call fCloseWorkBookWithoutSave(wbSource)
    
    Application.AutomationSecurity = msoAutomationSecurityByUI
    
    If fCheckIfGotBusinessError Then Err.Clear
    If fCheckIfUnCapturedExceptionAbnormalError Then End
    
    
    MsgBox "done"
End Sub


Function fCompareDictionaryKeys(dictBase As Dictionary, dictThis As Dictionary) As Dictionary
    Dim dictOut As Dictionary
    Dim i As Long
    Dim sKey As String
    
    Set dictOut = New Dictionary
    
    'missed from right one
    For i = 0 To dictBase.Count - 1
        sKey = dictBase.Keys(i)
        
        If Not dictThis.Exists(sKey) Then
            dictOut.Add DELETED_FROM_NEW_VERSION & DELIMITER & sKey, dictBase.Items(i) + 1
        Else
            'dictOut.Add SAME_IN_BOTH & DELIMITER & sKey, dictBase.Items(i) + 1 & DELIMITER & dictThis(sKey) + 1
            dictThis.Remove sKey
        End If
    Next
    
'    Dim iBlankColNum As Integer
'    If dictBase.Count > 0 Then iBlankColNum = UBound(Split(dictBase.Keys(0), DELIMITER)) - LBound(Split(dictBase.Keys(0), DELIMITER)) + 1
'    If dictThis <= 0 And dictThis.Count > 0 Then iBlankColNum = UBound(Split(dictThis.Keys(0), DELIMITER)) - LBound(Split(dictThis.Keys(0), DELIMITER)) + 1
    
    'missed from LEFT one
    For i = 0 To dictThis.Count - 1
        sKey = dictThis.Keys(i)
        
        'If Not dictBase.Exists(sKey) Then
            'dictOut.Add "新版本有而基础版本中没有" & String(DELIMITER, iBlankColNum) & sKey, dictThis.Items(i) + 1
            dictOut.Add NEWLY_ADDED_IN_NEW_VERSION & DELIMITER & sKey, dictThis.Items(i) + 1
        'End If
    Next
    
    Set fCompareDictionaryKeys = dictOut
    Set dictOut = Nothing
End Function

Function fCompareDictionaryKeysAndSingleItem(dictBase As Dictionary, dictThis As Dictionary) As Dictionary
    Dim dictOut As Dictionary
    Dim i As Long
    Dim sKey As String
    Dim sValue As String
    
    Set dictOut = New Dictionary
    
    'missed from right one
    For i = 0 To dictBase.Count - 1
        sKey = dictBase.Keys(i)
        
        If Not dictThis.Exists(sKey) Then
            dictOut.Add DELETED_FROM_NEW_VERSION & DELIMITER & sKey, dictBase.Items(i) & DELIMITER & "新版本中没有设置"
        Else
            If dictBase.Items(i) <> dictThis(sKey) Then
                dictOut.Add BOTH_HAVE_BUT_DIFF_VALUE & DELIMITER & sKey, dictBase.Items(i) & DELIMITER & dictThis(sKey)
            Else
                'dictOut.Add SAME_IN_BOTH & DELIMITER & sKey, dictBase.Items(i) & DELIMITER & dictThis(sKey)
            End If
            
            dictThis.Remove sKey
        End If
    Next
    
    'missed from LEFT one
    For i = 0 To dictThis.Count - 1
        sKey = dictThis.Keys(i)
        
        'If Not dictBase.Exists(sKey) Then
            dictOut.Add NEWLY_ADDED_IN_NEW_VERSION & DELIMITER & sKey, dictThis.Items(i) & DELIMITER & "基础版本中没有设置"
        'End If
    Next
    
    Set fCompareDictionaryKeysAndSingleItem = dictOut
    Set dictOut = Nothing
End Function

Function fCompareDictionaryKeysAndMultipleItems(ByRef dictBase As Dictionary, ByRef dictThis As Dictionary) As Dictionary
    Dim dictOut As Dictionary
    Dim i As Long
    Dim sKey As String
    Dim sValue As String
    
    Set dictOut = New Dictionary
    
    'missed from right one
    For i = 0 To dictBase.Count - 1
        sKey = dictBase.Keys(i)
        
        If Not dictThis.Exists(sKey) Then
            dictOut.Add DELETED_FROM_NEW_VERSION & DELIMITER & sKey, dictBase.Items(i) & vbLf & "新版本中没有设置"
        Else
            If dictBase.Items(i) <> dictThis(sKey) Then
                dictOut.Add BOTH_HAVE_BUT_DIFF_VALUE & DELIMITER & sKey, dictBase.Items(i) & vbLf & dictThis(sKey)
            Else
                'dictOut.Add SAME_IN_BOTH & DELIMITER & sKey, dictBase.Items(i) & vbLf & dictThis(sKey)
            End If
            dictThis.Remove sKey
        End If
    Next
    
    'missed from LEFT one
    For i = 0 To dictThis.Count - 1
        sKey = dictThis.Keys(i)
        
        'If Not dictBase.Exists(sKey) Then
            dictOut.Add NEWLY_ADDED_IN_NEW_VERSION & DELIMITER & sKey, dictThis.Items(i) & vbLf & "基础版本中没有设置"
        'End If
    Next
    
    Set fCompareDictionaryKeysAndMultipleItems = dictOut
    Set dictOut = Nothing
End Function
Function fFindSheetBySheetCodeName(wb As Workbook, shtToMatch As Worksheet) As Worksheet
    Dim shtMatched As Worksheet
    
    Dim shtEach As Worksheet
    
    For Each shtEach In wb.Worksheets
        If shtEach.CodeName = shtToMatch.CodeName Then
            Set shtMatched = shtEach
            Exit For
        End If
    Next
    
    If shtMatched Is Nothing Then fErr shtToMatch.CodeName & " cannot be found in the opened macro file."
    Set fFindSheetBySheetCodeName = shtMatched
    Set shtMatched = Nothing
End Function

Function fAutoFileterAllSheets()
    fResetAutoFilter shtCompanyNameReplace
    fResetAutoFilter shtHospital
    fResetAutoFilter shtHospitalReplace
    fResetAutoFilter shtSalesRawDataRpt
    fResetAutoFilter shtSalesInfos
    fResetAutoFilter shtProductMaster
    fResetAutoFilter shtProductNameReplace
    fResetAutoFilter shtProductProducerReplace
    fResetAutoFilter shtProductSeriesReplace
    fResetAutoFilter shtProductUnitRatio
    fResetAutoFilter shtProductProducerMaster
    fResetAutoFilter shtProductNameMaster
    fResetAutoFilter shtProfit
    fResetAutoFilter shtSelfSalesOrder
    fResetAutoFilter shtSelfSalesPreDeduct
    fResetAutoFilter shtSelfPurchaseOrder
    fResetAutoFilter shtSalesManMaster
    fResetAutoFilter shtFirstLevelCommission
    fResetAutoFilter shtSecondLevelCommission
    fResetAutoFilter shtSalesManCommConfig
    fResetAutoFilter shtSelfInventory
    fResetAutoFilter shtInventoryRawDataRpt
    fResetAutoFilter shtImportCZL2SalesCompSales
    fResetAutoFilter shtCZLSales2CompRawData
    fResetAutoFilter shtCZLSales2Companies
    fResetAutoFilter shtCZLInvDiff
    fResetAutoFilter shtPromotionProduct
    fResetAutoFilter shtSalesCompInvUnified
    fResetAutoFilter shtSalesCompInvCalcd
    fResetAutoFilter shtSalesCompInvDiff
    fResetAutoFilter shtProductTaxRate
    fResetAutoFilter shtRefund
End Function

Function fResetAutoFilter(sht As Worksheet)
    sht.Rows(1).AutoFilter
    sht.Rows(1).AutoFilter
End Function


Sub subMain_RefreshAllPvTables()
    ThisWorkbook.RefreshAll
    fShowAndActiveSheet shtPV
End Sub

Sub subMain_InvisibleHideCurrentSheet()
    If shtMainMenu.CodeName = ActiveSheet.CodeName Then Exit Sub
    
    'If ThisWorkbook.Worksheets.Count > 1 Then
        fVeryHideSheet ActiveSheet
    'End If
End Sub

Function fGetReplaceUnifyErrorRowCount_SCompSalesInfo() As Long
    fGetReplaceUnifyErrorRowCount_SCompSalesInfo = CLng(fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_SALES_INFO"))
End Function
Function fSetReplaceUnifyErrorRowCount_SCompSalesInfo(ByVal rowCnt As Long) As Long
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_SALES_INFO", CStr(rowCnt))
End Function

Function fGetReplaceUnifyErrorRowCount_SalesInventory() As Long
    fGetReplaceUnifyErrorRowCount_SalesInventory = CLng(fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_COMPNAY_INVENTORY"))
End Function
Function fSetReplaceUnifyErrorRowCount_SCompInventory(ByVal rowCnt As Long) As Long
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_COMPNAY_INVENTORY", CStr(rowCnt))
End Function

Function fGetReplaceUnifyErrorRowCount_CZLSales2Comp() As Long
    fGetReplaceUnifyErrorRowCount_CZLSales2Comp = CLng(fGetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_CZL_SALES_2_COMPANIES"))
End Function
Function fSetReplaceUnifyErrorRowCount_CZLSales2Comp(ByVal rowCnt As Long) As Long
    Call fSetSpecifiedConfigCellValue(shtSysConf, "[Facility For Testing]", "Value", "Setting Item ID=REPLACE_UNIFY_ERR_ROW_COUNT_CZL_SALES_2_COMPANIES", CStr(rowCnt))
End Function

Sub subMain_CloneMacro()
    If Workbooks.Count <= 0 Then MsgBox "Please open the macro first.": Exit Sub
    
    Dim wbSource As Workbook
    Dim wbTarget As Workbook
     
    Dim sSourceMacro As String
    Dim sTargetMacro As String
    Dim sExportParentFolder As String
    Dim sSourceCodeFolder_Left As String
    Dim sSourceCodeFolder_Right As String
    Dim bSameFileName As Boolean
    
    'On Error GoTo error_handling
    
    Call fInitialization
    
    FrmCloneMacro.Show
    If gsRtnValueOfForm <> CONST_SUCCESS Then GoTo error_handling
    
    sSourceMacro = fGetValue(RANGE_CloneMacro_Source)
    sTargetMacro = fGetValue(RANGE_CloneMacro_Target)
    
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    
    Set wbSource = Workbooks(fGetFileBaseName(sSourceMacro))
    
    If fFileExists(sTargetMacro) Then
        Dim wb As Workbook
        Dim bAlreadyOpened As Boolean
        bAlreadyOpened = False
        
        If fExcelFileIsOpen(sTargetMacro, wb) Then
            If UCase(wb.FullName) <> UCase(sTargetMacro) Then
                fErr "Another excel file with the same file name has already been open, please close it first" & vbCr _
                 & "File Name: " & wb.Name & vbCr & vbCr _
                 & "File fullpath: " & wb.FullName
            Else
                Set wbTarget = wb
                bAlreadyOpened = True
            End If
        End If
        
        If Not bAlreadyOpened Then
            fMsgBox "The macro already exists, please open it first.": fErr
        End If
    Else
        Set wbTarget = Workbooks.Add(xlWBATWorksheet)
        Call wbTarget.SaveAs(sTargetMacro, xlOpenXMLWorkbookMacroEnabled)
    End If
    
    Call fCopyAllItemsToAnotherMacroCloneMacro(wbSource, wbTarget)
    
    Call fListAllExternalLinks(wbTarget)
    Call fListAllNames(wbTarget)
    Call fListAllValidation(wbTarget)
    
    wbTarget.Save
    wbTarget.Activate
     
error_handling:
    Set wbSource = Nothing
    Set wbTarget = Nothing
    If gsRtnValueOfForm <> CONST_SUCCESS Then GoTo reset_excel_options
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    fMsgBox "Done, you may still need to check the sheet [" & POTENTIOAL_UNSOLVED_EXTERNAL_LINES & "] to check if any external link exists and need to be fixed manually. ", vbExclamation
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Sub subMain_GenCodeSnippet()
    On Error GoTo error_handling
    
    Call fInitialization
    
    frmCodeSnippet.Show
    If gsRtnValueOfForm = CONST_CANCEL Then fErr
     
    Dim sContent As String
    
    If gsRtnValueOfForm = "ENUM" Then
        sContent = fGenCodeSnippet_Enum_BaseOnSelection
'    ElseIf gsRtnValueOfForm = "FOR_LOOP" Then
'        sContent = fGenCodeSnippet_ForLoop
'    ElseIf gsRtnValueOfForm = "SELECT_CASE" Then
'        sContent = fGenCodeSnippet_SELECT_CASE
    Else
        fErr "gsRtnValueOfForm not covered in subMain_GenCodeSnippet: " & gsRtnValueOfForm
    End If
    
    ClipBoard_SetData sContent
error_handling:
    If gsRtnValueOfForm <> CONST_SUCCESS Then GoTo reset_excel_options
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    fMsgBox "Done, you may still need to check the sheet [" & POTENTIOAL_UNSOLVED_EXTERNAL_LINES & "] to check if any external link exists and need to be fixed manually. ", vbExclamation
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
'    Application.EnableEvents = True
'    Application.DisplayAlerts = True
End Sub
Private Function fGenCodeSnippet_Enum_BaseOnSelection() As String
    Dim sContent As String
    Dim rgSelection As Range
    Dim i As Integer
    Dim aCell As Range
    Dim sElem As String
    Dim aRg As Range
    
    If Workbooks.Count > 0 Then Set rgSelection = Selection
    
    sContent = "Enum PleaseChangeIt"
    sContent = sContent & vbCr & "    [_first] = 1"
    
    If Not rgSelection Is Nothing Then
        i = 0
        If rgSelection.Rows.Count > 1 And rgSelection.Columns.Count <= 1 Then
            For Each aRg In rgSelection.Rows
                If Len(Trim(aRg.Cells(1, 1).value)) > 0 Then
                    sElem = Trim(aRg.Cells(1, 1).value)
                    i = i + 1
                    If InStr(sElem, " ") > 0 Or InStr(sElem, "_") > 0 Or InStr(sElem, "\") > 0 Or InStr(sElem, "/") > 0 Then
                        sElem = WorksheetFunction.Proper(sElem)
                    End If
                    sElem = Replace(sElem, "_", "")
                    sElem = Replace(sElem, "/", "")
                    sElem = Replace(sElem, "\", "")
                    sElem = Replace(sElem, " ", "")
                    
                    sContent = sContent & vbCr _
                             & Left("    " & sElem & " = " & i & Space(50), 50) & "'" & fNum2Letter(i)
                End If
            Next
            
            sContent = sContent & vbCr & "    [_last] = " & sElem
        ElseIf rgSelection.Rows.Count <= 1 And rgSelection.Columns.Count > 1 Then
            For Each aRg In rgSelection.Columns
                If Len(Trim(aRg.Cells(1, 1).value)) > 0 Then
                    sElem = Trim(aRg.Cells(1, 1).value)
                    i = i + 1
                    If InStr(sElem, " ") > 0 Or InStr(sElem, "_") > 0 Or InStr(sElem, "\") > 0 Or InStr(sElem, "/") > 0 Then
                        sElem = WorksheetFunction.Proper(sElem)
                    End If
                    sElem = Replace(sElem, "_", "")
                    sElem = Replace(sElem, "/", "")
                    sElem = Replace(sElem, "\", "")
                    sElem = Replace(sElem, " ", "")
                    
                    sContent = sContent & vbCr _
                             & Left("    " & sElem & " = " & i & Space(50), 50) & "'" & fNum2Letter(i)
                End If
            Next
            
            sContent = sContent & vbCr & "    [_last] = " & sElem
        Else
            For Each aCell In rgSelection.Cells
                If Len(Trim(aCell.value)) > 0 Then
                    sElem = Trim(aCell.value)
                    i = i + 1
                    If InStr(sElem, " ") > 0 Or InStr(sElem, "_") > 0 Or InStr(sElem, "\") > 0 Or InStr(sElem, "/") > 0 Then
                        sElem = WorksheetFunction.Proper(sElem)
                    End If
                    sElem = Replace(sElem, " ", "")
                    
                    sContent = sContent & vbCr _
                             & Left("    " & sElem & " = " & i & Space(50), 50) & "'" & fNum2Letter(i)
                End If
            Next
            sContent = sContent & vbCr & "    [_last] = " & sElem
        End If
    Else
        For i = 1 To 20
            sContent = sContent & vbCr _
                     & Left("    " & "Col_" & i & " = " & i & Space(50), 50) & "'" & fNum2Letter(i)
        Next
        sContent = sContent & vbCr & "    [_last] = " & "Col_" & i
    End If
    
    sContent = sContent & vbCr & "End Enum"
    fGenCodeSnippet_Enum_BaseOnSelection = sContent
End Function

Function fListAllExternalLinks(wb As Workbook)
    Dim rg As Range
    Dim eachcell As Range
    Dim dict As Dictionary
    Dim sht As Worksheet
    
    Set dict = New Dictionary
    For Each sht In wb.Worksheets
        On Error Resume Next
        Set rg = sht.UsedRange.SpecialCells(xlCellTypeFormulas)
    
        If rg Is Nothing Then GoTo next_one
        
        For Each eachcell In rg.Cells
            If InStr(1, eachcell.Formula, "[") > 0 Then
                If eachcell.MergeCells Then
                    If eachcell.Address = eachcell.MergeArea.Cells(1, 1).Address Then
                        dict.Add "Formula: " & DELIMITER _
                                & "'" & sht.Name & DELIMITER _
                                & "'" & eachcell.MergeArea.Address & DELIMITER _
                                & "'" & eachcell.Formula, ""
                    End If
                Else
                    dict.Add "Formula: " & DELIMITER _
                            & "'" & sht.Name & DELIMITER _
                            & "'" & eachcell.Address & DELIMITER _
                            & "'" & eachcell.Formula, ""
                End If
            End If
        Next
        
next_one:
    Next
    
    Dim extLink
    For Each extLink In wb.LinkSources
        If Not fAreSame(CStr(extLink), ThisWorkbook.FullName) Then
            dict.Add "Link: " & DELIMITER _
                    & "'" & sht.Name & DELIMITER _
                    & "'" & extLink & DELIMITER _
                    & "", ""
        End If
    Next
    On Error GoTo 0
    
    Call fPastePotentialRiskLinkstoManuallyHandleSheet(dict, wb, Array("Type", "sheet name", "cell", "formula/link name"))
    Set dict = Nothing
End Function
Function fListAllNames(wb As Workbook)
    Dim eachName As Name
    Dim sName As String
    Dim dict As Dictionary
    Set dict = New Dictionary
    Dim sht As Worksheet
    
    For Each sht In wb.Worksheets
        For Each eachName In sht.Names
            sName = eachName.Name
            
            If InStr(1, sName, "!") > 0 Then sName = Split(sName, "!")(1)
            On Error Resume Next
                dict.Add "Name: " & DELIMITER _
                        & "'" & sName & DELIMITER _
                        & "'" & eachName.Parent.Name & DELIMITER _
                        & "'" & eachName.RefersTo & DELIMITER _
                        & "'" & eachName.value & DELIMITER _
                        & "'" & eachName.RefersToLocal, ""
            On Error GoTo 0
        Next
    Next
    
    For Each eachName In wb.Names
        sName = eachName.Name
        
        If InStr(1, sName, "!") > 0 Then sName = Split(sName, "!")(1)
        On Error Resume Next
            dict.Add "Name: " & DELIMITER _
                    & "'" & sName & DELIMITER _
                    & "'" & eachName.Parent.Name & DELIMITER _
                    & "'" & eachName.RefersTo & DELIMITER _
                    & "'" & eachName.value & DELIMITER _
                    & "'" & eachName.RefersToLocal, ""
        On Error GoTo 0
    Next
    
    Call fPastePotentialRiskLinkstoManuallyHandleSheet(dict, wb, Array("Type", "name's name", "whhere name is", "refers to", "value", "RefersToLocal"))
    Set dict = Nothing
End Function
Function fListAllValidation(wb As Workbook)
    Dim rg As Range
    Dim eachcell As Range
    Dim dict As Dictionary
    Dim sht As Worksheet
    Dim vD As Validation
    Dim vdType As XlDVType
    Dim sType As String
    
    Set dict = New Dictionary
    For Each sht In wb.Worksheets
        On Error Resume Next
        Set rg = sht.UsedRange.SpecialCells(xlCellTypeAllValidation)
    
        If rg Is Nothing Then GoTo next_one
        
        For Each eachcell In rg.Cells
            Set vD = eachcell.Validation
            vdType = vD.Type
             
            Select Case vdType
                Case xlValidateCustom
                    sType = "xlValidateCustom"
                Case xlValidateDate
                    sType = "xlValidateDate"
                Case xlValidateDecimal
                    sType = "xlValidateDecimal"
                Case xlValidateInputOnly
                    sType = "xlValidateInputOnly"
                    GoTo next_one
                Case xlValidateList
                    sType = "xlValidateList"
                Case xlValidateTextLength
                    sType = "xlValidateTextLength"
                Case xlValidateTime
                    sType = "xlValidateTime"
                Case xlValidateWholeNumber
                    sType = "xlValidateWholeNumber"
                Case Else
                    sType = "Unknow validation Type"
            End Select
            
            If eachcell.MergeCells Then
                If eachcell.Address = eachcell.MergeArea.Cells(1, 1).Address Then
                    dict.Add "Validation: " & DELIMITER _
                            & "'" & sht.Name & DELIMITER _
                            & "'" & eachcell.MergeArea.Address & DELIMITER _
                            & "'" & vD.Formula1 & DELIMITER _
                            & "'" & vD.Formula2 & DELIMITER _
                            & "'" & sType, ""
                End If
            Else
                dict.Add "Validation: " & DELIMITER _
                        & "'" & sht.Name & DELIMITER _
                        & "'" & eachcell.Address & DELIMITER _
                        & "'" & vD.Formula1 & DELIMITER _
                        & "'" & vD.Formula2 & DELIMITER _
                        & "'" & sType, ""
            End If
        
        Next
        On Error GoTo 0
next_one:
    Next
     
    Call fPastePotentialRiskLinkstoManuallyHandleSheet(dict, wb, Array("Type", "sheet name", "cell", "formula1", "formula2", "validation type"))
    Set dict = Nothing
End Function

Function fPastePotentialRiskLinkstoManuallyHandleSheet(dict As Dictionary, wb As Workbook, arrHeader)
    If dict.Count <= 0 Then Exit Function
    
    Dim shtLog As Worksheet
    If Not fSheetExists(POTENTIOAL_UNSOLVED_EXTERNAL_LINES, , wb) Then
        Set shtLog = fAddNewSheet(POTENTIOAL_UNSOLVED_EXTERNAL_LINES, wb)
    Else
        Set shtLog = wb.Worksheets(POTENTIOAL_UNSOLVED_EXTERNAL_LINES)
    End If
    
    Dim lRow As Long
    
    lRow = fGetValidMaxRow(shtLog)
    If lRow = 0 Then
        lRow = 1
    Else
        lRow = lRow + 2
    End If
    shtLog.Cells(lRow, 1).Resize(1, ArrLen(arrHeader)).value = arrHeader
    
    Call fPasteAppendDictionaryToSheet(dict, shtLog)
    Set shtLog = Nothing
End Function

Function fCopyAllItemsToAnotherMacroCloneMacro(wbSource As Workbook, wbTarget As Workbook)
    Dim vbSourcePrj As VBIDE.VBProject
    Dim vbTargetPrj As VBIDE.VBProject
    
    Set vbSourcePrj = wbSource.VBProject
    Set vbTargetPrj = wbTarget.VBProject
    
    '========== clean first ==============================
    Dim sTmpShtName As String
    sTmpShtName = fGenRandomUniqueString()
    Call fAddNewSheet(sTmpShtName, wbTarget)
    
    Dim sht As Worksheet
    For Each sht In wbTarget.Worksheets
        If Not fAreSame(sht.Name, sTmpShtName) Then
            sht.Visible = xlSheetVisible
            sht.Delete
        End If
    Next
    
    Dim vbcomp As VBComponent
    For Each vbcomp In vbTargetPrj.VBComponents
        If vbcomp.Type = vbext_ct_StdModule _
        Or vbcomp.Type = vbext_ct_ClassModule _
        Or vbcomp.Type = vbext_ct_MSForm Then
            Call vbTargetPrj.VBComponents.Remove(vbcomp)
        Else
            '
        End If
    Next
    '=====================================================
    
    '========== export all module's source code ==========
    Dim sSourceFolder As String
    sSourceFolder = wbSource.Path & "\TempFolder_SourceCode"
    If fFolderExists(sSourceFolder) Then
        fGetFSO
        gFSO.DeleteFolder sSourceFolder, True
    End If
    MkDir sSourceFolder
    
    For Each vbcomp In vbSourcePrj.VBComponents
        If vbcomp.Type = vbext_ct_StdModule _
        Or vbcomp.Type = vbext_ct_ClassModule _
        Or vbcomp.Type = vbext_ct_MSForm Then
            vbcomp.Export sSourceFolder & "\" & vbcomp.Name & ".bas"
        End If
    Next
    Set vbcomp = Nothing
    '=====================================================
    
    '================ copy all sheets first ==============
    For Each sht In wbSource.Worksheets
        If fSheetExistsByCodeName(sht.CodeName, , wbTarget) Then
            Dim sTmpOrig As String
            sTmpOrig = sTmpShtName
            sTmpShtName = fGenRandomUniqueString()
            Call fAddNewSheet(sTmpShtName, wbTarget)
            Call fDeleteSheet(sTmpOrig, wbTarget)
        End If
        
        sht.Visible = xlSheetVisible
        sht.Copy after:=wbTarget.Worksheets(wbTarget.Worksheets.Count())
    Next
    
    Call fDeleteSheet(sTmpShtName, wbTarget)
    '=====================================================
    
    '================ import all modules    ==============
    On Error Resume Next
    Dim aFile As File
    Dim aModule As VBComponent
    For Each aFile In gFSO.GetFolder(sSourceFolder).Files
        Set aModule = vbTargetPrj.VBComponents(gFSO.GetBaseName(aFile.Name))
        If Err.Number = 0 Then
            If aModule Is Nothing Then
                vbTargetPrj.VBComponents.Import aFile.Path
                Call fSleep(10)
            End If
        Else
            Err.Clear
            vbTargetPrj.VBComponents.Import aFile.Path
            Call fSleep(10)
        End If
    Next
    On Error GoTo 0
    '=====================================================
    
    '================ add all reference     ==============
    On Error Resume Next
    Dim eachRef As Reference
    
    For Each eachRef In vbSourcePrj.References
        Dim tmpRef As Reference
        Set tmpRef = vbTargetPrj.References(eachRef.Name)
        If Err.Number = 0 Then
            If tmpRef Is Nothing Then
                Call vbTargetPrj.References.AddFromGuid(eachRef.GUID, eachRef.Major, eachRef.Minor)
                Call fSleep(10)
            End If
        Else
            Err.Clear
            Call vbTargetPrj.References.AddFromGuid(eachRef.GUID, eachRef.Major, eachRef.Minor)
            Call fSleep(10)
        End If
    Next
    
    On Error GoTo 0
    '=====================================================
    
    '================ change link           ==============
    On Error Resume Next
    Call wbTarget.ChangeLink(wbSource.Name, wbTarget.Name)
    On Error GoTo 0
    '=====================================================
    
    '================ thisworkbook module code==============
    Dim sCodeScript As String
    Dim sourceCodeM As CodeModule
    Dim targetCodeM As CodeModule
    
    Set sourceCodeM = vbSourcePrj.VBComponents("ThisWorkbook").CodeModule
    Set targetCodeM = vbTargetPrj.VBComponents("ThisWorkbook").CodeModule
    
    sCodeScript = sourceCodeM.Lines(1, sourceCodeM.CountOfLines)
    targetCodeM.DeleteLines 1, targetCodeM.CountOfLines
    targetCodeM.AddFromString sCodeScript
    
    Set sourceCodeM = Nothing
    Set targetCodeM = Nothing
    '=====================================================
    
    Set vbSourcePrj = Nothing
    Set vbTargetPrj = Nothing
End Function
  
Function fAreSame(sA As String, sB As String, Optional bCaseSensitive As Boolean = False, Optional bTrim As Boolean = True) As Boolean
    If bCaseSensitive Then
        If bTrim Then
            fAreSame = CBool(Trim(sA) = Trim(sB))
        Else
            fAreSame = CBool(sA = sB)
        End If
    Else
        If bTrim Then
            fAreSame = CBool(UCase(Trim(sA)) = UCase(Trim(sB)))
        Else
            fAreSame = CBool(UCase(sA) = UCase(sB))
        End If
    End If
End Function

Function fSleep(howlong As Long)
    Dim i As Long
    
    howlong = howlong * 300
    For i = 0 To howlong
        DoEvents
    Next
End Function
