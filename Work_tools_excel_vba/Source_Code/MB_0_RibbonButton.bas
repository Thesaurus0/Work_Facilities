Attribute VB_Name = "MB_0_RibbonButton"
Option Explicit
Option Base 1

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

 
  
