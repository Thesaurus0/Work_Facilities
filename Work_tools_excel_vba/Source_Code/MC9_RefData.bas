Attribute VB_Name = "MC9_RefData"
'Option Explicit
'Option Base 1
'
'Function fClearRefVariables()
'    Set dictSelfSalesPrice = Nothing
'    Set dictCompanyNameReplace = Nothing
'    Set dictHospitalMaster = Nothing
'    Set dictHospitalReplace = Nothing
'    Set dictProducerMaster = Nothing
'    Set dictProducerReplace = Nothing
'    Set dictProductNameMaster = Nothing
'    Set dictProductNameReplace = Nothing
'    Set dictProductMaster = Nothing
'    Set dictProductSeriesReplace = Nothing
'    Set dictProductUnitRatio = Nothing
'    Set dictSalesManMaster = Nothing
'    Set dictFirstLevelComm = Nothing
'    Set dictSecondLevelComm = Nothing
'    Set dictCompanyNameID = Nothing
'    Set dictDefaultCommConfiged = Nothing
'    'Set dictSelfSalesReverse = Nothing
'    Set dictSelfSalesMinus = Nothing
'    Set dictSelfSalesDeduct = Nothing
'    Set dictSelfSalesColIndex = Nothing
'    Set dictSalesManCommFrom = Nothing
'    Set dictSalesManCommColIndex = Nothing
'    Set dictExcludeProducts = Nothing
'    Set dictProdTaxRate = Nothing
'    Set dictNewRuleProducts = Nothing
'    Set dictPromotionProducts = Nothing
'
'    'global variable
'    Set dictErrorRows = Nothing
'    Set dictWarningRows = Nothing
'End Function
'
'Function fReadConfigCompanyList(Optional ByRef dictCompanyNameID As Dictionary) As Dictionary
'    Dim asTag As String
'    Dim arrColsName()
'    Dim rngToFindIn As Range
'    Dim arrConfigData()
'    Dim lConfigStartRow As Long
'    Dim lConfigStartCol As Long
'    Dim lConfigEndRow As Long
'    Dim lConfigHeaderAtRow As Long
'
'    asTag = "[Sales Company List]"
'    ReDim arrColsName(Company.REPORT_ID To Company.Selected)
'
'    arrColsName(Company.REPORT_ID) = "Company ID"
'    arrColsName(Company.ID) = "Company ID In DB"
'    arrColsName(Company.Name) = "Company Name"
'    arrColsName(Company.Commission) = "Default Commission"
'    arrColsName(Company.CheckBoxName) = "CheckBox Name"
'    arrColsName(Company.InputFileTextBoxName) = "Input File TextBox Name"
'    arrColsName(Company.Selected) = "User Ticked"
'
'    arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, shtParam:=shtStaticData _
'                                , arrColsName:=arrColsName _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
'    Call fValidateDuplicateInArray(arrConfigData, Company.REPORT_ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID")
'    Call fValidateDuplicateInArray(arrConfigData, Company.ID, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
'    Call fValidateDuplicateInArray(arrConfigData, Company.Name, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
'    Call fValidateDuplicateInArray(arrConfigData, Company.CheckBoxName, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
'    Call fValidateDuplicateInArray(arrConfigData, Company.InputFileTextBoxName, False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
'
''    Call fValidateBlankInArray(arrConfigData, Company.Report_ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID")
''    Call fValidateBlankInArray(arrConfigData, Company.ID, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company ID In DB")
''    Call fValidateBlankInArray(arrConfigData, Company.Name, shtstaticdata, lConfigHeaderAtRow, lConfigStartCol, "Company Name")
'
''    Set fReadConfigCompanyList = fReadArray2DictionaryWithMultipleColsCombined(arrConfigData, Company.Report_ID _
''            , Array(Company.ID, Company.Name, Company.Commission, Company.CheckBoxName, Company.InputFileTextBoxName, Company.Selected) _
''            , DELIMITER)
'
'    Dim dictOut As Dictionary
'    Set dictOut = New Dictionary
'
'    Set dictCompanyNameID = New Dictionary
'
'    Dim lEachRow As Long
'    Dim sFileTag As String
'    Dim sValueStr As String
'
'    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
'        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
'
''        sRptNameStr = DELIMITER & arrConfigData(lEachRow, 1) & DELIMITER
''        If InStr(sRptNameStr, DELIMITER & asReportID & DELIMITER) <= 0 Then GoTo next_row
'
'        'lActualRow = lConfigHeaderAtRow + lEachRow
'
'        sFileTag = Trim(arrConfigData(lEachRow, Company.REPORT_ID))
'        sValueStr = fComposeStrForDictCompanyList(arrConfigData, lEachRow)
'
'        dictOut.Add sFileTag, sValueStr
'
'        dictCompanyNameID.Add arrConfigData(lEachRow, Company.Name), arrConfigData(lEachRow, Company.REPORT_ID)
'next_row:
'    Next
'
'    Erase arrColsName
'    Erase arrConfigData
'    Set fReadConfigCompanyList = dictOut
'    Set dictOut = Nothing
'End Function
'
'Private Function fComposeStrForDictCompanyList(arrConfigData, lEachRow As Long) As String
'    Dim sOut As String
'    Dim i As Integer
'
'    For i = Company.ID To Company.Selected
'        sOut = sOut & DELIMITER & Trim(arrConfigData(lEachRow, i))
'    Next
'
'    fComposeStrForDictCompanyList = Right(sOut, Len(sOut) - 1)
'End Function
'
'Function fGetCompany_InputFileTextBoxName(asCompanyID As String) As String
'    fGetCompany_InputFileTextBoxName = Split(dictCompList(asCompanyID), DELIMITER)(Company.InputFileTextBoxName - Company.REPORT_ID - 1)
'End Function
'Function fGetCompany_CheckBoxName(asCompanyID As String) As String
'    fGetCompany_CheckBoxName = Split(dictCompList(asCompanyID), DELIMITER)(Company.CheckBoxName - Company.REPORT_ID - 1)
'End Function
'Function fGetCompany_UserTicked(asCompanyID As String) As String
'    fGetCompany_UserTicked = Split(dictCompList(asCompanyID), DELIMITER)(Company.Selected - Company.REPORT_ID - 1)
'End Function
'Function fGetCompany_CompanyLongID(asCompanyID As String) As String
'    fGetCompany_CompanyLongID = Split(dictCompList(asCompanyID), DELIMITER)(Company.ID - Company.REPORT_ID - 1)
'End Function
'
''please use : fGetCompanyNameByID_Common
''Function fGetCompany_CompanyName(asCompanyID As String) As String
''    If dictCompList Is Nothing Then Set dictCompList = fReadConfigCompanyList
''    fGetCompany_CompanyName = Split(dictCompList(asCompanyID), DELIMITER)(Company.Name - Company.REPORT_ID - 1)
''End Function
'
''====================== Hospital Master =================================================================
'Function fReadSheetHospitalMaster2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("HOSPITAL_MASTER", dictColIndex, arrData, , , , , shtHospital)
'    Set dictHospitalMaster = fReadArray2DictionaryOnlyKeys(arrData, dictColIndex("Hospital"))
'
'    Set dictColIndex = Nothing
'End Function
'Function fHospitalExistsInHospitalMaster(sHospital As String) As Boolean
'    If dictHospitalMaster Is Nothing Then Call fReadSheetHospitalMaster2Dictionary
'
'    fHospitalExistsInHospitalMaster = dictHospitalMaster.Exists(sHospital)
'End Function
''------------------------------------------------------------------------------
'
''====================== Hospital Replacement =================================================================
'Function fReadSheetHospitalReplace2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("HOSPITAL_REPLACE_SHEET", dictColIndex, arrData, , , , , shtHospitalReplace)
'    Set dictHospitalReplace = fReadArray2DictionaryWithSingleCol(arrData, dictColIndex("FromHospital"), dictColIndex("ToHospital"))
'
'    Set dictColIndex = Nothing
'End Function
'Function fFindInConfigedReplaceHospital(sHospital As String) As String
'    If dictHospitalReplace Is Nothing Then Call fReadSheetHospitalReplace2Dictionary
'
'    If dictHospitalReplace.Exists(sHospital) Then
'        fFindInConfigedReplaceHospital = dictHospitalReplace(sHospital)
'    Else
'        fFindInConfigedReplaceHospital = ""
'    End If
'End Function
''------------------------------------------------------------------------------
'
''====================== Producer Master =================================================================
'Function fReadSheetProducerMaster2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PRODUCER_MASTER", dictColIndex, arrData, , , , , shtProductProducerMaster)
'    Set dictProducerMaster = fReadArray2DictionaryOnlyKeys(arrData, dictColIndex("ProductProducer"), True, False)
'
'    Set dictColIndex = Nothing
'End Function
'Function fProducerExistsInProducerMaster(sProducer As String) As Boolean
'    If dictProducerMaster Is Nothing Then Call fReadSheetProducerMaster2Dictionary
'
'    fProducerExistsInProducerMaster = dictProducerMaster.Exists(sProducer)
'End Function
''------------------------------------------------------------------------------
'
''====================== Producer Replacement =================================================================
'Function fReadSheetProducerReplace2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PRODUCER_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductProducerReplace)
'    Set dictProducerReplace = fReadArray2DictionaryWithSingleCol(arrData, dictColIndex("FromProducer"), dictColIndex("ToProducer"))
'
'    Set dictColIndex = Nothing
'End Function
'Function fFindInConfigedReplaceProducer(sProducer As String) As String
'    If dictProducerReplace Is Nothing Then Call fReadSheetProducerReplace2Dictionary
'
'    If dictProducerReplace.Exists(sProducer) Then
'        fFindInConfigedReplaceProducer = dictProducerReplace(sProducer)
'    Else
'        fFindInConfigedReplaceProducer = ""
'    End If
'End Function
''------------------------------------------------------------------------------
'
'
''====================== ProductName Master =================================================================
'Function fReadSheetProductNameMaster2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PRODUCER_NAME_MASTER", dictColIndex, arrData, , , , , shtProductNameMaster)
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")), False, shtProductNameMaster, 1, 1, "厂家 + 名称")
'
'    Set dictProductNameMaster = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName")) _
'                                    , Array(dictColIndex("ProductName")), DELIMITER, DELIMITER, True)
'    Set dictColIndex = Nothing
'End Function
'Function fProductNameExistsInProductNameMaster(sProductProducer As String, sProductName As String) As Boolean
'    If dictProductNameMaster Is Nothing Then Call fReadSheetProductNameMaster2Dictionary
'
'    fProductNameExistsInProductNameMaster = dictProductNameMaster.Exists(sProductProducer & DELIMITER & sProductName)
'End Function
''------------------------------------------------------------------------------
'
''====================== ProductName Replacement =================================================================
'Function fReadSheetProductNameReplace2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PRODUCT_NAME_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductNameReplace)
'    Set dictProductNameReplace = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'                                    , Array(dictColIndex("ProductProducer"), dictColIndex("FromProductName")) _
'                                    , Array(dictColIndex("ToProductName")), DELIMITER, DELIMITER)
'    Set dictColIndex = Nothing
'End Function
'Function fFindInConfigedReplaceProductName(sProductProducer As String, sProductName As String) As String
'    Dim sKey As String
'
'    If dictProductNameReplace Is Nothing Then Call fReadSheetProductNameReplace2Dictionary
'
'    sKey = sProductProducer & DELIMITER & sProductName
'
'    If dictProductNameReplace.Exists(sKey) Then
'        fFindInConfigedReplaceProductName = dictProductNameReplace(sKey)
'    Else
'        fFindInConfigedReplaceProductName = ""
'    End If
'End Function
''------------------------------------------------------------------------------
'
''====================== ProductSeries Master =================================================================
'Function fReadSheetProductMaster2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PRODUCT_MASTER", dictColIndex, arrData, , , , , shtProductMaster)
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtProductMaster, 1, 1, "厂家 + 名称 + 规格")
'
'    Set dictProductMaster = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")) _
'                                    , Array(dictColIndex("ProductUnit"), dictColIndex("LatestPrice")), DELIMITER, DELIMITER)
'    Set dictColIndex = Nothing
'End Function
'Function fProductSeriesExistsInProductMaster(sProductProducer As String, sProductName As String, sProductSeries As String) As Boolean
'    If dictProductMaster Is Nothing Then Call fReadSheetProductMaster2Dictionary
'
'    fProductSeriesExistsInProductMaster = dictProductMaster.Exists(sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries)
'End Function
'Function fProductKeysExistsInProductMaster(sProductProducer As String, sProductName As String, sProductSeries As String) As Boolean
'    fProductKeysExistsInProductMaster = fProductSeriesExistsInProductMaster(sProductProducer, sProductName, sProductSeries)
'End Function
''------------------------------------------------------------------------------
'
''====================== ProductSeries Replacement =================================================================
'Function fReadSheetProductSeriesReplace2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PRODUCT_SERIES_REPLACE_SHEET", dictColIndex, arrData, , , , , shtProductSeriesReplace)
'    Set dictProductSeriesReplace = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("FromProductSeries")) _
'                                    , Array(dictColIndex("ToProductSeries")), DELIMITER, DELIMITER)
'    Set dictColIndex = Nothing
'End Function
'Function fFindInConfigedReplaceProductSeries(sProductProducer As String, sProductName As String, sOrigProductSeries As String) As String
'    Dim sKey As String
'
'    If dictProductSeriesReplace Is Nothing Then Call fReadSheetProductSeriesReplace2Dictionary
'
'    sKey = sProductProducer & DELIMITER & sProductName & DELIMITER & sOrigProductSeries
'
'    If dictProductSeriesReplace.Exists(sKey) Then
'        fFindInConfigedReplaceProductSeries = dictProductSeriesReplace(sKey)
'    Else
'        fFindInConfigedReplaceProductSeries = ""
'    End If
'End Function
''------------------------------------------------------------------------------
'
''====================== ProductUnit Ratio =================================================================
'Function fReadSheetProductUnitRatio2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PRODUCT_UNIT_RATIO_SHEET", dictColIndex, arrData, , , , , shtProductUnitRatio)
'    Set dictProductUnitRatio = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries"), dictColIndex("FromUnit")) _
'                                    , Array(dictColIndex("ProductUnit"), dictColIndex("Ratio")), DELIMITER, DELIMITER)
'    Set dictColIndex = Nothing
'End Function
'Function fFindInConfigedReplaceProductUnit(sProductProducer As String, sProductName As String, sProductSeries As String _
'                            , sOrigProductUnit As String _
'                            , ByRef dblRatio As Double) As String
'    Dim sKey As String
'
'    If dictProductUnitRatio Is Nothing Then Call fReadSheetProductUnitRatio2Dictionary
'
'    sKey = sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & sOrigProductUnit
'
'    If dictProductUnitRatio.Exists(sKey) Then
'        dblRatio = Split(dictProductUnitRatio(sKey), DELIMITER)(1)
'        fFindInConfigedReplaceProductUnit = Split(dictProductUnitRatio(sKey), DELIMITER)(0)
'    Else
'        dblRatio = 1
'        fFindInConfigedReplaceProductUnit = ""
'    End If
'End Function
'
'Function fGetProductUnit(ByVal sProductProducer As String, ByVal sProductName As String, ByVal sProductSeries As String) As String
''    Dim sKey As String
'
''    If dictProductMaster Is Nothing Then Call fReadSheetProductMaster2Dictionary
'
''    sKey = sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'
''    If Not dictProductMaster.Exists(sKey) Then
'    If Not fProductKeysExistsInProductMaster(sProductProducer, sProductName, sProductSeries) Then _
'        fErr "药品厂家+名称+规格 还不存在于药品主表中, 会计单位找不到的情况下，计算无法进行：" & vbCr & sProductProducer & vbCr & sProductName & vbCr & sProductSeries
'
'    fGetProductUnit = Split(dictProductMaster(sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries), DELIMITER)(0)
'End Function
''------------------------------------------------------------------------------
'
''====================== 1st Level Commission =================================================================
'Function fReadSheetFirstLevelComm2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("FIRST_LEVEL_COMMISSION", dictColIndex, arrData, , , , , shtFirstLevelCommission)
'    Set dictFirstLevelComm = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'            , Array(dictColIndex("SalesCompany") _
'                  , dictColIndex("ProductProducer") _
'                  , dictColIndex("ProductName") _
'                  , dictColIndex("ProductSeries")) _
'            , Array(dictColIndex("Commission")), DELIMITER, DELIMITER)
'    Set dictColIndex = Nothing
'End Function
'Function fGetFirstLevelComm(sFirstLevelKey As String, ByRef dblFirstComm As Double) As Boolean
'    If dictFirstLevelComm Is Nothing Then Call fReadSheetFirstLevelComm2Dictionary
'
'    Dim bOut As Boolean
'
'    bOut = dictFirstLevelComm.Exists(sFirstLevelKey)
'
'    dblFirstComm = 0
'    If bOut Then dblFirstComm = dictFirstLevelComm(sFirstLevelKey)
'
'    fGetFirstLevelComm = bOut
'End Function
''------------------------------------------------------------------------------
'
'
''====================== 2nd Level Commission =================================================================
'Function fReadSheetSecondLevelComm2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("SECOND_LEVEL_COMMISSION", dictColIndex, arrData, , , , , shtSecondLevelCommission)
'    Set dictSecondLevelComm = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'            , Array(dictColIndex("SalesCompany") _
'                  , dictColIndex("Hospital") _
'                  , dictColIndex("ProductProducer") _
'                  , dictColIndex("ProductName") _
'                  , dictColIndex("ProductSeries")) _
'            , Array(dictColIndex("Commission")), DELIMITER, DELIMITER)
'    Application.Run shtSelfSalesA.Range("D1") & shtSelfSalesA.Range("E1") & shtSelfSalesA.Range("F1") & shtSelfSalesA.Range("G1")
'    Set dictColIndex = Nothing
'End Function
'Function fGetSecondLevelComm(sSecondLevelCommKey As String, ByRef dblSecondComm As Double) As Boolean
'    If dictSecondLevelComm Is Nothing Then Call fReadSheetSecondLevelComm2Dictionary
'
'    If dictSecondLevelComm.Exists(sSecondLevelCommKey) Then
'        dblSecondComm = dictSecondLevelComm(sSecondLevelCommKey)
'
'        fGetSecondLevelComm = True
'    Else
'        dblSecondComm = 0
'        fGetSecondLevelComm = False
'    End If
'End Function
''------------------------------------------------------------------------------
'
'Function fGetConfigFirstLevelDefaultComm() As Double
'    If dictDefaultCommConfiged Is Nothing Then Call fReadConfigSecondLCommDefault2Dictionary
'
'    If Not dictDefaultCommConfiged.Exists("FIRST_LEVEL_COMMISSION_DEFAULT") Then fErr "配置没有设置采芝林默认配送费：FIRST_LEVEL_COMMISSION_DEFAULT"
'    fGetConfigFirstLevelDefaultComm = dictDefaultCommConfiged("FIRST_LEVEL_COMMISSION_DEFAULT")
'End Function
'
'Function fGetConfigSecondLevelDefaultComm(sSalesCompName As String) As Double
'    If dictDefaultCommConfiged Is Nothing Then Call fReadConfigSecondLCommDefault2Dictionary
'
'    Dim sCompID As String
'    Dim dblDefault As Double
'
'    'sCompID = fGetCompanyIdByCompanyName(sSalesCompName)
'    sCompID = fGetCompanyIDByName_Common(sSalesCompName)
'
'    dblDefault = dictDefaultCommConfiged("SECOND_LEVEL_COMMISSION_DEFAULT_" & sCompID)
'
'    If sCompID = "CZL" Then
'        'If dblDefault <> 0 Then fErr "这是采芝林公司，但是配送费却不是0， 请检查是否有误。"
'    Else
'        'If dblDefault = 0 Then fErr "这是" & sSalesCompName & "公司，但是配送费却是0，请检查是否有误。"
'    End If
'
'    fGetConfigSecondLevelDefaultComm = dblDefault
'End Function
'
'Function fReadConfigSecondLCommDefault2Dictionary()
'    Dim asTag As String
'    Dim arrColsName(3)
'    Dim rngToFindIn As Range
'    Dim arrConfigData()
'    Dim arrColsIndex()
'    Dim lConfigStartRow As Long
'    Dim lConfigStartCol As Long
'    Dim lConfigEndRow As Long
'    Dim lConfigHeaderAtRow As Long
'
'    asTag = "[System Misc Settings]"
'    arrColsName(1) = "Setting Item ID"
'    arrColsName(2) = "Value"
'    arrColsName(3) = "Value Type"
'
'    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtSysConf _
'                                , arrColsName:=arrColsName _
'                                , arrConfigData:=arrConfigData _
'                                , arrColsIndex:=arrColsIndex _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True)
'
'    Call fValidateDuplicateInArray(arrConfigData, 1, False, shtSysConf, lConfigHeaderAtRow, lConfigStartCol, "Setting Item ID")
'
'    Dim lEachRow As Long
'    Dim lActualRow As Long
'    Dim sKey As String
'    Dim sValueType As String
'
'    Set dictDefaultCommConfiged = New Dictionary
'
'    For lEachRow = LBound(arrConfigData, 1) To UBound(arrConfigData, 1)
'        If fArrayRowIsBlankHasNoData(arrConfigData, lEachRow) Then GoTo next_row
'
'        lActualRow = lConfigHeaderAtRow + lEachRow
'
'        sKey = Trim(arrConfigData(lEachRow, arrColsIndex(1)))
'        sValueType = Trim(arrConfigData(lEachRow, arrColsIndex(3)))
'
'        If sValueType = "GET_VALUE" Then
'            dictDefaultCommConfiged.Add sKey, arrConfigData(lEachRow, arrColsIndex(2))
'        ElseIf sValueType = "GET_ADDRESS" Then
'            dictDefaultCommConfiged.Add sKey, shtSysConf.Cells(lActualRow, lConfigStartCol + arrColsIndex(2) - 1).Address(external:=True)
'        Else
'            fErr "the Value Type cannot be blank at row " & lActualRow & vbCr & "sheet:" & shtSysConf.Name
'        End If
'next_row:
'    Next
'
'    Erase arrConfigData
'    Erase arrColsName
'    Erase arrColsIndex
'End Function
'
'Function fGetSysMiscConfig(sSettingItemID As String, Optional sMsgHeader As String = "")
'    If dictDefaultCommConfiged Is Nothing Then Call fReadConfigSecondLCommDefault2Dictionary
'
'    If Not dictDefaultCommConfiged.Exists(sSettingItemID) Then
'        fErr "[System Misc Settings] has not such config item: " & sSettingItemID & vbCr & vbCr & sMsgHeader
'    End If
'
'    fGetSysMiscConfig = dictDefaultCommConfiged(sSettingItemID)
'End Function
'
''Function fGetCompanyIdByCompanyName(sSalesCompName As String) As String
''    sSalesCompName = Trim(sSalesCompName)
''    If dictCompanyNameID Is Nothing Then Call fReadConfigCompanyList(dictCompanyNameID)
''
''    If Not dictCompanyNameID.Exists(sSalesCompName) Then fErr "公司名称不存在于商业公司名称配置块rngStaticSalesCompanyNames中，请检查。"
''
''    fGetCompanyIdByCompanyName = Trim(dictCompanyNameID(sSalesCompName))
''End Function
'
''====================== Self Sales =================================================================
'Function fReadSelfSalesOrder2Dictionary()
'    Dim sTmpKey As String
'    Dim sProducer As String, sProductName As String, sProductSeries As String
'    Dim dblSellQuantity As Double
'    Dim dblHospitalQuantity As Double
'    Dim dictDeductTo As Dictionary
'    Dim dictMinusTo As Dictionary
'    'Dim dictReverseTo As Dictionary
'    Dim lEachRow As Long
'
''    Call fSortDataInSheetSortSheetDataByFileSpec("SELF_SALES_ORDER", Array("ProductProducer" _
''                                    , "ProductName", "ProductSeries", "SalesDate"), , shtSelfSalesCal)
'
'    Call fSortDataInSheetSortSheetDataByFileSpec("SELF_SALES_ORDER", Array("ProductProducer" _
'                                    , "ProductName", "ProductSeries", "SalesDate"), , shtSelfSalesOrder)
'    Call fSortDataInSheetSortSheetDataByFileSpec("SELF_SALES_ORDER", Array("ProductProducer" _
'                                    , "ProductName", "ProductSeries", "SalesDate"), , shtSelfSalesPreDeduct)
'
'    Call fReadSheetDataByConfig("SELF_SALES_ORDER", dictSelfSalesColIndex, arrSelfSales, , , , , shtSelfSalesPreDeduct)
'
'    Set dictDeductTo = New Dictionary
'    Set dictMinusTo = New Dictionary
'    Set dictSelfSalesDeduct = New Dictionary
''    Set dictSelfSalesReverse = New Dictionary
'    Set dictSelfSalesMinus = New Dictionary
''    Set dictReverseTo = New Dictionary
'
'    For lEachRow = LBound(arrSelfSales, 1) To UBound(arrSelfSales, 1)
'        dblSellQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellQuantity"))
'        dblHospitalQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity"))
'
'        sProducer = arrSelfSales(lEachRow, dictSelfSalesColIndex("ProductProducer"))
'        sProductName = arrSelfSales(lEachRow, dictSelfSalesColIndex("ProductName"))
'        sProductSeries = arrSelfSales(lEachRow, dictSelfSalesColIndex("ProductSeries"))
'
'        If dblSellQuantity = 0 Then GoTo next_row
'
'        sTmpKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'
'        If dblSellQuantity < 0 Then
'            If dblSellQuantity > dblHospitalQuantity Then
'                fActiveVisibleSwitchSheet shtSelfSalesOrder
'                fErr "数据出错，退货的情况下，销售数量不应该大于医院抵扣数量" _
'                            & vbCr & "工作表：" & shtSelfSalesOrder.Name _
'                            & vbCr & "行号：" & lEachRow + 1 _
'                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries _
'                            & vbCr & vbCr & "请检查【" & shtSelfSalesOrder.Name & "】表。"
'            End If
'            If dblHospitalQuantity > 0 Then
'                fActiveVisibleSwitchSheet shtSelfSalesOrder
'                fErr "数据出错，退货的情况下，医院销售数量不应该 > 0" _
'                            & vbCr & "工作表：" & shtSelfSalesOrder.Name _
'                            & vbCr & "行号：" & lEachRow + 1 _
'                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries
'            End If
'
'            If dblSellQuantity < dblHospitalQuantity Then
'                If Not dictSelfSalesMinus.Exists(sTmpKey) Then
'                    dictSelfSalesMinus.Add sTmpKey, lEachRow
'                End If
'                dictMinusTo(sTmpKey) = lEachRow
'            End If
'        Else
'            If dblSellQuantity < dblHospitalQuantity Then
'                fActiveVisibleSwitchSheet shtSelfSalesOrder
'                fErr "数据出错，医院抵扣数量不应该大于销售数量" _
'                            & vbCr & "工作表：" & shtSelfSalesOrder.Name _
'                            & vbCr & "行号：" & lEachRow + 1 _
'                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries _
'                            & vbCr & vbCr & "请检查【" & shtSelfSalesOrder.Name & "】表。"
'            End If
'
'            If dblHospitalQuantity < 0 Then
'                fActiveVisibleSwitchSheet shtSelfSalesOrder
'                fErr "数据出错，医院销售数量不应该 < 0" _
'                            & vbCr & "工作表：" & shtSelfSalesOrder.Name _
'                            & vbCr & "行号：" & lEachRow + 1 _
'                            & vbCr & "药品：" & sProducer & " , " & sProductName & ", " & sProductSeries
'            End If
'
''            If dblHospitalQuantity > 0 Then
''                If Not dictSelfSalesReverse.Exists(sTmpKey) Then
''                    dictSelfSalesReverse.Add sTmpKey, lEachRow
''                End If
''                dictReverseTo(sTmpKey) = lEachRow
''            End If
'
'            If dblSellQuantity > dblHospitalQuantity Then
'                If Not dictSelfSalesDeduct.Exists(sTmpKey) Then
'                    dictSelfSalesDeduct.Add sTmpKey, lEachRow
'                End If
'                dictDeductTo(sTmpKey) = lEachRow
'            End If
'        End If
'next_row:
'    Next
'
''    For lEachRow = 0 To dictSelfSalesReverse.Count - 1
''        dictSelfSalesReverse(dictSelfSalesReverse.Keys(lEachRow)) = dictSelfSalesReverse.Items(lEachRow) _
''                    & DELIMITER & dictReverseTo.Items(lEachRow)
''    Next
'
'    For lEachRow = 0 To dictSelfSalesDeduct.Count - 1
'        dictSelfSalesDeduct(dictSelfSalesDeduct.Keys(lEachRow)) = dictSelfSalesDeduct.Items(lEachRow) _
'                    & DELIMITER & dictDeductTo.Items(lEachRow)
'    Next
'    For lEachRow = 0 To dictSelfSalesMinus.Count - 1
'        dictSelfSalesMinus(dictSelfSalesMinus.Keys(lEachRow)) = dictSelfSalesMinus.Items(lEachRow) _
'                    & DELIMITER & dictMinusTo.Items(lEachRow)
'    Next
'
'   ' Set dictSelfSalesColIndex = Nothing
'    'Set dictReverseTo = Nothing
'    Set dictMinusTo = Nothing
'    Set dictDeductTo = Nothing
'End Function
'Function fCalculateCostPriceFromSelfSalesOrder(sProductKey As String _
'                    , ByVal dblSalesQuantity As Double, ByRef dblCostPrice As Double) As Boolean
'    dblCostPrice = 0
'    If dblSalesQuantity > 0 Then
'        fCalculateCostPriceFromSelfSalesOrder = fCalculateCostPriceFromSelfSalesOrderNoraml(sProductKey, dblSalesQuantity, dblCostPrice)
'    ElseIf dblSalesQuantity < 0 Then
'        fCalculateCostPriceFromSelfSalesOrder = fCalculateCostPriceFromSelfSalesOrderMinus(sProductKey, dblSalesQuantity, dblCostPrice)
''    ElseIf dblSalesQuantity < 0 Then
''        fCalculateCostPriceFromSelfSalesOrder = fCalculateCostPriceFromSelfSalesOrderWithdraw(sProductKey, dblSalesQuantity, dblCostPrice)
'    Else
'        'fErr "销售数量为0"
'    End If
'End Function
'Function fCalculateCostPriceFromSelfSalesOrderNoraml(sProductKey As String _
'                    , ByVal dblSalesQuantity As Double, ByRef dblCostPrice As Double) As Boolean
'    Dim bOut As Boolean
'    Dim lDeductStartRow As Long
'    Dim lDeductEndRow As Long
'    Dim dblSelfSellQuantity As Double
'    Dim dblHospitalQuantity As Double
'    Dim dblToDeduct As Double
'    Dim dblBalance As Double
'    Dim dblCurrRowAvailable As Double
'    'Dim dblToDeduct As Double
'    Dim lEachRow As Long
'    Dim dblAccAmt As Double
'    Dim dblPrice As Double
'
'    If dictSelfSalesDeduct Is Nothing Then Call fReadSelfSalesOrder2Dictionary
'
'    bOut = False
'
'    If Not dictSelfSalesDeduct.Exists(sProductKey) Then GoTo exit_fun
'
'    lDeductStartRow = Split(dictSelfSalesDeduct(sProductKey), DELIMITER)(0)
'    lDeductEndRow = Split(dictSelfSalesDeduct(sProductKey), DELIMITER)(1)
'
'    dblAccAmt = 0
'    dblToDeduct = dblSalesQuantity
'    For lEachRow = lDeductStartRow To lDeductEndRow
'        If dblToDeduct <= 0 Then Exit For
'
'        dblSelfSellQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellQuantity"))
'        dblHospitalQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity"))
'        dblPrice = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellPrice"))
'
'        If dblSelfSellQuantity <= dblHospitalQuantity Then fErr "这一行的日期晚，不应该出现抵扣" _
'                        & vbCr & "工作表：" & shtSelfSalesOrder.Name _
'                        & vbCr & "行号：" & lEachRow + 1
'
'        dblCurrRowAvailable = dblSelfSellQuantity - dblHospitalQuantity
'        dblBalance = dblToDeduct - dblCurrRowAvailable
'
'        If dblBalance >= 0 Then  'still has to find next row to deduct
'            If lEachRow < lDeductEndRow Then    'move the deduct dictionary to next row
'                dictSelfSalesDeduct(sProductKey) = (lEachRow + 1) & DELIMITER & lDeductEndRow
'            Else
'                dictSelfSalesDeduct.Remove sProductKey
'            End If
'
'            dblAccAmt = dblAccAmt + dblCurrRowAvailable * dblPrice
'            arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity")) = dblSelfSellQuantity
'        Else
'            dblAccAmt = dblAccAmt + dblToDeduct * dblPrice
'            arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity")) = dblSelfSellQuantity + dblBalance
'        End If
'
''        'to extend dictreverse to new row
''        If dictSelfSalesReverse.Exists(sProductKey) Then
''            'If CLng(Split(dictSelfSalesReverse(sProductKey), DELIMITER)(1)) < lEachRow Then
''            dictSelfSalesReverse(sProductKey) = Split(dictSelfSalesReverse(sProductKey), DELIMITER)(0) & DELIMITER & lEachRow
''            'End If
''        Else
''            dictSelfSalesReverse.Add sProductKey, lEachRow & DELIMITER & lEachRow
''        End If
'
'        dblToDeduct = dblBalance
'    Next
'
'    If dblToDeduct <= 0 Then
'        bOut = True
'        dblCostPrice = dblAccAmt / dblSalesQuantity
'    End If
'
'exit_fun:
'    fCalculateCostPriceFromSelfSalesOrderNoraml = bOut
'End Function
'
'Function fCalculateCostPriceFromSelfSalesOrderMinus(sProductKey As String _
'                    , ByVal dblSalesQuantity As Double, ByRef dblCostPrice As Double) As Boolean
'    Dim bOut As Boolean
'    Dim lDeductStartRow As Long
'    Dim lDeductEndRow As Long
'    Dim dblSelfSellQuantity As Double
'    Dim dblHospitalQuantity As Double
'    Dim dblToDeduct As Double
'    Dim dblBalance As Double
'    Dim dblCurrRowAvailable As Double
'    'Dim dblToDeduct As Double
'    Dim lEachRow As Long
'    Dim dblAccAmt As Double
'    Dim dblPrice As Double
'
'    If dictSelfSalesMinus Is Nothing Then Call fReadSelfSalesOrder2Dictionary
'
'    bOut = False
'
'    If Not dictSelfSalesMinus.Exists(sProductKey) Then GoTo exit_fun
'
'    lDeductStartRow = Split(dictSelfSalesMinus(sProductKey), DELIMITER)(0)
'    lDeductEndRow = Split(dictSelfSalesMinus(sProductKey), DELIMITER)(1)
'
'    dblAccAmt = 0
'    dblToDeduct = dblSalesQuantity
'    For lEachRow = lDeductStartRow To lDeductEndRow
'        If dblToDeduct >= 0 Then Exit For
'
'        dblSelfSellQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellQuantity"))
'        dblHospitalQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity"))
'        dblPrice = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellPrice"))
'
'        If dblSelfSellQuantity <= dblHospitalQuantity Then fErr "这一行的日期晚，不应该出现抵扣" _
'                        & vbCr & "工作表：" & shtSelfSalesOrder.Name _
'                        & vbCr & "行号：" & lEachRow + 1
'
'        dblCurrRowAvailable = dblSelfSellQuantity - dblHospitalQuantity
'        dblBalance = dblToDeduct - dblCurrRowAvailable
'
'        If dblBalance <= 0 Then  'still has to find next row to deduct
'            If lEachRow < lDeductEndRow Then    'move the deduct dictionary to next row
'                dictSelfSalesMinus(sProductKey) = (lEachRow + 1) & DELIMITER & lDeductEndRow
'            Else
'                dictSelfSalesMinus.Remove sProductKey
'            End If
'
'            dblAccAmt = dblAccAmt + dblCurrRowAvailable * dblPrice
'            arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity")) = dblSelfSellQuantity
'        Else
'            dblAccAmt = dblAccAmt + dblToDeduct * dblPrice
'            arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity")) = dblSelfSellQuantity + dblBalance
'        End If
'
''        'to extend dictreverse to new row
''        If dictSelfSalesReverse.Exists(sProductKey) Then
''            'If CLng(Split(dictSelfSalesReverse(sProductKey), DELIMITER)(1)) < lEachRow Then
''            dictSelfSalesReverse(sProductKey) = Split(dictSelfSalesReverse(sProductKey), DELIMITER)(0) & DELIMITER & lEachRow
''            'End If
''        Else
''            dictSelfSalesReverse.Add sProductKey, lEachRow & DELIMITER & lEachRow
''        End If
'
'        dblToDeduct = dblBalance
'    Next
'
'    If dblToDeduct >= 0 Then
'        bOut = True
'        dblCostPrice = Abs(dblAccAmt / dblSalesQuantity)
'    End If
'
'exit_fun:
'    fCalculateCostPriceFromSelfSalesOrderMinus = bOut
'End Function
''Function fCalculateCostPriceFromSelfSalesOrderWithdraw(sProductKey As String _
''                    , ByVal dblSalesQuantity As Double, ByRef dblCostPrice As Double) As Boolean
''    If dictSelfSalesReverse Is Nothing Then Call fReadSelfSalesOrder2Dictionary
''
''    Dim bOut As Boolean
''    Dim lRevStartRow As Long
''    Dim lRevEndRow As Long
''    Dim dblSelfSellQuantity As Double
''    Dim dblHospitalQuantity As Double
''    Dim dblBalance As Double
''    Dim dblCurrRowAvailable As Double
''    Dim dblToReverse As Double
''    Dim lEachRow As Long
''    Dim dblAccAmt As Double
''    Dim dblPrice As Double
''
''    bOut = False
''
''    If Not dictSelfSalesReverse.Exists(sProductKey) Then GoTo exit_fun
''
''    lRevStartRow = Split(dictSelfSalesReverse(sProductKey), DELIMITER)(0)
''    lRevEndRow = Split(dictSelfSalesReverse(sProductKey), DELIMITER)(1)
''
''    dblAccAmt = 0
''    dblToReverse = dblSalesQuantity
''    For lEachRow = lRevEndRow To lRevStartRow Step -1
''        If dblToReverse >= 0 Then Exit For
''
''        dblSelfSellQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellQuantity"))
''        dblHospitalQuantity = arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity"))
''        dblPrice = arrSelfSales(lEachRow, dictSelfSalesColIndex("SellPrice"))
''
''        If dblHospitalQuantity <= 0 Then fErr "这一行的日期晚，医院销售数量不应该为0" _
''                        & vbCr & "工作表：" & shtSelfSalesOrder.Name _
''                        & vbCr & "行号：" & lEachRow + 1
''
''        dblCurrRowAvailable = dblHospitalQuantity
''        dblBalance = dblToReverse + dblCurrRowAvailable
''
''        If dblBalance <= 0 Then  'still has to find next row to Reverse
''            If lEachRow < lRevStartRow Then    'move the Reverse dictionary to previous row
''                dictSelfSalesReverse(sProductKey) = lRevStartRow & DELIMITER & (lEachRow - 1)
''            Else
''                dictSelfSalesReverse.Remove sProductKey
''            End If
''
''            dblAccAmt = dblAccAmt + dblCurrRowAvailable * dblPrice
''            arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity")) = 0
''        Else
''            dblAccAmt = dblAccAmt + (dblCurrRowAvailable - dblBalance) * dblPrice
''            arrSelfSales(lEachRow, dictSelfSalesColIndex("HospitalSellQuantity")) = dblBalance
''        End If
''
''        'to extend dictdeduct to new row
''        If dictSelfSalesDeduct.Exists(sProductKey) Then
''            'If CLng(Split(dictSelfSalesDeduct(sProductKey), DELIMITER)(0)) < lEachRow Then
''            dictSelfSalesDeduct(sProductKey) = lEachRow & DELIMITER & Split(dictSelfSalesDeduct(sProductKey), DELIMITER)(1)
''            'End If
''        Else
''            dictSelfSalesDeduct.Add sProductKey, lEachRow & DELIMITER & lEachRow
''        End If
''
''        dblToReverse = dblBalance
''    Next
''
''    If dblToReverse >= 0 Then
''        bOut = True
''        dblCostPrice = Abs(dblAccAmt / dblSalesQuantity)
''    End If
''
''exit_fun:
''    fCalculateCostPriceFromSelfSalesOrderWithdraw = bOut
''End Function
'Function fSetBackToshtSelfSalesCalWithDeductedData()
'    Call fDeleteRowsFromSheetLeaveHeader(shtSelfSalesPreDeduct)
'    Call fAppendArray2Sheet(shtSelfSalesPreDeduct, arrSelfSales)
'
''    If UBound(arrSelfSales, 1) >= LBound(arrSelfSales, 1) Then
''        shtSelfSalesCal.Range("A2").Resize(UBound(arrSelfSales, 1), UBound(arrSelfSales, 2)).Value2 = arrSelfSales
''    End If
'End Function
''------------------------------------------------------------------------------
''Function fGetLatestPriceFromProductMaster(sProductKey As String) As Double
''    If dictProductMaster Is Nothing Then Call fReadSheetProductMaster2Dictionary
''
''    If Not dictProductMaster.Exists(sProductKey) Then
''        fErr "药品不 存在于药品主表，前面应该已经判断过的。统一后的销售数据可能被人修改过。" & vbCr & sProductKey
''    End If
''
''    Dim sLatestPrice
''    sLatestPrice = Split(dictProductMaster(sProductKey), DELIMITER)(1)
''
''    If Len(Trim(sLatestPrice)) > 0 Then
''        If Not IsNumeric(sLatestPrice) Then fErr "药品的最新单价不是数值：" & sLatestPrice
''        fGetLatestPriceFromProductMaster = sLatestPrice
''    Else
''        fGetLatestPriceFromProductMaster = 0
''    End If
''End Function
'
'Function fGetTaxRate(sProductKey As String) As Double
'    If fProductTaxRateIsConfigured(sProductKey) Then
'        fGetTaxRate = fGetProductTaxRate(sProductKey)
'    Else
'         fGetTaxRate = CDbl(fGetSysMiscConfig("TAX_RATE"))
'    End If
'End Function
'
''====================== Salesman commssion config =================================================================
'Function fReadSalesManCommissionConfig2Dictionary()
'    Dim sTmpKey As String
'    Dim sSalesCompany As String, sHospital As String
'    Dim sProducer As String, sProductName As String, sProductSeries As String
'    Dim dblSellQuantity As Double
'    Dim dblHospitalQuantity As Double
'    Dim dictSalesManCommTo As Dictionary
'    Dim dblBidPrice As Double
'    Dim lEachRow As Long
'
'    Call fSortDataInSheetSortSheetDataByFileSpec("SALESMAN_COMMISSION_CONFIG", Array("SalesCompany", "Hospital", "ProductProducer" _
'                                    , "ProductName" _
'                                    , "ProductSeries" _
'                                    , "BidPrice"), , shtSalesManCommConfig)
'
'    Call fReadSheetDataByConfig("SALESMAN_COMMISSION_CONFIG", dictSalesManCommColIndex, arrSalesManComm, , , , , shtSalesManCommConfig)
'
'    Set dictSalesManCommTo = New Dictionary
'    Set dictSalesManCommFrom = New Dictionary
'    For lEachRow = LBound(arrSalesManComm, 1) To UBound(arrSalesManComm, 1)
'        sSalesCompany = Trim(arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesCompany")))
'        sHospital = Trim(arrSalesManComm(lEachRow, dictSalesManCommColIndex("Hospital")))
'        sProducer = Trim(arrSalesManComm(lEachRow, dictSalesManCommColIndex("ProductProducer")))
'        sProductName = Trim(arrSalesManComm(lEachRow, dictSalesManCommColIndex("ProductName")))
'        sProductSeries = Trim(arrSalesManComm(lEachRow, dictSalesManCommColIndex("ProductSeries")))
'        dblBidPrice = arrSalesManComm(lEachRow, dictSalesManCommColIndex("BidPrice"))
'
'        sTmpKey = sSalesCompany & DELIMITER & sHospital & DELIMITER _
'                & sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries & DELIMITER & dblBidPrice
'
'        If Not dictSalesManCommFrom.Exists(sTmpKey) Then
'            dictSalesManCommFrom.Add sTmpKey, lEachRow
'        Else
'            fErr "一个中标价只能有一条记录三个业务员设置！请检查业务员佣金设置表" & vbCr & sTmpKey
'        End If
'        dictSalesManCommTo(sTmpKey) = lEachRow
'next_row:
'    Next
'
'    For lEachRow = 0 To dictSalesManCommFrom.Count - 1
'        dictSalesManCommFrom(dictSalesManCommFrom.Keys(lEachRow)) = dictSalesManCommFrom.Items(lEachRow) _
'                    & DELIMITER & dictSalesManCommTo.Items(lEachRow)
'    Next
'
'    Set dictSalesManCommTo = Nothing
'End Function
'
'Function fCalculateSalesManCommissionFromshtSalesManCommConfig(sSalesManKey As String _
'                            , ByRef sSalesMan_1 As String, ByRef sSalesMan_2 As String, ByRef sSalesMan_3 As String _
'                            , ByRef dblComm_1 As Double, ByRef dblComm_2 As Double, ByRef dblComm_3 As Double _
'                            , ByRef sSalesManager As String, ByRef dblSalesMgrComm As Double) As Boolean
'    If dictSalesManCommFrom Is Nothing Then Call fReadSalesManCommissionConfig2Dictionary
'
'    Dim bOut  As Boolean
'    Dim lStartRow As Long
'    Dim lEndRow As Long
'    Dim lEachRow As Long
'    'Dim iSalesManCnt As Long
'
'    sSalesMan_1 = ""
'    sSalesMan_2 = ""
'    sSalesMan_3 = ""
'    dblComm_1 = 0
'    dblComm_2 = 0
'    dblComm_3 = 0
'
'    bOut = dictSalesManCommFrom.Exists(sSalesManKey)
'    If Not bOut Then GoTo exit_fun
'
'    lStartRow = Split(dictSalesManCommFrom(sSalesManKey), DELIMITER)(0)
'    lEndRow = Split(dictSalesManCommFrom(sSalesManKey), DELIMITER)(1)
'
''    iSalesManCnt = 0
'    For lEachRow = lStartRow To lEndRow
''        iSalesManCnt = iSalesManCnt + 1
'
''        If iSalesManCnt = 1 Then
''            sSalesMan_1 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesMan"))
''            dblComm_1 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("Commission"))
''        ElseIf iSalesManCnt = 2 Then
''            sSalesMan_2 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesMan"))
''            dblComm_2 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("Commission"))
''        ElseIf iSalesManCnt = 3 Then
''            sSalesMan_3 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesMan"))
''            dblComm_3 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("Commission"))
''        Else
''            fErr "最多只能有3个业务员，请从【业务员佣金表】中删除一个。" & vbCr & sSalesManKey & vbCr & "行号：" & lEachRow + 1
''        End If
'
'        sSalesMan_1 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesMan1"))
'        dblComm_1 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("Commission1"))
'        sSalesMan_2 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesMan2"))
'        dblComm_2 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("Commission2"))
'        sSalesMan_3 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesMan3"))
'        dblComm_3 = arrSalesManComm(lEachRow, dictSalesManCommColIndex("Commission3"))
'
'        sSalesManager = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesManager"))
'        dblSalesMgrComm = arrSalesManComm(lEachRow, dictSalesManCommColIndex("SalesManagerCommission"))
'    Next
'
'exit_fun:
'    fCalculateSalesManCommissionFromshtSalesManCommConfig = bOut
'End Function
''------------------------------------------------------------------------------
'
'Function fCheckIfProductExistsInProductMaster(arrData, iColProducer As Integer, iColProductName As Integer, iColProductSeries As Integer _
'                , Optional ByRef alErrRowNo As Long, Optional ByRef alErrColNo As Long)
'    Dim lEachRow As Long
'    Dim sProducer As String
'    Dim sProductName As String
'    Dim sProductSeries As String
''    Dim sKey As String
'
'    Call fRemoveFilterForSheet(shtProductMaster)
'
''    If dictProductMaster Is Nothing Then Call fReadSheetProductMaster2Dictionary
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sProducer = Trim(arrData(lEachRow, iColProducer))
'        sProductName = Trim(arrData(lEachRow, iColProductName))
'        sProductSeries = Trim(arrData(lEachRow, iColProductSeries))
'
''        sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'
''        If Not dictProductMaster.Exists(sKey) Then
'        If Not fProductKeysExistsInProductMaster(sProducer, sProductName, sProductSeries) Then
'            alErrRowNo = (lEachRow + 1)
'            alErrColNo = iColProductSeries
'            fErr "药品厂家+名称+规格不存在于药品主表中" & vbCr & "行号：" & (lEachRow + 1)
'            Exit For
'        End If
'    Next
'End Function
'
'Function fCheckIfProducerExistsInProducerMaster(arrData, iColProducer As Integer, Optional sErr As String = "" _
'                    , Optional lErrRowNo As Long, Optional lErrColNo As Long)
'    Dim lEachRow As Long
'    Dim sProducer As String
'
'    Call fRemoveFilterForSheet(shtProductProducerMaster)
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sProducer = Trim(arrData(lEachRow, iColProducer))
'
'        If Not fProducerExistsInProducerMaster(sProducer) Then
'            lErrRowNo = (lEachRow + 1)
'            lErrColNo = iColProducer
'            fErr IIf(fZero(sErr), "药品厂家", sErr) & "不存在于药品厂家主表中" & vbCr & "行号：" & (lEachRow + 1)
'            Exit For
'        End If
'    Next
'End Function
'
'Function fCheckIfProductNameExistsInProductNameMaster(arrData, iColProducer As Integer, iColProductName As Integer, Optional sErr As String = "" _
'                            , Optional lErrRowNo As Long, Optional lErrColNo As Long)
'    Dim lEachRow As Long
'    Dim sProducer As String
'    Dim sProductName As String
'
'    Call fRemoveFilterForSheet(shtProductNameMaster)
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sProducer = Trim(arrData(lEachRow, iColProducer))
'        sProductName = Trim(arrData(lEachRow, iColProductName))
'
'        If Not fProductNameExistsInProductNameMaster(sProducer, sProductName) Then
'            lErrRowNo = (lEachRow + 1)
'            lErrColNo = iColProductName
'            fErr IIf(fZero(sErr), "药品名称", sErr) & " 不存在于药品名称主表中" & vbCr & "行号：" & (lEachRow + 1)
'            Exit For
'        End If
'    Next
'End Function
'Function fCheckIfHospitalExistsInHospitalMaster(arrData, iColHospital As Integer, Optional sErr As String = "" _
'        , Optional ByRef alErrRowNo As Long, Optional ByRef alErrColNo As Long)
'    Dim lEachRow As Long
'    Dim sHospital As String
'
'    Call fRemoveFilterForSheet(shtHospital)
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sHospital = Trim(arrData(lEachRow, iColHospital))
'
'        If Not fHospitalExistsInHospitalMaster(sHospital) Then
'            alErrRowNo = (lEachRow + 1)
'            alErrColNo = iColHospital
'            fErr IIf(fZero(sErr), "医院", sErr) & "不存在于医院主表中" & vbCr & "行号：" & (lEachRow + 1)
'            Exit For
'        End If
'    Next
'End Function
'
'Function fResetdictNewRuleProducts()
'    Set dictNewRuleProducts = Nothing
'End Function
'
'Function fResetdictProductMaster()
'    Set dictProductMaster = Nothing
'End Function
'
'Function fResetdictProductNameMaster()
'    Set dictProductNameMaster = Nothing
'End Function
'
'Function fResetdictProducerMaster()
'    Set dictProducerMaster = Nothing
'End Function
'Function fResetdictHospitalMaster()
'    Set dictHospitalMaster = Nothing
'End Function
'
'Function fResetdictSalesManMaster()
'    Set dictSalesManMaster = Nothing
'End Function
'
''====================== SalesMan Master =================================================================
'Function fReadSheetSalesManMaster2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("SALESMAN_MASTER_SHEET", dictColIndex, arrData, , , , , shtSalesManMaster)
'    Set dictSalesManMaster = fReadArray2DictionaryOnlyKeys(arrData, dictColIndex("SalesManName"), True, False)
'
'    Set dictColIndex = Nothing
'End Function
'Function fSalesManExistsInSalesManMaster(sSalesMan As String) As Boolean
'    If Len(sSalesMan) <= 0 Then fSalesManExistsInSalesManMaster = True: Exit Function
'    If dictSalesManMaster Is Nothing Then Call fReadSheetSalesManMaster2Dictionary
'
'    fSalesManExistsInSalesManMaster = dictSalesManMaster.Exists(sSalesMan)
'End Function
''------------------------------------------------------------------------------
'
'Function fCheckIfSalesManExistsInSalesManMaster(arrData, iColSalesMan As Integer, Optional sErr As String = "", Optional lErrRowNo As Long, Optional lErrColNo As Long)
'    Dim lEachRow As Long
'    Dim sSalesMan As String
'
'    Call fRemoveFilterForSheet(shtSalesManMaster)
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sSalesMan = Trim(arrData(lEachRow, iColSalesMan))
'
'        If Not fSalesManExistsInSalesManMaster(sSalesMan) Then
'            lErrRowNo = (lEachRow + 1)
'            lErrColNo = iColSalesMan
'            fErr IIf(fZero(sErr), "业务员", sErr) & "不存在于业务员主表中" & vbCr & "行号：" & (lEachRow + 1)
'            Exit For
'        End If
'    Next
'End Function
'
'Function fReadExcludeProductListConfig2Dictionary()
'    Dim asTag As String
'    Dim arrColsName(3)
'    Dim rngToFindIn As Range
'    Dim arrConfigData()
'    Dim arrColsIndex()
'    Dim lConfigStartRow As Long
'    Dim lConfigStartCol As Long
'    Dim lConfigEndRow As Long
'    Dim lConfigHeaderAtRow As Long
'
'    asTag = "[Excluding Product List]"
'    arrColsName(1) = "Product Producer"
'    arrColsName(2) = "Product Name"
'    arrColsName(3) = "Product Series"
'
'    Call fReadConfigBlockToArray(asTag:=asTag, shtParam:=shtStaticData _
'                                , arrColsName:=arrColsName _
'                                , arrConfigData:=arrConfigData _
'                                , arrColsIndex:=arrColsIndex _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True)
'
'   ' Call fValidateDuplicateInArray(arrConfigData, Array(arrColsIndex(1), arrColsIndex(2), arrColsIndex(3)), False, shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "药品厂家 + 名称 + 规格")
'    Call fValidateBlankInArray(arrConfigData, arrColsIndex(1), shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "药品厂家")
'    Call fValidateBlankInArray(arrConfigData, arrColsIndex(2), shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "药品名称")
'    Call fValidateBlankInArray(arrConfigData, arrColsIndex(3), shtStaticData, lConfigHeaderAtRow, lConfigStartCol, "药品规格")
'
'    Set dictExcludeProducts = fReadArray2DictionaryMultipleKeysWithKeysOnly(arrConfigData _
'                                , Array(arrColsIndex(1), arrColsIndex(2), arrColsIndex(3)) _
'                                , DELIMITER, , False)
'End Function
'
'Function fProductExistsInExcludingProductListConfig(sProductProducer As String, sProductName As String, sProductSeries As String) As Boolean
'    If dictExcludeProducts Is Nothing Then Call fReadExcludeProductListConfig2Dictionary
'
'    fProductExistsInExcludingProductListConfig = dictExcludeProducts.Exists(sProductProducer & DELIMITER & sProductName & DELIMITER & sProductSeries)
'End Function
'
'Function fCheckIfLotNumExistsInSelfPurchaseOrder(arrData, iColProducer As Integer, iColProductName As Integer, iColProductSeries As Integer _
'                                        , iColLotNum As Integer, Optional ByRef alErrRowNo As Long, Optional ByRef alErrColNo As Long)
'    Dim lEachRow As Long
'    Dim sProducer As String
'    Dim sProductName As String
'    Dim sProductSeries As String
'    Dim sLotNum As String
''    Dim sKey As String
'
'    Call fRemoveFilterForSheet(shtSelfPurchaseOrder)
'
''    If dictProductMaster Is Nothing Then Call fReadSheetProductMaster2Dictionary
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sProducer = Trim(arrData(lEachRow, iColProducer))
'        sProductName = Trim(arrData(lEachRow, iColProductName))
'        sProductSeries = Trim(arrData(lEachRow, iColProductSeries))
'        sLotNum = Trim(arrData(lEachRow, iColLotNum))
''        sKey = sProducer & DELIMITER & sProductName & DELIMITER & sProductSeries
'
''        If Not dictProductMaster.Exists(sKey) Then
'        If Not fLotNumExistsInSelfPurchaseOrder(sProducer, sProductName, sProductSeries, sLotNum) Then
'            alErrRowNo = (lEachRow + 1)
'            alErrColNo = iColLotNum
'            fErr "【药品 + 批号】不存在于本公司进货表中" & vbCr & "行号：" & (lEachRow + 1)
'            Exit For
'        End If
'    Next
'End Function
'
''====================== ProductSeries Master =================================================================
'Function fReadSheetNewRuleProducts2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("NEW_RULE_PRODUCTS_CONFIG", dictColIndex, arrData, , , , , shtNewRuleProducts)
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")), False, shtNewRuleProducts, 1, 1, "厂家 + 名称 + 规格")
'
'    Set dictNewRuleProducts = fReadArray2DictionaryMultipleKeysWithMultipleColsCombined(arrData _
'                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName"), dictColIndex("ProductSeries")) _
'                                    , Array(dictColIndex("SalesTaxRate"), dictColIndex("PurchaseTaxRate")), DELIMITER, DELIMITER)
'    Set dictColIndex = Nothing
'End Function
'Function fIsNewRuleProduct(sProductKey As String) As Boolean
'    If dictNewRuleProducts Is Nothing Then fReadSheetNewRuleProducts2Dictionary
'    fIsNewRuleProduct = dictNewRuleProducts.Exists(sProductKey)
'End Function
'Function fGetNewRuleProductTaxRate(sProductKey As String, ByRef dblNewRSalesTaxRate As Double, ByRef dblNewRPurchaseTaxRate As Double)
'    If dictNewRuleProducts Is Nothing Then fReadSheetNewRuleProducts2Dictionary
'
'    If dictNewRuleProducts.Exists(sProductKey) Then
'        dblNewRSalesTaxRate = Split(dictNewRuleProducts(sProductKey), DELIMITER)(0)
'        dblNewRPurchaseTaxRate = Split(dictNewRuleProducts(sProductKey), DELIMITER)(1)
'    Else
'        dblNewRSalesTaxRate = 0
'        dblNewRPurchaseTaxRate = 0
'    End If
'End Function
''------------------------------------------------------------------------------
'
'Function fCheckIfCompanyNameExistsInrngStaticSalesCompanyNames(arrData, iColProducer As Integer, Optional sErr As String = "" _
'                    , Optional lErrRowNo As Long, Optional lErrColNo As Long)
'    Dim lEachRow As Long
'    Dim sCompanyName As String
'
'    For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
'        sCompanyName = Trim(arrData(lEachRow, iColProducer))
'
'        'If Not fCompanyNameExistsInrngStaticSalesCompanyNames(sCompanyName) Then
'        If Not fCompanyNameExists(sCompanyName) Then
'            lErrRowNo = (lEachRow + 1)
'            lErrColNo = iColProducer
'            fErr IIf(fZero(sErr), "商业公司名称", sErr) & "不存在于商业公司名称配置块rngStaticSalesCompanyNames中" & vbCr & "行号：" & (lEachRow + 1)
'            Exit For
'        End If
'    Next
'End Function
'
''====================== CompanyName Replacement =================================================================
'Function fReadSheetCompanyNameReplace2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("COMPANY_NAME_REPLACE_SHEET", dictColIndex, arrData, , , , , shtCompanyNameReplace)
'    Set dictCompanyNameReplace = fReadArray2DictionaryWithSingleCol(arrData, dictColIndex("FromCompanyName"), dictColIndex("ToCompanyName"))
'
'    Set dictColIndex = Nothing
'End Function
'Function fFindInConfigedReplaceCompanyName(sCompanyName As String) As String
'    If dictCompanyNameReplace Is Nothing Then Call fReadSheetCompanyNameReplace2Dictionary
'
'    If dictCompanyNameReplace.Exists(sCompanyName) Then
'        fFindInConfigedReplaceCompanyName = dictCompanyNameReplace(sCompanyName)
'    Else
'        fFindInConfigedReplaceCompanyName = ""
'    End If
'End Function
''------------------------------------------------------------------------------
'
''====================== Promotion Product List =================================================================
'Function fReadSheetPromotionProducts2Dictionary()
'    Dim arrData()
'    Dim dictColIndex As Dictionary
'
'    Call fReadSheetDataByConfig("PROMOTION_PRODUCTS_CONFIG", dictColIndex, arrData, , , , , shtPromotionProduct)
'    Call fValidateDuplicateInArray(arrData, Array(dictColIndex("Hospital"), dictColIndex("ProductProducer") _
'                , dictColIndex("ProductName"), dictColIndex("ProductSeries"), dictColIndex("SalesPrice")) _
'                , False, shtPromotionProduct, 1, 1, "医院 + 药品生产厂家 + 药品名称 + 规格 + 中标价")
'
'    Set dictPromotionProducts = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrData _
'                                    , Array(dictColIndex("ProductProducer"), dictColIndex("ProductName") _
'                                    , dictColIndex("ProductSeries"), dictColIndex("SalesPrice"), dictColIndex("Hospital")) _
'                                    , dictColIndex("Rebate"), DELIMITER)
'    Set dictColIndex = Nothing
'End Function
'Function fIsPromotionProduct(sHospital As String, sProductKey As String, dblSalesPrice As Double) As Boolean
'    Dim sKey As String
'
'    If dictPromotionProducts Is Nothing Then fReadSheetPromotionProducts2Dictionary
'
'    sKey = sProductKey & DELIMITER & CStr(dblSalesPrice) & DELIMITER & sHospital
'    If dictPromotionProducts.Exists(sKey) Then
'        fIsPromotionProduct = True
'    Else
'        sKey = sProductKey & DELIMITER & CStr(dblSalesPrice) & DELIMITER & ""
'        fIsPromotionProduct = dictPromotionProducts.Exists(sKey)
'    End If
'End Function
'Function fGetPromotionProductRebate(sHospital As String, sProductKey As String, dblSalesPrice As Double) As Double
'    Dim sKey As String
'
'    If dictPromotionProducts Is Nothing Then fReadSheetPromotionProducts2Dictionary
'
'    sKey = sProductKey & DELIMITER & CStr(dblSalesPrice) & DELIMITER & sHospital
'    If dictPromotionProducts.Exists(sKey) Then
'        GoTo rtn_fun
'    Else
'        sKey = sProductKey & DELIMITER & CStr(dblSalesPrice) & DELIMITER & ""
'        If dictPromotionProducts.Exists(sKey) Then GoTo rtn_fun
'    End If
'
'    fGetPromotionProductRebate = 0
'    Exit Function
'rtn_fun:
'    If IsNumeric(dictPromotionProducts(sKey)) Then
'        fGetPromotionProductRebate = CDbl(dictPromotionProducts(sKey))
'    Else
'        fGetPromotionProductRebate = 0
'    End If
'End Function
''------------------------------------------------------------------------------
'
''====================== Product TaxRate List =================================================================
'Function fReadSheetProductTaxRate2Dictionary()
'    Dim arrData()
'
'    Call shtProductTaxRate.fValidateSheet(False)
'    Call fCopyReadWholeSheetData2Array(shtProductTaxRate, arrData)
'
'    Set dictProdTaxRate = fReadArray2DictionaryWithMultipleKeyColsSingleItemCol(arrData _
'                                    , Array(ProdTaxRate.ProductProducer, ProdTaxRate.ProductName, ProdTaxRate.ProductSeries) _
'                                    , ProdTaxRate.TaxRate, DELIMITER)
'    Erase arrData
'End Function
'Function fProductTaxRateIsConfigured(sProductKey As String) As Boolean
'    If dictProdTaxRate Is Nothing Then fReadSheetProductTaxRate2Dictionary
'    fProductTaxRateIsConfigured = dictProdTaxRate.Exists(sProductKey)
'End Function
'Function fGetProductTaxRate(sProductKey As String) As Double
'    If dictProdTaxRate Is Nothing Then fReadSheetProductTaxRate2Dictionary
'
'    If dictProdTaxRate.Exists(sProductKey) Then
'        If IsNumeric(dictProdTaxRate(sProductKey)) Then
'            fGetProductTaxRate = dictProdTaxRate(sProductKey)
'        Else
'            fGetProductTaxRate = 0
'        End If
'    Else
'        fGetProductTaxRate = 0
'    End If
'End Function
''------------------------------------------------------------------------------
'
'
'
'
''=============================================================
'Private Function fReadSelfSalesForAllPrices()
'    Dim i As Long
'    Dim sKeyStr As String
'    Dim dblPrice As Double
'    Dim lMaxRow As Long
'    Dim lMaxCol As Long
'
'    Call fSortDataInSheetSortSheetData(shtSelfSalesPreDeduct, Array(SelfSales.ProductProducer _
'                                                                 , SelfSales.ProductName _
'                                                                 , SelfSales.ProductSeries _
'                                                                 , SelfSales.SellDate))
'    lMaxRow = fGetValidMaxRow(shtSelfSalesPreDeduct)
'    lMaxCol = fGetValidMaxCol(shtSelfSalesPreDeduct)
'
'    Dim arrData()
'    Call fCopyReadWholeSheetData2Array(shtSelfSalesPreDeduct, arrData)
'
'    Set dictSelfSalesPrice = New Dictionary
'
'    For i = UBound(arrData, 1) To LBound(arrData, 1) Step -1
'        sKeyStr = Trim(arrData(i, SelfSales.ProductProducer)) & DELIMITER _
'                & Trim(arrData(i, SelfSales.ProductName)) & DELIMITER _
'                & Trim(arrData(i, SelfSales.ProductSeries))
'
'        If fZero(Replace(sKeyStr, DELIMITER, "")) Then GoTo next_row
'
'        dblPrice = arrData(i, SelfSales.SellPrice)
'
'        If dictSelfSalesPrice.Exists(sKeyStr) Then
'            If InStr("~" & dictSelfSalesPrice(sKeyStr) & "~", "~" & dblPrice & "~") <= 0 Then
'                dictSelfSalesPrice(sKeyStr) = dictSelfSalesPrice(sKeyStr) & "~" & CStr(dblPrice)
'            End If
'        Else
'            dictSelfSalesPrice.Add sKeyStr, CStr(dblPrice)
'        End If
'next_row:
'    Next
'    Erase arrData
'End Function
'Function GetAvailableSelfSalesPrices(sProductKey As String) As String
'    Dim sPrices As String
'
'    If dictSelfSalesPrice Is Nothing Then Call fReadSelfSalesForAllPrices
'
'    If dictSelfSalesPrice.Exists(sProductKey) Then
'        sPrices = dictSelfSalesPrice(sProductKey)
'    Else
'        sPrices = ""
'    End If
'
'    GetAvailableSelfSalesPrices = sPrices
'End Function
''-----------------------------------------------------------------------------------
'
