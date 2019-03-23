Attribute VB_Name = "Common_Local"
Option Explicit
Option Base 1
 

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

    rng.Parent.UsedRange.value = rng.Parent.UsedRange.value
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
    Call fExcelFileOpenedToCloseIt(sExcelFileFullPath)
    
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
        , Destination:=shtTo.Range("$A$1"))
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
    
    If fZero(rgTarget.value) Then
        fErr "asTableLevelConf cannot be blank in " & shtParent.Name & vbCr & "range: " & rngConfigBlock.Address
    End If
    
    fGetTableLevelConfig = Trim(rgTarget.value)
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
    
    If Left(aValue, 1) = "гд" Then
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
    shtOutput.Range("A1").Resize(1, lMaxCol).value = arrHeader
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
            shtOutput.Columns(dictRptColIndex.Items(i)).Delete Shift:=xlToLeft
        End If
next_col:
    Next
End Function

Function fSaveWorkBookNotClose(wb As Workbook)
    wb.CheckCompatibility = False
    wb.Save
    wb.CheckCompatibility = True
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
    shtOutput.UsedRange.Delete Shift:=xlUp
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
                    arrData(lEachRow, 1) = Format(arrData(lEachRow, 1), sFormat)
                Next
                
                shtOutput.Cells(lRowFrom, lEachCol).Resize(UBound(arrData, 1), 1).value = arrData
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
            sKeyColsFormula = sKeyColsFormula & ","
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

Function fSetConditionFormatForBorder(ByRef shtParam As Worksheet, Optional lMaxCol As Long = 0 _
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
            sKeyColsFormula = sKeyColsFormula & ","
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
    
'    sAddr = fGetSysMiscConfig("REPORT_HEADER_LINE_COLOR")
'    lColor = fGetRangeFromExternalAddress(sAddr).Interior.Color
    
    With rgTarget
        .Font.Bold = True
        .Interior.Color = 8696052 'lColor
        
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
        
        If Err.Number = vbObjectError + CONFIG_ERROR_NUMBER Then
        Else
            fMsgBox "Error has occurred:" _
                    & vbCr & vbCr _
                    & "Error Number: " & Err.Number & vbCr _
                    & "Error Description:" & Err.Description
        End If
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
    
    shtParam.Cells(alHeaderAtRow, 1).Resize(1, iV).value = arrHeaderHorizontal
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

Function fGetValue(sRgName As String)
    fGetValue = ThisWorkbook.Worksheets(1).Range(sRgName).value
End Function
Function fSetValue(sRgName As String, aValue)
    ThisWorkbook.Worksheets(1).Range(sRgName).value = aValue
    ThisWorkbook.Save
End Function
