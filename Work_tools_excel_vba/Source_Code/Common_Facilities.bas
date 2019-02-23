Attribute VB_Name = "Common_Facilities"
Option Explicit
Option Base 1

'Dim dblStartTime As Double
Dim bCopyType As String
Dim sCopyStageRangeAddr As String
'Dim sCopiedAddress As String

Sub sub_ResetOnError_Initialize()
    Err.Clear
     
    'On Error GoTo err_exit
     
    Call fEnableExcelOptionsAll
     
      
err_exit:
    Err.Clear
    End
End Sub
'
'Function fStartTimer()
'    dblStartTime = Timer
'End Function
'Function fHowLong(Optional ByVal asPrefix As String = "")
'    Debug.Print IIf(Len(asPrefix) > 0, asPrefix & vbTab, "") & format(Timer - dblStartTime, "0000.00000000")
'End Function

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
        shtDataStage.Range(PASTE_START_CELL).value = sCopiedStr
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
        myData.SetText CStr(shtDataStage.Range(PASTE_START_CELL).value)
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
    
    If Workbooks.Count <= 0 Then Exit Sub
    
    For Each obj In ActiveSheet.DrawingObjects
        sStr = sStr & vbCr & obj.Name
    Next
     
    Set obj = Nothing
    
    MsgBox sStr
End Sub

Sub sub_ExportSourceCodeToFolder(Optional wb As Workbook)
    Dim wbIsnothing   As Boolean
    Dim sFolder As String
'    Dim sMsg As String
    'Dim iCnt As Long
'    Dim vbProj As VBIDE.VBProject
'    Dim vbComp As VBIDE.VBComponent

    On Error GoTo clear_exit
    
    wbIsnothing = CBool(wb Is Nothing)
    If wbIsnothing Then
        Set wb = ActiveWorkbook
    End If
    
    Call fWorkbookVBProjectIsProteced(wb)
    
    If Len(wb.Path) <= 0 Then GoTo clear_exit
    
'    Set vbProj = wb.VBProject
    sFolder = wb.Path & "\" & "Source_Code\"
        
    Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sFolder)
    Call fExportSourceCodeToFolder(sFolder, wb)
'    iCnt = vbProj.VBComponents.Count
    
'    fGetFSO
'
'    sMsg = sMsg & vbCr & vbCr & sFolder
'
'    If Not gFSO.FolderExists(sFolder) Then Call gFSO.CreateFolder(sFolder)
    
    'call fCheckPath(sfolder, true)
    'Call fDeleteAllFilesFromFolder(sFolder)
    
'    iCnt = 0
'    For Each vbComp In vbProj.VBComponents
'        'If UCase(vbcomp.Name) Like "SHEET*" Then GoTo Next_mod
'        'If vbComp.Type = 1 Or vbComp.Type = 3 Or vbComp.Type = 100 Then
'            vbComp.Export sFolder & vbComp.Name & ".bas"
'        'End If
'Next_mod:
'    Next
    
'    Set vbComp = Nothing
'    Set vbProj = Nothing
    
    MsgBox "Done!"
clear_exit:
    If wbIsnothing Then Set wb = Nothing
End Sub

Function fListAllFunctionsInLocalFile(shtOutput As Worksheet)
'    Dim sFileProp As String
'    Dim iLeftFileNum As Integer
'    Dim iRightFileNum As Integer
'    Dim lLeftFileRowCnt As Long
'    Dim lRightFileRowCnt As Long
    
    Dim dictFunsInFile As Dictionary
    Dim arrFileLines
    Dim iCnt As Long
    
    
    On Error GoTo error_handling
    
    Call fInitialization
    
    Call fBackupTextFileWithDefaultFileName(SOURCE_CODE_LIBRARY_FILE)
    Call fAppendBlankLineToTheEndOfTextFile(SOURCE_CODE_LIBRARY_FILE)
    Call fTrimTrailingBlanksForTextFile(SOURCE_CODE_LIBRARY_FILE)
    Call fDeleteMultipleBlankLinesFromTextFile(SOURCE_CODE_LIBRARY_FILE)
    arrFileLines = fReadTextFileAllLinesToArray(SOURCE_CODE_LIBRARY_FILE)
    Set dictFunsInFile = fGetAllFunctionListFromFile(arrFileLines, SOURCE_CODE_LIBRARY_FILE)

    Dim arrData()
    arrData = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictFunsInFile)

    Call fWriteArray2Sheet(shtOutput, arrData, 1)
    
    Erase arrData
error_handling:
    Erase arrFileLines
    Set dictFunsInFile = Nothing
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Function

Sub subMain_ListAllFunctionsInFileLibrary()
    Dim wbTmp As Workbook
    Set wbTmp = Workbooks.Add
    
    Dim shtOutput As Worksheet
    
    If Not fAddTmpSheetInWorkbook(shtOutput, wbTmp) Then GoTo exit_sub
    
    Call fListAllFunctionsInLocalFile(shtOutput)
exit_sub:
    Set shtOutput = Nothing
    Set wbTmp = Nothing
End Sub

Sub sub_ExportSourceCodeToFolder_ThisWorkbook()
     Call sub_ExportSourceCodeToFolder(ThisWorkbook)
End Sub
Sub sub_ListAllFunctionsOfActiveWorkbook()
    If Workbooks.Count <= 0 Then Exit Sub
    If Len(ActiveWorkbook.Path) <= 0 Then Exit Sub

    Dim shtOutput As Worksheet
    If Not fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtOutput) Then Exit Sub
    
    Dim arrModules()
    Dim arrFunctions()
    
    arrModules = fGetListAllModulesOfWorkbook()
    arrFunctions = fGetListAllSubFunctionsInWorkbook(arrModules)
    
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
    
    shtMenu.Range("A1").value = sNotifi
    shtMenu.Range("A1").Font.size = iFontSize
    shtMenu.Range("A1").Font.Color = iColor
    shtMenu.Range("A1").Font.Bold = bBold
    
    shtMenuCompInvt.Range("A1").value = sNotifi
    shtMenuCompInvt.Range("A1").Font.size = iFontSize
    shtMenuCompInvt.Range("A1").Font.Color = iColor
    shtMenuCompInvt.Range("A1").Font.Bold = bBold
End Function

'*************************************************************************

Function fGetListAllModulesOfWorkbook(Optional wb As Workbook) As Variant
    Dim arrOut()
    Dim iCnt As Long
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    
    If wb Is Nothing Then
        Set vbProj = ActiveWorkbook.VBProject
    Else
        Set vbProj = wb.VBProject
    End If
    
    iCnt = vbProj.VBComponents.Count
    ReDim arrOut(1 To iCnt, 3)
    
    iCnt = 0
    For Each vbComp In vbProj.VBComponents
        iCnt = iCnt + 1
        arrOut(iCnt, 1) = "Modules"
        arrOut(iCnt, 2) = fVBEComponentTypeToString(vbComp.Type)
        arrOut(iCnt, 3) = vbComp.Name
    Next
    
    fGetListAllModulesOfWorkbook = arrOut
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

Function fGetListAllSubFunctionsInWorkbook(arrModules(), Optional wb As Workbook) As Variant
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
    
    If wb Is Nothing Then
        Set vbProj = ActiveWorkbook.VBProject
    Else
        Set vbProj = wb.VBProject
    End If
    
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
    fGetListAllSubFunctionsInWorkbook = arrOut
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
Function fAddTmpSheetInWorkbook(shtTmp As Worksheet, Optional wb As Workbook) As Boolean
    fAddTmpSheetInWorkbook = fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtTmp, wb)
End Function
Function fGetTmpSheetInWorkbookWhenNotExistsCreateIt(shtTmp As Worksheet, Optional wb As Workbook) As Boolean
    Dim sTmp As String
    Dim response As VbMsgBoxResult
    
    If wb Is Nothing Then Set wb = ActiveWorkbook
    
    sTmp = "tmpOutput"
    
    If fSheetExists(sTmp, , wb) Then
        wb.Worksheets(sTmp).Activate
        
        response = MsgBox("There is an existing sheet " & sTmp & ", to delete it, please press yes" _
                    & vbCr, vbCritical + vbYesNoCancel)
        If response = vbNo Then
            Set shtTmp = wb.Worksheets(sTmp)
        ElseIf response = vbYes Then    'vbYes
            Call fDeleteSheet(sTmp)
            Set shtTmp = fAddNewSheet(sTmp, wb)
        Else
            fGetTmpSheetInWorkbookWhenNotExistsCreateIt = False
            Exit Function
        End If
    Else
        Set shtTmp = fAddNewSheet(sTmp, wb)
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
        
        gsEnv = rgIntersect.value
        
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
    
    If Selection.Rows.Count > 1 Then
        maxNum = Selection.Rows.Count
    Else
        maxNum = InputBox("How many letters to you want to generate? (either number or letter is ok, e.g., 20 or AF)", "Max Number letter")
    End If
    
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
    
    If lMax <= 0 Then
        fMsgBox "the number you input must be greater than 0 "
        Exit Sub
    End If
    
    Dim i As Long
    ReDim arrList(1 To lMax, 1)
    For i = 1 To lMax
        arrList(i, 1) = fNum2Letter(i)
    Next
    
    If Selection.Rows.Count > 1 Then
        Selection.value = arrList
    Else
        ActiveCell.Resize(UBound(arrList, 1), 1).value = arrList
    End If
    Erase arrList
End Sub

Sub sub_GenNumberList()
    Dim maxNum
    Dim lMax As Long
    Dim sMaxcol As String
    Dim arrList()
    
    If Not fPromptToOverWrite() Then Exit Sub
    
    If Selection.Rows.Count > 1 Then
        maxNum = Selection.Rows.Count
    Else
        maxNum = InputBox("How many letters to you want to generate? ( e.g., 20 , 100)", "Max Number")
    End If
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
    
    If Selection.Rows.Count > 1 Then
        Selection.value = arrList
    Else
        ActiveCell.Resize(UBound(arrList, 1), 1).value = arrList
    End If
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
'            .Width = 720 'contains 30×20 icons
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
    aActiveCellValue = ActiveCell.value
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
        fMsgBox "²»ÄÜÑ¡¶àÐÐ£¬Ö»ÄÜÑ¡Ò»ÐÐ¡£"
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
            
            If IsNumeric(eachCell.value) Then
                rgData.AutoFilter Field:=eachCell.Column _
                                , Criteria1:=eachCell.value _
                                , Operator:=xlAnd
            Else
                rgData.AutoFilter Field:=eachCell.Column _
                                , Criteria1:="=*" & eachCell.value & "*" _
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
    sSelectContent = ActiveCell.value
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
'        fMsgBox "²»ÄÜÑ¡¶àÐÐ£¬Ö»ÄÜÑ¡Ò»ÐÐ¡£"
'        End
'    End If
    
    sFirstValue = rngSelected.Cells(1, 1).value

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
    
    Debug.Print rngSelected.Cells(1, 1).value
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
    If Workbooks.Count <= 0 Then Exit Sub
    
    Dim shtEach As Worksheet
    Dim strAll As String
    
    For Each shtEach In ActiveWorkbook.Worksheets
        If Len(strAll) <= 0 Then
            strAll = shtEach.Name
        Else
            strAll = strAll & vbCr & shtEach.Name
        End If
    Next
    
    Set shtEach = Nothing
    
    MsgBox strAll
End Sub

Sub subMain_ListAllSheetsCodeName()
    If Workbooks.Count <= 0 Then Exit Sub
    
    Dim shtEach As Worksheet
    Dim strAll As String
    
    For Each shtEach In ActiveWorkbook.Worksheets
        If Len(strAll) <= 0 Then
            strAll = shtEach.CodeName
        Else
            strAll = strAll & vbCr & shtEach.CodeName
        End If
    Next
    
    Set shtEach = Nothing
    MsgBox strAll
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

Function fGetSavedValue(sRange As String)
    fGetSavedValue = ThisWorkbook.Worksheets(1).Range(sRange).value
End Function

Function fSetSavedValue(sRange As String, aValue)
    ThisWorkbook.Worksheets(1).Range(sRange).value = aValue
End Function
