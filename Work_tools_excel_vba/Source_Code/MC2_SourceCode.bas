Attribute VB_Name = "MC2_SourceCode"
Option Explicit
Option Base 1

Sub subMain_ValidateMacroWithLocal(Optional wb As Workbook)
    Dim sFileProp As String
    Dim iLeftFileNum As Integer
    Dim iRightFileNum As Integer
    Dim lLeftFileRowCnt As Long
    Dim lRightFileRowCnt As Long
    
    Dim dictFunsInFile As Dictionary
    Dim arrFileLines
    Dim iCnt As Long
    
    Dim vbP As VBIDE.VBProject
    Dim vbcomp As VBIDE.VBComponent
    Dim codeM As VBIDE.CodeModule
    
    On Error GoTo error_handling
    
    Call fInitialization
    
    Call fBackupTextFileWithDefaultFileName(SOURCE_CODE_LIBRARY_FILE)
    Call fAppendBlankLineToTheEndOfTextFile(SOURCE_CODE_LIBRARY_FILE)
    Call fTrimTrailingBlanksForTextFile(SOURCE_CODE_LIBRARY_FILE)
    Call fDeleteMultipleBlankLinesFromTextFile(SOURCE_CODE_LIBRARY_FILE)
    arrFileLines = fReadTextFileAllLinesToArray(SOURCE_CODE_LIBRARY_FILE)
    Set dictFunsInFile = fGetAllFunctionListFromFile(arrFileLines, SOURCE_CODE_LIBRARY_FILE)

    If wb Is Nothing Then
        If Len(ActiveWorkbook.Path) <= 0 Then GoTo error_handling
        Set vbP = ActiveWorkbook.VBProject
    Else
        If Len(wb.Path) <= 0 Then GoTo error_handling
        Set vbP = wb.VBProject
    End If
        
    Dim dictFunsInMacro  As Dictionary
    Dim sFunName As String
    Dim sModuleType As String
    Dim j As Long
    Dim k As Long
    Dim lStartLine As Long
    Dim lEndLine As Long
    Dim lFileStartLine As Long
    Dim lFileEndLine As Long
        
    iLeftFileNum = FreeFile
    Open COMPARE_TMP_FILE_LEFT For Output As #iLeftFileNum
    
    iRightFileNum = FreeFile
    Open COMPARE_TMP_FILE_RIGHT For Output As #iRightFileNum
    
    lLeftFileRowCnt = 0
    lRightFileRowCnt = 0
    
    For Each vbcomp In vbP.VBComponents
        sModuleType = fGetComponentTypeToString(vbcomp.Type)
        
        If sModuleType = "UserForm" Then GoTo next_comp
        
        Set codeM = vbcomp.CodeModule
        
        Set dictFunsInMacro = fGetAllSubFunctionsOfAModule(vbcomp)
        
        For j = 0 To dictFunsInMacro.Count - 1
            sFunName = dictFunsInMacro.Keys(j)
            lStartLine = Split(dictFunsInMacro.Items(j), DELIMITER)(2)
            lEndLine = Split(dictFunsInMacro.Items(j), DELIMITER)(3)
            
            If dictFunsInFile.Exists(sFunName) Then
                sFileProp = dictFunsInFile(sFunName)
                
                If InStr(sFileProp, "|") > 0 Then
                    sFileProp = Split(sFileProp, "|")(0)
                    MsgBox sFunName * "  in the file libarary has more than one instance, please click the other button to validate the file first"
                End If
                
                lFileStartLine = Split(sFileProp, ",")(0)
                lFileEndLine = Split(sFileProp, ",")(1)
                
                If Not fMacroSectionIsSameAsTextFile(codeM, lStartLine, lEndLine, arrFileLines, lFileStartLine, lFileEndLine) Then
                    For k = lStartLine To lEndLine
                        Print #iLeftFileNum, codeM.Lines(k, 1)
                        lLeftFileRowCnt = lLeftFileRowCnt + 1
                    Next
                    For k = lFileStartLine To lFileEndLine
                        Print #iRightFileNum, arrFileLines(k)
                        lRightFileRowCnt = lRightFileRowCnt + 1
                    Next
                    
                    iCnt = iCnt + 1
                    
                    For k = 1 To lLeftFileRowCnt - lRightFileRowCnt
                        Print #iRightFileNum, ""
                        lRightFileRowCnt = lRightFileRowCnt + 1
                    Next
                    For k = 1 To lRightFileRowCnt - lLeftFileRowCnt
                        Print #iLeftFileNum, ""
                        lLeftFileRowCnt = lLeftFileRowCnt + 1
                    Next
                End If
            Else
                For k = lStartLine To lEndLine
                    Print #iLeftFileNum, codeM.Lines(k, 1)
                    lLeftFileRowCnt = lLeftFileRowCnt + 1
                Next
                For k = lStartLine To lEndLine
                    Print #iRightFileNum, ""
                    lRightFileRowCnt = lRightFileRowCnt + 1
                Next
                iCnt = iCnt + 1
            End If
        Next
next_comp:
    Next
    
    If iCnt > 0 Then
        Shell """" & BEYOND_COMPARE_EXE & """ """ & COMPARE_TMP_FILE_LEFT & """  """ & COMPARE_TMP_FILE_RIGHT & """", vbMaximizedFocus
        MsgBox iCnt & " functions are found that are different from local libarary file" & vbCr & vbCr & SOURCE_CODE_LIBRARY_FILE, vbExclamation
    Else
        MsgBox "all are same"
    End If
    
error_handling:
    Close #iLeftFileNum
    Close #iRightFileNum
    
    Erase arrFileLines
    Set dictFunsInFile = Nothing
    Set dictFunsInMacro = Nothing
    Set vbP = Nothing
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub
Function fGetAllSubFunctionsOfAModule(vbcomp As VBIDE.VBComponent) As Dictionary
    Dim dictOut As Dictionary
    Dim codeM As CodeModule
    Dim lLineNum As Long
    Dim sFunName As String
    Dim procKind As VBIDE.vbext_ProcKind
    Dim sScope As String
    Dim lStartLine As Long
    Dim lBodyLine As Long
    Dim lCntLine As Long
    Dim lEndLine As Long
    
    Set dictOut = New Dictionary
    
    Set codeM = vbcomp.CodeModule
    
    lLineNum = codeM.CountOfDeclarationLines + 1
    
    Do Until lLineNum >= codeM.CountOfLines
        sFunName = codeM.ProcOfLine(lLineNum, procKind)
        sScope = fGetSubFunctionDeclareScope(codeM, sFunName, procKind)
        
        lStartLine = codeM.ProcStartLine(sFunName, procKind)
        lBodyLine = codeM.ProcBodyLine(sFunName, procKind)
        lCntLine = codeM.ProcCountLines(sFunName, procKind)
        
        lEndLine = fFindSubFunctionActualEndLine(codeM, lBodyLine, lBodyLine + lCntLine - (lBodyLine - lStartLine) - 1)
        
        If Not dictOut.Exists(sFunName) Then    'property let, abc, property get abc
            dictOut.Add sFunName, sScope & DELIMITER & vbcomp.Name & DELIMITER & lBodyLine & DELIMITER & lEndLine
        End If
        
        lLineNum = lLineNum + lCntLine + 1
    Loop
    
    Set codeM = Nothing
    
    Set fGetAllSubFunctionsOfAModule = dictOut
    Set dictOut = Nothing
End Function
Function fGetSubFunctionDeclareScope(codeM As CodeModule, sFunName As String, procKind As VBIDE.vbext_ProcKind) As String
    Dim sProcScope As String
    Dim sDeclare As String
    Dim lProcDeclareLne As Long

    On Error Resume Next
    lProcDeclareLne = codeM.ProcBodyLine(sFunName, procKind)
    If Err.Number <> 0 Then
        sProcScope = "Sub/Function Not found"
        GoTo exit_fun
    End If
    
    sDeclare = UCase(Trim(codeM.Lines(lProcDeclareLne, 1)))
    
    If Len(sDeclare) <= 0 Then
        sProcScope = "sub/fun not found"
    Else
        If Left(sDeclare, Len("PUBLIC ")) = "PUBLIC " Then
            sProcScope = "Public"
        ElseIf Left(sDeclare, Len("PRIVATE ")) = "PRIVATE " Then
            sProcScope = "Private"
        ElseIf Left(sDeclare, Len("FRIEND ")) = "FRIEND " Then
            sProcScope = "Friend"
        Else
            sProcScope = "Default"
        End If
    End If

exit_fun:
    fGetSubFunctionDeclareScope = sProcScope
End Function
Function fFindSubFunctionActualEndLine(codeM As CodeModule, lBodyLine As Long, lRoughEndLine As Long) As Long
    Dim lActualEndLine As Long
    Dim lEachRow As Long
    Dim sLineContent As String
    
    lActualEndLine = 0
    For lEachRow = lRoughEndLine To lBodyLine - 1 Step -1
        sLineContent = Trim(codeM.Lines(lEachRow, 1))
        
        If Left(sLineContent, 1) = "'" Then GoTo next_row
                
        sLineContent = UCase(sLineContent)
        
        If Left(sLineContent, Len("END ")) = "END " Then
            lActualEndLine = lEachRow
            Exit For
        End If
next_row:
    Next
    
    If lActualEndLine <= 0 Then fErr "lActualEndLine = 0"
    fFindSubFunctionActualEndLine = lActualEndLine
End Function


Function fMacroSectionIsSameAsTextFile(codeM As VBIDE.CodeModule, lMacroStartLine As Long, lMacroEndLine As Long _
               , arrFileLines, lFileStartLine As Long, lFileEndLine As Long) As Boolean
    Dim lMacroEachRow As Long
    Dim lFileEachRow As Long
    Dim sMacroLine As String
    Dim sFileLine As String
    Dim bOut As Boolean
    
    bOut = True
    lFileEachRow = 0
    For lMacroEachRow = 0 To lMacroEndLine - lMacroStartLine
        sMacroLine = UCase(Replace(codeM.Lines(lMacroStartLine + lMacroEachRow, 1), " ", ""))
        
        If Len(sMacroLine) <= 0 Or Left(sMacroLine, 1) = "'" Then GoTo next_row
        
        sFileLine = UCase(Replace(arrFileLines(lFileStartLine + lFileEachRow), " ", ""))
        
        Do While (Len(sFileLine) <= 0 Or Left(sFileLine, 1) = "'") And lFileEachRow <= lFileEndLine - lFileStartLine
            lFileEachRow = lFileEachRow + 1
            sFileLine = UCase(Replace(arrFileLines(lFileStartLine + lFileEachRow), " ", ""))
        Loop
        
        If sMacroLine <> sFileLine Then
            bOut = False
            Exit For
        End If
        
        lFileEachRow = lFileEachRow + 1
next_row:
    Next
    
    If bOut Then
        If lFileEachRow <= lFileEndLine - lFileStartLine Then
            sFileLine = UCase(Replace(arrFileLines(lFileStartLine + lFileEachRow), " ", ""))
            
            Do While (Len(sFileLine) <= 0 Or Left(sFileLine, 1) = "'") And lFileEachRow <= lFileEndLine - lFileStartLine
                lFileEachRow = lFileEachRow + 1
                sFileLine = UCase(Replace(arrFileLines(lFileStartLine + lFileEachRow), " ", ""))
            Loop
            
            If lFileEachRow <= lFileEndLine - lFileStartLine Then bOut = False
        End If
    End If
    fMacroSectionIsSameAsTextFile = bOut
End Function

Function fGetComponentTypeToString(compType As VBIDE.vbext_ComponentType) As String
    Dim sOut As String
    
    Select Case compType
        Case vbext_ComponentType.vbext_ct_ActiveXDesigner
            sOut = "ActiveXDesigner"
        Case vbext_ComponentType.vbext_ct_ClassModule
            sOut = "ClassModule"
        Case vbext_ComponentType.vbext_ct_Document
            sOut = "DocumentModule"
        Case vbext_ComponentType.vbext_ct_MSForm
            sOut = "UserForm"
        Case vbext_ComponentType.vbext_ct_StdModule
            sOut = "StandardModule"
        Case Else
            sOut = "Unknow Type: " & CStr(compType)
    End Select
    
    fGetComponentTypeToString = sOut
End Function

Sub subMainValidateSourceCodeFile()
    On Error GoTo error_handling
    
    Call fInitialization
    
    Call fBackupTextFileWithDefaultFileName(SOURCE_CODE_LIBRARY_FILE)
    
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
    
    sTmpFile = fGetFileNetName(sFileFullPath, True) & "." & Format(Now(), "yyyymmddHHMMSS")
    
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
                    Print #iLeftFileNum, "'========== left ==================== line " & lStartLine + 1 & " - " & lEndLine + 1 & " ========================="
                    
                    For k = lStartLine To lEndLine
                        Print #iLeftFileNum, arrFileLines(k)
                    Next
                Else
                    Print #iRightFileNum, "'==========right ==================== line " & lStartLine + 1 & " - " & lEndLine + 1 & " ========================="
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
    
    sTmpFile = fGetFileNetName(sFileFullPath, True) & "." & Format(Now(), "yyyymmddHHMMSS")
    
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
    
    sTmpFile = fGetFileNetName(sFileFullPath, True) & "." & Format(Now(), "yyyymmddHHMMSS")
    
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

Function fBackupTextFileWithDefaultFileName(sFileFullPath As String)
    Dim sTmpFile As String
    Dim lEachRow As Long
    Dim sBackupFolder As String
    
    sBackupFolder = fGetFileParentFolder(sFileFullPath) & "backup\"
    
    Call fCheckPath(sBackupFolder, True)
    
    sTmpFile = sBackupFolder & fGetFileNetName(sFileFullPath) & "." & Timer() * 100 & ".bak"
    
    fGetFSO
    Call gFSO.CopyFile(sFileFullPath, sTmpFile, True)
    
End Function

Function fInitialization()
    Err.Clear
    gErrNum = 0
    gErrMsg = ""
    
    Call fDisableExcelOptionsAll
    
    Application.ScreenUpdating = False
    
    'If Workbooks.Count > 0 Then Call fRemoveFilterForAllSheets
End Function

Function fExportSourceCodeToFolder(sFolderExportedTo As String, Optional wb As Workbook)
    Dim vbP As VBIDE.VBProject
    Dim vbcomp As VBIDE.VBComponent
    
    If wb Is Nothing Then
        Set vbP = ActiveWorkbook.VBProject
    Else
        Set vbP = wb.VBProject
    End If
    
    sFolderExportedTo = fCheckPath(sFolderExportedTo)
    
    For Each vbcomp In vbP.VBComponents
        vbcomp.Export sFolderExportedTo & vbcomp.Name & ".bas"
    Next
    
    Set vbcomp = Nothing
    Set vbP = Nothing
End Function

Sub subMain_Compare2MacroFiles()
    Dim sFilePath_Left As String
    Dim sFilePath_Right As String
    
    sFilePath_Left = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).value)
    sFilePath_Right = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).value)
    
    If fFileExists(sFilePath_Left) And fFileExists(sFilePath_Right) _
    And fExactExcelFileIsopen(sFilePath_Left) And fExactExcelFileIsopen(sFilePath_Right) Then
        subMain_Compare2MacroFiles_AfterOpen2Macros
    Else
        subMain_Compare2MacroFiles_AllInOne
    End If
End Sub
Sub subMain_Compare2MacroFiles_AfterOpen2Macros()
    Dim sFilePath_Left As String
    Dim sFilePath_Right As String
    Dim sExportParentFolder As String
    Dim sSourceCodeFolder_Left As String
    Dim sSourceCodeFolder_Right As String
    
    'On Error GoTo error_handling
    
    Call fInitialization
    
    sFilePath_Left = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).value)
    sFilePath_Right = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).value)
    
    sExportParentFolder = fGetFileParentFolder(sFilePath_Left) & "SourceCodeCompare_TempFolder\"
    sSourceCodeFolder_Left = sExportParentFolder & fGetFileNetName(sFilePath_Left)
    sSourceCodeFolder_Left = fCheckPath(sSourceCodeFolder_Left, True)
    
    sExportParentFolder = fGetFileParentFolder(sFilePath_Right) & "SourceCodeCompare_TempFolder\"
    sSourceCodeFolder_Right = sExportParentFolder & fGetFileNetName(sFilePath_Right)
    sSourceCodeFolder_Right = fCheckPath(sSourceCodeFolder_Right, True)
    
    Dim bAlreadyOpenedLeft As Boolean
    Dim bAlreadyOpenedRight As Boolean
    Dim wbLeft As Workbook
    Dim wbRight As Workbook
    
    If Not (fExactExcelFileIsopen(sFilePath_Left, wbLeft) And fExactExcelFileIsopen(sFilePath_Right, wbRight)) Then
        fErr "Please take the first step."
    End If
    
    bAlreadyOpenedLeft = ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyOpened)
    bAlreadyOpenedRight = ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyOpened)
    
    Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sSourceCodeFolder_Left)
    Call fExportSourceCodeToFolder(sSourceCodeFolder_Left, wbLeft)
    If Not bAlreadyOpenedLeft Then Call fCloseWorkBookWithoutSave(wbLeft)
    
    Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sSourceCodeFolder_Right)
    Call fExportSourceCodeToFolder(sSourceCodeFolder_Right, wbRight)
    If Not bAlreadyOpenedRight Then Call fCloseWorkBookWithoutSave(wbRight)
     
    Shell """" & BEYOND_COMPARE_EXE & """ /fileviewer=""Folder Compare"" /filters=-*frx """ & sSourceCodeFolder_Left & """  """ & sSourceCodeFolder_Right & """", vbMaximizedFocus
    
error_handling:
    'If gsRtnValueOfForm <> CONST_SUCCESS Then GoTo reset_excel_options
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub
Sub subMain_Compare2MacroFiles_AllInOne()
    Dim sFilePath_Left As String
    Dim sFilePath_Right As String
    Dim sExportParentFolder As String
    Dim sSourceCodeFolder_Left As String
    Dim sSourceCodeFolder_Right As String
    Dim bSameFileName As Boolean
    
    On Error GoTo error_handling
    
    Call fInitialization
    
    FrmCompareTwoMacroFiles.Show
    If gsRtnValueOfForm <> CONST_SUCCESS Then GoTo error_handling
    
    sFilePath_Left = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).value)
    sFilePath_Right = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).value)
    
    sExportParentFolder = fGetFileParentFolder(sFilePath_Left) & "SourceCodeCompare_TempFolder\"
    sSourceCodeFolder_Left = sExportParentFolder & fGetFileNetName(sFilePath_Left)
    sSourceCodeFolder_Left = fCheckPath(sSourceCodeFolder_Left, True)
    
    sExportParentFolder = fGetFileParentFolder(sFilePath_Right) & "SourceCodeCompare_TempFolder\"
    sSourceCodeFolder_Right = sExportParentFolder & fGetFileNetName(sFilePath_Right)
    sSourceCodeFolder_Right = fCheckPath(sSourceCodeFolder_Right, True)
    
    Dim bAlreadyOpenedLeft As Boolean
    Dim bAlreadyOpenedRight As Boolean
    Dim wbLeft As Workbook
    Dim wbRight As Workbook
    Dim bAlreadyExportedLeft As Boolean
    Dim bAlreadyExportedRight As Boolean
    
    bSameFileName = UCase(fGetFileBaseName(sFilePath_Left)) = UCase(fGetFileBaseName(sFilePath_Right))
    
    If bSameFileName Then
        bAlreadyExportedLeft = ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyExported)
        bAlreadyExportedRight = ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyExported)
    
        If Not bAlreadyExportedLeft Then
            Set wbLeft = fOpenWorkbook(sFilePath_Left, bAlreadyOpenedLeft)
            ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyOpened) = bAlreadyOpenedLeft
            'Call Hook
            Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sSourceCodeFolder_Left)
            Call fExportSourceCodeToFolder(sSourceCodeFolder_Left, wbLeft)
            Call fCloseWorkBookWithoutSave(wbLeft)
            
            ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyExported).value = True
        End If
        
        Set wbRight = fOpenWorkbook(sFilePath_Right, bAlreadyOpenedRight)
        ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyOpened) = bAlreadyOpenedRight
        'Call Hook
        Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sSourceCodeFolder_Right)
        Call fExportSourceCodeToFolder(sSourceCodeFolder_Right, wbRight)
        Call fCloseWorkBookWithoutSave(wbRight)
    Else
        'left file
        Set wbLeft = fOpenWorkbook(sFilePath_Left, bAlreadyOpenedLeft)
        ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroAlreadyOpened) = bAlreadyOpenedLeft
        'right file
        Set wbRight = fOpenWorkbook(sFilePath_Right, bAlreadyOpenedRight)
        ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroAlreadyOpened) = bAlreadyOpenedRight
        
        'Call Hook
        Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sSourceCodeFolder_Left)
        Call fExportSourceCodeToFolder(sSourceCodeFolder_Left, wbLeft)
        
        'Call Hook
        Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sSourceCodeFolder_Right)
        Call fExportSourceCodeToFolder(sSourceCodeFolder_Right, wbRight)
        
        If (Not bAlreadyOpenedLeft) Then Call fCloseWorkBookWithoutSave(wbLeft)
        If Not bAlreadyOpenedRight Then Call fCloseWorkBookWithoutSave(wbRight)
    End If
     
    Shell """" & BEYOND_COMPARE_EXE & """ /fileviewer=""Folder Compare"" /filters=-*frx """ & sSourceCodeFolder_Left & """  """ & sSourceCodeFolder_Right & """", vbMaximizedFocus
    
error_handling:
    If gsRtnValueOfForm <> CONST_SUCCESS Then GoTo reset_excel_options
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub
Sub subMain_Compare2MacroFiles_2ndStep()

End Sub

Sub subMain_CompareWithCommonLibFolder()
    Dim sTargetMacro As String
    Dim wbTarget As Workbook
    Dim arrLibFiles
    Dim dictCommonModules As Dictionary
    'Dim dictIgnore As Dictionary
    Dim response As VbMsgBoxResult
    Dim i As Integer
    Dim sModuleFileFullPath As String
    Dim sModuleName As String
    
    On Error GoTo error_handling
    
    Call fInitialization
    
    Call fSetSavedValue(RANGE_SyncWithCommLibWhichFunction, "COMPARE_WITH_COMMON_LIB")
    FrmSyncModulesFromLibFiles.Show
    If gsRtnValueOfForm <> CONST_SUCCESS Then fErr
    ThisWorkbook.Save
    
    sTargetMacro = fGetSavedValue(RANGE_TargetMacroToSyncWithCommLib)
    Set wbTarget = fOpenWorkbook(sTargetMacro, , , , , False)
    Call fWorkbookVBProjectIsProteced(wbTarget)
    
    'sCommLibFolder = fGetSavedValue(RANGE_CommonLibFolderSelected)
    arrLibFiles = Split(fGetSavedValue(RANGE_CommonLibFilesSelected), vbCrLf)
    Set dictCommonModules = fFilterCommonLibFilesWithMacro(arrLibFiles, wbTarget)
    
    If dictCommonModules.Count <= 0 Then fErr "No modules are found in the macro matching the provided common lib files"
     
    Dim sSourceCodeFolder_Macro As String
    Dim sSourceCodeFolder_CommLib As String
'
'    sFilePath_Left = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_LeftMacroToCompare).value)
'    sFilePath_Right = Trim(ThisWorkbook.Worksheets(1).Range(RANGE_RightMacroToCompare).value)
    
    sSourceCodeFolder_Macro = fGetFileParentFolder(sTargetMacro) & "SourceCodeCompare_TempFolder\" & fGetFileNetName(sTargetMacro)
    
    If Not fFolderExists(sSourceCodeFolder_Macro) Then fErr "the macro's compare folder does not exist, please run the first function first" & vbCr & vbCr & sSourceCodeFolder_Macro
    
    sSourceCodeFolder_CommLib = fGetFileParentFolder(CStr(arrLibFiles(LBound(arrLibFiles))))
    Erase arrLibFiles
    
    Dim sLibFile As String
    Dim vbP As VBIDE.VBProject
    
    Set vbP = wbTarget.VBProject
    
    Call fDeleleteAllFilesFromFolderIfNotExistsCreateIt(sSourceCodeFolder_Macro)
    
    For i = 0 To dictCommonModules.Count - 1
        sModuleFileFullPath = dictCommonModules.Keys(i)
        sModuleName = dictCommonModules.Items(i)
        
        vbP.VBComponents(sModuleName).Export sSourceCodeFolder_Macro & fGetFileBaseName(sModuleFileFullPath)
    Next
    
    Shell """" & BEYOND_COMPARE_EXE & """ /fileviewer=""Folder Compare"" /filters=-*frx """ & sSourceCodeFolder_Macro & """  """ & sSourceCodeFolder_CommLib & """", vbMaximizedFocus
    
error_handling:
'    Erase arrFileLines
'    Set dictFunsInFile = Nothing
    Set vbP = Nothing
    Set wbTarget = Nothing
    Set dictCommonModules = Nothing
    'Set dictIgnore = Nothing
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    fMsgBox "done.", vbInformation
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub
Sub subMain_SyncWithSelfRevised()
    Dim sTargetMacro As String
    Dim wbTarget As Workbook
    Dim arrLibFiles
'    Dim dictCommonModules As Dictionary
'    Dim dictIgnore As Dictionary
    Dim response As VbMsgBoxResult
    Dim i As Integer
    Dim sModuleFileFullPath As String
    Dim sModuleName As String
    
    'On Error GoTo error_handling
    
    Call fInitialization
    
    FrmSyncModulesFromSelfRevised.Show
    If gsRtnValueOfForm <> CONST_SUCCESS Then fErr
    ThisWorkbook.Save
    
    sTargetMacro = fGetSavedValue(RANGE_TargetMacroToSyncWithCommLib)
    Set wbTarget = fOpenWorkbook(sTargetMacro, , , , , False)
    Call fWorkbookVBProjectIsProteced(wbTarget)
    
    'sCommLibFolder = fGetSavedValue(RANGE_CommonLibFolderSelected)
    arrLibFiles = Split(fGetSavedValue(RANGE_CommonLibFilesSelected), vbCrLf)
'    Set dictCommonModules = fFilterCommonLibFilesWithMacro(arrLibFiles, wbTarget, dictIgnore)

    Dim sCommonModMsg As String
    For i = LBound(arrLibFiles) To UBound(arrLibFiles)
        sModuleName = fReadModuleNameFromSourceCodeFile(CStr(arrLibFiles(i)))
        
        sCommonModMsg = sCommonModMsg & vbCr & sModuleName
    Next

    If ArrLen(arrLibFiles) > 0 Then
        response = MsgBox("The modules below are going to replace the existing ones, are you sure to continue?" & vbCr & vbCr & sCommonModMsg, vbYesNo + vbQuestion + vbDefaultButton1)
        If response <> vbYes Then fErr
    Else
        fErr "No modules are found"
    End If

    For i = LBound(arrLibFiles) To UBound(arrLibFiles)
        sModuleFileFullPath = arrLibFiles(i)
        
        Call fImportModuleToWorkbookFromSourceCodeFile(wbTarget, sModuleFileFullPath)
    Next
    
    Erase arrLibFiles
    
error_handling:
'    Erase arrFileLines
'    Set dictFunsInFile = Nothing
    Set wbTarget = Nothing
'    Set dictCommonModules = Nothing
'    Set dictIgnore = Nothing
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    fMsgBox "done.", vbInformation
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub

Sub subMain_SyncWithCommonLib()
    Dim sTargetMacro As String
    Dim wbTarget As Workbook
    Dim arrLibFiles
    Dim dictCommonModules As Dictionary
    Dim dictIgnore As Dictionary
    Dim response As VbMsgBoxResult
    Dim i As Integer
    Dim sModuleFileFullPath As String
    Dim sModuleName As String

    On Error GoTo error_handling

    Call fInitialization

    Call fSetSavedValue(RANGE_SyncWithCommLibWhichFunction, "SYNC_WITH_COMMON_LIB")
    FrmSyncModulesFromLibFiles.Show
    If gsRtnValueOfForm <> CONST_SUCCESS Then fErr
    ThisWorkbook.Save

    sTargetMacro = fGetSavedValue(RANGE_TargetMacroToSyncWithCommLib)
    Set wbTarget = fOpenWorkbook(sTargetMacro, , , , , False)
    Call fWorkbookVBProjectIsProteced(wbTarget)

    'sCommLibFolder = fGetSavedValue(RANGE_CommonLibFolderSelected)
    arrLibFiles = Split(fGetSavedValue(RANGE_CommonLibFilesSelected), vbCrLf)
    Set dictCommonModules = fFilterCommonLibFilesWithMacro(arrLibFiles, wbTarget, dictIgnore)
    Erase arrLibFiles

    If dictIgnore.Count > 0 Then
        response = MsgBox("Some of the library file you provided are not found in the macro, so they will be ingored, to continue?" & vbCr & vbCr & Join(dictIgnore.Items, vbCr), vbYesNo + vbQuestion + vbDefaultButton1)
        If response <> vbYes Then fErr
    End If
    If dictCommonModules.Count > 0 Then
        response = MsgBox("The modules below are going to replace the existing ones, are you sure to continue?" & vbCr & vbCr & Join(dictCommonModules.Items, vbCr), vbYesNo + vbQuestion + vbDefaultButton1)
        If response <> vbYes Then fErr
    Else
        fErr "No modules are found in the macro matching the provided common lib files"
    End If

    For i = 0 To dictCommonModules.Count - 1
        sModuleFileFullPath = dictCommonModules.Keys(i)
        sModuleName = dictCommonModules.Items(i)

        'Call fRemoveDeleteModuleIfExists(wbTarget, sModuleName)

        Call fImportModuleToWorkbookFromSourceCodeFile(wbTarget, sModuleFileFullPath, sModuleName)
    Next

error_handling:
'    Erase arrFileLines
'    Set dictFunsInFile = Nothing
    Set wbTarget = Nothing
    Set dictCommonModules = Nothing
    Set dictIgnore = Nothing

    If gErrNum <> 0 Then GoTo reset_excel_options

    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    fMsgBox "done.", vbInformation
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub

Function fImportModuleToWorkbookFromSourceCodeFile(wbTarget As Workbook, sModuleFileFullPath As String, Optional asModuleName As String)
    Dim sModuleNameInFile As String
    
    sModuleNameInFile = fReadModuleNameFromSourceCodeFile(sModuleFileFullPath)
    
    If Len(Trim(asModuleName)) Then
        Call fRemoveDeleteModuleIfExists(wbTarget, asModuleName)
        
        If sModuleNameInFile <> asModuleName Then
            Call fRemoveDeleteModuleIfExists(wbTarget, sModuleNameInFile)
        End If
    Else
        Call fRemoveDeleteModuleIfExists(wbTarget, sModuleNameInFile)
    End If
    
    wbTarget.VBProject.VBComponents.Import sModuleFileFullPath
    If Len(Trim(asModuleName)) > 0 Then
        If sModuleNameInFile <> asModuleName Then
            wbTarget.VBProject.VBComponents(sModuleNameInFile).Name = asModuleName
        End If
    End If
End Function
Function fRemoveDeleteModuleIfExists(wbTarget As Workbook, sModuleName As String)
    Dim vbcomp As VBIDE.VBComponent
    Dim sOldName As String
    
    If fModuleExistsInMacro(sModuleName, wbTarget) Then
        Set vbcomp = wbTarget.VBProject.VBComponents(sModuleName)
        sOldName = Left(sModuleName & Replace(CStr(Timer()), ".", ""), 30)
        vbcomp.Name = sOldName

        'Debug.Print "wbTarget.VBProject.VBComponents.Remove vbComp: " & sModuleName
        wbTarget.VBProject.VBComponents.Remove vbcomp
        
'        Call wbTarget.VBProject.VBComponents.Remove(wbTarget.VBProject.VBComponents(sModuleName))
        Set vbcomp = Nothing
        
'        If Not fModuleExistsInMacro(sModuleName, wbTarget) Then
'            Debug.Print "" & sModuleName & " was removed."
'        Else
'            Debug.Print "!!! " & sModuleName & " was not removed."
'        End If
    Else
        Debug.Print "!!! " & sModuleName & " does not exists in the workbook, please check."
    End If
    Set vbcomp = Nothing
End Function

Function fModuleExistsInMacro(sModuleName As String, Optional wb As Workbook, Optional ByRef outCodeM As CodeModule) As Boolean
    On Error Resume Next
    
    fModuleExistsInMacro = True
    
    Dim vbP As VBIDE.VBProject
    
    If wb Is Nothing Then
        Set vbP = ActiveWorkbook.VBProject
    Else
        Set vbP = wb.VBProject
    End If
    
    Set outCodeM = vbP.VBComponents(sModuleName).CodeModule
    If Err.Number <> 0 Then
        Err.Clear
        fModuleExistsInMacro = False
    End If

    Set vbP = Nothing
End Function
Function fFilterCommonLibFilesWithMacro(arrLibFiles, wb As Workbook, Optional ByRef dictToIgnore As Dictionary, Optional ByRef sCommonFileForMsg As String) As Dictionary
    Dim dictOut As Dictionary
    Dim vbP As VBIDE.VBProject
    Dim vbcomp As VBIDE.VBComponent
    Dim sModType As String
    Dim dictWbModules As Dictionary
    Dim i As Integer
    
    Set vbP = wb.VBProject
    Set dictOut = New Dictionary
    Set dictToIgnore = New Dictionary
    Set dictWbModules = New Dictionary
    sCommonFileForMsg = ""
    
    For Each vbcomp In vbP.VBComponents
        sModType = fVBEComponentTypeToString(vbcomp.Type)
        
        If sModType <> "Document" Then
            dictWbModules.Add UCase(vbcomp.Name), ""
        End If
    Next
    
    Dim sLibFile  As String
    Dim sLibFileExt  As String
    Dim sModuleNameInFile As String
    For i = LBound(arrLibFiles) To UBound(arrLibFiles)
        sLibFile = arrLibFiles(i)
        
        sLibFileExt = UCase(fGetFileExtension(sLibFile))
        
        If sLibFileExt <> "BAS" And sLibFileExt <> "CLS" And sLibFileExt <> "FRM" Then GoTo next_file
        
        sModuleNameInFile = fReadModuleNameFromSourceCodeFile(sLibFile)
        
        If dictWbModules.Exists(UCase(sModuleNameInFile)) Then
            dictOut.Add sLibFile, sModuleNameInFile '& DELIMITER & sLibFile
            sCommonFileForMsg = sCommonFileForMsg & vbCr & sModuleNameInFile
        Else
            dictToIgnore.Add sLibFile, fGetFileBaseName(sLibFile)
        End If
next_file:
    Next

    Set dictWbModules = Nothing
    Set vbcomp = Nothing
    Set vbP = Nothing
    
    Set fFilterCommonLibFilesWithMacro = dictOut
    Set dictOut = Nothing
End Function

Function fReadModuleNameFromSourceCodeFile(sLibFile As String) As String
    Dim arrFileLines
    Dim lEachLine As Long
    Dim sEachLine As String
    Dim sModuleName As String
    
    arrFileLines = fReadTextFileAllLinesToArray(sLibFile)
    
    For lEachLine = LBound(arrFileLines) To UBound(arrFileLines)
        sEachLine = arrFileLines(lEachLine)
        
        If Trim(sEachLine) Like "Attribute VB_Name*" Then
            sModuleName = Trim(Split(sEachLine, "=")(1))
            sModuleName = Trim(Split(sModuleName, """")(1))
            Exit For
        End If
    Next
    Erase arrFileLines
    
    If Len(sModuleName) <= 0 Then fErr "the source code file is invalid, becase the module name cannot be detected"
    fReadModuleNameFromSourceCodeFile = sModuleName
End Function

Sub subMain_ScanUselessFunctions()
    If Workbooks.Count <= 0 Then Exit Sub
    
    Dim wbTarget As Workbook
    Set wbTarget = ActiveWorkbook
    
    Call fScanUselessFunctions(wbTarget)
    
    Set wbTarget = Nothing
End Sub

Sub subMain_CommentOutScanUselessFunctions()
    If Workbooks.Count <= 0 Then Exit Sub
    
    Dim wbTarget As Workbook
    Set wbTarget = ActiveWorkbook
    
    'Call fCommentOutUselessFunctions(wbTarget)
    
    Set wbTarget = Nothing
End Sub

Function fScanUselessFunctions(wbTarget As Workbook)
    Dim sTmpOutput As String
    Dim shtOutput As Worksheet
    Dim lMaxRow As Long
    Dim arrHeader()
    
    On Error GoTo error_handling
    
    Call fInitialization
    
    sTmpOutput = "tmpOutput_ScanSourceCode"
    arrHeader = Array("Module", "Top Sub/Functions to exclude", "", "Exception Module", "Sub/Functions not being called", "Sub/Function Scope", "From Line", "To Line")
    
    If fSheetExists(sTmpOutput, shtOutput, wbTarget) Then
        Call fShowAndActiveSheet(shtOutput)
        
        If Not (Not fFindInWorksheet(shtOutput.Cells, CStr(arrHeader(1)), False, True) Is Nothing _
              And Not fFindInWorksheet(shtOutput.Cells, CStr(arrHeader(2)), False, True) Is Nothing _
              And Not fFindInWorksheet(shtOutput.Cells, CStr(arrHeader(4)), False, True) Is Nothing _
              And Not fFindInWorksheet(shtOutput.Cells, CStr(arrHeader(5)), False, True) Is Nothing _
              And Not fFindInWorksheet(shtOutput.Cells, CStr(arrHeader(6)), False, True) Is Nothing _
              And Not fFindInWorksheet(shtOutput.Cells, CStr(arrHeader(7)), False, True) Is Nothing _
              And Not fFindInWorksheet(shtOutput.Cells, CStr(arrHeader(8)), False, True) Is Nothing) Then
            If MsgBox("There is an existing sheet " & shtOutput.Name & ", to overwrite it, please press Yes" _
                    , vbYesNoCancel + vbCritical + vbDefaultButton1, vbQuestion) = vbYes Then
                shtOutput.Cells.Delete
                Call fWriteHeaderToSheet(shtOutput, arrHeader)
                MsgBox "Please input, and click the button again.", vbInformation
                GoTo error_handling
            Else
                fErr
            End If
        Else
            lMaxRow = fGetValidMaxRow(shtOutput)
            fGetRangeByStartEndPos(shtOutput, 2, 4, lMaxRow + 2, 15).ClearContents
            fGetRangeByStartEndPos(shtOutput, 2, 4, lMaxRow + 2, 15).ClearFormats
        End If
    Else
        Set shtOutput = fAddNewSheet(sTmpOutput, wbTarget)
        Call fFreezeSheet(shtOutput)
        
        Call fWriteHeaderToSheet(shtOutput, arrHeader)
        
        MsgBox "Please input, and click the button again.", vbInformation
        GoTo error_handling
    End If
    
    Dim arrOut()
    
    arrOut = fScanSourceCodeToFindUncalledFunctions(wbTarget)
    
    If ArrLen(arrOut) > 0 Then
        shtOutput.Cells(2, 4).Resize(ArrLen(arrOut), 5).value = arrOut
        Erase arrOut
    End If
    
    shtOutput.Columns.AutoFit
error_handling:
'    Erase arrFileLines
'    Set dictFunsInFile = Nothing
    Set wbTarget = Nothing
    Set shtOutput = Nothing
    'Set dictCommonModules = Nothing
    'Set dictIgnore = Nothing
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    fMsgBox "done.", vbInformation
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Function

Function fScanSourceCodeToFindUncalledFunctions(Optional wb As Workbook)
    Dim vbP As VBIDE.VBProject
    Dim vbcomp As VBIDE.VBComponent
    Dim codeM As VBIDE.CodeModule
    Dim prodKind As VBIDE.vbext_ProcKind
    Dim sModuleUserInput 'As String
    
    If wb Is Nothing Then
        Set vbP = ActiveWorkbook.VBProject
    Else
        Set vbP = wb.VBProject
    End If
    
    sModuleUserInput = InputBox("You count input the module if you want to scan one module by one module " & vbCr & "press cancel to scan all modules", "one module by one module", fGetSavedValue(RANGE_ScanUselessOnebyOneModule))
    sModuleUserInput = Trim(sModuleUserInput)
    
    If Len(sModuleUserInput) > 0 Then
        If Not fModuleExistsInMacro(CStr(sModuleUserInput), wb) Then
            fErr "the module you input does not exists, please check: " & sModuleUserInput
        End If
        
        Call fSetSavedValue(RANGE_ScanUselessOnebyOneModule, sModuleUserInput)
    End If
    
    Dim dictExclude As Dictionary
    Dim dictIsolated As Dictionary
    Dim dictFunList As Dictionary
    Dim sSubFunName As String
    Dim i As Long
    Dim sSubFunProp As String
    Dim sModuleName As String
    
    Set dictExclude = fReadExcludeModuleFunList(ActiveSheet)
    Set dictIsolated = New Dictionary
    
    For Each vbcomp In vbP.VBComponents
        If Len(sModuleUserInput) > 0 Then
            If UCase(vbcomp.Name) <> UCase(sModuleUserInput) Then GoTo next_module
        End If
        
        Set codm = vbcomp.CodeModule
        sModuleType = fGetComponentTypeToString(vbcomp.Type)
        
        Set dictFunList = fGetAllSubFunctionsOfAModule(vbcomp)
        
        For i = 0 To dictFunList.Count - 1
            sSubFunName = dictFunList.Keys(i)
            sSubFunProp = dictFunList.Items(i)
'            lStartLine = Split(sSubFunProp, DELIMITER)(2)
'            lEndLine = Split(sSubFunProp, DELIMITER)(3)
            
            If sSubFunName = "" Then
            ElseIf dictExclude.Exists(sModuleName & DELIMITER & sSubFunName) Then
                GoTo next_module
            End If
        Next
next_module:
    Next
    
    Set dictExclude = Nothing
    Set dictIsolated = Nothing
    Set vbP = Nothing
End Function

Function fReadExcludeModuleFunList(sht As Worksheet) As Dictionary
    Dim dictOut As Dictionary
    Dim lMaxRow As Long
    Dim lEachRow As Long
    Dim arrData()
    
    lMaxRow = fGetValidMaxRow(sht)
    
    Set dictOut = New Dictionary
    
    If lMaxRow > 2 Then
        arrData = fReadRangeDatatoArrayByStartEndPos(sht, 2, 1, lMaxRow, 2)
        
        For lEachRow = LBound(arrData, 1) To UBound(arrData, 1)
            sModule = Trim(arrData(lEachRow, 1))
            sfun = Trim(arrData(lEachRow, 2))
            
            If Len(sModule) > 0 And Len(sfun) > 0 Then
                If Not dictOut.Exists(sModule & DELIMITER & sfun) Then dictOut.Add sModule & DELIMITER & sfun, ""
            End If
        Next
    End If
    
    Erase arrData
    Set fReadExcludeModuleFunList = dictOut
    Set dictOut = Nothing
End Function

Function fWriteHeaderToSheet(sht As Worksheet, arrHeaders(), Optional alHeaderAtRow As Long = 1, Optional alHeaderFromCol As Long = 1)
    sht.Cells(alHeaderAtRow, alHeaderFromCol).Resize(1, ArrLen(arrHeaders)).value = fTranspose1DimenArrayTo2DimenArrayHorizontal(arrHeaders)
    
    Call fSetFormatBoldOrangeBorderForRangeEspeciallyForHeader(fGetRangeByStartEndPos(sht, 1, 1, 1, ArrLen(arrHeaders)))
    Call fSetConditionFormatForBorder(sht, , , , 1, True)
    
    sht.Columns.AutoFit
End Function
