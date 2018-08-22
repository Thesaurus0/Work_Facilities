Attribute VB_Name = "MZ_UT"
Option Explicit
Option Base 1

Sub AllUnitTest()
'    Dim asTag As String, rngToFindIn As Range _
'                                , arrConfigData() _
'                                , lConfigStartRow As Long _
'                                , lConfigStartCol As Long _
'                                , lConfigEndRow As Long _
'                                , lOutBlockHeaderAtRow As Long
'    Dim arrColsName()
'    Dim arrColsIndex()
'    Dim lConfigHeaderAtRow As Long
'
'    asTag = "[Input Files]"
'    arrColsName = Array("xxa", "Company ID", "Company Name")
'
'    Set rngToFindIn = ActiveSheet.Cells
'
'Call fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=activeshet.Cells, arrColsName:=arrColsName _
'                                , arrConfigData:=arrConfigData _
'                                , arrColsIndex:=arrColsIndex _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
 
'Call fReadConfigBlockToArray(asTag:=asTag, rngToFindIn:=ActiveSheet.Cells, arrColsName:=arrColsName _
'                                , arrConfigData:=arrConfigData _
'                                , arrColsIndex:=arrColsIndex _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
                       
'arrConfigData = fReadConfigBlockToArrayNet(asTag:=asTag, rngToFindIn:=ActiveSheet.Cells, arrColsName:=arrColsName _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                )
'arrConfigData = fReadConfigBlockToArrayValidated(asTag:=asTag, rngToFindIn:=rngToFindIn, arrColsName:=arrColsName _
'                                , lConfigStartRow:=lConfigStartRow _
'                                , lConfigStartCol:=lConfigStartCol _
'                                , lConfigEndRow:=lConfigEndRow _
'                                , lOutConfigHeaderAtRow:=lConfigHeaderAtRow _
'                                , abNoDataConfigThenError:=True _
'                                , arrKeyCols:=Array(2) _
'                                , bNetValues:=False _
'                                )
'    Dim arr()
'
'    'Debug.Print UBound(arr) & "-" & LBound(arr)
'   ' Debug.Print fArrayIsEmpty(arr)
'    'Debug.Print fGetArrayDimension(arr)
'    Dim a
'    Set a = ActiveCell.MergeArea
'
'    'Dim a
'    Set a = Selection
'
'    Debug.Print fGetValidMaxRowOfRange(Selection, True)
'
'    Dim bbb()
'    'bbb = fReadRangeDataToArray(Selection)

    'Debug.Print fGetSpecifiedConfigCellAddress(shtSysConf, "[Input Files]", "File Full Path", "Company ID = PW")
    'Debug.Print fGenRandomUniqueString
    'Debug.Assert fTrim(vbLf & " abcd " & vbCr) = "abcd"
    'Debug.Print fJoin(Selection.Value)
    
'    Dim arr
'    arr = fReadConfigWholeColsToArray(shtSysConf, "[Sales Company List]", Array("Company ID", "Company Name"), Array(1))
    
    'Call fReadConfigInputFiles
    
    'Call ThisWorkbook.fReadConfigGetAllCommandBars
'    Dim sAddr As String
'    sAddr = Range("A12:Z34").Address(ReferenceStyle:=xlR1C1, external:=True)
'    sAddr = Range("A12:Z34").Address(external:=True)
''    Debug.Print sAddr
''    Debug.Print fReplaceConvertR1C1ToA1(sAddr)
'
'    Dim rng As Range
'    Set rng = fGetRangeFromExternalAddress(sAddr)
'    Debug.Print rng.Address
    
    'Debug.Print fGetFileExtension("abce\ef\a\c\aaa.txt")
    
   ' Call fConvertFomulaToValueForSheetIfAny(ActiveSheet)
   fDeleteAllConditionFormatFromSheet ActiveSheet
   Call fSetConditionFormatForOddEvenLine(ActiveSheet, , , , Array(1), True)
   Call fSetConditionFormatForBorders(ActiveSheet, , , , Array(1), True)
End Sub

Sub testa()
'    Debug.Print Asc(" ")
'    Debug.Print Asc(vbCr)
'    Debug.Print Asc(vbLf)
'    Debug.Print Asc(vbCrLf)
'    Debug.Print Asc(vbNewLine)
'    Debug.Print Asc(vbTab)
    Dim aa
    aa = ActiveSheet.Range("c10:f20")
    
'    Dim bb(2, 4)
'    bb(0, 0) = "a"
'
'    Dim cc()
'    cc = Array()
'    Debug.Print LBound(aa, 1) & " - " & UBound(aa, 1)
'    Debug.Print LBound(aa, 2) & " - " & UBound(aa, 2)
'    Debug.Print LBound(bb, 1) & " - " & UBound(bb, 1)
'    Debug.Print LBound(bb, 2) & " - " & UBound(bb, 2)
'    Debug.Print LBound(cc, 1) & " - " & UBound(cc, 1)
'    Debug.Print LBound(cc, 2) & " - " & UBound(cc, 2)
    
'    Const DELI = " " & DELIMITER & " "
'    Dim f
'    'f = aa(0)
'    'Debug.Print Join(aa(3), "")
'    Dim s As String
'    Debug.Print fArrayIsEmptyOrNoData(s)
'    Dim a As String
'
'    a = "c"
'    'Debug.Print Switch(a = "a", 1, a = "b", 2, a = "c", 3, a = "e", 4)
'    Debug.Print Switch("c", 3, "e", 4)
    
'    Dim a
'    a = "[xxx]"
'    Debug.Print (a Like "[*]")

    Dim arr(1000)
    
    Dim i As Long
    
    For i = 1 To 1000
        arr(i) = Rnd() * 1000
    Next
    
    Call fSortArrayQuickSortDesc(arr)
    Call fSortArrayQuickSort(arr)
End Sub

Sub aaaa()
'    Dim a
'    'a = shtSalesRawDataRpt
'
'       'shtSalesRawDataRpt.Close
'
''    Call fHideAllSheetExcept("1", "2", "6", "24")
''    Dim rngAddr As String
''    rngAddr = fGetRangeByStartEndPos(shtProductMaster, 2, 1, 800, 1).Address(external:=True)
''    rngAddr = "=" & rngAddr
''    Call fSetValidationListForRange(fGetRangeByStartEndPos(shtProductProducerReplace, 2, 1, 1000, 1), rngAddr)
'
''    Dim sht ' As Worksheet
''     sht = Evaluate("shtSelfSalesCal")
''
''     fKeepCopyContent
''     Application.CutCopyMode = 0
''     fCopyFromKept
'    'End
'    fGetFSO
'    Debug.Print fCheckPath("F:\Github_Local_Repository\\Pharmacy_Excel_Tool_Macro\历史表\a.txt")
'    'a = Dir("F:\Github_Local_Repository\Pharmacy_Excel_Tool_Macro\历史表\a.txt", )
'   Debug.Print a
    End
End Sub

Function fWriteToNewTextFile(sFileFullPath As String, sContent As String) As Boolean
    Dim fileNum As Long
    
    fWriteToNewTextFile = False
    
    fileNum = FreeFile
    Open sFileFullPath For Output As #fileNum
    Print #fileNum, sContent
    Close #fileNum
    
    fWriteToNewTextFile = True
End Function

Function fLog(sContent As String)
   Dim fileNum As Long
     
    
    fileNum = FreeFile
    Open "F:\VBA_Orders\aaa.txt" For Append As #fileNum
    Print #fileNum, sContent
    Close #fileNum
     
End Function

Sub subMain_ExportToText()
    Dim arrData()
    Dim sFileBaseName As String
    Dim sContent As String
    Dim i As Long
    Dim sFileFullPath As String
    Dim iCnt As Long
    
    On Error GoTo 0
    
    Dim fileNum As Long
     
    
    Const par_folder = "F:\VBA_Orders\ExportText04\"
    
     
    'fGetFSO
    'gFSO.DeleteFolder par_folder, True
    
    Call fCheckPath(par_folder, True)
    
    Call fCopyReadWholeSheetData2Array(ActiveSheet, arrData)
    
    For i = LBound(arrData, 1) To UBound(arrData, 1)
        sFileBaseName = Trim(arrData(i, 4))
        
        sFileBaseName = fReplaceIlleagleCharInFileNameOneCharByOneChar(sFileBaseName)
            
        If Len(Trim(sFileBaseName)) <= 0 Then
            sFileBaseName = "标题为空_" & i
        Else
'            'Call fReplaceIlleagleCharInFileName(sFileBaseName)
'            sFileBaseName = fReplaceIlleagleCharInFileNameOneCharByOneChar(sFileBaseName)
'            If Len(sFileBaseName) <= 0 Then
'                sFileBaseName = "标题为空_" & i
'            Else
'            End If
        End If
        
        sFileFullPath = par_folder & sFileBaseName & ".txt"
        
        If fFileExists(sFileFullPath) Then
            sFileFullPath = par_folder & sFileBaseName & "_" & i & "_" & Format(Now(), "hhmmss") & ".txt"
        End If
        
        'On Error Resume Next
        If fWriteToNewTextFile(sFileFullPath, sFileBaseName & vbCr & vbCr & arrData(i, 8)) Then
            iCnt = iCnt + 1
        End If
        
'        If Err.Number <> 0 Then
'
'
'            sFileBaseName = fReplaceIlleagleCharInFileNameOneCharByOneChar(sFileBaseName)
'
'            If Len(sFileBaseName) <= 0 Then
'                sFileBaseName = "标题为空_" & i
'            End If
'
'            If fWriteToNewTextFile(sFileFullpath, sFileBaseName & vbCr & vbCr & arrData(i, 8)) Then
'                iCnt = iCnt + 1
'            End If
'
'            If Err.Number <> 0 Then
'                MsgBox "err: " & Err.Description
'                Err.Clear
'            End If
'        End If
        
        Debug.Print iCnt
        
'        If fFileExists(sFileFullpath) Then
'            'Call fLog(sFileFullpath)
'        Else
'           ' Call fLog("error: " & sFileFullpath)
'        End If
    Next
    
    Erase arrData
    
    MsgBox iCnt & " files are generated."
End Sub

Function fReplaceIlleagleCharInFileName(ByRef sFileName As String)
    sFileName = Replace(sFileName, "，", "")
    sFileName = Replace(sFileName, "?", "")
    sFileName = Replace(sFileName, "/", "-")
    sFileName = Replace(sFileName, "\", "-")
    sFileName = Replace(sFileName, "*", "")
End Function

Function fReplaceIlleagleCharInFileNameOneCharByOneChar(sFileName As String) As String
    Dim i As Long
    Dim sEach As String
    Dim sOut As String
    
    Call fReplaceIlleagleCharInFileName(sFileName)
    
    For i = 1 To Len(sFileName)
        sEach = Mid(sFileName, i, 1)
        
        If Asc(sEach) = 63 Then
            sOut = sOut & ""
        Else
            sOut = sOut & sEach
        End If
    Next
    
    fReplaceIlleagleCharInFileNameOneCharByOneChar = sOut
End Function

Sub test()
    Dim a As String
    Dim b As String
    a = "F:\VBA_Orders\ExportText\爱情玫瑰?-友情月季.txt"
    b = ActiveSheet.Range("D4857")
    
    Dim arr()
    ReDim arr(1 To Len(b))
    Dim i
    For i = 1 To Len(b)
    
        arr(i) = Mid(b, i, 1)
        
        If Asc(arr(i)) = 63 Then
            Debug.Print ""
        End If
    Next
    
    Call fReplaceIlleagleCharInFileName(a)
    Debug.Print a
End Sub
'Function f()
'
'    Dim reg, arr, d, n, m
'    arr = Range("a1:a" & Range("a65535").End(xlUp).Row)  '把A:A列放入数组
'    Set d = CreateObject("Scripting.Dictionary")  '申明字典
'    For n = 1 To UBound(arr)
'    d(1) = d(1) & arr(n, 1)  '把所有A列数据合并，并放入字典d(1)的item中
'    Next
'    Set reg = CreateObject("vbscript.regexp")  '申明正则
'    m = d.Item(1)
'    reg.Pattern = "[^\u4e00-\u9fa5]" '正则汉字判断公式
'    reg.Global = True
'    If reg.Replace(m, "") <> "" Then MsgBox "有汉字"  '判断并返回结果
'
'
'End Function
