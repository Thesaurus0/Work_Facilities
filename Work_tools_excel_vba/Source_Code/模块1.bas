Attribute VB_Name = "Ä£¿é1"
Sub ccaaafa()
'    ThisWorkbook.Worksheets(1).Visible = xlSheetVisible
'    ThisWorkbook.Worksheets(1).Delete
'    Dim a, b
'    a = dictNavigate.Keys
'    b = dictNavigate.Items
'
'    Dim c, d
'    c = dictWbListCurrPos.Keys
'    d = dictWbListCurrPos.Items
'
'    Dim e
'    e = lLastPosBeforeManualActive
    Dim a
    a = Timer()
End Sub
Sub ºê1()
Attribute ºê1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim sTmpFile As String
    Dim iFileNum
    
    sTmpFile = "H:\Work_Facilities\Work_tools_excel_vba\a.txt"
Call fCreateTextFileInUnicode(sTmpFile)
    
    iFileNum = FreeFile
    Open sTmpFile For Output As #iFileNum
     
                Print #iFileNum, "ÖÐ¹ú"
     Close #iFileNum
End Sub
