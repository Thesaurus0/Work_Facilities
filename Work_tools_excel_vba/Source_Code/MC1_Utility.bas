Attribute VB_Name = "MC1_Utility"
Option Explicit
Option Base 1

Dim StartTime As Single

Public ApplicationClass As ApplicationEventClass

Function fNavagatorInitialize()
    Set dictNavigate = New Dictionary
    Set dictWbListCurrPos = New Dictionary
    lLastPosBeforeManualActive = 0
    
    Call fConnectEventHandler
End Function
Function fConnectEventHandler()
    On Error Resume Next
    Set ApplicationClass = New ApplicationEventClass
End Function

Function fNavigateStackNextPositionToMoveBack(Optional ByRef sWb As String, Optional ByRef sSheet As String) As Long
    Dim lCurrPos As Long
    Dim lNextPos As Long
    Dim sRgAddr As String
    Dim sActiveWbName As String
    Dim i As Long
    
    If dictNavigate Is Nothing Then Call fNavagatorInitialize
    
    lNextPos = 0
    If dictNavigate.Count <= 0 Then Exit Function
    
    sActiveWbName = ActiveWorkbook.Name
    lCurrPos = dictWbListCurrPos(sActiveWbName)

    For i = lCurrPos - 1 To dictNavigate.Keys(0) Step -1
        If Not dictNavigate.Exists(i) Then GoTo next_pos
        
        sRgAddr = dictNavigate(i)
        sWb = Replace(Replace(Split(sRgAddr, "]")(0), "[", ""), "'", "")
        sSheet = Replace(Split(Split(sRgAddr, "]")(1), "!")(0), "'", "")
    
        If sWb <> sActiveWbName Then GoTo next_pos
        
        If fSheetExists(sSheet, , Workbooks(sWb)) Then
            If sSheet = Workbooks(sWb).ActiveSheet.Name Then
                dictNavigate.Remove (i)
                GoTo next_pos
            Else
                lNextPos = i
                Exit For
            End If
        Else
            dictNavigate.Remove (i)
        End If
next_pos:
    Next
    
    fNavigateStackNextPositionToMoveBack = lNextPos
End Function
Function fNavigateStackNextPositionToMoveForWard(Optional ByRef sWb As String, Optional ByRef sSheet As String) As Long
    Dim lCurrPos As Long
    Dim lNextPos As Long
    Dim sRgAddr As String
    Dim sActiveWbName As String
    Dim i As Long
    
    lNextPos = 0
    If dictNavigate.Count <= 0 Then Exit Function
    
    sActiveWbName = ActiveWorkbook.Name
    lCurrPos = dictWbListCurrPos(sActiveWbName)
    If lCurrPos >= dictNavigate.Keys(dictNavigate.Count - 1) Then Exit Function

    For i = lCurrPos + 1 To dictNavigate.Keys(dictNavigate.Count - 1)
        If Not dictNavigate.Exists(i) Then GoTo next_pos
        
        sRgAddr = dictNavigate(i)
        sWb = Replace(Replace(Split(sRgAddr, "]")(0), "[", ""), "'", "")
        sSheet = Replace(Split(Split(sRgAddr, "]")(1), "!")(0), "'", "")
    
        If sWb <> sActiveWbName Then GoTo next_pos
        
        If fSheetExists(sSheet, , Workbooks(sWb)) Then
            If sSheet = Workbooks(sWb).ActiveSheet.Name Then
                dictNavigate.Remove (i)
                GoTo next_pos
            Else
                lNextPos = i
                Exit For
            End If
        Else
            dictNavigate.Remove (i)
        End If
next_pos:
    Next
    
    fNavigateStackNextPositionToMoveForWard = lNextPos
End Function

Sub subMain_NavigateBack()
    Dim lNextPos As Long
    Dim sWb As String
    Dim sSheet As String
    
    lNextPos = fNavigateStackNextPositionToMoveBack(sWb, sSheet)
    
    If lNextPos > 0 Then
        Application.EnableEvents = False
        
        Workbooks(sWb).Activate
        Workbooks(sWb).Worksheets(sSheet).Visible = xlSheetVisible
        Workbooks(sWb).Worksheets(sSheet).Activate
        
        Application.EnableEvents = True
        
        dictWbListCurrPos(sWb) = lNextPos
        lLastPosBeforeManualActive = lNextPos
    End If
    
    Call RefreshRibbonControl("btnBack")
    Call RefreshRibbonControl("btnForward")
End Sub

Sub subMain_NavigateForward()
    Dim lNextPos As Long
    Dim sWb As String
    Dim sSheet As String
    
    lNextPos = fNavigateStackNextPositionToMoveForWard(sWb, sSheet)
    If lNextPos > 0 Then
        Application.EnableEvents = False
        
        Workbooks(sWb).Activate
        Workbooks(sWb).Worksheets(sSheet).Visible = xlSheetVisible
        Workbooks(sWb).Worksheets(sSheet).Activate
        
        Application.EnableEvents = True
        
        dictWbListCurrPos(sWb) = lNextPos
        lLastPosBeforeManualActive = lNextPos
    End If
    
    Call RefreshRibbonControl("btnBack")
    Call RefreshRibbonControl("btnForward")
End Sub

Function fStartTime()
    StartTime = Timer
End Function
Function fHowLong()
    Debug.Print "Total Time: " & Format(Timer - StartTime, "0.00000000000") & " seconds"
End Function

Function fShowAllVeryHideSheets(Optional wb As Workbook)
    Dim wbIsnothing   As Boolean
    
    wbIsnothing = CBool(wb Is Nothing)
    If wbIsnothing Then
        Set wb = ActiveWorkbook
    End If
    
    Dim sht As Worksheet
    
    For Each sht In wb.Worksheets
        sht.Visible = xlSheetVisible
    Next
    
    Set sht = Nothing
    If wbIsnothing Then Set wb = Nothing
End Function

Function fOpenAcitveWorkbookLocation(Optional wb As Workbook)
    Dim wbIsnothing   As Boolean
    
    wbIsnothing = CBool(wb Is Nothing)
    If wbIsnothing Then Set wb = ActiveWorkbook
    
    If Len(wb.Path) <= 0 Then Exit Function
    
    'Shell "explorer.exe /e,/select, " & wb.FullName, vbMaximizedFocus
    Call fOpenFile(wb.Path)
    
    If wbIsnothing Then Set wb = Nothing
End Function
Sub unprotected()
    If Hook Then
        MsgBox "VBA Project is unprotected!", vbInformation, "*****"
    End If
End Sub

