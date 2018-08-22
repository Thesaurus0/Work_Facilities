VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationEventClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents ExcelAppEvents As Application
Attribute ExcelAppEvents.VB_VarHelpID = -1

Private Sub Class_Initialize()
    Set ExcelAppEvents = Application
End Sub

Private Sub Class_Terminate()
    Set ExcelAppEvents = Nothing
End Sub
 
Private Sub ExcelAppEvents_SheetActivate(ByVal Sh As Object)
    If dictNavigate Is Nothing Then Call fNavagatorInitialize
    
    If dictNavigate.Count <= 0 Then
        dictNavigate.Add 1, ActiveCell.Address(external:=True)
    Else
        dictNavigate.Add dictNavigate.Keys(dictNavigate.Count - 1) + 1, ActiveCell.Address(external:=True)
    End If
    
    dictWbListCurrPos(Sh.Parent.Name) = dictNavigate.Keys(dictNavigate.Count - 1)
    lLastPosBeforeManualActive = 0
    
    Call RefreshRibbonControl("btnBack")
    Call RefreshRibbonControl("btnForward")
End Sub

Private Sub ExcelAppEvents_SheetDeactivate(ByVal Sh As Object)
    Dim sRgAddr As String
    Dim sLastRgAddr As String
    Dim sWb As String
    Dim sSheet As String
    Dim sLastWb As String
    Dim sLastSheet As String
    Dim sActiveWbName As String
    Dim i As Long
    
    If lLastPosBeforeManualActive <= 0 Then Exit Sub
    
    sActiveWbName = ActiveWorkbook.Name
    
    sLastRgAddr = dictNavigate(lLastPosBeforeManualActive)
    sLastWb = Replace(Replace(Split(sLastRgAddr, "]")(0), "[", ""), "'", "")
    sLastSheet = Replace(Split(Split(sLastRgAddr, "]")(1), "!")(0), "'", "")
    
    For i = lLastPosBeforeManualActive + 1 To dictNavigate.Keys(dictNavigate.Count - 1)
        If Not dictNavigate.Exists(i) Then GoTo next_pos
        
        sRgAddr = dictNavigate(i)
        sWb = Replace(Replace(Split(sRgAddr, "]")(0), "[", ""), "'", "")
        sSheet = Replace(Split(Split(sRgAddr, "]")(1), "!")(0), "'", "")
    
        If sWb = sLastWb Then
            dictNavigate.Remove i
        End If
next_pos:
    Next
End Sub

Private Sub ExcelAppEvents_WorkbookActivate(ByVal wb As Workbook)
    If dictWbListCurrPos Is Nothing Then Call fNavagatorInitialize
    If Not dictWbListCurrPos.Exists(wb.Name) Then dictWbListCurrPos.Add wb.Name, 1
    
    If dictNavigate.Count <= 0 Then
        dictNavigate.Add 1, ActiveCell.Address(external:=True)
    Else
        dictNavigate.Add dictNavigate.Keys(dictNavigate.Count - 1) + 1, ActiveCell.Address(external:=True)
    End If
End Sub
 
