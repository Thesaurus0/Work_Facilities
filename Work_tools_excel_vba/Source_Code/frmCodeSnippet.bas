VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCodeSnippet 
   Caption         =   "UserForm1"
   ClientHeight    =   7695
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13590
   OleObjectBlob   =   "frmCodeSnippet.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmCodeSnippet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 
Private Sub cbOK_Click()
    If optEnum.value Then
        gsRtnValueOfForm = "ENUM"
    ElseIf optFor.value Then
        gsRtnValueOfForm = "FOR_LOOP"
    ElseIf optSelectCase.value Then
        gsRtnValueOfForm = "SELECT_CASE"
    Else
    End If
    
    Call fSetValue(RANGE_CodeSnippet_AutoOk, cbAutoOk.value)
    Unload Me
End Sub

Private Sub optEnum_Click()
    If optEnum.value Then Call cbOK_Click
End Sub

Private Sub optFor_Click()
    If optFor.value Then Call cbOK_Click
End Sub

Private Sub optSelectCase_Click()
    If optSelectCase.value Then Call cbOK_Click
End Sub

Private Sub UserForm_Initialize()
    cbAutoOk.value = CBool(UCase(fGetValue(RANGE_CodeSnippet_AutoOk)) = "TRUE")
End Sub
