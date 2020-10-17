VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RegExp_"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mRegExp As RegExp
Private sPattern As String

 
Property Let Pattern(val As String)
    sPattern = val
End Property
'-----------------------------------------------------------
Private Function GetRegExp(Optional asPatten As String = "")
    If mRegExp Is Nothing Then
        Set mRegExp = New VBScript_RegExp_55.RegExp
        mRegExp.IgnoreCase = True
        mRegExp.Global = True
    End If
    
    If A.Nzero(asPatten) Then mRegExp.Pattern = asPatten
End Function

Function ReplaceDatePattern(ByRef sToReplace As String, aDate As Date) As String
    Dim oMatchCollection As MatchCollection
    Dim oMatch As match
    Dim lStartPos As Long
    Dim lEndPos As Long
    Dim lLen As Long
    Dim lPrevEndPos As Long
    
    Dim sDatePattern As String
    Dim sNewStr As String
    Dim sDate As String
        
    GetRegExp
    mRegExp.Pattern = "((yyyy)|(yy)|(mmm)|(mm)|(dd)|(hh)|(ss))+((\W_){0,1}((yyyy)|(yy)|(mmm)|(mm)|(dd)|(hh)|(ss))+)+"
    
    Set oMatchCollection = mRegExp.Execute(sToReplace)
    
    sNewStr = ""
    lPrevEndPos = 0
    For Each oMatch In oMatchCollection
        sDatePattern = oMatch.Value
        
        lStartPos = oMatch.FirstIndex + 1
        lEndPos = oMatch.FirstIndex + oMatch.length
        lLen = oMatch.length
        
        sDate = Format(aDate, sDatePattern)
        
        If Len(sNewStr) <= 0 Then
            sNewStr = Left(sToReplace, lStartPos - 1) & sDate
        Else
            sNewStr = sNewStr & Mid(sToReplace, lPrevEndPos + 1, lStartPos - lPrevEndPos - 1) & sDate
        End If
        
        lPrevEndPos = lEndPos
    Next
    
    If Len(sToReplace) > lPrevEndPos Then
        sNewStr = sNewStr & Right(sToReplace, Len(sToReplace) - lPrevEndPos)
    End If
    
    ReplaceDatePattern = sNewStr
    
    Set oMatch = Nothing
    Set oMatchCollection = Nothing
End Function

Private Sub Class_Initialize()
    Call GetRegExp
End Sub

Private Sub Class_Terminate()
    Set mRegExp = Nothing
End Sub
