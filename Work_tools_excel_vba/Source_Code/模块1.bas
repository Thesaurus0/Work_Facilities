Attribute VB_Name = "Ä£¿é1"
Option Explicit


#If Win64 Then

    Private Declare PtrSafe Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As LongPtr) As LongPtr
    Private Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As UUID) As LongPtr
    Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As LongPtr, ByVal dwId As LongPtr, ByRef riid As UUID, ByRef ppvObject As Object) As LongPtr

#Else

    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function IIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As UUID) As Long
    Private Declare Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long

#End If

Type UUID 'GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Const OBJID_NATIVEOM As LongPtr = &HFFFFFFF0

' Run as entry point of example
Public Sub Test()

Dim i As Long
Dim xlApps() As Application

    If GetAllExcelInstances(xlApps) Then
        For i = LBound(xlApps) To UBound(xlApps)
            If xlApps(i).Workbooks(1).Name <> ThisWorkbook.Name Then
                MsgBox (xlApps(i).Workbooks(1).Name)
            End If
        Next
    End If

End Sub

' Actual public facing function to be called in other code
Public Function GetAllExcelInstances(xlApps() As Application) As Long

On Error GoTo MyErrorHandler

Dim n As Long
#If Win64 Then
    Dim hWndMain As LongPtr
#Else
    Dim hWndMain As Long
#End If
Dim app As Application

    ' Cater for 100 potential Excel instances, clearly could be better
    ReDim xlApps(1 To 100)

    hWndMain = FindWindowEx(0&, 0&, "XLMAIN", vbNullString)

    Do While hWndMain <> 0
        Set app = GetExcelObjectFromHwnd(hWndMain)
        If Not (app Is Nothing) Then
            If n = 0 Then
                n = n + 1
                Set xlApps(n) = app
            ElseIf checkHwnds(xlApps, app.hwnd) Then
                n = n + 1
                Set xlApps(n) = app
            End If
        End If
        hWndMain = FindWindowEx(0&, hWndMain, "XLMAIN", vbNullString)
    Loop

    If n Then
        ReDim Preserve xlApps(1 To n)
        GetAllExcelInstances = n
    Else
        Erase xlApps
    End If

    Exit Function

MyErrorHandler:
    MsgBox "GetAllExcelInstances" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description

End Function

#If Win64 Then
    Private Function checkHwnds(xlApps() As Application, hwnd As LongPtr) As Boolean
#Else
    Private Function checkHwnds(xlApps() As Application, hwnd As Long) As Boolean
#End If

Dim i As Integer

    For i = LBound(xlApps) To UBound(xlApps)
        If xlApps(i).hwnd = hwnd Then
            checkHwnds = False
            Exit Function
        End If
    Next i

    checkHwnds = True

End Function

#If Win64 Then
    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As LongPtr) As Application
#Else
    Private Function GetExcelObjectFromHwnd(ByVal hWndMain As Long) As Application
#End If

On Error GoTo MyErrorHandler

#If Win64 Then
    Dim hWndDesk As LongPtr
    Dim hwnd As LongPtr
#Else
    Dim hWndDesk As Long
    Dim hwnd As Long
#End If
Dim strText As String
Dim lngRet As Long
Dim iid As UUID
Dim obj As Object

    hWndDesk = FindWindowEx(hWndMain, 0&, "XLDESK", vbNullString)

    If hWndDesk <> 0 Then

        hwnd = FindWindowEx(hWndDesk, 0, vbNullString, vbNullString)

        Do While hwnd <> 0

        strText = String$(100, Chr$(0))
        lngRet = CLng(GetClassName(hwnd, strText, 100))

        If Left$(strText, lngRet) = "EXCEL7" Then

            Call IIDFromString(StrPtr(IID_IDispatch), iid)

            If AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, iid, obj) = 0 Then 'S_OK

                Set GetExcelObjectFromHwnd = obj.Application
                Exit Function

            End If

        End If

        hwnd = FindWindowEx(hWndDesk, hwnd, vbNullString, vbNullString)
        Loop

        On Error Resume Next

    End If

    Exit Function
MyErrorHandler:
    MsgBox "GetExcelObjectFromHwnd" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description

End Function

Sub tesasdfasdgasdg()
    Dim dict As New Dictionary
     'Call fPastePotentialRiskLinkstoManuallyHandleSheet(dict, ActiveWorkbook, Array("Type", "sheet name", "cell", "formula1", "formula2", "validation type"))
     
     Call fListAllValidation(ActiveWorkbook)
End Sub

Sub aaaasdfasdf()
    Dim a
    a = ActiveWorkbook.LinkSources
    
    ActiveWorkbook.ChangeLink a(1), ActiveWorkbook.FullName
End Sub

Function fListAllExternalLinksbbbbbbbb()
    Dim rg As Range
    Dim eachcell As Range
    Dim dict As Dictionary
    Dim sht As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    Set dict = New Dictionary
    For Each sht In wb.Worksheets
        On Error Resume Next
        Set rg = sht.UsedRange.SpecialCells(xlCellTypeFormulas)
    
        If rg Is Nothing Then GoTo next_one
        
        For Each eachcell In rg.Cells
            If InStr(1, eachcell.Formula, "[") > 0 Then
                If eachcell.MergeCells Then
                    If eachcell.Address = eachcell.MergeArea.Cells(1, 1).Address Then
                        dict.Add "Formula: " & DELIMITER _
                                & "'" & sht.Name & DELIMITER _
                                & "'" & eachcell.MergeArea.Address & DELIMITER _
                                & "'" & eachcell.Formula, ""
                    End If
                Else
                    dict.Add "Formula: " & DELIMITER _
                            & "'" & sht.Name & DELIMITER _
                            & "'" & eachcell.Address & DELIMITER _
                            & "'" & eachcell.Formula, ""
                End If
            End If
        Next
        
next_one:
    Next
    
    Dim extLink
    For Each extLink In wb.LinkSources
        If Not fAreSame(CStr(extLink), ThisWorkbook.FullName) Then
            dict.Add "Link: " & DELIMITER _
                    & "'" & sht.Name & DELIMITER _
                    & "'" & extLink & DELIMITER _
                    & "", ""
        End If
    Next
    On Error GoTo 0
    
    Call fPastePotentialRiskLinkstoManuallyHandleSheet(dict, wb, Array("Type", "sheet name", "cell", "formula/link name"))
    Set dict = Nothing
End Function
Function fListAllNamesaaaaaaaaaaaaa()
    Dim eachName As Name
    Dim sName As String
    Dim dict As Dictionary
    Set dict = New Dictionary
    Dim sht As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    For Each sht In wb.Worksheets
        For Each eachName In sht.Names
            sName = eachName.Name
            
            If InStr(1, sName, "!") > 0 Then sName = Split(sName, "!")(1)
            On Error Resume Next
                dict.Add "Name: " & DELIMITER _
                        & "'" & sName & DELIMITER _
                        & "'" & eachName.Parent.Name & DELIMITER _
                        & "'" & eachName.RefersTo & DELIMITER _
                        & "'" & eachName.value & DELIMITER _
                        & "'" & eachName.RefersToLocal, ""
            On Error GoTo 0
        Next
    Next
    
    For Each eachName In wb.Names
        sName = eachName.Name
        
        If InStr(1, sName, "!") > 0 Then sName = Split(sName, "!")(1)
        On Error Resume Next
            dict.Add "Name: " & DELIMITER _
                    & "'" & sName & DELIMITER _
                    & "'" & eachName.Parent.Name & DELIMITER _
                    & "'" & eachName.RefersTo & DELIMITER _
                    & "'" & eachName.value & DELIMITER _
                    & "'" & eachName.RefersToLocal, ""
        On Error GoTo 0
    Next
    
    Call fPastePotentialRiskLinkstoManuallyHandleSheet(dict, wb, Array("Type", "name's name", "whhere name is", "refers to", "value", "RefersToLocal"))
    Set dict = Nothing
End Function

Function aaaaaaaaaaaaaaaaaaaaa()
    Dim rg As Range
    Dim eachcell As Range
    Dim dict As Dictionary
    Dim sht As Worksheet
    Dim vD As Validation
    Dim vdType As XlDVType
    Dim sType As String
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    
    Set dict = New Dictionary
    For Each sht In wb.Worksheets
        On Error Resume Next
        Set rg = sht.UsedRange.SpecialCells(xlCellTypeAllValidation)
    
        If rg Is Nothing Then GoTo next_one
        
        For Each eachcell In rg.Cells
            Set vD = eachcell.Validation
            vdType = vD.Type
             
            Select Case vdType
                Case xlValidateCustom
                    sType = "xlValidateCustom"
                Case xlValidateDate
                    sType = "xlValidateDate"
                Case xlValidateDecimal
                    sType = "xlValidateDecimal"
                Case xlValidateInputOnly
                    sType = "xlValidateInputOnly"
                    vD.Delete
                    GoTo next_one
                Case xlValidateList
                    sType = "xlValidateList"
                Case xlValidateTextLength
                    sType = "xlValidateTextLength"
                Case xlValidateTime
                    sType = "xlValidateTime"
                Case xlValidateWholeNumber
                    sType = "xlValidateWholeNumber"
                Case Else
                    sType = "Unknow validation Type"
            End Select
            
            
            If eachcell.MergeCells Then
                If eachcell.Address = eachcell.MergeArea.Cells(1, 1).Address Then
                    dict.Add "Validation: " & DELIMITER _
                            & "'" & sht.Name & DELIMITER _
                            & "'" & eachcell.MergeArea.Address & DELIMITER _
                            & "'" & vD.Formula1 & DELIMITER _
                            & "'" & vD.Formula2 & DELIMITER _
                            & "'" & sType, ""
                End If
            Else
                dict.Add "Validation: " & DELIMITER _
                        & "'" & sht.Name & DELIMITER _
                        & "'" & eachcell.Address & DELIMITER _
                        & "'" & vD.Formula1 & DELIMITER _
                        & "'" & vD.Formula2 & DELIMITER _
                        & "'" & sType, ""
            End If
        
        Next
        On Error GoTo 0
next_one:
    Next
     
    Call fPastePotentialRiskLinkstoManuallyHandleSheet(dict, wb, Array("Type", "sheet name", "cell", "formula1", "formula2", "validation type"))
    Set dict = Nothing
End Function

