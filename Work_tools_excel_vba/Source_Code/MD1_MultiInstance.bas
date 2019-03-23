Attribute VB_Name = "MD1_MultiInstance"
'------------- Code Module --------------

Option Explicit

Declare PtrSafe Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Declare PtrSafe Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare PtrSafe Function IIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef lpiid As UUID) As Long
Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, ByRef riid As UUID, ByRef ppvObject As Object) As Long

Type UUID 'GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Const OBJID_NATIVEOM As Long = &HFFFFFFF0

'------------- Form Module --------------
 
'Sub GetAllWorkbookWindowNames()
Sub Command1_Click()
    On Error GoTo MyErrorHandler

    Dim hWndMain As Long
    hWndMain = FindWindowEx(0&, 0&, "XLMAIN", vbNullString)

    Do While hWndMain <> 0
        GetWbkWindows hWndMain
        hWndMain = FindWindowEx(0&, hWndMain, "XLMAIN", vbNullString)
    Loop

    Exit Sub

MyErrorHandler:
    MsgBox "GetAllWorkbookWindowNames" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description
End Sub

Private Sub GetWbkWindows(ByVal hWndMain As Long)
    On Error GoTo MyErrorHandler

    Dim hWndDesk As Long
    hWndDesk = FindWindowEx(hWndMain, 0&, "XLDESK", vbNullString)

    If hWndDesk <> 0 Then
        Dim hwnd As Long
        hwnd = FindWindowEx(hWndDesk, 0, vbNullString, vbNullString)

        Dim strText As String
        Dim lngRet As Long
        Do While hwnd <> 0
            strText = String$(100, Chr$(0))
            lngRet = GetClassName(hwnd, strText, 100)

            If Left$(strText, lngRet) = "EXCEL7" Then
                GetExcelObjectFromHwnd hwnd
                Exit Sub
            End If

            hwnd = FindWindowEx(hWndDesk, hwnd, vbNullString, vbNullString)
            Loop

        On Error Resume Next
    End If

    Exit Sub

MyErrorHandler:
    MsgBox "GetWbkWindows" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description
End Sub

Public Function GetExcelObjectFromHwnd(ByVal hwnd As Long) As Boolean
    On Error GoTo MyErrorHandler

    Dim fOk As Boolean
    fOk = False

    Dim iid As UUID
    Call IIDFromString(StrPtr(IID_IDispatch), iid)

    Dim obj As Object
    If AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, iid, obj) = 0 Then 'S_OK
        Dim objApp As Excel.Application
        Set objApp = obj.Application
        Debug.Print objApp.Workbooks(1).Name

        Dim myWorksheet As Worksheet
        For Each myWorksheet In objApp.Workbooks(1).Worksheets
            Debug.Print "     " & myWorksheet.Name
            DoEvents
        Next

        fOk = True
    End If

    GetExcelObjectFromHwnd = fOk

    Exit Function

MyErrorHandler:
    MsgBox "GetExcelObjectFromHwnd" & vbCrLf & vbCrLf & "Err = " & Err.Number & vbCrLf & "Description: " & Err.Description
End Function
