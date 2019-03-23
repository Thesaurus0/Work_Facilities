Attribute VB_Name = "Common_API"
Option Explicit
'https://www.spreadsheet1.com/how-to-copy-strings-to-clipboard-using-excel-vba.html
'https://docs.microsoft.com/en-us/office/vba/access/concepts/windows-api/retrieve-information-from-the-clipboard
'Use An API To Put Text In Windows Clipboard
 
Public Const GHND = &H42
Public Const CF_TEXT = 1
Public Const MAXSIZE = 4096
#If Mac Then
    ' ignore
#Else
    #If VBA7 Then
        Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
        Declare PtrSafe Function CloseClipboard Lib "User32" () As Long
        Declare PtrSafe Function OpenClipboard Lib "User32" (ByVal hwnd As LongPtr) As LongPtr
        Declare PtrSafe Function EmptyClipboard Lib "User32" () As Long
        Declare PtrSafe Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As LongPtr
        Declare PtrSafe Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
        
        Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As LongPtr) As LongPtr
        Declare PtrSafe Function GetClipboardData Lib "user32.dll" (ByVal wFormat As LongPtr) As LongPtr
        Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As Long
    #Else
        Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
        Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
        Declare Function CloseClipboard Lib "User32" () As Long
        Declare Function OpenClipboard Lib "User32" (ByVal hwnd As Long) As Long
        Declare Function EmptyClipboard Lib "User32" () As Long
        Declare Function lstrcpy Lib "kernel32" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
        Declare Function SetClipboardData Lib "User32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
        Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal wFormat As Long) As Long
        Declare Function GetClipboardData Lib "user32.dll" (ByVal wFormat As Long) As Long
        Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    #End If
#End If
 

Function ClipBoard_SetData(MyString As String)
    #If Mac Then
        With New MSForms.DataObject
            .SetText MyString
            .PutInClipboard
        End With
    #Else
        #If VBA7 Then
            Dim hGlobalMemory As LongPtr
            Dim hClipMemory   As LongPtr
            Dim lpGlobalMemory    As LongPtr
        #Else
            Dim hGlobalMemory As Long
            Dim hClipMemory   As Long
            Dim lpGlobalMemory    As Long
        #End If

        Dim x                 As Long

        ' Allocate moveable global memory.
       '-------------------------------------------
       hGlobalMemory = GlobalAlloc(GHND, Len(MyString) + 1)

        ' Lock the block to get a far pointer
       ' to this memory.
       lpGlobalMemory = GlobalLock(hGlobalMemory)

        ' Copy the string to this global memory.
       lpGlobalMemory = lstrcpy(lpGlobalMemory, MyString)

        ' Unlock the memory.
       If GlobalUnlock(hGlobalMemory) <> 0 Then
            MsgBox "Could not unlock memory location. Copy aborted."
            GoTo OutOfHere2
        End If

        ' Open the Clipboard to copy data to.
       If OpenClipboard(0&) = 0 Then
            MsgBox "Could not open the Clipboard. Copy aborted."
            Exit Function
        End If

        ' Clear the Clipboard.
       x = EmptyClipboard()

        ' Copy the data to the Clipboard.
       hClipMemory = SetClipboardData(CF_TEXT, hGlobalMemory)

OutOfHere2:

        If CloseClipboard() = 0 Then
            MsgBox "Could not close Clipboard."
        End If
    #End If
End Function
Function ClipBoard_GetData()
    Dim MyString As String

    #If VBA7 Then
        Dim hClipMemory As LongPtr
        Dim lpClipMemory As LongPtr
        Dim RetVal  As LongPtr
    #Else
        Dim hClipMemory As Long
        Dim lpClipMemory As Long
        Dim RetVal As Long
    #End If

   If OpenClipboard(0&) = 0 Then
      MsgBox "Cannot open Clipboard. Another app. may have it open"
      Exit Function
   End If
          
   ' Obtain the handle to the global memory
   ' block that is referencing the text.
   hClipMemory = GetClipboardData(CF_TEXT)
   If IsNull(hClipMemory) Then
      MsgBox "Could not allocate memory"
      GoTo OutOfHere
   End If
 
   ' Lock Clipboard memory so we can reference
   ' the actual data string.
   lpClipMemory = GlobalLock(hClipMemory)
 
   If Not IsNull(lpClipMemory) Then
      MyString = Space$(MAXSIZE)
      RetVal = lstrcpy(MyString, lpClipMemory)
      RetVal = GlobalUnlock(hClipMemory)
       
      ' Peel off the null terminating character.
      MyString = Mid(MyString, 1, InStr(1, MyString, Chr$(0), 0) - 1)
   Else
      MsgBox "Could not lock memory to copy string from."
   End If
 
OutOfHere:
   RetVal = CloseClipboard()
   ClipBoard_GetData = MyString
End Function
