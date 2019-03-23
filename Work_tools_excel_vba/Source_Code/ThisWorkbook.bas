VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Set ApplicationClass = Nothing
    Set dictNavigate = Nothing
    Set dictWbListCurrPos = Nothing

    Application.OnKey "%{LEFT}"
    Application.OnKey "%{RIGHT}"
    Application.OnKey "^{ENTER}"
End Sub
 

Private Sub Workbook_Open()
    Call sub_WorkBookInitialization

    Application.OnKey "%{LEFT}", "subMain_NavigateBack"
    Application.OnKey "%{RIGHT}", "subMain_NavigateForward"
    Application.OnKey "^{ENTER}", "subMain_OpenAcitveWorkbookLocation"
    
    ThisWorkbook.Saved = True
End Sub

Public Sub sub_WorkBookInitialization()
    'Call fNavagatorInitialize
    
End Sub
