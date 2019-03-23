Attribute VB_Name = "MA_0_RibbonUI"
Option Explicit

#If VBA7 And Win64 Then  'Win64
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Public gRibbonObj_Utility As IRibbonUI
 

'=============================================================
Sub RefreshRibbonControl(sControl As String)
    If gRibbonObj_Utility Is Nothing Then Call fGetRibbonReference
    Call gRibbonObj_Utility.InvalidateControl(sControl)
End Sub
Sub subRefreshRibbon()
    fGetRibbonReference.Invalidate
End Sub
Sub Excel_Utility_Onload(ribbon As IRibbonUI)
  Set gRibbonObj_Utility = ribbon
  
  fCreateAddNameUpdateNameWhenExists "nmRibbonPointer", ObjPtr(ribbon)
  'Names("nmRibbonPointer").RefersTo = ObjPtr(ribbon)
  
  gRibbonObj_Utility.ActivateTab "ERP_2010"
  ThisWorkbook.Saved = True
End Sub
Function fGetRibbonReference() As IRibbonUI
    If Not gRibbonObj_Utility Is Nothing Then Set fGetRibbonReference = gRibbonObj_Utility: Exit Function
    
    Dim objRibbon As Object
    Dim lRibPointer As LongPtr
    
    lRibPointer = [nmRibbonPointer]
    CopyMemory objRibbon, lRibPointer, LenB(lRibPointer)
    
    Set fGetRibbonReference = objRibbon
    Set gRibbonObj_Utility = objRibbon
    Set objRibbon = Nothing
End Function

'---------------------------------------------------------------------
Sub Utility_onAction(control As IRibbonControl)
    Call fGetControlAttributes(control, "ACTION")
End Sub
Sub Utility_getImage(control As IRibbonControl, ByRef imageMso)
    Call fGetControlAttributes(control, "IMAGE", imageMso)
End Sub
Sub Utility_getLabel(control As IRibbonControl, ByRef label)
    Call fGetControlAttributes(control, "LABEL", label)
End Sub
Sub Utility_getSize(control As IRibbonControl, ByRef size)
    Call fGetControlAttributes(control, "SIZE", size)
End Sub
Sub Utility_getEnabled(control As IRibbonControl, ByRef val)
    If control.ID = "btnBack" Then
        val = CBool(fNavigateStackNextPositionToMoveBack > 0)
    ElseIf control.ID = "btnForward" Then
        val = CBool(fNavigateStackNextPositionToMoveForWard > 0)
    Else
        Call fGetControlAttributes(control, "ENABLED", val)
    End If
End Sub
Sub Utility_getShowImage(control As IRibbonControl, ByRef ShowImage)
    Call fGetControlAttributes(control, "SHOW_IMAGE", ShowImage)
End Sub
Sub Utility_getSupertip(control As IRibbonControl, ByRef Supertip)
    Call fGetControlAttributes(control, "SUPERTIP", Supertip)
End Sub
Sub Utility_getScreentip(control As IRibbonControl, ByRef screentip)
    Call fGetControlAttributes(control, "SCREENTIP", screentip)
End Sub

'================== toggle button common function===========================================
Sub ToggleButtonToSwitchSheet_onAction(control As IRibbonControl, pressed As Boolean)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If Not sht Is Nothing Then
        fToggleSheetVisibleFromUIRibbonControl pressed, sht, control
    End If
    Set sht = Nothing
End Sub

Sub ToggleButtonToSwitchSheet_getPressed(control As IRibbonControl, ByRef returnedVal)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If sht Is Nothing Then
        returnedVal = False
    Else
        returnedVal = (sht.Visible = xlSheetVisible And ActiveSheet Is sht)
    End If
End Sub
Function fGetSheetByUIRibbonTag(ByVal asButtonTag As String) As Worksheet
    Dim sht As Worksheet
    
    If fSheetExistsByCodeName(asButtonTag, sht) Then
        Set fGetSheetByUIRibbonTag = sht
    Else
        MsgBox "The button's Tag is not corresponding to any worksheet in this workbook, please check the customUI.xml you prepared," _
            & " The design thought is that the button's tag is the name of a sheet, so that the common function ToggleButtonToSwitchSheet_onAction/getPressed can get a worksheet."
    End If
    Set sht = Nothing
End Function
Function fToggleSheetVisibleFromUIRibbonControl(ByVal pressed As Boolean, sht As Worksheet, control As IRibbonControl)
    If pressed Then
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        End If
    Else
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        Else
            fVeryHideSheet sht
        End If
    End If
    
    'fGetRibbonReference.InvalidateControl (control.id)
    fGetRibbonReference.Invalidate
End Function

'---------------------------------------------------------------------


'================ dev facilities ==============================================
Sub btnGenNumberList_onAction(control As IRibbonControl)
    sub_GenNumberList
End Sub
Sub btnGenAlphabetList_onAction(control As IRibbonControl)
    sub_GenAlpabetList
End Sub
Sub btnListAllActiveXOnCurrSheet_onAction(control As IRibbonControl)
    Sub_ListActiveXControlOnActiveSheet
End Sub
Sub btnResetOnError_onAction(control As IRibbonControl)
    sub_ResetOnError_Initialize
End Sub
'------------------------------------------------------------------------------

Function fGetControlAttributes(control As IRibbonControl, sType As String, Optional ByRef val)
    If Not (sType = "LABEL" Or sType = "IMAGE" Or sType = "SIZE" Or sType = "ENABLED" Or sType = "ACTION" Or sType = "SHOW_IMAGE" Or sType = "SCREENTIP" Or sType = "SUPERTIP") Then
        fErr "wrong param to fGetControlAttributes: " & vbCr & "sType=" & sType & vbCr & "control=" & control.ID
    End If
    
    Select Case control.ID
        Case "btnBack"
            Select Case sType
                Case "LABEL":       val = "后退" & vbCr & "(Alt + 向左箭头)"
                Case "IMAGE":       val = "ScreenNavigatorBack"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                Case "ACTION":      Call subMain_NavigateBack
            End Select
        Case "btnForward"
            Select Case sType
                Case "LABEL":   val = "前进" & vbCr & "(Alt + 向右箭头)"
                Case "IMAGE":   val = "ScreenNavigatorForward"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                Case "ACTION":  Call subMain_NavigateForward
            End Select
        Case "btnFilterBySelected"
            Select Case sType
                Case "LABEL":   val = "以所选过滤"
                Case "IMAGE":   val = "FilterBySelection"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":  Call Sub_FilterBySelectedCells
            End Select
        Case "btnSortSheetBySelected"
            Select Case sType
                Case "LABEL":   val = "以所选排序"
                Case "IMAGE":   val = "SortUp"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":  Call sub_SortBySelectedCells
            End Select
            
        Case "btnRemoveFilter"
            Select Case sType
                Case "LABEL":   val = "清除过滤"
                Case "IMAGE":   val = "FilterClearAllFilters"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":  Call Sub_RemoveFilterForAcitveSheet
            End Select
        Case "btnShowAllVeryHideSheets"
            Select Case sType
                Case "LABEL":   val = "Show All Very Hide"
                Case "IMAGE":   val = "NameDefineMenu"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call fShowAllVeryHideSheets
            End Select
        Case "btnOpenFileLocation"
            Select Case sType
                Case "LABEL":   val = "Open File Location"
                Case "IMAGE":   val = "FileOpen"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_OpenAcitveWorkbookLocation
            End Select
        Case "btnCopyFileFullPath"
            Select Case sType
                Case "LABEL":   val = "WorkBook Full Path"
                Case "IMAGE":   val = "FileOpen"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_DisplayWorkbookFullPath
            End Select
        Case "btnBackupActiveWorkbook"
            Select Case sType
                Case "LABEL":   val = "Backup this WorkBook"
                Case "IMAGE":   val = "FileSaveAs"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_BackupActiveWorkbook
            End Select
            
        Case "btnExportSourceCode"
            Select Case sType
                Case "LABEL":   val = "Export Source Code"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call btnExportSourceCode_onAction
            End Select
        Case "btnExportSourceCodeThisAddIn"
            Select Case sType
                Case "LABEL":   val = "Export This AddIn Source Code"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call sub_ExportSourceCodeToFolder_ThisWorkbook
            End Select
        Case "btnValidateLocalSourceCodeFile"
            Select Case sType
                Case "LABEL":   val = "Validate Local Source Library"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMainValidateSourceCodeFile
            End Select
        Case "btnValidateMacroWithLocal"
            Select Case sType
                Case "LABEL":   val = "Validate Macro(Synchronize)"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_ValidateMacroWithLocal
            End Select
            
        Case "btnCompareWithCommLib"
            Select Case sType
                Case "LABEL":   val = "Compare with Comm Lib"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_CompareWithCommonLibFolder
            End Select
        Case "btnSyncWithCommLib"
            Select Case sType
                Case "LABEL":   val = "Sync Modules(from Common Lib)"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_SyncWithCommonLib
            End Select
        Case "btnSyncWithSelfRevised"
            Select Case sType
                Case "LABEL":   val = "Sync Modules(from Self Revised)"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_SyncWithSelfRevised
            End Select
        Case "btnListAllFunctions"
            Select Case sType
                Case "LABEL":   val = "List All Functions"
                Case "IMAGE":   val = "SmartArtAddBullet"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call sub_ListAllFunctionsOfActiveWorkbook
            End Select
        Case "btnCompareTwoMacros"
            Select Case sType
                Case "LABEL":   val = "Compare Two Macros"
                Case "IMAGE":   val = "SmartArtAddBullet"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_Compare2MacroFiles
            End Select
        Case "btnListAllSheetName"
            Select Case sType
                Case "LABEL":   val = "List All Sheets Name"
                Case "IMAGE":   val = "SmartArtAddBullet"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_ListAllSheets
            End Select
        Case "btnListAllSheetCodeName"
            Select Case sType
                Case "LABEL":   val = "List All Sheets CodeName"
                Case "IMAGE":   val = "SmartArtAddBullet"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_ListAllSheetsCodeName
            End Select
        Case "btnListAllFunctionsInLocalFile"
            Select Case sType
                Case "LABEL":   val = "List All Funs In File Lib"
                Case "IMAGE":   val = "SmartArtAddBullet"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_ListAllFunctionsInFileLibrary
            End Select
            
        Case "btnScanUselessFunctions"
            Select Case sType
                Case "LABEL":   val = "Scan Useless Functions"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_ScanUselessFunctions
            End Select
        Case "btnCommentOutScanUselessFunctions"
            Select Case sType
                Case "LABEL":   val = "Comment Out Useless Func"
                Case "IMAGE":   val = "ViewPrintLayoutView"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_CommentOutScanUselessFunctions
            End Select
            
        Case "btnCloneMacro"
            Select Case sType
                Case "LABEL":   val = "Clone Macro"
                Case "IMAGE":   val = "PageOrientationPortraitLandscape"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_CloneMacro
            End Select
        Case "btnGenCodeSnippet"
            Select Case sType
                Case "LABEL":   val = "Gen Code Snippet"
                Case "IMAGE":   val = "PageOrientationPortraitLandscape"
                Case "SIZE":        val = "false"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_GenCodeSnippet
            End Select
            
    End Select
End Function

Sub testaaaaa()
End Sub
