Attribute VB_Name = "Common_SysConfig"
Option Explicit
Option Base 1

Public Const BUSINESS_ERROR_NUMBER = 10000
Public Const CONFIG_ERROR_NUMBER = 20000
Public Const DELIMITER = "|"

Public gErrNum As Long
Public gErrMsg As String
Public gsRtnValueOfForm As String
Public Const CONST_FAIL = "FAIL"
Public Const CONST_SUCCESS = "SUCCESS"
Public Const CONST_CANCEL = "CANCEL"

'=======================================
Public gFSO As FileSystemObject
Public gRegExp As VBScript_RegExp_55.RegExp
Public Const PW_PROTECT_SHEET = "abcd1234"

Public gProBar As ProgressBar
'=======================================

Public dictNavigate As Dictionary
Public dictWbListCurrPos As Dictionary
Public lLastPosBeforeManualActive As Long
'=======================================

Public Const SOURCE_CODE_LIBRARY_FILE = "H:\Work_Facilities\Work_tools_excel_vba\[Important]_All_Common_Functions.bas"
'Public Const SOURCE_CODE_LIBRARY_FILE = "H:\Work_Facilities\Work_tools_excel_vba\a.txt"
Public Const COMPARE_TMP_FILE_LEFT = "H:\Work_Facilities\Often_Text_Files\Left.bas"
Public Const COMPARE_TMP_FILE_RIGHT = "H:\Work_Facilities\Often_Text_Files\Right.bas"
Public Const BEYOND_COMPARE_EXE = "D:\Program Files\Beyond Compare 4\BCompare.exe"
 
Public Const RANGE_LeftMacroToCompare = "A1"    'save the 1st macro file name user input
Public Const RANGE_RightMacroToCompare = "A2"    'save the 2nd macro file name user input
Public Const RANGE_LeftMacroAlreadyOpened = "A3"    'save the value indicating if the 1st macro file name is already opened
Public Const RANGE_RightMacroAlreadyOpened = "A4"    'save the value indicating if the 2nd macro file name is already opened
Public Const RANGE_LeftMacroAlreadyExported = "A5"    'save the value indicating if the 1st macro file name is already opened
Public Const RANGE_RightMacroAlreadyExported = "A6"    'save the value indicating if the 2nd macro file name is already opened
Public Const RANGE_TargetMacroToSyncWithCommLib = "A7"
Public Const RANGE_CommonLibFilesSelected = "A8"
Public Const RANGE_CommonLibFolderSelected = "A9"
Public Const RANGE_SyncWithCommLibWhichFunction = "A10"
Public Const RANGE_ScanUselessOnebyOneModule = "A11"
Public Const RANGE_CloneMacro_Source = "A12"
Public Const RANGE_CloneMacro_Target = "A13"
Public Const RANGE_CodeSnippet_AutoOk = "A14"
