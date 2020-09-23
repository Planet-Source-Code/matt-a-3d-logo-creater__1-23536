Attribute VB_Name = "Module1"
Public Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Public Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Public Type MENUITEMINFO
   cbSize As Long
   fMask As Long
   fType As Long
   fState As Long
   wID As Long
   hSubMenu As Long
   hbmpChecked As Long
   hbmpUnchecked As Long
   dwItemData As Long
   dwTypeData As String
   cch As Long
End Type
Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10
Public Const MFT_SEPARATOR = &H800
Public Const MFT_STRING = &H0
Public Const MFS_DEFAULT = &H1000
Public Const MFS_ENABLED = &H0
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal _
   hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As _
   MENUITEMINFO) As Long
Public Type RECT
   left As Long
   top As Long
   right As Long
   bottom As Long
End Type
Public Type TPMPARAMS
   cbSize As Long
   rcExclude As RECT
End Type
Public Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As Long, ByVal _
   fuFlags As Long, ByVal x As Long, ByVal y As Long, ByVal hWnd As Long, lptpm As _
   TPMPARAMS) As Long
Public Const TPM_LEFTALIGN = &H0
Public Const TPM_TOPALIGN = &H0
Public Const TPM_NONOTIFY = &H80
Public Const TPM_RETURNCMD = &H100
Public Const TPM_LEFTBUTTON = &H0
Public Type POINT_TYPE
   x As Long
   y As Long
End Type
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Public Declare Function SetRectEmpty Lib "user32.dll" (lpRect As RECT) As Long



