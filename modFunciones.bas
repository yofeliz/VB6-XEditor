Attribute VB_Name = "modFunciones"
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function CreateCaret Lib "user32" (ByVal hwnd As Long, ByVal hBitmap As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function ShowCaret Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetCaretBlinkTime Lib "user32" () As Long
Public Declare Function SetCaretBlinkTime Lib "user32" (ByVal wMSeconds As Long) As Long

Public Const EM_UNDO = &HC7
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_GETSEL = &HB0
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1

Public Const EM_FMTLINES = &HC8
Public Const EM_GETLINECOUNT = &HBA

