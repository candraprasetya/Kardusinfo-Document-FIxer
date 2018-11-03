Attribute VB_Name = "Module1"
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVS_EX_GRIDLINES = &H1
Public Const GWL_STYLE        As Long = (-16)
Public Const LVM_GETHEADER    As Long = (LVM_FIRST + 31)
Public Const HDS_BUTTONS      As Long = 2
Public Function SetFlatHeaders(LVhwnd As Long)
    Dim hHeader As Long
    Dim Style   As Long
    'get the handle to the listview header
    hHeader = SendMessage(LVhwnd, LVM_GETHEADER, 0, ByVal 0&)
    'set the new style
    Style = GetWindowLong(hHeader, GWL_STYLE)
    Style = Style And Not HDS_BUTTONS
    Call SetWindowLong(hHeader, GWL_STYLE, Style)
End Function
Public Function lvwStyle(lvStyle As ListView)
    Dim rStyle As Long
    Dim r As Long
    rStyle = SendMessageLong(lvStyle.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0&, 0&)
    rStyle = rStyle Xor LVS_EX_FULLROWSELECT Xor LVS_EX_GRIDLINES
    r = SendMessageLong(lvStyle.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, 0&, rStyle)
End Function

