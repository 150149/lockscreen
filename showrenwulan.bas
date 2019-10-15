Attribute VB_Name = "showrenwulan"
Option Explicit
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_WNDPROC = (-4)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Private Const WM_CLOSE = &H10
Private Const WM_DESTROY = &H2
 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Dim mlOldproc As Long
 
Private Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_SYSCOMMAND
            If wParam = SC_CLOSE Then
                SendMessage hwnd, WM_CLOSE, ByVal 0&, ByVal 0&
            End If
        Case WM_DESTROY
            SetWindowLong hwnd, GWL_WNDPROC, mlOldproc
    End Select
    WndProc = CallWindowProc(mlOldproc, hwnd, Msg, wParam, lParam)
End Function

Public Sub subclass(hwnd As Long)
    Dim lStyle As Long
    lStyle = GetWindowLong(hwnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX Or WS_SYSMENU
    SetWindowLong hwnd, GWL_STYLE, lStyle
    mlOldproc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WndProc)
End Sub
