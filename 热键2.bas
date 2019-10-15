Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal _
wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal id As Long, _
ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal Hwnd As Long, ByVal id As Long) _
As Long


    
Private preWinProc As Long
Private Modifiers As Long, uVirtKey As Long, idHotKey As Long

Private Const WM_HOTKEY = &H312
Private Const GWL_WNDPROC = (-4)

Public Enum ThreeKey
   CTRL = &H2
   SHIFT = &H4
   ALT = &H1
   NONE = &H0
End Enum

Private Type taLong
    ll As Long
End Type

Private Type t2Int
    lWord As Integer
    hWord As Integer
End Type


Private Function Wndproc(ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    If Msg = WM_HOTKEY Then
            Dim lp As taLong, i2 As t2Int
            lp.ll = lParam
            LSet i2 = lp
            If (i2.lWord = Modifiers) And i2.hWord = uVirtKey Then '最好自己写
Open "time1" For Output As #16
Print #16, "0"
Close #16
            End If
    End If
    Wndproc = CallWindowProc(preWinProc, Hwnd, Msg, wParam, lParam)
    
End Function


Public Sub RegHotKey(FormHwnd As Long, fiers As ThreeKey, vKey As Long, Optional HotKey As Long = 1)
    preWinProc = GetWindowLong(FormHwnd, GWL_WNDPROC)
    SetWindowLong FormHwnd, GWL_WNDPROC, AddressOf Wndproc
    idHotKey = HotKey
    Modifiers = fiers
    uVirtKey = vKey
    RegisterHotKey FormHwnd, idHotKey, Modifiers, uVirtKey
End Sub

Public Sub UnRegHotKey(FormHwnd As Long)
    SetWindowLong FormHwnd, GWL_WNDPROC, preWinProc
    Call UnregisterHotKey(FormHwnd, uVirtKey)
End Sub


