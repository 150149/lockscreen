Attribute VB_Name = "Module1"
Option Explicit

Private Declare Function CallNextHookEx _
                Lib "user32" (ByVal hHook As Long, _
                              ByVal nCode As Long, _
                              ByVal wParam As Long, _
                              lParam As Any) As Long

Private Declare Function SetWindowsHookEx _
                Lib "user32" _
                Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                           ByVal lpfn As Long, _
                                           ByVal hmod As Long, _
                                           ByVal dwThreadId As Long) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Declare Sub CopyMemory _
                Lib "kernel32" _
                Alias "RtlMoveMemory" (Destination As Any, _
                                       Source As Any, _
                                       ByVal Length As Long)

Private Type PKBDLLHOOKSTRUCT

    vkCode As Long

    scanCode As Long

    flags As Long

    time As Long

    dwExtraInfo As Long

End Type

Private Const WM_KEYDOWN = &H100

Private Const WM_SYSKEYDOWN = &H104

Private Const WM_KEYUP = &H101

Private Const WM_SYSKEYUP = &H105

Private Const VK_LWIN = &H5B

Private Const VK_RWIN = &H5C

Private Const HC_ACTION = 0

Private Const WH_KEYBOARD_LL = 13

Private lngHook As Long

'使用底层KeyboardHook拦截按键消息

Public Function LowLevelKeyboardProc(ByVal nCode As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long

    Dim blnHook As Boolean

    Dim p       As PKBDLLHOOKSTRUCT

    If nCode = HC_ACTION Then

        Select Case wParam

            Case WM_KEYDOWN, WM_SYSKEYDOWN, WM_KEYUP, WM_SYSKEYUP

                Call CopyMemory(p, ByVal lParam, Len(p))

                If p.vkCode = VK_LWIN Or p.vkCode = VK_RWIN Then '按下了左/右Win键

                    blnHook = True

                End If

            Case Else

                'do nothing

        End Select

    End If

    If blnHook Then

        LowLevelKeyboardProc = 1

    Else

        Call CallNextHookEx(WH_KEYBOARD_LL, nCode, wParam, lParam)

    End If

End Function

Public Sub HooK()

    lngHook = SetWindowsHookEx(WH_KEYBOARD_LL, AddressOf LowLevelKeyboardProc, App.hInstance, 0)

End Sub

Public Sub UnHooK()

    Call UnhookWindowsHookEx(lngHook)

End Sub

