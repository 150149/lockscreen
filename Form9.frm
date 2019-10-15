VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
If App.PrevInstance = True Then End
Form1.Hide
App.TaskVisible = False
Dim Modifiers As Long
    preWinProc = GetWindowLong(Me.Hwnd, GWL_WNDPROC)
    SetWindowLong Me.Hwnd, GWL_WNDPROC, AddressOf WndProc
    uVirtKey = vbKeyDelete
    RegisterHotKey Me.Hwnd, 1, Modifiers, uVirtKey
End Sub


Private Sub Form_Unload(Cancel As Integer)
SetWindowLong Me.Hwnd, GWL_WNDPROC, preWinProc
    UnregisterHotKey Me.Hwnd, uVirtKey   '取消系统级热键,释放资源
End
End Sub
