VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
If App.PrevInstance = True Then End
Form1.Hide
Timer1.Enabled = True
Timer1.Interval = 300
On Error Resume Next
Dim wsh
Set wsh = CreateObject("wscript.shell")
wsh.regwrite "HKLM\Software\Microsoft\Windows\Currentversion\Run\" & App.exeName, App.Path & "\" & App.exeName & ".exe", "REG_SZ"
End Sub

Private Function CheckExeIsRun(exeName As String) As Boolean
On Error GoTo Err
Dim WMI
Dim Obj
Dim Objs
CheckExeIsRun = False
Set WMI = GetObject("WinMgmts:")
Set Objs = WMI.InstancesOf("Win32_Process")
For Each Obj In Objs
If (InStr(UCase(exeName), UCase(Obj.Description)) <> 0) Then
CheckExeIsRun = True
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
End If
Next
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
Exit Function
Err:
If Not Objs Is Nothing Then Set Objs = Nothing
If Not WMI Is Nothing Then Set WMI = Nothing
End Function

Private Sub Timer1_Timer()
If CheckExeIsRun("启动程序定制c盘版.exe") Then
Else
Shell App.Path & "\启动程序定制c盘版.exe", vbNormalFocus
End If
End Sub
