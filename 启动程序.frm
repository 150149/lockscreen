VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   12090
   ClientTop       =   5775
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5025
   Begin VB.Timer Timer7 
      Left            =   3840
      Top             =   840
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1440
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Timer Timer6 
      Left            =   3720
      Top             =   360
   End
   Begin VB.Timer Timer5 
      Left            =   240
      Top             =   2400
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   1440
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Timer Timer4 
      Left            =   360
      Top             =   1800
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   1440
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.Timer Timer3 
      Left            =   240
      Top             =   1320
   End
   Begin VB.Timer Timer2 
      Left            =   120
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public g As String
Public f As String
Public time As String
Public time2 As String
Public time3 As String
Public timeopen As String
Public hot As String
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long '

Private Declare Function ShowWindow Lib "user32" (ByVal Hwnd As Long, ByVal nCmdShow As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Const TH32CS_SNAPPROCESS = &H2

Private Const TH32CS_SNAPheaplist = &H1

Private Const TH32CS_SNAPthread = &H4

Private Const TH32CS_SNAPmodule = &H8

Private Const TH32CS_SNAPall = TH32CS_SNAPPROCESS + TH32CS_SNAPheaplist + TH32CS_SNAPthread + TH32CS_SNAPmodule

Private Const MAX_PATH As Integer = 260

Private Const PROCESS_ALL_ACCESS = &H100000 + &HF0000 + &HFFF

Private Type PROCESSENTRY32

dwSize As Long

cntUseage As Long

th32ProcessID As Long

th32DefaultHeapID As Long

th32ModuleID As Long

cntThreads As Long

th32ParentProcessID As Long

pcPriClassBase As Long

swFlags As Long

szExeFile As String * 1024

End Type

Public Sub AntiKill()

On Error Resume Next

Dim MySnapHandle As Long

Dim hProcess As Long

Dim ProcessInfo As PROCESSENTRY32

Dim addr As Long, hMod As Long

Dim ASM(0) As Byte

Dim sProcess As String

ASM(0) = &HC3 'retn

hMod = GetModuleHandle("kernel32")

addr = GetProcAddress(hMod, "TerminateProcess")

'Debug.Print Hex(Addr)

MySnapHandle = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)

ProcessInfo.dwSize = Len(ProcessInfo)

If ProcessFirst(MySnapHandle, ProcessInfo) <> 0 Then

Do

sProcess = Left(LCase(ProcessInfo.szExeFile), InStr(ProcessInfo.szExeFile, ".") + 3)

If sProcess = "taskmgr.exe" Then

hProcess = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessInfo.th32ProcessID)

'Debug.Print hProcess

WriteProcessMemory hProcess, ByVal addr, ByVal VarPtr(ASM(0)), 1, 0&

'Debug.Print Err.LastDllError

CloseHandle hProcess

End If

Loop While ProcessNext(MySnapHandle, ProcessInfo) <> 0

End If

CloseHandle MySnapHandle

Err.Clear

End Sub

Private Sub Timer7_Timer()
hot = hot - 1
If hot = 5 Then
Timer7.Enabled = False
End If
End Sub


Private Sub Form_Load()
If App.PrevInstance = True Then End
Form1.Hide
HideCurrentProcess '隐藏进程
hot = 10
Timer7.Enabled = True
Timer7.Interval = 1000
App.TaskVisible = False
If Dir(App.Path & "\time2") = "" Then
Open "time2" For Output As #15
Print #15, ""
Close #15
End If
If Dir(App.Path & "\time1") = "" Then
Open "time1" For Output As #15
Print #15, "1"
Close #15
End If
If Dir(App.Path & "\time3") = "" Then
Open "time3" For Output As #15
Print #15, ""
Close #15
End If
If Dir(App.Path & "\timeopen") = "" Then
Open "timeopen" For Output As #15
Print #15, "0"
Close #15
End If
Open App.Path & "\time2" For Input As #14
Line Input #14, time
Close #14
Open App.Path & "\time3" For Input As #14
Line Input #14, time3
Close #14
Open App.Path & "\timeopen" For Input As #14
Line Input #14, timeopen
Close #14
If timeopen = "1" Then
Open App.Path & "\time1" For Output As #16
Print #16, "0"
Close #16
End If
Timer1.Enabled = True
Timer1.Interval = 300
Timer2.Enabled = True
Timer2.Interval = 300
Timer3.Enabled = True
Timer3.Interval = 800
Timer4.Enabled = True
Timer4.Interval = 1000
Timer5.Enabled = True
Timer5.Interval = 500
On Error Resume Next
Dim wsh
Set wsh = CreateObject("wscript.shell")
wsh.regwrite "HKLM\Software\Microsoft\Windows\Currentversion\Run\" & App.exeName, App.Path & "\" & App.exeName & ".exe", "REG_SZ"
Dim NTSrv As New ClsSrvCtrl
With NTSrv
       .Name = "150149进程保护"
       .Account = "LocalSystem"
       .Description = "150149进程保护"
       .DisplayName = "150149进程保护"
       .Command = "C:\Program Files\lockscreen\protector.exe"
       .Interact = SERVICE_INTERACT_WITH_DESKTOP
       .StartType = SERVICE_AUTO_START
   End With
    Call NTSrv.SetNTService
With NTSrv
       .Name = "150149进程保护2"
       .Account = "LocalSystem"
       .Description = "150149进程保护2"
       .DisplayName = "150149进程保护2"
       .Command = "C:\Program Files\lockscreen\启动程序定制c盘版.exe"
       .Interact = SERVICE_INTERACT_WITH_DESKTOP
       .StartType = SERVICE_AUTO_START
   End With
    Call NTSrv.SetNTService
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

Private Function fileStr(ByVal strFileName As String) As String
    On Error GoTo Err1
    Dim tempInput As String
    Open strFileName For Input As #1
    Do While Not EOF(1)
        Line Input #1, tempInput
        If Right(tempInput, 1) <> Chr(10) Then tempInput = tempInput & Chr(10)
        tempInput = Replace(tempInput, Chr(10), vbCrLf)
        fileStr = fileStr & tempInput
    Loop
    If fileStr <> "" Then fileStr = Left(fileStr, Len(fileStr) - 2)
    Close #1
    Exit Function
Err1:
    MsgBox "不存在该文件或该文件不能访问！", vbExclamation
End Function

Private Sub Timer1_Timer()
Open App.Path & "\time1" For Input As #10
Line Input #10, g
Close #10
Text1.Text = g
Text3.Text = time
Text2.Text = f
Text4.Text = time3
End Sub

Private Sub Timer2_Timer()
If CheckExeIsRun("lockscreen.exe") Then
f = "1"
Else
f = "0"
End If
If CheckExeIsRun("protector.exe") Then
Else
Shell App.Path & "\protector.exe", vbNormalFocus
End If
End Sub

Private Sub Timer3_Timer()
If f = "0" Then
If g = "0" Then
Shell App.Path & "\lockscreen.exe", vbNormalFocus
End If
End If
End Sub


Private Sub Timer4_Timer()
If time = "" Then
Else
time = time - 1
If time = "0" Then
Open "time1" For Output As #16
Print #16, "0"
Close #16
Timer4.Enabled = False
Timer6.Enabled = True
Timer6.Interval = 1000
Open "time2" For Input As #14
Line Input #14, time
Close #14
End If
End If
End Sub

Private Sub Timer5_Timer()
If time = "60" Then
Form2.Show
ElseIf time = "600" Then
Form3.Show
ElseIf time = "10" Then
Form4.Show
End If
End Sub

Private Sub Timer6_Timer()
If time3 = "" Then
Else
time3 = time3 - 1
If time3 = "0" Then
Open "time1" For Output As #16
Print #16, "1"
Close #16
Shell "taskkill /f /im lockscreen.exe"
Open "time3" For Input As #14
Line Input #14, time3
Close #14
Timer4.Enabled = True
Timer4.Interval = 1000
Timer6.Enabled = False
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell App.Path & "\lockscreen.exe", vbNormalFocus
Shell App.Path & "\protector.exe", vbNormalFocus
End
End Sub
