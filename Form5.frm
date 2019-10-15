VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "修改密码"
   ClientHeight    =   4020
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   7650
   LinkTopic       =   "Form2"
   ScaleHeight     =   4020
   ScaleWidth      =   7650
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton command2 
      Caption         =   "开机即刻自启动"
      Height          =   495
      Left            =   4560
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   480
      TabIndex        =   10
      Top             =   3480
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1800
      TabIndex        =   6
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   480
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   855
      Left            =   4560
      TabIndex        =   1
      Top             =   2520
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      MaxLength       =   16
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0C000&
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label Label7 
      Caption         =   "秒"
      Height          =   375
      Left            =   1320
      TabIndex        =   11
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "定时解锁--请换算成秒来填"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3000
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "不使用定时功能请留空,不足1小时请填0"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   3615
   End
   Begin VB.Label Label4 
      Caption         =   "分"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label3 
      Caption         =   "时"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "————定时启动设置————"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "请输入新密码"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public timeopen1 As String
Private Type MiYaoData
  Char As String '//元素值
  Num As Integer '//对照字母，产生顺序号
End Type
Private oShadow As New aShadow
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
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

Private Sub Command1_Click()
If Len(Text1.Text) < 2 Then
MsgBox "密码长度不足"
    Else
Dim m As String
m = Text1.Text
If m = "" Then
Else
Dim max As Integer '//密钥长度
Dim p(26) As MiYaoData '//密钥数据
Dim t(26) As Integer '//转换密钥数据的中间变量
Dim MingWen(10, 26) As String '//明文
Dim i, j, k As Integer
Dim MinAsc, Mindex As Integer '//当前最小字符ASCII值和其下标
Dim s, ss As String

s = "woshivbchengxu"
max = Len(s) '//获取密钥长度

'//逐个获取密钥字母（不应重复）
For i = 0 To max - 1
  p(i).Char = Mid(s, i + 1, 1)
  p(i).Num = max
Next

'//对应英文字母表，逐个填写密钥字母的顺序，为生成密文输出作准备
For i = 0 To max - 1
  '//查找当前小的密钥字符所在下标
  MinAsc = Asc("z") + 10 '//将一个大数作为当前最小值
  For j = 0 To max - 1
    If Asc(p(j).Char) < MinAsc And p(j).Num = max Then
       MinAsc = Asc(p(j).Char) '//更改最小值
       Mindex = j '//登记其下标
    End If
  Next
  '//当前最小值找到，修改数据（下标）
  p(Mindex).Num = i '//登记其下标
Next
'//将"p(i).num=j"的格式转换成"p(j).num=i"
'//防止数据覆盖，转换结果另存
For i = 0 To max - 1
  t(p(i).Num) = i
Next
'//将另存结果，写入原始数据中
For i = 0 To max - 1
  p(i).Num = t(i)
Next

'//去除明文中的空格
s = Text1.Text
s = Replace(s, " ", "")

'//明文依次写入二维数组
For i = 0 To Len(s) - 1
  MingWen((i \ max), (i Mod max)) = Mid(s, i + 1, 1)
Next

'//最后一行上的数据未满，用a,b,……填充
For j = (i - 1) Mod max + 1 To max - 1
  MingWen((i - 1) \ max, j) = Chr(96 - (i - 1) Mod max + j)
Next
m = ""
'//按密钥字母的顺序，转置明文，生成密文
For j = 0 To max - 1 '//初始明文矩阵的列范围
  ss = ""
  '//初始明文矩阵的行范围
  For i = 0 To (Len(s) + max - 1) \ max - 1
    ss = ss + UCase$(MingWen(i, p(j).Num)) '//转换成大写字母
  Next
  m = m + ss '//显示密文
Next

Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(Text1.Text) = 0 Then
        MsgBox "No Plaintext"
    Else
        If Len("woshivbchengxu") = 0 Then
            MsgBox "No Password"
        Else
            KeyBits = 128
            BlockBits = 128
            pass = GetPassword
                plaintext = StrConv(Text1.Text, vbFromUnicode)
            End If

#If SUPPORT_LEVEL Then
            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
#Else
            m_Rijndael.SetCipherKey pass, KeyBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
#End If
            m = HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
        End If
  MsgBox m
Open "xxx" For Output As #1
Print #1, m
Close #1
Dim hour As String
hour = Text2.Text
Dim minute As String
minute = Text3.Text
If hour = "" Or minute = "" Then
Open "time2" For Output As #2
Print #2, ""
MsgBox "关闭定时功能"
Close #2
Else
Dim time As Integer
time = hour * 3600 + minute * 60
Open "time2" For Output As #2
Print #2, time
Close #2
MsgBox time
End If
Dim unlocktime As String
unlocktime = Text4.Text
If unlocktime = "" Then
Open "time3" For Output As #2
Print #2, ""
MsgBox "关闭自动解锁"
Close #2
Unload Me
End
Else
Open "time3" For Output As #2
Print #2, unlocktime
Close #2
MsgBox unlocktime
End If
Unload Me
End
End If
End If
End Sub

Private Sub command2_Click()
If timeopen1 = "0" Then
Open "timeopen" For Output As #3
Print #3, "1"
Close #3
command2.Caption = "开机即刻自启动"
Open "timeopen" For Input As #3
Line Input #3, timeopen1
Close #3
ElseIf timeopen1 = "1" Then
Open "timeopen" For Output As #3
Print #3, "0"
Close #3
command2.Caption = "关闭开机自启动"
Open "timeopen" For Input As #3
Line Input #3, timeopen1
Close #3
End If
End Sub

Private Sub Form_Load()
If Dir(App.Path & "\timeopen") = "" Then
Open "timeopen" For Output As #3
Print #3, "0"
Close #3
Open "timeopen" For Input As #4
Line Input #4, timeopen1
Close #4
Else
Open "timeopen" For Input As #4
Line Input #4, timeopen1
Close #4
End If
If timeopen1 = "0" Then
command2.Caption = "关闭开机自启动"
ElseIf timeopen1 = "1" Then
command2.Caption = "开机即刻自启动"
End If
With oShadow
    If .Shadow(Me) Then
        .Depth = 7 '阴影宽度
        .Color = RGB(0, 0, 0) '阴影颜色
        .Transparency = 50 '阴影色深
    End If
 End With
End Sub

Private Sub Form_DblClick()
Unload Me
End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button = 1 Then
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0
End If
End Sub


