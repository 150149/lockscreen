VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "请输入密码"
   ClientHeight    =   3090
   ClientLeft      =   -60
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   2640
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   11.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   240
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C000&
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "验证码"
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "密码"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "请输入密码"
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public mm As String
Public jc As String
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
If Len(Text2.Text) < 2 Then
MsgBox "密码长度不够"
    Else
Open "xxx" For Input As #1
Line Input #1, mm
Close #1
Dim c As String
c = Text3.Text
Dim smi As String
Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long
            KeyBits = 128
            BlockBits = 128
            pass = GetPassword
                plaintext = StrConv(Text2.Text, vbFromUnicode)
#If SUPPORT_LEVEL Then
            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0, BlockBits
#Else
            m_Rijndael.SetCipherKey pass, KeyBits
            m_Rijndael.ArrayEncrypt plaintext, ciphertext, 0
#End If
            smi = HexDisplay(ciphertext, UBound(ciphertext) + 1, BlockBits \ 8)
If c = "" Then
MsgBox "验证码未填写"
ElseIf c = jc Then
If mm = smi Then
Form2.Show
Form1.Hide
Else
MsgBox "密码错误"
End If
Else
MsgBox "验证码错误"
End If
End If
End Sub


Private Sub Form_Load()
With oShadow
    If .Shadow(Me) Then
        .Depth = 7 '阴影宽度
        .Color = RGB(0, 0, 0) '阴影颜色
        .Transparency = 50 '阴影色深
    End If
 End With
Dim ja, jb As String
Randomize '必要代码，不然的话每次随机生成都是一样的数
ja = Int(Rnd * (9 - 0 + 1)) + 0 '其中m>n生成n~m间（包括n，m）的随机整数
jb = Int(Rnd * (9 - 0 + 1)) + 0 '其中m>n生成n~m间（包括n，m）的随机整数
jc = ja + jb
Label2 = ja & "+" & jb & "="
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
