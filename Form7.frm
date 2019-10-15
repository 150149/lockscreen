VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "提示"
   ClientHeight    =   2040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   3300
   LinkTopic       =   "Form2"
   ScaleHeight     =   2040
   ScaleWidth      =   3300
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   120
   End
   Begin VB.Label Label1 
      Caption         =   "60秒后锁定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2655
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sm As String
Const Hwndx = -1
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height
Dim XX As Long
XX = SetWindowPos(Me.Hwnd, Hwndx, 0, 0, 0, 0, 3)
sm = 3
Timer1.Enabled = True
Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
sm = sm - 1
If sm = 0 Then
Timer1.Enabled = False
sm = 5
Form2.Hide
End If
End Sub
