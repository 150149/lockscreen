VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "提示"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3810
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleMode       =   0  'User
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer2 
      Left            =   2400
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "秒后锁定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ss As Integer
Const Hwndx = -1
Private Declare Function SetWindowPos Lib "user32" (ByVal Hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Form_Load()
Me.Left = Screen.Width - Me.Width
Me.Top = Screen.Height - Me.Height
Dim XX As Long
XX = SetWindowPos(Me.Hwnd, Hwndx, 0, 0, 0, 0, 3)
Timer1.Enabled = True
Timer1.Interval = 1000
Timer2.Enabled = True
Timer2.Interval = 500
ss = 10
End Sub

Private Sub Timer1_Timer()
ss = ss - 1
If ss = 0 Then
Timer1.Enabled = False
Timer2.Enabled = False
ss = 10
Unload Me
End If
End Sub

Private Sub Timer2_Timer()
Label1.Caption = ss
End Sub


