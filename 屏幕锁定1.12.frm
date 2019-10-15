VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form4"
   ScaleHeight     =   3720
   ScaleWidth      =   9075
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command2 
      Caption         =   "密码文件伪装"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "屏幕锁定版本：1.12"
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "-------------------------------------------高级设置--------------------------------------------"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   8655
   End
   Begin VB.Label Label1 
      Caption         =   "请设置新的解锁密码"
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim c As String
c = Text1.Text
Open "xxx.txt" For Output As #2
Print #2, c
Close #2
Form4.Hide
End Sub

Private Sub Command2_Click()
Dim m, n As String
m = "bakba11"
n = "adbja235"
Open "abc" For Output As #3
Print #3, m
Close #3
Open "admin" For Output As #4
Print #4, n
Close #4
Dim q As String
q = bagavfkaufga4a54da47da34da64da3d4a34d6a4da3
Open "player" For Output As #5
Print #5, q
Close #5
End Sub
