VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "����"
   ClientHeight    =   3720
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form4"
   ScaleHeight     =   3720
   ScaleWidth      =   9075
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command3 
      Caption         =   "&t"
      Height          =   375
      Left            =   5760
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&m"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ���޸�"
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
   Begin VB.Label Label5 
      Caption         =   "��alt+t�˳������޸�"
      Height          =   255
      Left            =   840
      TabIndex        =   8
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Label Label2 
      Caption         =   "�����ļ�αװ��alt+m"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1320
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "���س���ȷ���޸�����"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "��Ļ�����汾��1.32"
      Height          =   375
      Left            =   7440
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "�������µĽ�������"
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
Open "xxx" For Output As #2
Print #2, c
Close #2
Text1.Text = ""
Form1.Show
Text1.SetFocus
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


Private Sub Command3_Click()
Form1.Show
Form4.Hide
End Sub
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Call Command1_Click  'ģ�ⵥ��Command1��ť
End If
End Sub
