VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "√‹¬Î¥ÌŒÛ¿≤"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form3"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  '∆¡ƒª÷––ƒ
   Begin VB.CommandButton Command1 
      Caption         =   "»∑∂®"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "√‹¬Î¥ÌŒÛ£¨«Î÷ÿ ‘"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   20.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Show
Form3.Hide
End Sub
