VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "∆¡ƒªÀ¯∂®√‹¬ÎΩ‚√‹≥Ã–Ú"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   6870
   StartUpPosition =   3  '¥∞ø⁄»± °
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ω‚√‹"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
#Const SUPPORT_LEVEL = 0     'Default=0
'Must be equal to SUPPORT_LEVEL in cRijndael

'An instance of the Class
Private m_Rijndael As New cRijndael


'Used to display what the program is doing in the Form's caption
Public Property Let Status(TheStatus As String)
    If Len(TheStatus) = 0 Then
        Me.Caption = App.Title
    Else
        Me.Caption = App.Title & " - " & TheStatus
    End If
    Me.Refresh
End Property


'Assign TheString to the Text property of TheTextBox if possible.  Otherwise give warning.
Private Sub DisplayString(TheTextBox As TextBox, ByVal TheString As String)
    If Len(TheString) < 65536 Then
        TheTextBox.Text = TheString
    Else
    End If
End Sub


Private Function GetPassword() As Byte()
    Dim data() As Byte
        data = StrConv("woshivbchengxu", vbFromUnicode)
        ReDim Preserve data(31)
    GetPassword = data
End Function

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

Private Sub Command1_Click()
 Dim pass()        As Byte
    Dim plaintext()   As Byte
    Dim ciphertext()  As Byte
    Dim KeyBits       As Long
    Dim BlockBits     As Long

    If Len(Text1.Text) = 0 Then
        MsgBox "No Ciphertext"
    Else
        If Len("woshivbchengxu") = 0 Then
            MsgBox "No Password"
        Else
            KeyBits = 128
            BlockBits = 128
            pass = GetPassword


            If HexDisplayRev(Text1.Text, ciphertext) = 0 Then
                MsgBox "Text not Hex data"

                Exit Sub
            End If

#If SUPPORT_LEVEL Then
            m_Rijndael.SetCipherKey pass, KeyBits, BlockBits
            If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0, BlockBits) <> 0 Then

                Exit Sub
            End If
#Else
            m_Rijndael.SetCipherKey pass, KeyBits
            If m_Rijndael.ArrayDecrypt(plaintext, ciphertext, 0) <> 0 Then
   
                Exit Sub
            End If
#End If
 

                DisplayString Text1, StrConv(plaintext, vbUnicode)
DisplayString Text1, HexDisplay(plaintext, UBound(plaintext) + 1, BlockBits \ 8)
        End If
    End If
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

