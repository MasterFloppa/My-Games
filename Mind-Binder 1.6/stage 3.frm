VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Level 3"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form4"
   Picture         =   "stage 3.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   85
      Left            =   840
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   85
      Left            =   840
      Top             =   840
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   " Level 3 "
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   62.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1455
      Left            =   8880
      TabIndex        =   9
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label9 
      BackColor       =   &H0012FCFC&
      Height          =   2175
      Left            =   12480
      TabIndex        =   8
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label8 
      BackColor       =   &H0012FCFC&
      Height          =   2175
      Left            =   9480
      TabIndex        =   7
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H0012FCFC&
      Height          =   2175
      Left            =   6480
      TabIndex        =   6
      Top             =   7200
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   12480
      TabIndex        =   5
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   9480
      TabIndex        =   4
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   6480
      TabIndex        =   3
      Top             =   4800
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   12480
      TabIndex        =   2
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   9480
      TabIndex        =   1
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   6480
      TabIndex        =   0
      Top             =   2400
      Width           =   2775
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x, flag, counter As Integer

Private Sub Form_Load()
Timer2.Enabled = False
flag = 1
x = 0
counter = 0

Label1.BackColor = &H12FCFC
Label2.BackColor = &H12FCFC
Label3.BackColor = &H12FCFC
Label4.BackColor = &H12FCFC
Label5.BackColor = &H12FCFC
Label6.BackColor = &H12FCFC
Label7.BackColor = &H12FCFC
Label8.BackColor = &H12FCFC
Label9.BackColor = &H12FCFC
End Sub

Private Sub label1_Click()
counter = counter + 1
Label1.Enabled = False
End Sub

Private Sub label2_Click()
flag = 0
End Sub

Private Sub label3_Click()
flag = 0
End Sub

Private Sub label4_Click()
counter = counter + 1
Label4.Enabled = False
End Sub

Private Sub label5_Click()
flag = 0
End Sub

Private Sub label6_Click()
counter = counter + 1
Label6.Enabled = False
End Sub

Private Sub Label7_Click()
flag = 0
End Sub

Private Sub Label8_Click()
counter = counter + 1
Label8.Enabled = False
End Sub

Private Sub Label9_Click()
counter = counter + 1
Label9.Enabled = False
End Sub

Private Sub Timer1_Timer()
x = x + 2
If x = 10 Then
Label9.BackColor = vbRed
ElseIf x = 12 Then
Label9.BackColor = &H12FCFC
ElseIf x = 14 Then
Label8.BackColor = vbRed
ElseIf x = 16 Then
Label8.BackColor = &H12FCFC
ElseIf x = 18 Then
Label1.BackColor = vbRed
ElseIf x = 20 Then
Label1.BackColor = &H12FCFC
ElseIf x = 22 Then
Label4.BackColor = vbRed
ElseIf x = 24 Then
Label4.BackColor = &H12FCFC
ElseIf x = 26 Then
Label6.BackColor = vbRed
ElseIf x = 28 Then
Label6.BackColor = &H12FCFC
ElseIf x = 30 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If counter = 5 And flag = 1 Then
MsgBox "Congartulation you cleared stage 3!"
Form5.Show
Form4.Hide
Timer2.Enabled = False

ElseIf flag = 0 Then
MsgBox "You got it wrong, try again"
Form12.Label2.Caption = "Sorry " + Form9.Text1 + "," + " you are not awarded"
Form12.BackColor = 8421504
Form12.Image1.Visible = False
Form12.Image2.Visible = True
Form12.Image3.Visible = False
Form12.Label2.ForeColor = &HC0&
Timer2.Enabled = False
Timer1.Enabled = False
Form4.Hide
Unload Form4
Unload Form3
Unload Form1
Unload Form9
Form12.Show

End If
End Sub



