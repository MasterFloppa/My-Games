VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Level 1"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   17940
   LinkTopic       =   "Form1"
   Picture         =   "stage 1.frx":0000
   ScaleHeight     =   8685
   ScaleWidth      =   17940
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   95
      Left            =   1080
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   95
      Left            =   1080
      Top             =   600
   End
   Begin VB.Label Label6 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   12000
      TabIndex        =   6
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label5 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   9120
      TabIndex        =   5
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   6240
      TabIndex        =   4
      Top             =   5760
      Width           =   2775
   End
   Begin VB.Label Label3 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   12000
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   9120
      TabIndex        =   2
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H0012FCFC&
      Height          =   2295
      Left            =   6240
      TabIndex        =   1
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   " Level 1 "
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
      Height          =   1575
      Left            =   8520
      TabIndex        =   0
      Top             =   1800
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
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
End Sub

Private Sub label2_Click()
flag = 0
End Sub

Private Sub label4_Click()
flag = 0
End Sub

Private Sub label1_Click()
counter = counter + 1
Label1.Enabled = False
End Sub

Private Sub label3_Click()
counter = counter + 1
Label3.Enabled = False
End Sub

Private Sub label6_Click()
counter = counter + 1
Label6.Enabled = False
End Sub

Private Sub label5_Click()
counter = counter + 1
Label5.Enabled = False
End Sub

Private Sub Timer1_Timer()
x = x + 2
If x = 10 Then
Label1.BackColor = vbRed
ElseIf x = 12 Then
Label1.BackColor = &H12FCFC
ElseIf x = 14 Then
Label5.BackColor = vbRed
ElseIf x = 16 Then
Label5.BackColor = &H12FCFC
ElseIf x = 18 Then
Label3.BackColor = vbRed
ElseIf x = 20 Then
Label3.BackColor = &H12FCFC
ElseIf x = 22 Then
Label6.BackColor = vbRed
ElseIf x = 24 Then
Label6.BackColor = &H12FCFC
ElseIf x = 26 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If counter = 4 And flag = 1 Then
MsgBox "Congartulation you cleared stage 1!"
Form3.Show
Form1.Hide
Timer2.Enabled = False

ElseIf flag = 0 Then
MsgBox "You got it wrong, try again"
Form12.Label2.Caption = "Sorry " + Form9.Text1.Text + "," + " you are not awarded"
Form12.BackColor = 8421504
Form12.Image1.Visible = False
Form12.Image2.Visible = True
Form12.Image3.Visible = False
Form12.Label2.ForeColor = &HC0&
Form12.Show
Timer2.Enabled = False
Timer1.Enabled = False
Form1.Hide
Unload Form1
Unload Form9


End If
End Sub
