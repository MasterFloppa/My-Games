VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00FF00FF&
   Caption         =   "Level 4"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form5"
   Picture         =   "stage 4.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   75
      Left            =   720
      Top             =   2040
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   720
      Top             =   1080
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level 4"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1455
      Left            =   9000
      TabIndex        =   12
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   5880
      TabIndex        =   11
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label12 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   13800
      TabIndex        =   10
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label11 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   11160
      TabIndex        =   9
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label10 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   8520
      TabIndex        =   8
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   5880
      TabIndex        =   7
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label8 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   13800
      TabIndex        =   6
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   11160
      TabIndex        =   5
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label6 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   8520
      TabIndex        =   4
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   5880
      TabIndex        =   3
      Top             =   4800
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   13800
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   11160
      TabIndex        =   1
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   8520
      TabIndex        =   0
      Top             =   2520
      Width           =   2535
   End
End
Attribute VB_Name = "Form5"
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
Label1.BackColor = &H400000
Label2.BackColor = &H400000
Label3.BackColor = &H400000
Label4.BackColor = &H400000
Label5.BackColor = &H400000
Label6.BackColor = &H400000
Label7.BackColor = &H400000
Label8.BackColor = &H400000
Label9.BackColor = &H400000
Label10.BackColor = &H400000
Label11.BackColor = &H400000
Label12.BackColor = &H400000
End Sub


Private Sub Label_Click()
flag = 0
End Sub

Private Sub label2_Click()
flag = 0
End Sub

Private Sub label6_Click()
flag = 0
End Sub

Private Sub Label7_Click()
flag = 0
End Sub

Private Sub Label10_Click()
flag = 0
End Sub

Private Sub Label12_Click()
flag = 0
End Sub

Private Sub Label11_Click()
counter = counter + 1
Label1.Enabled = False
End Sub

Private Sub Label9_Click()
counter = counter + 1
Label9.Enabled = False
End Sub

Private Sub Label8_Click()
counter = counter + 1
Label8.Enabled = False
End Sub

Private Sub label5_Click()
counter = counter + 1
Label5.Enabled = False
End Sub

Private Sub label4_Click()
counter = counter + 1
Label4.Enabled = False
End Sub

Private Sub label3_Click()
counter = counter + 1
Label3.Enabled = False
End Sub

Private Sub Timer1_Timer()
x = x + 2
If x = 10 Then
Label5.BackColor = &HFF00FF 'light purpule
ElseIf x = 12 Then
Label5.BackColor = &H400000    'dark blue
ElseIf x = 14 Then
Label3.BackColor = &HFF00FF
ElseIf x = 16 Then
Label3.BackColor = &H400000
ElseIf x = 18 Then
Label8.BackColor = &HFF00FF
ElseIf x = 20 Then
Label8.BackColor = &H400000
ElseIf x = 22 Then
Label11.BackColor = &HFF00FF
ElseIf x = 24 Then
Label11.BackColor = &H400000
ElseIf x = 26 Then
Label9.BackColor = &HFF00FF
ElseIf x = 28 Then
Label9.BackColor = &H400000
ElseIf x = 30 Then
Label4.BackColor = &HFF00FF
ElseIf x = 32 Then
Label4.BackColor = &H400000
ElseIf x = 34 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If counter = 6 And flag = 1 Then
MsgBox "Congartulation you cleared stage 4!"
Form6.Show
Form5.Hide
Timer2.Enabled = False

ElseIf flag = 0 Then
MsgBox "You got it wrong, try again"
Form12.Label2.Caption = " " + Form9.Text1.Text + "," + " you are awarded with a bronze medal!"
Form12.BackColor = &HF2&
Form12.Image1.Visible = False
Form12.Image2.Visible = False
Form12.Image3.Visible = True
Timer2.Enabled = False
Timer1.Enabled = False
Form5.Hide
Unload Form5
Unload Form4
Unload Form3
Unload Form1
Unload Form9
Form12.Show

End If
End Sub




