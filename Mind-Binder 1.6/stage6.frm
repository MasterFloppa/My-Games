VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Level 6"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   Picture         =   "stage6.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   58
      Left            =   1200
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Interval        =   58
      Left            =   1200
      Top             =   840
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level 6 "
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   56.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1215
      Left            =   9000
      TabIndex        =   12
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   13560
      TabIndex        =   11
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   10920
      TabIndex        =   10
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   8280
      TabIndex        =   9
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   5640
      TabIndex        =   8
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   13560
      TabIndex        =   7
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   10920
      TabIndex        =   6
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   8280
      TabIndex        =   5
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   5640
      TabIndex        =   4
      Top             =   5160
      Width           =   2655
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   13560
      TabIndex        =   3
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   10920
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   8280
      TabIndex        =   1
      Top             =   3000
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   2175
      Left            =   5640
      TabIndex        =   0
      Top             =   3000
      Width           =   2655
   End
End
Attribute VB_Name = "Form10"
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
Label1.BackColor = &H1D1D1D
Label2.BackColor = &H1D1D1D
Label3.BackColor = &H1D1D1D
Label4.BackColor = &H1D1D1D
Label5.BackColor = &H1D1D1D
Label6.BackColor = &H1D1D1D
Label7.BackColor = &H1D1D1D
Label8.BackColor = &H1D1D1D
Label9.BackColor = &H1D1D1D
Label10.BackColor = &H1D1D1D
Label11.BackColor = &H1D1D1D
Label12.BackColor = &H1D1D1D
End Sub

Private Sub Label7_Click()
flag = 0
End Sub

Private Sub Label11_Click()
flag = 0
End Sub

Private Sub Label10_Click()
flag = 0
End Sub

Private Sub Label9_Click()
flag = 0
End Sub

Private Sub label4_Click()
flag = 0
End Sub

Private Sub label2_Click()
flag = 0
End Sub

Private Sub Label12_Click()
counter = counter + 1
Label2.Enabled = False
End Sub

Private Sub Label8_Click()
counter = counter + 1
Label8.Enabled = False
End Sub

Private Sub label6_Click()
counter = counter + 1
Label6.Enabled = False
End Sub

Private Sub label5_Click()
counter = counter + 1
Label5.Enabled = False
End Sub

Private Sub label3_Click()
counter = counter + 1
Label3.Enabled = False
End Sub

Private Sub label1_Click()
counter = counter + 1
Label1.Enabled = False
End Sub

Private Sub Timer1_Timer()
x = x + 2
If x = 10 Then
Label1.BackColor = &HFF00&
ElseIf x = 12 Then
Label1.BackColor = &H1D1D1D
ElseIf x = 14 Then
Label3.BackColor = &HFF00&
ElseIf x = 16 Then
Label3.BackColor = &H1D1D1D
ElseIf x = 18 Then
Label6.BackColor = &HFF00&
ElseIf x = 20 Then
Label6.BackColor = &H1D1D1D
ElseIf x = 22 Then
Label12.BackColor = &HFF00&
ElseIf x = 24 Then
Label12.BackColor = &H1D1D1D
ElseIf x = 26 Then
Label5.BackColor = &HFF00&
ElseIf x = 28 Then
Label5.BackColor = &H1D1D1D
ElseIf x = 30 Then
Label12.BackColor = &HFF00&
ElseIf x = 32 Then
Label12.BackColor = &H1D1D1D
ElseIf x = 34 Then
Label8.BackColor = &HFF00&
ElseIf x = 36 Then
Label8.BackColor = &H1D1D1D
ElseIf x = 38 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If counter = 6 And flag = 1 Then
MsgBox "Congartulation you cleared stage 6!"
Form11.Show
Form10.Hide
Timer2.Enabled = False

ElseIf flag = 0 Then
MsgBox "You got it wrong, try again"
Form12.Label2.Caption = " " + Form9.Text1.Text + "," + " you are awarded with a silver cup!"
Form12.BackColor = &HFFFF80
Form12.Image1.Visible = True
Form12.Image2.Visible = False
Form12.Image3.Visible = False
Timer2.Enabled = False
Timer1.Enabled = False
Form10.Hide
Unload Form6
Unload Form5
Unload Form4
Unload Form3
Unload Form1
Unload Form9
Unload Form10
Form12.Show

End If
End Sub

