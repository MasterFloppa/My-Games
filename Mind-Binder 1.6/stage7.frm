VERSION 5.00
Begin VB.Form Form11 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Level 7"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18150
   LinkTopic       =   "Form11"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "stage7.frx":0000
   ScaleHeight     =   9630
   ScaleWidth      =   18150
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   55
      Left            =   480
      Top             =   2520
   End
   Begin VB.Timer Timer1 
      Interval        =   55
      Left            =   480
      Top             =   1800
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   13080
      TabIndex        =   16
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   10680
      TabIndex        =   15
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   8280
      TabIndex        =   14
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   5880
      TabIndex        =   13
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   13080
      TabIndex        =   12
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   10680
      TabIndex        =   11
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   8280
      TabIndex        =   10
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   5880
      TabIndex        =   9
      Top             =   5760
      Width           =   2415
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   13080
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   10680
      TabIndex        =   7
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   8280
      TabIndex        =   6
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   5880
      TabIndex        =   5
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   13080
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   10680
      TabIndex        =   3
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   8280
      TabIndex        =   2
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   1815
      Left            =   5880
      TabIndex        =   1
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Level 7"
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
      Height          =   1335
      Left            =   8640
      TabIndex        =   0
      Top             =   720
      Width           =   3975
   End
End
Attribute VB_Name = "Form11"
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
Label13.BackColor = &H1D1D1D
Label14.BackColor = &H1D1D1D
Label15.BackColor = &H1D1D1D
Label16.BackColor = &H1D1D1D
End Sub

Private Sub Label15_Click()
flag = 0
End Sub

Private Sub Label14_Click()
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

Private Sub Label8_Click()
flag = 0
End Sub

Private Sub label6_Click()
flag = 0
End Sub

Private Sub label4_Click()
flag = 0
End Sub

Private Sub label1_Click()
flag = 0
End Sub

Private Sub label2_Click()
counter = counter + 1
Label2.Enabled = False
End Sub

Private Sub label3_Click()
counter = counter + 1
Label3.Enabled = False
End Sub

Private Sub label5_Click()
counter = counter + 1
Label5.Enabled = False
End Sub

Private Sub Label7_Click()
counter = counter + 1
Label7.Enabled = False
End Sub

Private Sub Label12_Click()
counter = counter + 1
Label12.Enabled = False
End Sub

Private Sub Label13_Click()
counter = counter + 1
Label13.Enabled = False
End Sub

Private Sub Label16_Click()
counter = counter + 1
Label16.Enabled = False
End Sub

Private Sub Timer1_Timer()
x = x + 2
If x = 10 Then
Label12.BackColor = &HFF00&  'green
ElseIf x = 12 Then
Label12.BackColor = &H1D1D1D 'black
ElseIf x = 14 Then
Label7.BackColor = &HFF00&
ElseIf x = 16 Then
Label7.BackColor = &H1D1D1D
ElseIf x = 18 Then
Label13.BackColor = &HFF00&
ElseIf x = 20 Then
Label13.BackColor = &H1D1D1D
ElseIf x = 22 Then
Label3.BackColor = &HFF00&
ElseIf x = 24 Then
Label3.BackColor = &H1D1D1D
ElseIf x = 26 Then
Label2.BackColor = &HFF00&
ElseIf x = 28 Then
Label2.BackColor = &H1D1D1D
ElseIf x = 30 Then
Label16.BackColor = &HFF00&
ElseIf x = 32 Then
Label16.BackColor = &H1D1D1D
ElseIf x = 34 Then
Label5.BackColor = &HFF00&
ElseIf x = 36 Then
Label5.BackColor = &H1D1D1D
ElseIf x = 38 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If counter = 7 And flag = 1 Then
MsgBox "Congartulation you cleared stage 7!"
Form7.Show
Form11.Hide
Timer2.Enabled = False

ElseIf flag = 0 Then
Form12.Label2.Caption = " " + Form9.Text1.Text + "," + " you are awarded with a silver cup!"
Form12.BackColor = &HFFFF80
Form12.Image1.Visible = True
Form12.Image2.Visible = False
Form12.Image3.Visible = False
MsgBox "You got it wrong, try again"
Timer2.Enabled = False
Timer1.Enabled = False
Form11.Hide
Unload Form6
Unload Form5
Unload Form4
Unload Form3
Unload Form1
Unload Form9
Unload Form10
Unload Form11
Form12.Show

End If
End Sub


