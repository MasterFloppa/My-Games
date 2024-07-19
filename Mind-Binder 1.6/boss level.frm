VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00FF00FF&
   Caption         =   "Boss level"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   Picture         =   "boss level.frx":0000
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   50
      Left            =   1080
      Top             =   1920
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1080
      Top             =   1080
   End
   Begin VB.Label Label17 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   14880
      TabIndex        =   20
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label20 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   14880
      TabIndex        =   19
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label19 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   14880
      TabIndex        =   18
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label18 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   14880
      TabIndex        =   17
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "Boss level"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   9240
      TabIndex        =   16
      Top             =   480
      Width           =   4575
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   12480
      TabIndex        =   15
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   10080
      TabIndex        =   14
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   7680
      TabIndex        =   13
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label13 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   5280
      TabIndex        =   12
      Top             =   8040
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   12480
      TabIndex        =   11
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   10080
      TabIndex        =   10
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   7680
      TabIndex        =   9
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   5280
      TabIndex        =   8
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   12480
      TabIndex        =   7
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   10080
      TabIndex        =   6
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   7680
      TabIndex        =   5
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   5280
      TabIndex        =   4
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   12480
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   10080
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   7680
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      BorderStyle     =   1  'Fixed Single
      Height          =   1935
      Left            =   5280
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "Form7"
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

Label1.BackColor = vbBlack
Label2.BackColor = vbBlack
Label3.BackColor = vbBlack
Label4.BackColor = vbBlack
Label5.BackColor = vbBlack
Label6.BackColor = vbBlack
Label7.BackColor = vbBlack
Label8.BackColor = vbBlack
Label9.BackColor = vbBlack
Label10.BackColor = vbBlack
Label11.BackColor = vbBlack
Label12.BackColor = vbBlack
Label13.BackColor = vbBlack
Label14.BackColor = vbBlack
Label15.BackColor = vbBlack
Label16.BackColor = vbBlack
Label17.BackColor = vbBlack
Label18.BackColor = vbBlack
Label19.BackColor = vbBlack
Label20.BackColor = vbBlack
End Sub


Private Sub Labe14_Click()
flag = 0
End Sub

Private Sub Labe18_Click()
flag = 0
End Sub

Private Sub Labe19_Click()
flag = 0
End Sub

Private Sub Labe20_Click()
flag = 0
End Sub

Private Sub Labe13_Click()
flag = 0
End Sub

Private Sub Labe12_Click()
flag = 0
End Sub

Private Sub Labe10_Click()
flag = 0
End Sub

Private Sub Labe8_Click()
flag = 0
End Sub

Private Sub Label7_Click()
flag = 0
End Sub

Private Sub label5_Click()
flag = 0
End Sub

Private Sub label4_Click()
flag = 0
End Sub

Private Sub label1_Click()
counter = counter + 1
Label1.Enabled = False
End Sub

Private Sub label2_Click()
counter = counter + 1
Label2.Enabled = False
End Sub

Private Sub label3_Click()
counter = counter + 1
Label3.Enabled = False
End Sub

Private Sub label6_Click()
counter = counter + 1
Label6.Enabled = False
End Sub

Private Sub Label9_Click()
counter = counter + 1
Label9.Enabled = False
End Sub

Private Sub Label11_Click()
counter = counter + 1
Label11.Enabled = False
End Sub

Private Sub Label15_Click()
counter = counter + 1
Label15.Enabled = False
End Sub

Private Sub Label17_Click()
counter = counter + 1
Label17.Enabled = False
End Sub

Private Sub Label16_Click()
counter = counter + 1
Label16.Enabled = False
End Sub

Private Sub Timer1_Timer()
x = x + 2
If x = 10 Then
Label3.BackColor = &HFF00FF
ElseIf x = 12 Then
Label3.BackColor = vbBlack
ElseIf x = 14 Then
Label11.BackColor = &HFF00FF
ElseIf x = 16 Then
Label11.BackColor = vbBlack
ElseIf x = 18 Then
Label2.BackColor = &HFF00FF
ElseIf x = 20 Then
Label2.BackColor = vbBlack
ElseIf x = 22 Then
Label16.BackColor = &HFF00FF
ElseIf x = 24 Then
Label16.BackColor = vbBlack
ElseIf x = 26 Then
Label15.BackColor = &HFF00FF
ElseIf x = 28 Then
Label15.BackColor = vbBlack
ElseIf x = 30 Then
Label6.BackColor = &HFF00FF
ElseIf x = 32 Then
Label6.BackColor = vbBlack
ElseIf x = 34 Then
Label1.BackColor = &HFF00FF
ElseIf x = 36 Then
Label1.BackColor = vbBlack
ElseIf x = 38 Then
Label9.BackColor = &HFF00FF
ElseIf x = 40 Then
Label9.BackColor = vbBlack
ElseIf x = 42 Then
Label17.BackColor = &HFF00FF
ElseIf x = 44 Then
Label17.BackColor = vbBlack
ElseIf x = 46 Then
Timer2.Enabled = True
Timer1.Enabled = False
End If

End Sub

Private Sub Timer2_Timer()
If counter = 9 And flag = 1 Then
MsgBox "Congartulation you cleared Boss level!!!"
Form8.Show
Form7.Hide
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
Form6.Hide
Unload Form7
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

