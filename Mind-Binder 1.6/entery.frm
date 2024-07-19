VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00131395&
   Caption         =   "Entery"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13605
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form9"
   Picture         =   "entery.frx":0000
   ScaleHeight     =   7395
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7080
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFF00&
      Caption         =   "Start"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Corbel"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9600
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5880
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H001B1BCF&
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   8520
      TabIndex        =   0
      Top             =   4710
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter your name"
      BeginProperty Font 
         Name            =   "Agency FB"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   6720
      TabIndex        =   1
      Top             =   3360
      Width           =   8295
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c As Integer
Private Sub Command1_Click()
c = Len(Text1.Text)
If c > 10 Then
MsgBox "Too long name"
Else
Form1.Show
Form8.Label2.Caption = Text1.Text + ","
Form9.Hide
End If
End Sub

Private Sub Form_Load()
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Form2.Show
Form9.Hide
End Sub


Private Sub Text1_Change()
c = Len(Text1.Text)
If c >= 3 Then
Command1.Enabled = True
ElseIf c < 3 Then
Command1.Enabled = False
End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
c = Len(Text1.Text)
If c > 10 Then
MsgBox "Too long name"
Else
Call Command1_Click
End If
End If
End Sub
