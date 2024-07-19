VERSION 5.00
Begin VB.Form Form8 
   BackColor       =   &H0012FCFC&
   BorderStyle     =   0  'None
   Caption         =   "Cong form"
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19140
   LinkTopic       =   "Form8"
   ScaleHeight     =   9660
   ScaleWidth      =   19140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF2838&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8400
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000F2&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Height          =   1695
      Left            =   13680
      TabIndex        =   5
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Congratulations"
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   11295
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   12960
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "You Cleared The World of Mind-Binder!  And got a golden cup. Hope you didnt lose your brains..."
      BeginProperty Font 
         Name            =   "Harrington"
         Size            =   62.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   2520
      TabIndex        =   0
      Top             =   2280
      Width           =   15255
   End
   Begin VB.Image Image1 
      Height          =   3570
      Left            =   8640
      Picture         =   "congform.frx":0000
      Top             =   3360
      Width           =   3180
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Form8.Hide
Form2.Show
Unload Form1
Unload Form3
Unload Form4
Unload Form5
Unload Form6
Unload Form7
Unload Form9
Unload Form10
Unload Form11
End Sub

