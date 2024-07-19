VERSION 5.00
Begin VB.Form Form12 
   BackColor       =   &H000000F2&
   Caption         =   "Award giving"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form12"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H000FF1F7&
      Caption         =   "Restart "
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000FF1F7&
      Caption         =   "Exit "
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   2055
   End
   Begin VB.Image Image3 
      Height          =   4350
      Left            =   9240
      Picture         =   "medal.frx":0000
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   3000
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0005FEFE&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   240
      TabIndex        =   0
      Top             =   3720
      Width           =   19935
   End
   Begin VB.Image Image2 
      Height          =   10335
      Left            =   2640
      Picture         =   "medal.frx":2792
      Stretch         =   -1  'True
      Top             =   120
      Width           =   16455
   End
   Begin VB.Image Image1 
      Height          =   12765
      Left            =   2640
      Picture         =   "medal.frx":DE6C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16365
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
Form12.Hide
Unload Form7
Unload Form6
Unload Form5
Unload Form4
Unload Form3
Unload Form1
Unload Form9
Unload Form10
Unload Form11
Unload Form12
Form2.Show
End Sub







