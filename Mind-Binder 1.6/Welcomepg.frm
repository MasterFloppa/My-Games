VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FF8080&
   Caption         =   "Mind-binder"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   Picture         =   "Welcomepg.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H8000000D&
      Caption         =   "About "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7560
      Width           =   3735
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000C000&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8880
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "Click here for instructions "
      BeginProperty Font 
         Name            =   "Onyx"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   3855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C000&
      Caption         =   "Next "
      BeginProperty Font 
         Name            =   "Tempus Sans ITC"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   6015
   End
   Begin VB.Image Image1 
      Height          =   1560
      Left            =   9120
      Picture         =   "Welcomepg.frx":8B344
      Stretch         =   -1  'True
      Top             =   8640
      Width           =   3420
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "By: Samrat"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1095
      Left            =   13800
      TabIndex        =   4
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Mind: dont lose your brains!!!!!!!"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   5040
      TabIndex        =   1
      Top             =   2520
      Width           =   12855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to the world of mind-binder"
      BeginProperty Font 
         Name            =   "Niagara Solid"
         Size            =   84
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1815
      Left            =   3240
      TabIndex        =   0
      Top             =   720
      Width           =   15615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form9.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
MsgBox "This is the world of mind-binder,which is full of colorful tiles...                                   Click on the tiles that have blinked, in order to proceed to next stages.                    Please don't rush there is no time limit.                                                                           Be aware not to lose your brains!!!"
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
MsgBox "The 'World of Mind-Binder' was created by Samrat Gupta in September 2017 "
End Sub

