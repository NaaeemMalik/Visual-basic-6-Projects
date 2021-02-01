VERSION 5.00
Begin VB.Form Form9 
   Caption         =   "CGPA Calculator"
   ClientHeight    =   9075
   ClientLeft      =   585
   ClientTop       =   645
   ClientWidth     =   14775
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form9"
   ScaleHeight     =   9075
   ScaleWidth      =   14775
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Height          =   2655
      Left            =   1320
      Picture         =   "Form9.frx":324A
      ScaleHeight     =   2595
      ScaleWidth      =   11115
      TabIndex        =   5
      Top             =   6000
      Width           =   11175
   End
   Begin VB.PictureBox Picture2 
      Height          =   3135
      Left            =   1320
      Picture         =   "Form9.frx":F9AC
      ScaleHeight     =   3075
      ScaleWidth      =   11115
      TabIndex        =   3
      Top             =   2400
      Width           =   11175
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   11160
      MaskColor       =   &H00FF8080&
      Picture         =   "Form9.frx":1846A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Go To Back Semister"
      Top             =   240
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   5280
      Picture         =   "Form9.frx":200F4
      ScaleHeight     =   975
      ScaleWidth      =   2895
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   3015
      Index           =   2
      Left            =   12960
      TabIndex        =   6
      Top             =   6120
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   2415
      Index           =   1
      Left            =   12720
      TabIndex        =   4
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Contect  Us  Here :"
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1920
      Width           =   10095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "This Programe  Is Developed By "
      Height          =   615
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Form9.BackColor = vbYellow
End Sub


Private Sub aboutme_Click()
Form9.Hide
Form1.Show
End Sub


Private Sub Command2_Click()
Form9.Hide
Form1.Show
End Sub

