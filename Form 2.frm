VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "W Dec"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "W inc"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "H Dec"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "hight inc"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Height = Form1.Height + 50
End Sub

Private Sub Command2_Click()
Form1.Height = Form1.Height - 50
End Sub

Private Sub Command3_Click()
Form1.Width = Form1.Width + 50
End Sub

Private Sub Command4_Click()
Form1.Width = Form1.Width - 50
End Sub

Private Sub Form_Load()
Form1.BackColor = vbYellow
End Sub
