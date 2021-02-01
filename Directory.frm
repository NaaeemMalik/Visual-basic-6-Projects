VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000B&
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   Icon            =   "Directory.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   4965
      Left            =   3120
      TabIndex        =   2
      Top             =   0
      Width           =   3975
   End
   Begin VB.DirListBox Dir1 
      Height          =   4365
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   2655
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Form1.BackColor = vbWhite
End Sub
