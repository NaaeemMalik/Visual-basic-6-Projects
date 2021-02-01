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
   Begin VB.CommandButton Command2 
      Caption         =   "Open"
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Text            =   "c:\"
      Top             =   1440
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Text            =   "E:\Bohemia\Ek Tera Pyar .mp4"
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim F1 As File
Dim F

Private Sub Command1_Click()
Set F1 = FSO.GetFile(Text1.Text)
F = Text2.Text
Print F1.Size
F1.Copy (F)

End Sub

Private Sub Form_Load()
Shell ("D:\Visual basic\C GPA Calculator....exe"), windowstyle = vbMaximizedFocus

End Sub
