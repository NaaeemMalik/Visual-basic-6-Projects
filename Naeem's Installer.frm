VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Naeem's Installer"
   ClientHeight    =   5295
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   10500
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   6840
      TabIndex        =   6
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   4
      Left            =   480
      TabIndex        =   4
      Text            =   "\Adobe Photoshop 7.0\Setup.exe"
      Top             =   4080
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   3
      Left            =   480
      TabIndex        =   3
      Text            =   "\Audio_Realtek_6.0.1\Setup.exe"
      Top             =   3240
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   480
      TabIndex        =   2
      Text            =   "\vlc-2.2.1-win32.exe"
      Top             =   2400
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Text            =   "\Winrer 4.0.exe"
      Top             =   1560
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Text            =   "\avast_free_antivirus_setup.exe"
      Top             =   720
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "Defult Folder is : "
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim F As Object


Private Sub File1_Click()
Shell (File1.Path & "\" & File1.FileName)
End Sub

Private Sub Form_DblClick()
On Error GoTo error

For i = 0 To Text1.UBound
Shell (App.Path & Text1(i).Text), windowstyle = vbMaximizedFocus
Next



Exit Sub
error:
End Sub

Private Sub Form_Load()
On Error GoTo error
Label1.Caption = Label1.Caption & App.Path

File1.Path = (App.Path)


For i = 0 To Text1.UBound
MsgBox App.Path & Text1(i).Text
Shell (App.Path & Text1(i).Text), windowstyle = vbMaximizedFocus
Next



Exit Sub
error:
MsgBox i + 1 & "th  File  Name       " & App.Path & Text1(i).Text & "       Is  Not  Correct ", vbCritical, "Naeem's Installer"
End Sub

