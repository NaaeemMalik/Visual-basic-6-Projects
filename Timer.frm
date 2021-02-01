VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Timer"
   ClientHeight    =   3540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   5625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4320
      TabIndex        =   6
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   5415
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3600
      Top             =   2160
   End
   Begin VB.Label Label4 
      Height          =   735
      Left            =   1800
      TabIndex        =   7
      Top             =   2880
      Width           =   3015
   End
   Begin VB.Label Label3 
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   0
      Left            =   3000
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j, k

Private Sub Command1_Click()
i = 0
j = 0
k = 0
Label1(0).Caption = i
Label1(1).Caption = j
Label1(2).Caption = k
Label2.Caption = Time
Label3.Caption = "" & Date
End Sub

Private Sub Form_Load()
i = 0
j = 0
k = 0
Label1(0).Caption = i
Label1(1).Caption = j
Label1(2).Caption = k
Label2.Caption = Time
Label3.Caption = "" & Date
End Sub

Private Sub Timer1_Timer()
i = i + 1
Label1(0).Caption = i
Label1(1).Caption = j
Label1(2).Caption = k

If i = 60 Then
i = 0
j = j + 1
End If
If j = 60 Then
j = 0
k = k + 1
End If
Label4.Caption = Time
End Sub
