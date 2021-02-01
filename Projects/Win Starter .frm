VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Book Antiqua"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1080
      Top             =   1920
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4440
      TabIndex        =   0
      Top             =   3840
      Width           =   7095
   End
   Begin VB.Label Label3 
      Caption         =   "If The Name Matches I Will Tell You Otherwise Computer Will Shutdown"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   15255
   End
   Begin VB.Label Label2 
      Caption         =   "Or Write Else Your Name The Computer Will Shutdown"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      TabIndex        =   3
      Top             =   9840
      Width           =   13935
   End
   Begin VB.Label f 
      Caption         =   "If You Do Not Write "
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   2
      Top             =   9000
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "Type Here Your Name :"
      Height          =   975
      Left            =   4440
      TabIndex        =   1
      Top             =   2760
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S, M, G, Shutdown, T, H(100), i
Dim FSO As New FileSystemObject
Dim F1, F2 As File
Dim SF1, SF2 As TextStream

Private Sub Form_Load()
On Error Resume Next
Timer1.Enabled = True
If FSO.FileExists("D:\Books\Funny Books\Sample.txt") = False Then
FSO.CreateTextFile ("D:\Books\Funny Books\Sample.txt")
End If
If FSO.FileExists("D:\Books\Funny Books\History.txt") = False Then
FSO.CreateTextFile ("D:\Books\Funny Books\History.txt")
End If

i = 1
End Sub

Private Sub Timer1_Timer()
Set F1 = FSO.GetFile("D:\Books\Funny Books\Sample.txt")
Set SF1 = F1.OpenAsTextStream(ForReading)
S = S + 1
a = Checker()
If S > 30 And G = True Then
If Shutdown = False Then
Shell "shutdown.exe -s"
Shutdown = True
'MsgBox "ShuttingDown"
End If
End If

If G = False Then
If Shutdown = True Then
Shell "shutdown.exe -a"
Shutdown = False
End If
Set F2 = FSO.GetFile("D:\Books\Funny Books\History.txt")
Set SF2 = F2.OpenAsTextStream(ForAppending)
SF2.WriteLine Text1.Text
SF2.WriteLine Now
SF2.WriteLine (" ")
Timer1.Enabled = False
Form1.Visible = False
End If
End Sub

Function Checker()
T = Text1.Text
Do
H(i) = Trim(SF1.ReadLine)
If Text1.Text = H(i) Then
G = False
Exit Do
Else
G = True
End If
i = 1 + i
Loop Until SF1.AtEndOfLine = True
i = 1
End Function
