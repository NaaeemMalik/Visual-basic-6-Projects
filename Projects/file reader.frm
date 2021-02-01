VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   4500
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2658.749
   ScaleMode       =   0  'User
   ScaleWidth      =   5859.022
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtU 
      Height          =   3705
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   5925
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Read"
      Height          =   390
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Write"
      Height          =   390
      Left            =   2100
      TabIndex        =   3
      Top             =   3960
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3600
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3960
      Width           =   2325
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FSO As New FileSystemObject
Dim F1 As File
Dim SF1 As TextStream
Dim S(1000)
Dim i, SN, SV, GM, j


Private Sub cmdCancel_Click()
FSO.CreateTextFile (App.Path & "\Sample.txt")
Set F1 = FSO.GetFile(App.Path & "\Sample.txt")
Set SF1 = F1.OpenAsTextStream(ForAppending)
SF1.Write (txtU.Text)
End Sub

Private Sub cmdOK_Click()
Set F1 = FSO.GetFile(App.Path & "\Sample.txt")
Set SF1 = F1.OpenAsTextStream
Do
i = i + 1
S(i) = SF1.ReadLine
If (i Mod 4 <> 1) Then

GM = Val(GM) + Val(S(i)) 'Marks Adding Code

Else
j = j + 1
If i > 2 Then
SV = GM / 3
MsgBox SN & " Has Got " & SV & " Average "
GM = 0
SN = S(i)
Else
SN = S(i)
End If
End If
If SF1.AtEndOfStream = True Then
SV = GM / 3
MsgBox SN & " Has Got " & SV & " Average "
GM = 0
End If

Loop Until SF1.AtEndOfStream = True
i = 0
End Sub

Private Sub Form_Load()
On Error Resume Next
If FSO.FileExists = False Then
FSO.CreateTextFile (App.Path & "\Sample.txt")
End If

Set F1 = FSO.GetFile(App.Path & "\Sample.txt")
Set SF1 = F1.OpenAsTextStream(ForReading)
txtU.Text = SF1.ReadAll
j = 1
End Sub
