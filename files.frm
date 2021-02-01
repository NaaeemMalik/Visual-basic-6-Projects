VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3750
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   1440
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A, B, C, D



Private Sub Form_Load()
Form1.BackColor = vbWhite
End Sub

Private Sub Timer1_Timer()
A = 1 + A
If A = 0 Then
MsgBox ("You Are Now Roking With The Best")
Form1.BackColor = vbWhite

ElseIf A = 1 Then
Form1.BackColor = vbGreen

ElseIf A = 2 Then
Form1.BackColor = vbRed

ElseIf A = 3 Then
Form1.BackColor = vbYellow
A = -1
End If

End Sub
