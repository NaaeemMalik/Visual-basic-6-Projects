VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "di"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   960
      Top             =   960
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Dim conn As New ADODB.Connection

conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=C:\wamp\sample.mdb"

If conn.State Then
MsgBox "Connected Successfully"
End If
conn.Close

End Sub
