VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public n1 As Double, n2 As Double, opr As Integer

Public Function backspace()

Dim a As Double
a = Len(result) - 1
result = Mid(result, 1, a)
End Function

Private Sub c_Click()
result = "0"
n1 = 0
n2 = 0
End Sub

Private Sub equal_Click()
On Error GoTo error
n2 = CDbl(result)
On Error GoTo error
If opr = 0 Then
On Error GoTo error
result = n1 / n2
ElseIf opr = 1 Then
result = n1 * n2
ElseIf opr = 2 Then
result = n1 - n2
Else
result = n1 + n2
error:
MsgBox Err.Description
End If

End Sub

Private Sub multiplicate_Click()
result = n1 * n2
End Sub

Private Sub plus_Click()
result = n1 + n2
End Sub

Private Sub point_Click()
If InStr(result, ".") = 0 Then result = result & "."
End Sub

Private Sub Q_Click()
End
End Sub

Private Sub Timer1_Timer()
result.BackColor = vbGreen
Text1.Text = n1
Text2.Text = n2

Form1.Caption = "           " & Time & "                                                                    Naeem's Calculator                     " & Date

If opr = 0 Then
Text3.Text = "/"
ElseIf opr = 1 Then
Text3.Text = "*"
ElseIf opr = 2 Then
Text3.Text = "-"
ElseIf opr = 3 Then
Text3.Text = "+"
Else
Text3.Text = ""
End If
End Sub

Private Sub x1_Click()
If result = 0 Then
MsgBox "cannot find 1/x"
Else
result = 1 / result
End If
End Sub


Private Sub Command15_Click(Index As Integer)
'result = Index'

result = result & Index
'End If'
End Sub

Private Sub cos_Click()
result = Math.cos(CDbl(result))
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
equal.BackColor = vbWhite
Form1.BackColor = vbYellow
End Sub

Private Sub log_Click()
If result < 0 Then
MsgBox ("Cannot find log of this number")
Else
result = Math.log(CDbl(result))
End If
End Sub

Private Sub op_Click(Index As Integer)
n1 = result
opr = Index
result = ""
End Sub

Private Sub percent_Click()
result = n1 * CDbl(result) / 100
End Sub

Private Sub sin_Click()
result = Math.sin(CDbl(result))
End Sub

Private Sub sqrt_Click()
If result <= 0 Then
MsgBox ("Cannot find Squareroot of of negative value")
Else
result = Sqr(result)
End If
End Sub

Private Sub tan_Click()
result = Math.tan(CDbl(result))
End Sub

Private Sub x2_Click()
result = result * result
End Sub

Private Sub x3_Click()
result = result * result * result
End Sub

