Attribute VB_Name = "Module1"
Sub Get_2p()
UserForm1.Hide
Dim pget(0 To 2) As Double
Dim point1 As Variant
Dim point2 As Variant
Dim dist As Variant
Dim text As String

pget(0) = 2#: pget(1) = 2#: pget(2) = 0#

point1 = ThisDrawing.Utility.GetPoint(pget)
point2 = ThisDrawing.Utility.GetPoint(pget)

Open "C:\result.txt" For Append As #1
Write #1, UserForm1.TextBox1.Value & ";" & UserForm1.TextBox3.Value * ((point2(0) - point1(0)) ^ 2 + (point2(1) - point1(1)) ^ 2) ^ 0.5 & ";" & UserForm1.TextBox2.Value
Close #1
MsgBox "X1: " & point1(0) & " Y1: " & point1(1) & "X2: " & point2(0) & " Y2: " & point2(1)
UserForm1.Show
End Sub


