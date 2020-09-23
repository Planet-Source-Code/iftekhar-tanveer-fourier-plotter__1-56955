Attribute VB_Name = "Module1"
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public getVal As New Class1
Public stopped As Boolean

Const pi = 3.14159265358979
Public xx1 As Single, xx2 As Single, yy1 As Single, yy2 As Single


Public Function f(X As Single)
Dim n As Integer, returnVal As Single, bn As Single
Dim Omega As Single

Omega = 6.28318530717958 / Val(Form2.Text3.Text)

If Form2.txta0 = "" And Form2.txtAn = "" And Form2.txtBn = "" Then
    
    Form2.Script1.Run "main", Val(Form2.Text1.Text), X
    f = getVal.out
    
Else
    For n = 1 To Val(Form2.Text1.Text)
        Form2.Script1.Run "main", n, X
        returnVal = returnVal + (getVal.an1 * Cos(n * Omega * X)) + (getVal.bn1 * Sin(n * Omega * X))
    Next
    Form2.Script1.Run "main2", n, X
    f = (getVal.a01 / 2) + returnVal
End If

End Function

Public Sub setpoint()
xx1 = Val(Form2.p1x.Text)
yy1 = Val(Form2.p1y.Text)
xx2 = Val(Form2.p2x.Text)
yy2 = Val(Form2.p2y.Text)

End Sub


Function StartDoc(DocName As String) As Long

Dim Scr_hDC As Long
Scr_hDC = GetDesktopWindow()
'change "Open" to "Explore" to bring up file explorer
StartDoc = ShellExecute(Form2.hDC, "Open", DocName, "", App.Path, 3)

End Function

