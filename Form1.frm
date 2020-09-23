VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fourier Plotter"
   ClientHeight    =   4665
   ClientLeft      =   1665
   ClientTop       =   495
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   10260
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      Visible         =   0   'False
      X1              =   0
      X2              =   10260
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Visible         =   0   'False
      X1              =   2265
      X2              =   2265
      Y1              =   0
      Y2              =   4665
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
Form2.Show
Me.Top = 0

Form2.Top = Me.Top + Me.Height
Form2.Left = Me.Left + (Me.Width - Form2.Width) / 2

End Sub
Public Sub draw(colr As Long, drawLines As Boolean)
Dim X As Single
Me.PSet (xx1, f(xx1)), colr
    
    If drawLines Then
    
    For X = xx1 To xx2 Step Val(Form2.Spac.Text)
        Me.Line -(X, f(X)), colr
        DoEvents
        If stopped Then Exit For
    Next
    
    Else
    
    For X = xx1 To xx2 Step Val(Form2.Spac.Text)
        Me.PSet (X, f(X)), colr
        DoEvents
        If stopped Then Exit For
    Next
    
    End If
    
    

End Sub
Public Sub Clear()
Me.Cls

Me.Scale (xx1, yy1)-(xx2, yy2)
Me.Line (0, yy1)-(0, yy2)
Me.Line (xx1, 0)-(xx2, 0)
End Sub

Private Sub Form_Load()
setpoint
Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Caption = "Fourier Plotter:" & "x = " & Format(X, "0.00000") & " : y = " & Format(Y, "0.00000")
DoEvents
End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub
