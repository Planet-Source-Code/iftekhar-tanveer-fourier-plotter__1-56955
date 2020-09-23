VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   1845
   ClientWidth     =   12480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   12480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command7 
      Caption         =   "A word from me"
      Height          =   330
      Left            =   10695
      TabIndex        =   33
      Top             =   5280
      Width           =   1620
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   45
      ScaleHeight     =   5775
      ScaleWidth      =   12345
      TabIndex        =   16
      Top             =   -5610
      Visible         =   0   'False
      Width           =   12345
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Draw lines between points"
         Height          =   300
         Left            =   5895
         TabIndex        =   41
         Top             =   2970
         Width           =   3090
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5265
         TabIndex        =   40
         Text            =   "0"
         Top             =   4755
         Width           =   1500
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Point Locator ( Use this to locate an actual point on Graph ) "
         Height          =   1905
         Left            =   1860
         TabIndex        =   36
         Top             =   3720
         Width           =   7950
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1770
            TabIndex        =   39
            Text            =   "0"
            Top             =   1035
            Width           =   1500
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Locate"
            Height          =   435
            Left            =   6195
            TabIndex        =   38
            Top             =   1260
            Width           =   1485
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "I want to locate the point                                                    (                     ,                     )"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   1635
            TabIndex        =   37
            Top             =   315
            Width           =   3480
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10125
         ScaleHeight     =   735
         ScaleWidth      =   1230
         TabIndex        =   30
         Top             =   2685
         Width           =   1230
      End
      Begin VB.HScrollBar HScroll2 
         Height          =   270
         LargeChange     =   10
         Left            =   5880
         Max             =   150
         Min             =   10
         TabIndex        =   29
         Top             =   2565
         Value           =   80
         Width           =   3105
      End
      Begin VB.TextBox p2y 
         Height          =   285
         Left            =   7650
         TabIndex        =   27
         Text            =   "-3.14159265358979"
         Top             =   1290
         Width           =   2010
      End
      Begin VB.TextBox Spac 
         Height          =   285
         Left            =   8205
         MaxLength       =   5
         TabIndex        =   26
         Text            =   "0.08"
         Top             =   2160
         Width           =   780
      End
      Begin VB.TextBox p2x 
         Height          =   285
         Left            =   5550
         TabIndex        =   24
         Text            =   "15.70796326794895"
         Top             =   1290
         Width           =   1980
      End
      Begin VB.TextBox p1y 
         Height          =   285
         Left            =   7650
         TabIndex        =   21
         Text            =   "3.14159265358979"
         Top             =   960
         Width           =   2010
      End
      Begin VB.TextBox p1x 
         Height          =   285
         Left            =   5535
         TabIndex        =   20
         Text            =   "-15.70796326794895"
         Top             =   960
         Width           =   2010
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Back"
         Height          =   330
         Left            =   11190
         TabIndex        =   18
         Top             =   1365
         Width           =   1140
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "(0,0)"
         Height          =   270
         Left            =   3105
         TabIndex        =   32
         Top             =   1740
         Width           =   570
      End
      Begin VB.Line Line2 
         X1              =   3075
         X2              =   3075
         Y1              =   795
         Y2              =   2610
      End
      Begin VB.Line Line1 
         X1              =   1710
         X2              =   4455
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         Height          =   285
         Left            =   10125
         TabIndex        =   31
         Top             =   2445
         Width           =   900
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Space between Pixels"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5820
         TabIndex        =   28
         Top             =   2160
         Width           =   2325
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   ", ,"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   7545
         TabIndex        =   25
         Top             =   960
         Width           =   120
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   ") ) "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   9675
         TabIndex        =   23
         Top             =   885
         Width           =   120
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "P1 (  P2 ("
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   4965
         TabIndex        =   22
         Top             =   900
         Width           =   600
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   $"Form2.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   1710
         TabIndex        =   19
         Top             =   795
         Width           =   2745
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   2625
         Top             =   2790
         Width           =   840
      End
      Begin VB.Shape Shape2 
         FillStyle       =   0  'Solid
         Height          =   150
         Left            =   1845
         Top             =   3330
         Width           =   2505
      End
      Begin VB.Shape Shape4 
         FillStyle       =   0  'Solid
         Height          =   2145
         Left            =   1560
         Top             =   645
         Width           =   3060
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form2.frx":009A
      Left            =   5130
      List            =   "Form2.frx":00C2
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   990
      Width           =   5895
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1560
      Left            =   8835
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Text            =   "Form2.frx":023D
      Top             =   1755
      Width           =   3510
   End
   Begin VB.TextBox txtBn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4905
      Width           =   5250
   End
   Begin VB.TextBox txtAn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4995
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   3345
      Width           =   5160
   End
   Begin VB.TextBox txta0 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3420
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1425
      Width           =   5160
   End
   Begin MSScriptControlCtl.ScriptControl Script1 
      Left            =   9765
      Top             =   5265
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.TextBox Text1 
      Height          =   255
      Left            =   11475
      MaxLength       =   4
      TabIndex        =   5
      Top             =   3945
      Width           =   690
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   11835
      Top             =   5175
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   9135
      Max             =   1000
      Min             =   1
      TabIndex        =   0
      Top             =   4245
      Value           =   800
      Width           =   3135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Settings ..."
      Height          =   330
      Left            =   11190
      TabIndex        =   17
      Top             =   1365
      Width           =   1140
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   330
      Left            =   11190
      TabIndex        =   4
      Top             =   1050
      Width           =   1140
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   330
      Left            =   11190
      TabIndex        =   3
      Top             =   735
      Width           =   1140
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Stop"
      Height          =   330
      Left            =   11190
      TabIndex        =   9
      Top             =   420
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw"
      Height          =   330
      Left            =   11190
      TabIndex        =   2
      Top             =   105
      Width           =   1140
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6225
      TabIndex        =   34
      Text            =   "6.28318530717958"
      Top             =   2415
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5955
      TabIndex        =   35
      Top             =   2415
      Width           =   360
   End
   Begin VB.Image Image5 
      Height          =   810
      Left            =   5070
      Picture         =   "Form2.frx":0255
      Top             =   1920
      Width           =   1080
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Preset Functions f(x) ="
      Height          =   285
      Left            =   3450
      TabIndex        =   15
      Top             =   1020
      Width           =   1635
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Additional scripting"
      Height          =   450
      Left            =   8895
      TabIndex        =   13
      Top             =   1515
      Width           =   1860
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   ";"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9825
      TabIndex        =   11
      Top             =   330
      Width           =   240
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Any Periodic function f(x) of a period T can be represented by the Fourier Series:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      TabIndex        =   10
      Top             =   135
      Width           =   2580
   End
   Begin VB.Shape Shape1 
      Height          =   5805
      Left            =   15
      Top             =   0
      Width           =   12450
   End
   Begin VB.Image Image4 
      Height          =   1425
      Left            =   15
      Picture         =   "Form2.frx":3027
      Stretch         =   -1  'True
      Top             =   4380
      Width           =   4905
   End
   Begin VB.Image Image3 
      Height          =   1800
      Left            =   45
      Picture         =   "Form2.frx":1CE45
      Top             =   2625
      Width           =   4965
   End
   Begin VB.Image Image2 
      Height          =   1620
      Left            =   120
      Picture         =   "Form2.frx":3A167
      Top             =   780
      Width           =   3360
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   2670
      Picture         =   "Form2.frx":4BD29
      Top             =   45
      Width           =   7275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fourier Summation upto (n = 1 to ................)"
      Height          =   270
      Left            =   9135
      TabIndex        =   1
      Top             =   3960
      Width           =   3180
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()

Select Case Combo1.ListIndex

Case Is = 0         'Sea Wave
txta0.Text = "(1 - Exp(-2 * pi)) / pi"
txtAn.Text = "((1 - Exp(-2 * pi)) / pi)/ (1 + n ^ 2)"
txtBn.Text = "((1 - Exp(-2 * pi) * n) / pi) / (1 + n ^ 2)"
Text2.Text = "pi = 3.14159265358979"
Text3.Text = "6.28318530717958"
Check1.Value = 0

Case Is = 1         'primitive Sawtooth
txta0.Text = 0
txtAn.Text = 0
txtBn.Text = "((-1) ^ (n + 1)) / n"
Text2.Text = "pi = 3.14159265358979"
Text3.Text = "6.28318530717958"
Check1.Value = 0

Case Is = 2         'Square Wave

txta0.Text = 0
txtAn.Text = 0
txtBn.Text = "2*k*(1-cos(n*pi))/(n*pi)"
Text2.Text = "pi = 3.14159265358979" & vbCrLf & "const k = 1.5"
Text3.Text = "6.28318530717958"
Check1.Value = 1

Case Is = 3         'Saw Tooth
txta0.Text = "a/2"
txtAn.Text = "a / ((pi * n) ^ 2) * ((-1) ^ n - 1)"
txtBn.Text = "a / (pi * n * (-1) ^ (n + 1))"
Text2.Text = "pi = 3.14159265358979" & vbCrLf & "const a = 1.5"
Text3.Text = "6.28318530717958"
Check1.Value = 1

Case Is = 4         'Triangular Wave
txta0.Text = 0
txtAn.Text = "4 * (1 - (-1) ^ n) / (n * pi) ^ 2"
txtBn.Text = 0
Text2.Text = "pi = 3.14159265358979"
Text3.Text = "6.28318530717958"
Check1.Value = 0

Case Is = 5         'Half Wave Rectifier
txta0.Text = ""
txtAn.Text = ""
txtBn.Text = ""
Text2.Text = "pi = 3.14159265358979" & vbCrLf & "b = 1.5" & vbCrLf & "p = 0" & vbCrLf & _
"For i = 1 To n" & vbCrLf & "p = p + Cos(2 * i * X) / (4 * i ^ 2 - 1)" & vbCrLf & "Next" & vbCrLf & _
"out = b / pi + 0.5 * b * Sin(X) - 2 * b * p / pi"
Text3.Text = "6.28318530717958"
Check1.Value = 1


Case Is = 6          'Full Wave Rectifier
txta0.Text = "4*b/pi"
txtAn.Text = "-2 / (4 * n ^ 2 - 1)"
txtBn.Text = "0"
Text2.Text = "pi = 3.14159265358979" & vbCrLf & " b=1.5"
Text3.Text = "6.28318530717958"
Check1.Value = 1

Case Is = 7          'Minar
txta0.Text = "4*b/pi"
txtAn.Text = "2 * b * ((-1) ^ n - 1) / (pi * n * (n + 2))"
txtBn.Text = "0"
Text2.Text = "pi = 3.14159265358979" & vbCrLf & " b=1.5"
Text3.Text = "6.28318530717958"
Check1.Value = 0

Case Is = 8          ' 1/x
txta0.Text = ""
txtAn.Text = ""
txtBn.Text = ""
Text2.Text = "if x<>0 then out = 1/x"
Text3.Text = "6.28318530717958"
Check1.Value = 1

Case Is = 9          ' 1/x^2
txta0.Text = ""
txtAn.Text = ""
txtBn.Text = ""
Text2.Text = "if x<>0 then out = 1/(x^2)"
Text3.Text = "6.28318530717958"
Check1.Value = 1

Case Is = 10          ' One Dimentional Heat Flow
txta0.Text = ""
txtAn.Text = ""
txtBn.Text = ""
Text2.Text = "pi = 3.14159265358979" & vbCrLf & "l = pi" & vbCrLf & "c=1" & vbCrLf & " t = 0" & vbCrLf & _
"For i = 1 To n" & vbCrLf & _
"p = p + (Sin(i * pi / 2) * Sin(i * pi * x / l) * Exp(-t * (c * i * pi / l))) / (i ^ 2)" & vbCrLf & _
"Next" & vbCrLf & "p = p * 4 * l / (pi ^ 2)" & vbCrLf & "out = p"

Text3.Text = "6.28318530717958"
Check1.Value = 1

End Select

End Sub

Private Sub Command1_Click()
On Error GoTo hell

Dim codeText As String

codeText = "Sub Main(n,x)" & vbCrLf & Text2.Text

If txtAn.Text <> "" Then codeText = codeText & vbCrLf & "Output.an1 = " & txtAn.Text
If txtBn.Text <> "" Then codeText = codeText & vbCrLf & "Output.bn1 = " & txtBn.Text
codeText = codeText & vbCrLf & "End Sub" & vbCrLf & "Sub main2(n,x)" & vbCrLf & Text2.Text
If txta0.Text <> "" Then codeText = codeText & vbCrLf & "Output.a01 = " & txta0.Text
codeText = codeText & vbCrLf & "End sub"


Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
        
stopped = False

If Check1.Value = 0 Then
    Form1.draw Picture1.BackColor, False
Else
    Form1.draw Picture1.BackColor, True
End If

Exit Sub
hell:
End Sub

Private Sub Command2_Click()
Form1.Clear
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
stopped = True
End Sub

Private Sub Command5_Click()
Picture2.Top = 20
Picture2.Left = 40
Picture2.Visible = True
End Sub

Private Sub Command6_Click()
setpoint
Picture2.Visible = False
Form1.Line1.Visible = False
Form1.Line2.Visible = False
End Sub

Private Sub Command7_Click()
Form3.Show
End Sub

Private Sub Command8_Click()
Form1.Line1.X1 = Val(Text4.Text)
Form1.Line1.X2 = Val(Text4.Text)
Form1.Line2.Y1 = Val(Text5.Text)
Form1.Line2.Y2 = Val(Text5.Text)
Form1.Line1.Visible = True
Form1.Line2.Visible = True

End Sub

Private Sub Command9_Click()
Dim r As Long
r = StartDoc(App.Path & "\VBSCRIPT.CHM")

End Sub

Private Sub Form_Load()
Text1.Text = HScroll1.Value
Combo1.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
End Sub

Private Sub HScroll1_Change()
Text1.Text = HScroll1.Value

End Sub

Private Sub HScroll1_Scroll()
Text1.Text = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
Spac.Text = HScroll2.Value / 1000
End Sub

Private Sub HScroll2_Scroll()
Spac.Text = HScroll2.Value / 1000
End Sub

Private Sub p1x_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then p1x_LostFocus
End Sub

Private Sub p1x_LostFocus()
codeText = "Sub main()" & vbCrLf & "pi = 3.14159265358979" & vbCrLf & _
"out = " & p1x.Text & vbCrLf & "End Sub"

Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
Script1.Run "main"
p1x.Text = getVal.out

End Sub

Private Sub p1y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then p1y_LostFocus

End Sub

Private Sub p1y_LostFocus()
codeText = "Sub main()" & vbCrLf & "pi = 3.14159265358979" & vbCrLf & _
"out = " & p1y.Text & vbCrLf & "End Sub"

Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
Script1.Run "main"
p1y.Text = getVal.out

End Sub

Private Sub p2x_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then p2x_LostFocus

End Sub

Private Sub p2x_LostFocus()
codeText = "Sub main()" & vbCrLf & "pi = 3.14159265358979" & vbCrLf & _
"out = " & p2x.Text & vbCrLf & "End Sub"

Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
Script1.Run "main"
p2x.Text = getVal.out

End Sub

Private Sub p2y_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then p2y_LostFocus

End Sub

Private Sub p2y_LostFocus()
codeText = "Sub main()" & vbCrLf & "pi = 3.14159265358979" & vbCrLf & _
"out = " & p2y.Text & vbCrLf & "End Sub"

Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
Script1.Run "main"
p2y.Text = getVal.out

End Sub

Private Sub Picture1_Click()
On Error GoTo hell
cd.ShowColor
Picture1.BackColor = cd.Color

hell:
End Sub

Private Sub ScriptControl1_Error()

End Sub

Private Sub Script1_Error()

MsgBox "Error occured in one of the given expressions." & vbCrLf & _
"The following code generates an error:" & vbCrLf & _
Script1.Error.Text & vbCrLf & vbCrLf & "The Error Description was: " _
& Script1.Error.Description & vbCrLf & "To remove the error, check the equations" _
, vbCritical, "Error number:" & Script1.Error.Number
 

End Sub

Private Sub Text1_LostFocus()

If Val(Text1.Text) <= 1000 And Val(Text1.Text) > 0 Then
        HScroll1.Value = Val(Text1.Text)
    
ElseIf Val(Text1.Text) > 0 Then
    
        k = MsgBox("Calculating terms more than 1000 will take much long time. Do you want to continue anyway?", vbYesNo Or vbQuestion, "Fourier")
        
        If k = vbYes Then
            Exit Sub
        Else
            Text1.Text = HScroll1.Value
        End If
    
Else
    
        MsgBox "Invalid Number", , "Fourier"
        Text1.Text = HScroll1.Value
End If

End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
Combo1.ListIndex = Combo1.ListCount - 1

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text3_LostFocus
End Sub

Private Sub Text3_LostFocus()

codeText = "Sub main()" & vbCrLf & "pi = 3.14159265358979" & vbCrLf & _
"out = " & Text3.Text & vbCrLf & "End Sub"

Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
Script1.Run "main"
Text3.Text = getVal.out

If Val(Text3.Text) = 0 Then MsgBox "Period Must NOT be Zero", vbCritical, _
"Fourier": Text3.Text = 6.28318530717958

End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text4_LostFocus

End Sub

Private Sub Text4_LostFocus()
codeText = "Sub main()" & vbCrLf & "pi = 3.14159265358979" & vbCrLf & _
"out = " & Text4.Text & vbCrLf & "End Sub"

Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
Script1.Run "main"
Text4.Text = getVal.out

End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text5_LostFocus

End Sub

Private Sub Text5_LostFocus()
codeText = "Sub main()" & vbCrLf & "pi = 3.14159265358979" & vbCrLf & _
"out = " & Text5.Text & vbCrLf & "End Sub"

Script1.Reset
Script1.AddObject "Output", getVal, True
Script1.AddCode codeText
Script1.Run "main"
Text5.Text = getVal.out

End Sub

Private Sub txta0_KeyDown(KeyCode As Integer, Shift As Integer)
Combo1.ListIndex = Combo1.ListCount - 1

End Sub

Private Sub txtAn_KeyDown(KeyCode As Integer, Shift As Integer)
Combo1.ListIndex = Combo1.ListCount - 1

End Sub

Private Sub txtBn_KeyDown(KeyCode As Integer, Shift As Integer)
Combo1.ListIndex = Combo1.ListCount - 1

End Sub

