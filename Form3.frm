VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A word from me"
   ClientHeight    =   3915
   ClientLeft      =   1605
   ClientTop       =   3420
   ClientWidth     =   10290
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   10290
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   9030
      TabIndex        =   3
      Top             =   3150
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":0000
      Height          =   555
      Left            =   3615
      TabIndex        =   4
      Top             =   3150
      Width           =   3630
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "chayanitbd@yahoo.com     go2chayan@gmail.com"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   330
      Left            =   4395
      TabIndex        =   2
      Top             =   2595
      Width           =   4665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":0087
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   3615
      TabIndex        =   1
      Top             =   690
      Width           =   6360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fourier Plotter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4815
      TabIndex        =   0
      Top             =   45
      Width           =   3525
   End
   Begin VB.Image Image1 
      Height          =   3015
      Left            =   225
      Picture         =   "Form3.frx":032F
      Stretch         =   -1  'True
      Top             =   435
      Width           =   2970
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

