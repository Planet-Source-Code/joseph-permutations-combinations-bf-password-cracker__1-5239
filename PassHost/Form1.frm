VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   735
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2040
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Text            =   "PASS"
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "password"
      Height          =   615
      Left            =   1080
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "STATUS"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Enabled = False
A = InputBox("Input password", pASSWORD)
If A = Text1.Text Then
    Text2.Text = Text2.Text & vbCrLf & A & "pASSWORD SUCCESS"
    Command1.Enabled = False
End If

    
End Sub
