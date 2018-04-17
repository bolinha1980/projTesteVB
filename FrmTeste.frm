VERSION 5.00
Begin VB.Form FrmTeste 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdMult 
      Caption         =   "Multiplicacao"
      Height          =   405
      Left            =   1260
      TabIndex        =   5
      Top             =   660
      Width           =   1125
   End
   Begin VB.CommandButton CmdSoma 
      Caption         =   "Soma"
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   660
      Width           =   1125
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   270
      Width           =   1125
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   270
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Valor 2"
      Height          =   195
      Left            =   1290
      TabIndex        =   4
      Top             =   60
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Valor 1"
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   30
      Width           =   495
   End
End
Attribute VB_Name = "FrmTeste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CmdMult_Click()
    Dim CT As New ClassTeste
    MsgBox CT.FuncaoRetornaMult(Val(Text1.Text), Val(Text2.Text))
End Sub

Private Sub CmdSoma_Click()
    Dim CT As New ClassTeste
    MsgBox CT.FuncaoRetornaSoma(Val(Text1.Text), Val(Text2.Text))
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    End If
End Sub
