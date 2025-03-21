VERSION 5.00
Begin VB.Form P3_HUB 
   Caption         =   "Entrar"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   2745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BTN_CONECTAR 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox TXT_Nome 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "TESTE"
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label LB_Nome 
      Caption         =   "Nome"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "P3_HUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN_CONECTAR_Click()
If TXT_Nome.Text = "" Then Exit Sub
Dim novoForm As New P3_Cliente

novoForm.LB_Nome.Caption = TXT_Nome.Text
novoForm.Show

'if P3_Server_CHAT.WindowState

P3_Server_CHAT.Show

Unload Me
End Sub
