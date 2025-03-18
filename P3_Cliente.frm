VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form P3_Cliente 
   Caption         =   "Cliente"
   ClientHeight    =   2130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   ScaleHeight     =   2130
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1320
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton BTN_ENVIAR 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox TXT_Mensagem 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label LB_Nome 
      Caption         =   "Nome"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "P3_Cliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
'TXT_CHAT.Text = TXT_CHAT.Text + vbCrLf + "[" + LB_Nome.Caption + "]" + " entrou no chat!"
'Winsock1.Connect
End Sub

