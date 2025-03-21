VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form ExibirIP 
   BackColor       =   &H0080C0FF&
   Caption         =   "Exibir IP"
   ClientHeight    =   2010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1560
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton BTN_ExibirIP 
      Caption         =   "Exibir IP"
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label LB_IP 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3135
   End
End
Attribute VB_Name = "ExibirIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN_ExibirIP_Click()

    LB_IP.Caption = Winsock1.LocalIP

End Sub

