VERSION 5.00
Begin VB.Form MENU 
   BackColor       =   &H00C0E0FF&
   Caption         =   "MENU LETS"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton BTN_BaixarPag 
      Caption         =   "Baixar WebPage"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton BTN_ExibirIP 
      Caption         =   "Exibir IP"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton BTN_MiniServer 
      Caption         =   "Mini Server"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "MENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN_BaixarPag_Click()
    DwnldWebPage.Show
End Sub

Private Sub BTN_ExibirIP_Click()
    ExibirIP.Show
End Sub

Private Sub BTN_MiniServer_Click()
    ServerLet.Show
End Sub
