VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form DwnldWebPage 
   BackColor       =   &H0080C0FF&
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox BarraSTATUS 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   4440
      Width           =   7695
   End
   Begin VB.TextBox TXT_Porta 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "80"
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton BTN_Conectar 
      Caption         =   "Conectar"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton BTN_Limpar 
      Caption         =   "Limpar"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox TXT_Host 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "example.com"
      Top             =   1080
      Width           =   5655
   End
   Begin VB.TextBox TXT_WebPage 
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1680
      Width           =   7215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7080
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Somente chamadas HTTP são aceitas!    Ou seja, a porta deve ser 80"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label LB_Porta 
      BackColor       =   &H0080C0FF&
      Caption         =   "Porta"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LB_Host 
      BackColor       =   &H0080C0FF&
      Caption         =   "Host"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   495
   End
End
Attribute VB_Name = "DwnldWebPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BTN_Conectar_Click()

If TXT_Porta.Text = "443" Then Exit Sub

If Winsock1.State <> sckClosed Then Winsock1.Close

Winsock1.RemoteHost = TXT_Host.Text
Winsock1.RemotePort = TXT_Porta.Text
Winsock1.Connect

End Sub

Private Sub BTN_Limpar_Click()
TXT_WebPage.Text = ""
BarraSTATUS.BackColor = vbYellow
BarraSTATUS.Text = "Esperando Requisição!"
End Sub

Private Sub Form_Load()
BarraSTATUS.BackColor = vbYellow
BarraSTATUS.Text = "Esperando Requisição!"
End Sub

Private Sub Winsock1_Connect()

Dim strComando As String

strComando = "GET / HTTP/1.0" & vbCrLf
strComando = strComando & "Host: " & TXT_Host.Text & vbCrLf
strComando = strComando & "Accept: */*" & vbCrLf
strComando = strComando & "Accept: text/html" & vbCrLf
strComando = strComando & vbCrLf
Debug.Print (strComando)

Winsock1.SendData strComando
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim webData As String
Winsock1.GetData webData, vbString
TXT_WebPage.Text = webData

BarraSTATUS.BackColor = vbGreen
BarraSTATUS.Text = "HTML recebido de " + TXT_Host.Text
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print (Description)
Debug.Print (Winsock1.State)

BarraSTATUS.BackColor = vbRed
BarraSTATUS.Text = "Erro!" + vbCrLf + Description

Winsock1.Close
End Sub
