VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Begin VB.Form P3_Server_CHAT 
   BackColor       =   &H000080FF&
   Caption         =   "CHAT"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4920
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "127.0.0.1"
      LocalPort       =   31016
   End
   Begin VB.TextBox TXT_CHAT 
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "P3_Server_CHAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
