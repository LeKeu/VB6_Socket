VERSION 5.00
Begin VB.Form JogoDaVelha 
   Caption         =   "Jogo da Velha"
   ClientHeight    =   3765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXT_STATUS 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   0
      TabIndex        =   12
      Top             =   3480
      Width           =   4695
   End
   Begin VB.CommandButton BTN_O 
      Caption         =   "O"
      Height          =   375
      Left            =   4200
      TabIndex        =   11
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton BTN_X 
      Caption         =   "X"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton BTN_Zerar 
      Caption         =   "Zerar"
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton BTN_8 
      Height          =   975
      Left            =   2280
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton BTN_7 
      Height          =   975
      Left            =   1200
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton BTN_6 
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton BTN_5 
      Height          =   975
      Left            =   2280
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton BTN_4 
      Height          =   975
      Left            =   1200
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton BTN_3 
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton BTN_2 
      Height          =   975
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BTN_1 
      Height          =   975
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton BTN_0 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "JogoDaVelha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PodeEscolherSimbolo As Boolean
Dim JogoComecou As Boolean
Dim JogoAcabou As Boolean
Dim InicioJogo As Boolean

Dim SimboloJogadorPrincipal As String
Dim SimboloJogadorIA As String

Dim JogadorFoiUltimo As Boolean
Dim IAFoiUltimo As Boolean

Dim MelhorMov As Integer
Dim TemMovsSobrandoBool As Boolean

Dim Botoes(2, 2) As Integer
'0 é jogador | 1 é IA | -1 é vazio
Dim Pontuacao(2, 2) As Integer
Dim RetornoEval As Integer
Dim RetornoMiniMax As Integer

Private Function AcharMelhorMovimento() As Integer()
    Dim BestVal As Integer
    BestVal = -1000
    
    Dim BestMove(1) As Integer
    BestMove(0) = -1 ' linha
    BestMove(1) = -1 ' coluna
    
    Dim MoveVal As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            If Botoes(i, j) = -1 Then
                Botoes(i, j) = 0
                MoveVal = MiniMax(0, False)
                Botoes(i, j) = -1
                
                If MoveVal > BestVal Then
                    BestMove(0) = i ' linha
                    BestMove(1) = j ' coluna
                    BestVal = MoveVal
                End If
            End If
        Next j
    Next i
    
    AcharMelhorMovimento = BestMove
End Function

Private Function MiniMax(ByVal Depth As Integer, ByVal EhMax As Boolean) As Integer
    Dim PontuacaoEval As Integer
    
    PontuacaoEval = Evaluate()
    
    If PontuacaoEval = 10 Then
        MiniMax = PontuacaoEval
        Exit Function
    End If
    
    If PontuacaoEval = -10 Then
        MiniMax = PontuacaoEval
        Exit Function
    End If
    
    TemMovsSobrando
    
    If TemMovsSobrandoBool = False Then
        RetornoMiniMax = 0
        Exit Function
    End If
    
    If (EhMax = True) Then
        Dim MelhorMax As Integer
        MelhorMax = -1000
        
        Dim i As Integer
        Dim j As Integer
        
        For i = 0 To 2
            For j = 0 To 2
                ' checar se o espaço está vazio
                If Botoes(i, j) = -1 Then
                    Botoes(i, j) = 0 ' faz o mov
                    
                    ' chama recursivamente uau
                    MelhorMax = Max(MelhorMax, MiniMax(Depth + 1, Not EhMax))
                    
                    Botoes(i, j) = -1 ' faz o mov
                End If
            Next j
        Next i
        MiniMax = Melhor
    End If
    
    If (EhMax = False) Then
        Dim MelhorMin As Integer
        MelhorMin = 1000
        
        Dim i2 As Integer
        Dim j2 As Integer
        
        For i2 = 0 To 2
            For j2 = 0 To 2
                ' checar se o espaço está vazio
                If Botoes(i2, j2) = -1 Then
                    Botoes(i2, j2) = 0 ' faz o mov
                    
                    ' chama recursivamente uau
                    MelhorMin = Min(MelhorMin, MiniMax(Depth + 1, Not EhMax))
                    
                    Botoes(i2, j2) = -1 ' faz o mov
                End If
            Next j2
        Next i2
    End If
    
End Function

Private Function Evaluate() As Integer
    Dim Linha_e As Integer
    Dim Coluna_e As Integer
    
    For Linha_e = 0 To 2
        If Botoes(Linha_e, 0) = Botoes(Linha_e, 1) And Botoes(Linha_e, 1) = Botoes(Linha_e, 2) Then
            If (Botoes(Linha_e, 0)) = 0 Then
                Evaluate = 10
                Exit Function
            End If
            
            If (Botoes(Linha_e, 0)) = 1 Then
                Evaluate = -10
                Exit Function
            End If
            
        End If
    Next Linha_e
    
    For Coluna_e = 0 To 2
        If Botoes(0, Coluna_e) = Botoes(1, Coluna_e) And Botoes(1, Coluna_e) = Botoes(2, Coluna_e) Then
            If (Botoes(0, Coluna_e)) = 0 Then
                Evaluate = 10
                Exit Function
            End If
            
            If (Botoes(0, Coluna_e)) = 1 Then
                Evaluate = -10
                Exit Function
            End If
            
        End If
    Next Coluna_e
    
    If Botoes(0, 0) = Botoes(1, 1) And Botoes(1, 1) = Botoes(2, 2) Then
        If (Botoes(0, 0)) = 0 Then
                Evaluate = 10
                Exit Function
            End If
            
            If (Botoes(0, 0)) = 1 Then
                Evaluate = -10
                Exit Function
            End If
    End If
    
    If Botoes(0, 2) = Botoes(1, 1) And Botoes(1, 1) = Botoes(2, 0) Then
        If (Botoes(0, 2)) = 0 Then
                Evaluate = 10
                Exit Function
            End If
            
            If (Botoes(0, 2)) = 1 Then
                Evaluate = -10
                Exit Function
            End If
    End If
    
    Evaluate = 0 ' caso nenhum tenha ganhado
End Function

Private Sub TemMovsSobrando()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            If Botoes(i, j) = -1 Then
                TemMovsSobrandoBool = True
                Exit Sub
            End If
        Next j
    Next i
    TemMovsSobrandoBool = False
End Sub

Private Sub IniciarBotoes()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            Debug.Print "Definindo Botoes(" & i & "," & j & ") = -1"
            Botoes(i, j) = -1
        Next j
    Next i
End Sub

Private Sub BTN_Zerar_Click()
    Unload Me
    Load Me
    Me.Show
End Sub

Private Sub Form_Load()
    PodeEscolherSimbolo = True
    JogoComecou = False
    
    IniciarBotoes
    DesativarBotoesJogo
End Sub

Private Sub BTN_O_Click()
    If PodeEscolherSimbolo = True And JogoComecou = False Then
        SimboloJogadorPrincipal = "O"
        SimboloJogadorIA = "X"
        PodeEscolherSimbolo = False
        
        BTN_X.Enabled = False
        BTN_O.Enabled = False
        InicioJogo = True
        
        Jogar
    End If
End Sub

Private Sub BTN_X_Click()
    If PodeEscolherSimbolo = True And JogoComecou = False Then
        SimboloJogadorPrincipal = "X"
        SimboloJogadorIA = "O"
        PodeEscolherSimbolo = False
        
        BTN_X.Enabled = False
        BTN_O.Enabled = False
        InicioJogo = True
        
        Jogar
    End If
End Sub

Private Sub Jogar()
    JogoComecou = True
    
    If InicioJogo = False Then
        If JogadorFoiUltimo Then IAJoga
        If IAFoiUltimo Then JogadorJoga
    End If
    
    If InicioJogo = True Then
        Dim Vez As Integer
        'Vez = Int(Rnd * 2)
        Vez = 0
        Select Case Vez
            Case 0
                JogadorJoga
            Case 1
                IAJoga
        End Select
        InicioJogo = False
    End If
    
End Sub

Sub JogadorJoga()
    AtivarBotoesJogo
    TXT_STATUS.Text = "Sua vez!"
    TXT_STATUS.BackColor = vbGreen
    
    JogadorFoiUltimo = True
    IAFoiUltimo = False
    
    'Jogar
End Sub

Sub IAJoga()
    DesativarBotoesJogo
    TXT_STATUS.Text = "Vez da IA!"
    TXT_STATUS.BackColor = vbYellow
    
    IAFoiUltimo = True
    JogadorFoiUltimo = False
    
    Dim BestMove(1) As Integer
    Dim TempMove() As Integer

    TempMove = AcharMelhorMovimento()

    BestMove(0) = TempMove(0)
    BestMove(1) = TempMove(1)
    
    Debug.Print "Melhor linha: " & BestMove(0)
    Debug.Print "Melhor coluna: " & BestMove(1)
    
    Dim Resultado As String
    Resultado = "" + CStr(BestMove(0)) + CStr(BestMove(1))
    
    Select Case Resultado
        Case "00"
            Botoes(0, 0) = 1
            BTN_0.Caption = SimboloJogadorIA
        Case "01"
            Botoes(0, 1) = 1
            BTN_1.Caption = SimboloJogadorIA
        Case "02"
            Botoes(0, 2) = 1
            BTN_2.Caption = SimboloJogadorIA
        Case "10"
            Botoes(1, 0) = 1
            BTN_3.Caption = SimboloJogadorIA
        Case "11"
            Botoes(1, 1) = 1
            BTN_4.Caption = SimboloJogadorIA
        Case "12"
            Botoes(1, 2) = 1
            BTN_5.Caption = SimboloJogadorIA
        Case "20"
            Botoes(2, 0) = 1
            BTN_6.Caption = SimboloJogadorIA
        Case "21"
            Botoes(2, 1) = 1
            BTN_7.Caption = SimboloJogadorIA
        Case "22"
            Botoes(2, 2) = 1
            BTN_8.Caption = SimboloJogadorIA
    End Select
    
    'Jogar
End Sub

Private Sub DesativarBotoesJogo()
    BTN_0.Enabled = False
    BTN_1.Enabled = False
    BTN_2.Enabled = False
    BTN_3.Enabled = False
    BTN_4.Enabled = False
    BTN_5.Enabled = False
    BTN_6.Enabled = False
    BTN_7.Enabled = False
    BTN_8.Enabled = False
End Sub

Private Sub AtivarBotoesJogo()
    BTN_0.Enabled = True
    BTN_1.Enabled = True
    BTN_2.Enabled = True
    BTN_3.Enabled = True
    BTN_4.Enabled = True
    BTN_5.Enabled = True
    BTN_6.Enabled = True
    BTN_7.Enabled = True
    BTN_8.Enabled = True
End Sub

Private Sub BTN_0_Click()
    If BTN_0.Caption = "" Then
        BTN_0.Caption = SimboloJogadorPrincipal
        Botoes(0, 0) = 0
        Jogar
    End If
End Sub

Private Sub BTN_1_Click()
    If BTN_1.Caption = "" Then
        BTN_1.Caption = SimboloJogadorPrincipal
        Botoes(0, 1) = 0
        Jogar
    End If
End Sub

Private Sub BTN_2_Click()
    If BTN_2.Caption = "" Then
        BTN_2.Caption = SimboloJogadorPrincipal
        Botoes(0, 2) = 0
        Jogar
    End If
End Sub

Private Sub BTN_3_Click()
    If BTN_3.Caption = "" Then
        BTN_3.Caption = SimboloJogadorPrincipal
        Botoes(1, 0) = 0
        Jogar
    End If
End Sub

Private Sub BTN_4_Click()
    If BTN_4.Caption = "" Then
        BTN_4.Caption = SimboloJogadorPrincipal
        Botoes(1, 1) = 0
        Jogar
    End If
End Sub

Private Sub BTN_5_Click()
    If BTN_5.Caption = "" Then
        BTN_5.Caption = SimboloJogadorPrincipal
        Botoes(1, 2) = 0
        Jogar
    End If
End Sub

Private Sub BTN_6_Click()
    If BTN_6.Caption = "" Then
        BTN_6.Caption = SimboloJogadorPrincipal
        Botoes(2, 0) = 0
        Jogar
    End If
End Sub

Private Sub BTN_7_Click()
    If BTN_7.Caption = "" Then
        BTN_7.Caption = SimboloJogadorPrincipal
        Botoes(2, 1) = 0
        Jogar
    End If
End Sub

Private Sub BTN_8_Click()
    If BTN_8.Caption = "" Then
        BTN_8.Caption = SimboloJogadorPrincipal
        Botoes(2, 2) = 0
        Jogar
    End If
End Sub

Private Function Max(ByVal a As Integer, ByVal b As Integer) As Integer
    If a > b Then
        Max = a
    Else
        Max = b
    End If
End Function

Private Function Min(ByVal a As Integer, ByVal b As Integer) As Integer
    If a < b Then
        Min = a
    Else
        Min = b
    End If
End Function
