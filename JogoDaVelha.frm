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
   Begin VB.TextBox TXT_HELP 
      Alignment       =   2  'Center
      Height          =   2055
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
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
Dim InicioJogo As Boolean

Dim SimboloJogadorPrincipal As String
Dim SimboloJogadorIA As String

Dim Botoes(2, 2) As Integer
Dim score As Integer
'+1 é jogador | -1 é IA | 0 é vazio

Private Function IsMovesLeft() As Boolean
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            If Botoes(i, j) = 0 Then
                IsMovesLeft = True
                Exit Function
            End If
        Next j
    Next i
    IsMovesLeft = False
End Function

Private Function IsBoardFull() As Boolean
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            If Botoes(i, j) = 0 Then
                IsBoardFull = False
                Exit Function
            End If
        Next j
    Next i
    IsBoardFull = True
End Function

Private Function Get_Best_Move() As Integer()
    Dim best_score As Integer
    best_score = -1000
    
    Dim BestMove(1) As Integer
    BestMove(0) = -1
    BestMove(1) = -1
        
    Dim i As Integer
    Dim j As Integer
    
    Dim Score_getBestMove As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            If Botoes(i, j) = 0 Then
                Botoes(i, j) = -1
                Score_getBestMove = MiniMax(0, False)
                Botoes(i, j) = 0
                ' teste
                If Score_getBestMove > best_score Then
                    best_score = Score_getBestMove
                    BestMove(0) = i
                    BestMove(1) = j
                End If
            End If
        Next j
    Next i
    Get_Best_Move = BestMove
End Function

Private Function MiniMax(ByRef Depth As Integer, ByRef is_maximizing As Boolean) As Integer
    Dim scoreEval As Integer
    scoreEval = Evaluate()
    
    If scoreEval = 10 Then ' player ganhou
        MiniMax = -1
        Exit Function
    End If
    
    If scoreEval = -10 Then ' ia ganhou
        MiniMax = 1
        Exit Function
    End If
    
    If IsMovesLeft() = False Then
        MiniMax = 0
        Exit Function
    End If
    
    If is_maximizing Then
        Dim bestScore_Max As Integer
        bestScore_Max = -1000
        Dim score As Integer
        
        For i = 0 To 2
            For j = 0 To 2
                If Botoes(i, j) = 0 Then
                    Botoes(i, j) = -1
                    score = MiniMax(Depth + 1, False)
                    Botoes(i, j) = 0
                    bestScore_Max = Max(score, bestScore_Max)
                End If
            Next j
        Next i
        MiniMax = bestScore_Max
        Exit Function
    Else
        Dim bestScore_Min As Integer
        bestScore_Min = 1000
        Dim score2 As Integer
        
        For i = 0 To 2
            For j = 0 To 2
                If Botoes(i, j) = 0 Then
                    Botoes(i, j) = 1
                    score2 = MiniMax(Depth + 1, True)
                    Botoes(i, j) = 0
                    bestScore_Min = Min(score2, bestScore_Min)
                End If
            Next j
        Next i
        MiniMax = bestScore_Min
        Exit Function
    End If
    
    
End Function

Private Sub Jogar()
    Dim resultado As Integer
    resultado = Evaluate()
    
    If resultado = 10 Then
        MsgBox "Parabéns! Você venceu!"
        DesativarBotoesJogo
        Exit Sub
    ElseIf resultado = -10 Then
        MsgBox "Você perdeu para a IA!"
        DesativarBotoesJogo
        Exit Sub
    ElseIf IsMovesLeft() = False Then
        MsgBox "Empate!"
        DesativarBotoesJogo
        Exit Sub
    End If
    
    Dim BestMoveAI() As Integer
    BestMoveAI = Get_Best_Move()
    
    Debug.Print "Melhor linha: " & BestMoveAI(0)
    Debug.Print "Melhor coluna: " & BestMoveAI(1)
    
    Debug.Print "==================================="
    
    If BestMoveAI(0) >= 0 And BestMoveAI(0) <= 2 And BestMoveAI(1) >= 0 And BestMoveAI(1) <= 2 Then
        Botoes(BestMoveAI(0), BestMoveAI(1)) = -1
        
        Select Case BestMoveAI(0) & BestMoveAI(1)
            Case "00"
                BTN_0.Caption = SimboloJogadorIA
            Case "01"
                BTN_1.Caption = SimboloJogadorIA
            Case "02"
                BTN_2.Caption = SimboloJogadorIA
            Case "10"
                BTN_3.Caption = SimboloJogadorIA
            Case "11"
                BTN_4.Caption = SimboloJogadorIA
            Case "12"
                BTN_5.Caption = SimboloJogadorIA
            Case "20"
                BTN_6.Caption = SimboloJogadorIA
            Case "21"
                BTN_7.Caption = SimboloJogadorIA
            Case "22"
                BTN_8.Caption = SimboloJogadorIA
        End Select
        
        resultado = Evaluate()
        
        If resultado = 10 Then
            MsgBox "Parabéns! Você venceu!"
            DesativarBotoesJogo
        ElseIf resultado = -10 Then
            MsgBox "Você perdeu para a IA!"
            DesativarBotoesJogo
        ElseIf IsMovesLeft() = False Then
            MsgBox "Empate!"
            DesativarBotoesJogo
        End If
    Else
        Debug.Print "Erro: Coordenadas inválidas retornadas pelo algoritmo Minimax"
    End If
End Sub


Private Function Evaluate() As Integer
    Dim row As Integer
    Dim col As Integer
    
    ' linhas
    For row = 0 To 2
        If Botoes(row, 0) = Botoes(row, 1) And Botoes(row, 1) = Botoes(row, 2) Then
            If Botoes(row, 0) = 1 Then Evaluate = 10: Exit Function

            If Botoes(row, 0) = -1 Then Evaluate = -10: Exit Function

        End If
    Next row
    
    ' colunas
    For col = 0 To 2
        If Botoes(0, col) = Botoes(1, col) And Botoes(1, col) = Botoes(2, col) Then
            If Botoes(0, col) = 1 Then Evaluate = 10: Exit Function
            
            If Botoes(0, col) = -1 Then Evaluate = -10: Exit Function
        End If
    Next col
    
    'diagonal principal
    If Botoes(0, 0) = Botoes(1, 1) And Botoes(1, 1) = Botoes(2, 2) Then
        If Botoes(0, 0) = 1 Then Evaluate = 10: Exit Function
        If Botoes(0, 0) = -1 Then Evaluate = -10: Exit Function
    End If
    
    ' diagonal secondaria
    If Botoes(0, 2) = Botoes(1, 1) And Botoes(1, 1) = Botoes(2, 0) Then
        If Botoes(0, 2) = 1 Then Evaluate = 10: Exit Function
        If Botoes(0, 2) = -1 Then Evaluate = -10: Exit Function
    End If
    
    Evaluate = 0
    
End Function

' ==========================================================
' ==========================================================
' ==========================================================
' ==========================================================
' ==========================================================

Private Sub IniciarJogo()
    AtivarBotoesJogo
End Sub

Private Sub IniciarBotoes()
    Dim i As Integer
    Dim j As Integer
    
    For i = 0 To 2
        For j = 0 To 2
            Botoes(i, j) = 0
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
        
        IniciarJogo
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
        
        IniciarJogo
    End If
End Sub

Private Sub AtualizarBarraStatus(ByRef texto As String, ByRef cor As ColorConstants)
    BarraSTATUS.Text = texto
    BarraSTATUS.BackColor = cor
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
        Botoes(0, 0) = 1
        Jogar
    End If
End Sub

Private Sub BTN_1_Click()
    If BTN_1.Caption = "" Then
        BTN_1.Caption = SimboloJogadorPrincipal
        Botoes(0, 1) = 1
        Jogar
    End If
End Sub

Private Sub BTN_2_Click()
    If BTN_2.Caption = "" Then
        BTN_2.Caption = SimboloJogadorPrincipal
        Botoes(0, 2) = 1
        Jogar
    End If
End Sub

Private Sub BTN_3_Click()
    If BTN_3.Caption = "" Then
        BTN_3.Caption = SimboloJogadorPrincipal
        Botoes(1, 0) = 1
        Jogar
    End If
End Sub

Private Sub BTN_4_Click()
    If BTN_4.Caption = "" Then
        BTN_4.Caption = SimboloJogadorPrincipal
        Botoes(1, 1) = 1
        Jogar
    End If
End Sub

Private Sub BTN_5_Click()
    If BTN_5.Caption = "" Then
        BTN_5.Caption = SimboloJogadorPrincipal
        Botoes(1, 2) = 1
        Jogar
    End If
End Sub

Private Sub BTN_6_Click()
    If BTN_6.Caption = "" Then
        BTN_6.Caption = SimboloJogadorPrincipal
        Botoes(2, 0) = 1
        Jogar
    End If
End Sub

Private Sub BTN_7_Click()
    If BTN_7.Caption = "" Then
        BTN_7.Caption = SimboloJogadorPrincipal
        Botoes(2, 1) = 1
        Jogar
    End If
End Sub

Private Sub BTN_8_Click()
    If BTN_8.Caption = "" Then
        BTN_8.Caption = SimboloJogadorPrincipal
        Botoes(2, 2) = 1
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

Private Sub PrintarMatriz()
    TXT_HELP.Text = ""
    For i = 0 To 2
        For j = 0 To 2
            TXT_HELP.Text = TXT_HELP.Text & CStr(Botoes(i, j))
        Next j
        TXT_HELP.Text = TXT_HELP.Text & vbCrLf
    Next i
End Sub
