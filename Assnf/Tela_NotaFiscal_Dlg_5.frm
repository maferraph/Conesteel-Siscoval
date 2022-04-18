VERSION 5.00
Begin VB.Form Tela_NotaFiscal_Dlg_5 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comentários para a Nota Fiscal"
   ClientHeight    =   2670
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5565
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_1 
      Caption         =   "Comentários de nota fiscal:"
      Height          =   2652
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5532
      Begin VB.CommandButton BT_Deletar 
         Caption         =   "&Deletar"
         Height          =   732
         Left            =   1560
         Picture         =   "Tela_NotaFiscal_Dlg_5.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Apaga este comentário."
         Top             =   1800
         Width           =   732
      End
      Begin VB.CommandButton BT_Editar 
         Caption         =   "&Editar"
         Height          =   732
         Left            =   840
         Picture         =   "Tela_NotaFiscal_Dlg_5.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Edita comentário existente."
         Top             =   1800
         Width           =   732
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   732
         Left            =   3840
         Picture         =   "Tela_NotaFiscal_Dlg_5.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancela operação de inserir/editar comentários."
         Top             =   1800
         Width           =   732
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   4680
         Picture         =   "Tela_NotaFiscal_Dlg_5.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Volta ao assistente da nota fiscal."
         Top             =   1800
         Width           =   732
      End
      Begin VB.TextBox TXT_Comentario 
         Height          =   612
         HideSelection   =   0   'False
         Left            =   120
         MaxLength       =   150
         ScrollBars      =   3  'Both
         TabIndex        =   1
         ToolTipText     =   "Comentários da nota fiscal."
         Top             =   1080
         Width           =   5292
      End
      Begin VB.HScrollBar BH_1 
         Height          =   252
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   5292
      End
      Begin VB.CommandButton BT_Novo 
         Caption         =   "&Novo"
         Height          =   732
         Left            =   120
         Picture         =   "Tela_NotaFiscal_Dlg_5.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Novo comentário."
         Top             =   1800
         Width           =   732
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "Apa&gar"
         Height          =   732
         Left            =   3120
         Picture         =   "Tela_NotaFiscal_Dlg_5.frx":163A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Apaga campos."
         Top             =   1800
         Width           =   732
      End
      Begin VB.CommandButton BT_Inserir 
         Caption         =   "&Inserir"
         Height          =   732
         Left            =   2400
         Picture         =   "Tela_NotaFiscal_Dlg_5.frx":1A7C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Insere comentário na nota fiscal."
         Top             =   1800
         Width           =   732
      End
      Begin VB.Label LB_Comentario 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   192
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   780
      End
      Begin VB.Label LB_Linha 
         AutoSize        =   -1  'True
         Caption         =   "Linha"
         Height          =   192
         Left            =   120
         TabIndex        =   10
         ToolTipText     =   "Altere os valores da barra de rolagem para navegar entre as linhas da nota fiscal."
         Top             =   240
         Width           =   384
      End
   End
End
Attribute VB_Name = "Tela_NotaFiscal_Dlg_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NOMEAPLIC As String = "Assistente de Comentários de Nota Fiscal"
Private Sub BH_1_Change()
    On Error GoTo ERRO_SISCOVAL
    'verifica qual o tipo de dado da nota
    VerificaLinha
    LB_Linha.Caption = "Linha " & Str(BH_1.Value + 1) & "/" & Str(BH_1.Max + 1)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Comentario.Text = ""
    TXT_Comentario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Comentario.Text = Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 1)
    AtivaTelaEmEdicao (False)
    VerificaLinha
    BT_Voltar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    Dim cResp As String
    cResp = MsgBox("Você tem certeza que deseja remover este comentário ?", vbInformation + vbYesNo + vbDefaultButton1, "Remover Comentários")
    If cResp = vbYes Then
        Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 1) = ""
        Tela_NotaFiscal.FG_2.TextMatrix(BH_1.Value + 1, 1) = ""
        TXT_Comentario.Text = ""
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (True)
    TXT_Comentario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Inserir_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (False)
    If TXT_Comentario.Text <> "" Then
        Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 1) = Trim(TXT_Comentario.Text)
        Tela_NotaFiscal.FG_2.TextMatrix(BH_1.Value + 1, 1) = "CT"
        BT_Novo.Enabled = False
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
    Else
        Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 1) = ""
        Tela_NotaFiscal.FG_2.TextMatrix(BH_1.Value + 1, 1) = ""
        BT_Novo.Enabled = True
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
    End If
    BT_Voltar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (True)
    TXT_Comentario.Text = ""
    TXT_Comentario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    'Verifica se ja tem comentários
    For I = 20 To 1 Step -1
        If Tela_NotaFiscal.FG_1.TextMatrix(I, 1) <> "" And _
           Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = "CT" Then
            Tela_NotaFiscal.BT_Comentarios.Caption = "Remover comentários"
            Exit For
        ElseIf Tela_NotaFiscal.FG_1.TextMatrix(I, 1) <> "" And _
           Tela_NotaFiscal.FG_2.TextMatrix(I, 1) <> "CT" And _
           I = 1 Then
            Tela_NotaFiscal.BT_Comentarios.Caption = "Incluir comentários"
            Exit For
        End If
    Next I
    Unload Tela_NotaFiscal_Dlg_5
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    LB_Linha.Caption = "Linha 1/20"
    BH_1.Max = 19
    AtivaTelaEmEdicao (False)
    VerificaLinha
    DLL_FUNCS.RegistraEvento "Abrir Comentários de Notas Fiscais", Tela_NotaFiscal.TXT_NF.Text
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Comentario_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Comentario.SelLength = Len(TXT_Comentario.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Comentario_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then
        TXT_Comentario.Text = Trim(TXT_Comentario.Text)
        If Len(TXT_Comentario.Text) = 0 Then
            TXT_Comentario.Text = ""
        Else
            TXT_Comentario.Text = Right(TXT_Comentario.Text, Len(TXT_Comentario.Text) - 1)
        End If
        BT_Inserir.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub

Private Sub Old()
    On Error GoTo ERRO_SISCOVAL
    Dim Resp
    Dim NumLinha, NumErro As Integer
    NumLinha = NumErro = 0
    
    'Se já existe os comentarios
    If Tela_NotaFiscal.BT_Comentarios.Caption = "Remover comentários" Then
        For I = 20 To 1 Step -1
            If Tela_NotaFiscal.FG_1.TextMatrix(I, 1) <> "" And _
               Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = "CT" Then
                Tela_NotaFiscal.FG_1.TextMatrix(I, 1) = ""
                Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = ""
                Exit For
            End If
        Next I
        Tela_NotaFiscal.BT_Comentarios.Caption = "Incluir comentários"
    
    'Inclui comentarios
    ElseIf Tela_NotaFiscal.BT_Comentarios.Caption = "Incluir comentários" Then
        'verfica se existem linhas em branco
         For I = 1 To 20
            If Tela_NotaFiscal.FG_1.TextMatrix(I, 1) = "" And _
               Tela_NotaFiscal.FG_1.TextMatrix(I, 2) = "" Then
                Exit For
            ElseIf I = 20 Then
                RespMsg = MsgBox("Não existem mais linha em branco para inserir comentários.", vbOKOnly, "Assistente de Comentários")
                Exit Sub
            End If
        Next I
        Resp = InputBox("Digite os comentários que você deseja inserir na nota fiscal:", "Inserir comentários")
        If Resp = "" Then
            Exit Sub
        End If
        For I = 20 To 1 Step -1
            If Tela_NotaFiscal.FG_1.TextMatrix(I, 1) = "" And _
               Tela_NotaFiscal.FG_1.TextMatrix(I, 2) = "" Then
                NumLinha = I
                Exit For
            Else
                If I = 1 Then
                    NumErro = 21
                End If
            End If
        Next I
        If NumErro = 21 Then
            RespMsg = MsgBox("Não existem mais linha em branco para serem preenchidas no quadro dados do produto na nota fiscal.", vbOKOnly, "Assistente de Comentários")
            Exit Sub
        End If
        Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 1) = Trim(Resp)
        Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 1) = "CT"
        Tela_NotaFiscal.BT_Comentarios.Caption = "Remover comentários"
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Sub AtivaTelaEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Valor = True Then
        BT_Novo.Enabled = False
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        BT_Inserir.Enabled = True
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
        TXT_Comentario.Enabled = True
        LB_Comentario.Enabled = True
        BH_1.Enabled = False
        LB_Linha.Enabled = False
    Else
        BT_Novo.Enabled = True
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
        BT_Inserir.Enabled = False
        BT_Apagar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Voltar.Enabled = True
        TXT_Comentario.Enabled = False
        LB_Comentario.Enabled = False
        BH_1.Enabled = True
        LB_Linha.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub VerificaLinha()
    On Error GoTo ERRO_SISCOVAL
    'verifica qual o tipo de dado da nota
    If Tela_NotaFiscal.FG_2.TextMatrix(BH_1.Value + 1, 1) <> "CT" And _
       Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 1) = "" And _
       Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 2) = "" Then
        BT_Novo.Enabled = True
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        TXT_Comentario.Text = "não existe comentário nesta linha"
    ElseIf Tela_NotaFiscal.FG_2.TextMatrix(BH_1.Value + 1, 1) <> "CT" And _
       Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 1) <> "" Or _
       Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 2) <> "" Then
        BT_Novo.Enabled = False
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        TXT_Comentario.Text = "impossível inserir comentários nesta linha porque ela contêm dados"
    Else
        BT_Novo.Enabled = False
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
        TXT_Comentario.Text = Tela_NotaFiscal.FG_1.TextMatrix(BH_1.Value + 1, 1)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
