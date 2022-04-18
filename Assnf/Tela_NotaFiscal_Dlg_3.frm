VERSION 5.00
Begin VB.Form Tela_NotaFiscal_Dlg_3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione o bando para inserir Depósito em conta corrente na nota fiscal"
   ClientHeight    =   2295
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Tela 
      Caption         =   "Dados sobre a Conta:"
      Height          =   2292
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5292
      Begin VB.CommandButton BT_Inserir 
         Caption         =   "&Inserir"
         Height          =   732
         Left            =   3240
         Picture         =   "Tela_NotaFiscal_Dlg_3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Insere na nota fiscal"
         Top             =   1200
         Width           =   732
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   4440
         Picture         =   "Tela_NotaFiscal_Dlg_3.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Volta à tela da nota fiscal."
         Top             =   1200
         Width           =   732
      End
      Begin VB.ListBox LT_NomeConta 
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Selecione o nome de uma conta para inserir na nota fiscal"
         Top             =   240
         Width           =   1452
      End
      Begin VB.TextBox TXT_NomeBanco 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   7
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   480
         Width           =   2292
      End
      Begin VB.TextBox TXT_Bairro 
         Enabled         =   0   'False
         Height          =   288
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   6
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   480
         Width           =   1092
      End
      Begin VB.TextBox TXT_Agencia 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   1080
         Width           =   1092
      End
      Begin VB.TextBox TXT_ContaCorrente 
         Enabled         =   0   'False
         Height          =   288
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   4
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   1680
         Width           =   1092
      End
      Begin VB.Label LB_NomeBanco 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Banco:"
         Height          =   192
         Left            =   1680
         TabIndex        =   11
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label LB_Bairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   192
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   468
      End
      Begin VB.Label LB_Agencia 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
         Height          =   192
         Left            =   1680
         TabIndex        =   9
         Top             =   840
         Width           =   636
      End
      Begin VB.Label LB_ContaCorrente 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
         Height          =   192
         Left            =   1680
         TabIndex        =   8
         Top             =   1440
         Width           =   1104
      End
   End
End
Attribute VB_Name = "Tela_NotaFiscal_Dlg_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NOMEAPLIC As String = "Assistente de Bancos para Notas Fiscais"
Private Sub BT_Inserir_Click()
    On Error GoTo ERRO_SISCOVAL
    Dim NumLinha, NumErro As Integer
    NumLinha = NumErro = 0
    If LT_NomeConta.ListIndex = -1 Then
        RespMsg = MsgBox("Não foi selecionada nenhuma conta para inserir na nota fiscal.", vbOKOnly, "Assistente de Bancos")
        LT_NomeConta.SetFocus
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
        RespMsg = MsgBox("Não existem mais linha em branco para serem preenchidas no quadro dados do produto na nota fiscal.", vbOKOnly, "Assistente de Bancos")
        Unload Tela_NotaFiscal_Dlg_3
        Exit Sub
    End If

    Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 1) = "DEPÓSITO EM C/C: " & Trim(TXT_NomeBanco.Text) & " - " & Trim(TXT_Bairro.Text) & " - Agência: " & Trim(TXT_Agencia.Text) & " - C/C: " & Trim(TXT_ContaCorrente.Text)
    Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 1) = "DP"
    Unload Tela_NotaFiscal_Dlg_3
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Tela_NotaFiscal.BT_DepositoBancario.Caption = "Incluir C/C Banco"
    Unload Tela_NotaFiscal_Dlg_3
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Activate()
    On Error GoTo ERRO_SISCOVAL
    LT_NomeConta.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abrindo Tabela
    Set DLL_BD.BDSIS_TBBAN = DLL_BD.BDSIS.OpenRecordset("Bancos")
    'Abrindo campos
    Set DLL_BD.BDSIS_TBBAN_CPNMC = DLL_BD.BDSIS_TBBAN.Fields("Nome da Conta")
    Set DLL_BD.BDSIS_TBBAN_CPNMB = DLL_BD.BDSIS_TBBAN.Fields("Nome do Banco")
    Set DLL_BD.BDSIS_TBBAN_CPBAI = DLL_BD.BDSIS_TBBAN.Fields("Bairro")
    Set DLL_BD.BDSIS_TBBAN_CPAGE = DLL_BD.BDSIS_TBBAN.Fields("Agência")
    Set DLL_BD.BDSIS_TBBAN_CPCON = DLL_BD.BDSIS_TBBAN.Fields("Conta Corrente")
    DLL_BD.BDSIS_TBBAN.Index = "Nome da Conta"

    'DLL_BD.BDSIS_TBBAN.MoveFirst
    Do While Not DLL_BD.BDSIS_TBBAN.EOF
        If DLL_BD.BDSIS_TBBAN_CPNMC.Value <> "" Then
            LT_NomeConta.AddItem (DLL_BD.BDSIS_TBBAN_CPNMC.Value)
        End If
        DLL_BD.BDSIS_TBBAN.MoveNext
    Loop
    DLL_FUNCS.RegistraEvento "Abrir Bancos para Notas Fiscais", Tela_NotaFiscal.TXT_NF.Text
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Terminate()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_NotaFiscal_Dlg_3
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_NomeConta_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NomeConta.ListIndex = -1 Then
        Exit Sub
    End If
    DLL_BD.BDSIS_TBBAN.MoveFirst
    DLL_BD.BDSIS_TBBAN.Seek "=", LT_NomeConta.Text
    If DLL_BD.BDSIS_TBBAN.NoMatch Then
        RespMsg = MsgBox("Ocorreu erro durante a procura do nome da conta.", vbOKOnly, "Assistente de Bancos")
        Exit Sub
    Else
        TXT_NomeBanco.Text = DLL_BD.BDSIS_TBBAN_CPNMB.Value
        TXT_Bairro.Text = DLL_BD.BDSIS_TBBAN_CPBAI.Value
        TXT_Agencia.Text = DLL_BD.BDSIS_TBBAN_CPAGE.Value
        TXT_ContaCorrente.Text = DLL_BD.BDSIS_TBBAN_CPCON.Value
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
