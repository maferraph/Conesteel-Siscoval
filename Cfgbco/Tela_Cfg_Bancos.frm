VERSION 5.00
Begin VB.Form Tela_Cfg_Bancos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastros de Bancos"
   ClientHeight    =   3675
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5385
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Tela 
      Caption         =   "Dados sobre a Conta:"
      Height          =   3612
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   5292
      Begin VB.TextBox TXT_ContaCorrente 
         Height          =   288
         Left            =   4080
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   1800
         Width           =   1092
      End
      Begin VB.TextBox TXT_Agencia 
         Height          =   288
         Left            =   2880
         MaxLength       =   15
         TabIndex        =   4
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   1800
         Width           =   1092
      End
      Begin VB.TextBox TXT_Bairro 
         Height          =   288
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   1800
         Width           =   1092
      End
      Begin VB.TextBox TXT_NomeConta 
         Height          =   288
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   960
         Width           =   1212
      End
      Begin VB.TextBox TXT_NomeBanco 
         Height          =   288
         Left            =   3000
         MaxLength       =   30
         TabIndex        =   2
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   960
         Width           =   2172
      End
      Begin VB.ListBox LT_NomeConta 
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1452
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   4440
         Picture         =   "Tela_Cfg_Bancos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   732
         Left            =   3720
         Picture         =   "Tela_Cfg_Bancos.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cancela edição."
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "Apa&gar"
         Height          =   732
         Left            =   3000
         Picture         =   "Tela_Cfg_Bancos.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Salvar 
         Caption         =   "&Salvar"
         Height          =   732
         Left            =   2280
         Picture         =   "Tela_Cfg_Bancos.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Deletar 
         Caption         =   "&Deletar"
         Height          =   732
         Left            =   1560
         Picture         =   "Tela_Cfg_Bancos.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Apagar um cadastro de uma empresa"
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Editar 
         Caption         =   "&Editar"
         Height          =   732
         Left            =   840
         Picture         =   "Tela_Cfg_Bancos.frx":1412
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Editar os cadastros de empresas existentes."
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Novo 
         Caption         =   "&Novo"
         Height          =   732
         Left            =   120
         Picture         =   "Tela_Cfg_Bancos.frx":1854
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Novo cadastro de empresa"
         Top             =   2760
         Width           =   732
      End
      Begin VB.Label LB_ContaCorrente 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente:"
         Height          =   192
         Left            =   4080
         TabIndex        =   19
         Top             =   1560
         Width           =   1104
      End
      Begin VB.Label LB_Agencia 
         AutoSize        =   -1  'True
         Caption         =   "Agência:"
         Height          =   192
         Left            =   2880
         TabIndex        =   18
         Top             =   1560
         Width           =   636
      End
      Begin VB.Label LB_Bairro 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   192
         Left            =   1680
         TabIndex        =   17
         Top             =   1560
         Width           =   468
      End
      Begin VB.Label LB_NomeContaLista 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Conta:"
         Height          =   192
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1164
      End
      Begin VB.Label LB_NomeConta 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Conta:"
         Height          =   192
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   1164
      End
      Begin VB.Label LB_NomeBanco 
         AutoSize        =   -1  'True
         Caption         =   "Nome do Banco:"
         Height          =   192
         Left            =   3000
         TabIndex        =   14
         Top             =   720
         Width           =   1212
      End
   End
End
Attribute VB_Name = "Tela_Cfg_Bancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** VARIÁVEIS DLL's ****************
Dim DLL_BD As Scvbd.Classe_Scvbd
Dim DLL_CARGA As Scvcarr.Classe_Scvcarr
Dim DLL_FUNCS As Scvfunc.Classe_Scvfunc

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Cadastros de Bancos"
Dim I, J As Integer
Dim RespMsg, Resp
Dim ModoEdicao As Boolean
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_NomeConta.Text = ""
    TXT_NomeBanco.Text = ""
    TXT_Bairro.Text = ""
    TXT_Agencia.Text = ""
    TXT_ContaCorrente.Text = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaCampos (False)
    AtivaBotoesEmEdicao (False)
    BT_Apagar.Value = True
    ModoEdicao = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NomeConta.ListIndex = -1 Then
        MsgBox ("Selecione primeiro uma conta na lista")
        LT_NomeConta.SetFocus
        Exit Sub
    End If
    AtivaCampos (False)
    Resp = MsgBox("Deseja realmente remover esta conta ?", vbYesNo, "Remover conta cadastrada")
    If Resp = vbYes Then
        DLL_BD.BDSIS_TBBAN.Seek "=", LT_NomeConta.Text
        If DLL_BD.BDSIS_TBBAN.NoMatch Then
            MsgBox ("Erro ao procurar esta conta")
            Exit Sub
        Else
            DLL_BD.BDSIS_TBBAN.Delete
            DLL_FUNCS.RegistraEvento "Deletar - Cadastro de Bancos", LT_NomeConta.Text
        End If
        LT_NomeConta.Clear
        DLL_BD.BDSIS_TBBAN.MoveFirst
        Do While Not DLL_BD.BDSIS_TBBAN.EOF
            If DLL_BD.BDSIS_TBBAN_CPNMC.Value <> "" Then
                LT_NomeConta.AddItem (DLL_BD.BDSIS_TBBAN_CPNMC.Value)
            End If
            DLL_BD.BDSIS_TBBAN.MoveNext
        Loop
        BT_Apagar.Value = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    ModoEdicao = True
    If LT_NomeConta.ListIndex = -1 Then
        MsgBox ("Selecione primeiro uma conta na lista")
        LT_NomeConta.SetFocus
        Exit Sub
    End If
    AtivaCampos (True)
    AtivaBotoesEmEdicao (True)
    TXT_NomeConta.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    ModoEdicao = False
    AtivaCampos (True)
    AtivaBotoesEmEdicao (True)
    BT_Apagar.Value = True
    TXT_NomeConta.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    If TXT_NomeConta.Text = "" Then
        MsgBox ("Não foi digitado o nome da conta.")
        TXT_NomeConta.SetFocus
        Exit Sub
    ElseIf TXT_NomeBanco.Text = "" Then
        MsgBox ("Não foi digitado o nome do banco.")
        TXT_NomeBanco.SetFocus
        Exit Sub
    ElseIf TXT_Bairro.Text = "" Then
        MsgBox ("Não foi digitado o nome do bairro.")
        TXT_Bairro.SetFocus
        Exit Sub
    ElseIf TXT_Agencia.Text = "" Then
        MsgBox ("Não foi digitado o número da agência.")
        TXT_Agencia.SetFocus
        Exit Sub
    ElseIf TXT_ContaCorrente.Text = "" Then
        MsgBox ("Não foi digitado o número da conta corrente.")
        TXT_ContaCorrente.SetFocus
        Exit Sub
    End If
    DLL_BD.BDSIS_TBBAN.MoveFirst
    Do While Not DLL_BD.BDSIS_TBBAN.EOF
        If Trim(DLL_BD.BDSIS_TBBAN_CPNMC.Value) = Trim(TXT_NomeConta.Text) Then
            MsgBox ("Já existe esse nome de conta... digite outro.")
            TXT_NomeConta.SetFocus
            Exit Sub
        End If
        DLL_BD.BDSIS_TBBAN.MoveNext
    Loop
    AtivaCampos (False)
    AtivaBotoesEmEdicao (False)
    If ModoEdicao = True Then
        DLL_BD.BDSIS_TBBAN.Edit
        DLL_BD.BDSIS_TBBAN_CPNMC.Value = TXT_NomeConta.Text
        DLL_BD.BDSIS_TBBAN_CPNMB.Value = TXT_NomeBanco.Text
        DLL_BD.BDSIS_TBBAN_CPBAI.Value = TXT_Bairro.Text
        DLL_BD.BDSIS_TBBAN_CPAGE.Value = TXT_Agencia.Text
        DLL_BD.BDSIS_TBBAN_CPCON.Value = TXT_ContaCorrente.Text
        DLL_BD.BDSIS_TBBAN.Update
    Else
        DLL_BD.BDSIS_TBBAN.AddNew
        DLL_BD.BDSIS_TBBAN_CPNMC.Value = TXT_NomeConta.Text
        DLL_BD.BDSIS_TBBAN_CPNMB.Value = TXT_NomeBanco.Text
        DLL_BD.BDSIS_TBBAN_CPBAI.Value = TXT_Bairro.Text
        DLL_BD.BDSIS_TBBAN_CPAGE.Value = TXT_Agencia.Text
        DLL_BD.BDSIS_TBBAN_CPCON.Value = TXT_ContaCorrente.Text
        DLL_BD.BDSIS_TBBAN.Update
    End If
    LT_NomeConta.Clear
    DLL_BD.BDSIS_TBBAN.MoveFirst
    Do While Not DLL_BD.BDSIS_TBBAN.EOF
        If DLL_BD.BDSIS_TBBAN_CPNMC.Value <> "" Then
            LT_NomeConta.AddItem (DLL_BD.BDSIS_TBBAN_CPNMC.Value)
        End If
        DLL_BD.BDSIS_TBBAN.MoveNext
    Loop
    DLL_FUNCS.RegistraEvento "Salvar - Cadastro de Bancos", TXT_NomeConta.Text
    BT_Apagar.Value = True
    ModoEdicao = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_Bancos
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (5)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Bancos...")
    If DLL_BD.AbreTabela_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Bancos...")
    If DLL_BD.AbreCampos_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega lista de bancos
    DLL_CARGA.CarregaTexto ("Carregando lista de bancos...")
    Do While Not DLL_BD.BDSIS_TBBAN.EOF
        If DLL_BD.BDSIS_TBBAN_CPNMC.Value <> "" Then
            LT_NomeConta.AddItem (DLL_BD.BDSIS_TBBAN_CPNMC.Value)
        End If
        DLL_BD.BDSIS_TBBAN.MoveNext
    Loop
    DLL_FUNCS.RegistraEvento "Abrir Cadastro de Bancos", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    AtivaCampos (False)
    AtivaBotoesEmEdicao (False)
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_Bancos
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
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
        RespMsg = MsgBox("Ocorreu erro durante a procura do nome da conta.")
        Exit Sub
    Else
        TXT_NomeConta.Text = DLL_BD.BDSIS_TBBAN_CPNMC.Value
        TXT_NomeBanco.Text = DLL_BD.BDSIS_TBBAN_CPNMB.Value
        TXT_Bairro.Text = DLL_BD.BDSIS_TBBAN_CPBAI.Value
        TXT_Agencia.Text = DLL_BD.BDSIS_TBBAN_CPAGE.Value
        TXT_ContaCorrente.Text = DLL_BD.BDSIS_TBBAN_CPCON.Value
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Agencia_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Agencia.SelLength = Len(TXT_Agencia.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Agencia_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_ContaCorrente.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Bairro_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Bairro.SelLength = Len(TXT_Bairro.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Bairro_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Agencia.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ContaCorrente_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ContaCorrente.SelLength = Len(TXT_ContaCorrente.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ContaCorrente_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        BT_Salvar.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NomeBanco_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_NomeBanco.SelLength = Len(TXT_NomeBanco.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NomeBanco_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Bairro.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NomeConta_GotFocus()
    TXT_NomeConta.SelLength = Len(TXT_NomeConta.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NomeConta_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_NomeBanco.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NomeConta_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_NomeConta.Text = UCase(TXT_NomeConta.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
 
 

'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub AtivaCampos(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    TXT_NomeConta.Enabled = Valor
    TXT_NomeBanco.Enabled = Valor
    TXT_Bairro.Enabled = Valor
    TXT_Agencia.Enabled = Valor
    TXT_ContaCorrente.Enabled = Valor
    LB_NomeConta.Enabled = Valor
    LB_NomeBanco.Enabled = Valor
    LB_Bairro.Enabled = Valor
    LB_Agencia.Enabled = Valor
    LB_ContaCorrente.Enabled = Valor
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub AtivaBotoesEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Valor = True Then
        BT_Novo.Enabled = False
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        BT_Salvar.Enabled = True
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
    Else
        BT_Novo.Enabled = True
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
        BT_Salvar.Enabled = False
        BT_Apagar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Voltar.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
