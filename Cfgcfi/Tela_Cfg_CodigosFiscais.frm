VERSION 5.00
Begin VB.Form Tela_Cfg_CodigosFiscais 
   AutoRedraw      =   -1  'True
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Códigos Fiscais"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5295
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Tela 
      Caption         =   "Códigos Fiscais:"
      Height          =   3612
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5292
      Begin VB.ListBox LT_Codigos 
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   2052
      End
      Begin VB.OptionButton RD_Entrada 
         Caption         =   "Entrada"
         Height          =   252
         Left            =   3960
         TabIndex        =   13
         Top             =   2040
         Width           =   852
      End
      Begin VB.OptionButton RD_Saida 
         Caption         =   "Saída"
         Height          =   252
         Left            =   2760
         TabIndex        =   12
         Top             =   2040
         Width           =   972
      End
      Begin VB.TextBox TXT_CFOP 
         Height          =   288
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   10
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   1320
         Width           =   2772
      End
      Begin VB.TextBox TXT_NaturezaOperacao 
         Height          =   288
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   8
         ToolTipText     =   "Alíquota para Classificação Fiscal A."
         Top             =   600
         Width           =   2772
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   4440
         Picture         =   "Tela_Cfg_CodigosFiscais.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   732
         Left            =   3720
         Picture         =   "Tela_Cfg_CodigosFiscais.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancela edição."
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "Apa&gar"
         Height          =   732
         Left            =   3000
         Picture         =   "Tela_Cfg_CodigosFiscais.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Salvar 
         Caption         =   "&Salvar"
         Height          =   732
         Left            =   2280
         Picture         =   "Tela_Cfg_CodigosFiscais.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Deletar 
         Caption         =   "&Deletar"
         Height          =   732
         Left            =   1560
         Picture         =   "Tela_Cfg_CodigosFiscais.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Apagar um cadastro de uma empresa"
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Editar 
         Caption         =   "&Editar"
         Height          =   732
         Left            =   840
         Picture         =   "Tela_Cfg_CodigosFiscais.frx":1412
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Editar os cadastros de empresas existentes."
         Top             =   2760
         Width           =   732
      End
      Begin VB.CommandButton BT_Novo 
         Caption         =   "&Novo"
         Height          =   732
         Left            =   120
         Picture         =   "Tela_Cfg_CodigosFiscais.frx":1854
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Novo cadastro de empresa"
         Top             =   2760
         Width           =   732
      End
      Begin VB.Label LB_CFOP 
         AutoSize        =   -1  'True
         Caption         =   "C.F.O.P.:"
         Height          =   192
         Left            =   2400
         TabIndex        =   11
         Top             =   1080
         Width           =   612
      End
      Begin VB.Label LB_NaturezaOperacao 
         AutoSize        =   -1  'True
         Caption         =   "Natureza da Operação:"
         Height          =   192
         Left            =   2400
         TabIndex        =   9
         Top             =   360
         Width           =   1680
      End
   End
End
Attribute VB_Name = "Tela_Cfg_CodigosFiscais"
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
Const NOMEAPLIC As String = "Códigos Fiscais"
Dim I, J, NumReg As Integer
Dim RespMsg, Resp
Dim ModoEdicao As Boolean
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaCampos
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaCampos
    AtivaBotoesEmEdicao (False)
    AtivaTelaEmEdicao (False)
    LT_Codigos.Enabled = True
    BT_Novo.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Codigos.SelCount = 0 Then
        MsgBox ("Lista de códigos deve ter um índice selecionado.")
        LT_Codigos.SetFocus
        Exit Sub
    End If
    Resp = MsgBox("Deseja remover este item do banco de dados ?", vbQuestion + vbYesNo)
    If Resp = vbYes Then
        DLL_BD.BDSIS_TBCDF.Delete
        DLL_FUNCS.RegistraEvento "Deletar - Códigos Fiscais", LT_Codigos.List(LT_Codigos.ListIndex)
        LT_Codigos.RemoveItem (LT_Codigos.ListIndex)
        CarregaLista
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Codigos.SelCount = 0 Then
        MsgBox ("Lista de códigos deve ter um índice selecionado.")
        LT_Codigos.SetFocus
        Exit Sub
    End If
    LT_Codigos.Enabled = False
    AtivaBotoesEmEdicao (True)
    AtivaTelaEmEdicao (True)
    ModoEdicao = True
    TXT_NaturezaOperacao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaBotoesEmEdicao (True)
    AtivaTelaEmEdicao (True)
    LimpaCampos
    LT_Codigos.Enabled = False
    ModoEdicao = False
    TXT_NaturezaOperacao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    If ModoEdicao = True Then
        DLL_BD.BDSIS_TBCDF.Edit
    ElseIf ModoEdicao = False Then
        DLL_BD.BDSIS_TBCDF.AddNew
        DLL_BD.BDSIS_TBCDF_CPNRG.Value = NumReg + 1
    End If
    DLL_BD.BDSIS_TBCDF_CPNTO.Value = TXT_NaturezaOperacao.Text
    DLL_BD.BDSIS_TBCDF_CPCFO.Value = TXT_CFOP.Text
    If RD_Saida.Value = True Then
        DLL_BD.BDSIS_TBCDF_CPTIP.Value = "S"
    ElseIf RD_Entrada.Value = True Then
        DLL_BD.BDSIS_TBCDF_CPTIP.Value = "E"
    End If
    DLL_BD.BDSIS_TBCDF.Update
    NumReg = NumUltReg
    CarregaLista
    LimpaCampos
    AtivaBotoesEmEdicao (False)
    AtivaTelaEmEdicao (False)
    LT_Codigos.Enabled = True
    DLL_FUNCS.RegistraEvento "Salvar - Códigos Fiscais", TXT_NaturezaOperacao.Text
    BT_Novo.SetFocus
    ModoEdicao = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_CodigosFiscais
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
    DLL_CARGA.CarregaTexto ("Abrindo tabela Códigos Fiscais...")
    If DLL_BD.AbreTabela_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Códigos Fiscais...")
    If DLL_BD.AbreCampos_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega lista
    DLL_CARGA.CarregaTexto ("Carregando listas...")
    CarregaLista
        
    AtivaBotoesEmEdicao (False)
    AtivaTelaEmEdicao (False)
    ModoEdicao = False
    NumReg = NumUltReg
    DLL_FUNCS.RegistraEvento "Abrir Códigos Fiscais", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_CodigosFiscais
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Codigos_Click()
    On Error GoTo ERRO_SISCOVAL
    DLL_BD.BDSIS_TBCDF.MoveFirst
    Do While Not DLL_BD.BDSIS_TBCDF.EOF
        If DLL_BD.BDSIS_TBCDF_CPNRG.Value = LT_Codigos.Text Then
            TXT_NaturezaOperacao.Text = DLL_BD.BDSIS_TBCDF_CPNTO.Value
            TXT_CFOP.Text = DLL_BD.BDSIS_TBCDF_CPCFO.Value
            If DLL_BD.BDSIS_TBCDF_CPTIP.Value = "S" Then
                RD_Saida.Value = True
            ElseIf DLL_BD.BDSIS_TBCDF_CPTIP.Value = "E" Then
                RD_Entrada.Value = True
            End If
            Exit Do
        End If
        DLL_BD.BDSIS_TBCDF.MoveNext
    Loop
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Entrada_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_CFOP.Text <> "" Then
        BT_Salvar.SetFocus
    ElseIf KeyAscii = 13 And RD_Saida.Value = False And RD_Entrada.Value = False Then
        MsgBox ("Campo tipo deve ser preenchido")
        RD_Entrada.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Saida_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_CFOP.Text <> "" Then
        BT_Salvar.SetFocus
    ElseIf KeyAscii = 13 And RD_Saida.Value = False And RD_Entrada.Value = False Then
        MsgBox ("Campo tipo deve ser preenchido")
        RD_Saida.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CFOP_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_CFOP.Text <> "" Then
        RD_Saida.SetFocus
    ElseIf KeyAscii = 13 And TXT_CFOP.Text = "" Then
        MsgBox ("Campo C.F.O.P. deve ser preenchido")
        TXT_CFOP.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NaturezaOperacao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_NaturezaOperacao.Text <> "" Then
        TXT_CFOP.SetFocus
    ElseIf KeyAscii = 13 And TXT_NaturezaOperacao.Text = "" Then
        MsgBox ("Campo natureza do operação deve ser preenchido")
        TXT_NaturezaOperacao.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub AtivaBotoesEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    ' Valor = True -> Habilita todos controles
    If Valor = True Then
        BT_Novo.Enabled = False
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        BT_Voltar.Enabled = False
        BT_Salvar.Enabled = True
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
    Else
        BT_Novo.Enabled = True
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
        BT_Voltar.Enabled = True
        BT_Salvar.Enabled = False
        BT_Apagar.Enabled = False
        BT_Cancelar.Enabled = False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub AtivaTelaEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    ' Valor = True -> Habilita todos controles
    TXT_NaturezaOperacao.Enabled = Valor
    TXT_CFOP.Enabled = Valor
    RD_Saida.Enabled = Valor
    RD_Entrada.Enabled = Valor
    LB_NaturezaOperacao.Enabled = Valor
    LB_CFOP.Enabled = Valor
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub LimpaCampos()
    On Error GoTo ERRO_SISCOVAL
    TXT_NaturezaOperacao.Text = ""
    TXT_CFOP.Text = ""
    RD_Saida.Value = False
    RD_Entrada.Value = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub CarregaLista()
    On Error GoTo ERRO_SISCOVAL
    LT_Codigos.Clear
    DLL_BD.BDSIS_TBCDF.MoveFirst
    While Not DLL_BD.BDSIS_TBCDF.EOF
        If DLL_BD.BDSIS_TBCDF_CPNTO.Value <> "" Then
            LT_Codigos.AddItem (DLL_BD.BDSIS_TBCDF_CPNRG.Value)
        End If
        DLL_BD.BDSIS_TBCDF.MoveNext
    Wend
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function NumUltReg() As Integer
    On Error GoTo ERRO_SISCOVAL
    DLL_BD.BDSIS_TBCDF.MoveLast
    NumUltReg = DLL_BD.BDSIS_TBCDF_CPNRG.Value
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
