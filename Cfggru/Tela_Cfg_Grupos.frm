VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Cfg_Grupos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações dos grupos"
   ClientHeight    =   4065
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   5535
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar BP_1 
      Height          =   252
      Left            =   3840
      TabIndex        =   19
      Top             =   3840
      Width           =   1692
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton BT_Deletar 
      Caption         =   "&Deletar"
      Height          =   732
      Left            =   1560
      Picture         =   "Tela_Cfg_Grupos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Apaga grupo."
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton BT_Editar 
      Caption         =   "&Editar"
      Height          =   732
      Left            =   840
      Picture         =   "Tela_Cfg_Grupos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Edita grupo."
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   732
      Left            =   3840
      Picture         =   "Tela_Cfg_Grupos.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancela operação."
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "Apa&gar"
      Height          =   732
      Left            =   3120
      Picture         =   "Tela_Cfg_Grupos.frx":0B8E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Limpa todos os campos."
      Top             =   3000
      Width           =   732
   End
   Begin VB.Frame FR_1 
      Caption         =   "Descrição do Grupo"
      Height          =   2772
      Left            =   2040
      TabIndex        =   12
      Top             =   120
      Width           =   3372
      Begin VB.TextBox TXT_Tipo 
         Height          =   288
         Left            =   120
         MaxLength       =   30
         TabIndex        =   11
         ToolTipText     =   "Tipo do grupo."
         Top             =   2280
         Width           =   3132
      End
      Begin VB.TextBox TXT_Valor 
         Height          =   288
         Left            =   120
         MaxLength       =   30
         TabIndex        =   10
         ToolTipText     =   "Valor deste grupo."
         Top             =   1680
         Width           =   3132
      End
      Begin VB.TextBox TXT_Descricao 
         Height          =   288
         Left            =   120
         MaxLength       =   30
         TabIndex        =   9
         ToolTipText     =   "Descrição do grupo."
         Top             =   1080
         Width           =   3132
      End
      Begin VB.TextBox TXT_Grupo 
         Height          =   288
         Left            =   120
         MaxLength       =   20
         TabIndex        =   8
         ToolTipText     =   "Nome do grupo."
         Top             =   480
         Width           =   3132
      End
      Begin VB.Label LB_Tipo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   192
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   372
      End
      Begin VB.Label LB_Valor 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   192
         Left            =   120
         TabIndex        =   15
         Top             =   1440
         Width           =   420
      End
      Begin VB.Label LB_Grupo 
         AutoSize        =   -1  'True
         Caption         =   "Grupo:"
         Height          =   192
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   480
      End
      Begin VB.Label LB_Descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   192
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   780
      End
   End
   Begin VB.ListBox LT_Grupo 
      Height          =   2400
      ItemData        =   "Tela_Cfg_Grupos.frx":0FD0
      Left            =   120
      List            =   "Tela_Cfg_Grupos.frx":0FD2
      Sorted          =   -1  'True
      TabIndex        =   0
      ToolTipText     =   "Lista de grupos existentes."
      Top             =   360
      Width           =   1812
   End
   Begin VB.CommandButton BT_Novo 
      Caption         =   "&Novo"
      Height          =   732
      Left            =   120
      Picture         =   "Tela_Cfg_Grupos.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Novo grupo."
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   732
      Left            =   2400
      Picture         =   "Tela_Cfg_Grupos.frx":163E
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salva informações."
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   4680
      Picture         =   "Tela_Cfg_Grupos.frx":1A80
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Volta à tela principal."
      Top             =   3000
      Width           =   732
   End
   Begin MSComctlLib.StatusBar BS_1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   18
      Top             =   3816
      Width           =   5532
      _ExtentX        =   9763
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label LB_ListaGrupo 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Grupos"
      Height          =   192
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1128
   End
End
Attribute VB_Name = "Tela_Cfg_Grupos"
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
Const NOMEAPLIC As String = "Configurações dos grupos"
Dim ModoEdicao As Boolean
Dim I, J As Integer
Dim RespMsg, cResp
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    If ModoEdicao = False Then TXT_Grupo.Text = ""
    TXT_Descricao.Text = ""
    TXT_Valor.Text = ""
    TXT_Tipo.Text = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    BT_Apagar.Value = True
    LT_Grupo.ListIndex = -1
    AtivaTelaEmEdicao (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Grupo.ListIndex = -1 Then
        MsgBox ("Selecione primeiro um grupo na lista.")
        LT_Grupo.SetFocus
        Exit Sub
    End If
    cResp = MsgBox("A exclusão de um grupo pode afetar totalmente o banco de dados de estoque; só faça isso se você tem certeza absoluta de proceder esta operação. Você deseja realmente deletar um ítem de grupo ?", vbInformation + vbYesNo + vbDefaultButton2, "Deletar grupo")
    If cResp = vbYes Then
        DLL_BD.BDSIS_TBGRU.Delete
        DLL_FUNCS.RegistraEvento "Deletar - Configurações de Grupos", TXT_Grupo.Text
        LT_Grupo.RemoveItem (LT_Grupo.ListIndex)
        BT_Apagar.Value = True
        LT_Grupo.ListIndex = -1
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (True)
    ModoEdicao = True
    TXT_Grupo.Enabled = False
    LB_Grupo.Enabled = False
    TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (True)
    ModoEdicao = False
    TXT_Grupo.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Grupo.Text = "" Then
        MsgBox ("É necessário preencher todos os campos...")
        TXT_Grupo.SetFocus
        Exit Sub
    ElseIf TXT_Descricao.Text = "" Then
        MsgBox ("É necessário preencher todos os campos...")
        TXT_Descricao.SetFocus
        Exit Sub
    ElseIf TXT_Valor.Text = "" Then
        MsgBox ("É necessário preencher todos os campos...")
        TXT_Valor.SetFocus
        Exit Sub
    ElseIf TXT_Tipo.Text = "" Then
        MsgBox ("É necessário preencher todos os campos...")
        TXT_Tipo.SetFocus
        Exit Sub
    End If
    TelaEmEspera (True)
    ResetaBP (1)
    CarregaBSEP ("Salvando grupo...")
    If ModoEdicao = False Then
        DLL_BD.BDSIS_TBGRU.AddNew
    Else
        DLL_BD.BDSIS_TBGRU.Edit
    End If
    DLL_BD.BDSIS_TBGRU_CPGRU.Value = TXT_Grupo.Text
    DLL_BD.BDSIS_TBGRU_CPDES.Value = TXT_Descricao.Text
    DLL_BD.BDSIS_TBGRU_CPVAL.Value = TXT_Valor.Text
    DLL_BD.BDSIS_TBGRU_CPTIP.Value = TXT_Tipo.Text
    DLL_BD.BDSIS_TBGRU.Update
    DLL_FUNCS.RegistraEvento "Salvar - Configurações de Grupos", TXT_Grupo.Text
    ResetaBSEP
    If ModoEdicao = False Then LT_Grupo.AddItem (TXT_Grupo.Text)
    BT_Apagar.Value = True
    TXT_Grupo.Text = ""
ERRO_SISCOVAL:
    AtivaTelaEmEdicao (False)
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_Grupos
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    
    'Abre tela carregamento
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (5)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Usuários")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Usuários")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega lista de grupos
    DLL_CARGA.CarregaTexto ("Carregando lista de grupos...")
    LT_Grupo.Clear
    If DLL_BD.BDSIS_TBGRU.RecordCount > 0 Then
        DLL_BD.BDSIS_TBGRU.MoveFirst
        Do While Not DLL_BD.BDSIS_TBGRU.EOF
            LT_Grupo.AddItem (DLL_BD.BDSIS_TBGRU_CPGRU.Value)
            DLL_BD.BDSIS_TBGRU.MoveNext
        Loop
    Else
        MsgBox ("Não existe nenhum grupo ainda cadastrado.")
    End If
                    
    DLL_FUNCS.RegistraEvento "Abrir Configurações de Grupos", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    AtivaTelaEmEdicao (False)
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_Grupos
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Grupo_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Grupo.ListIndex = -1 Then Exit Sub
    DLL_BD.BDSIS_TBGRU.Seek "=", LT_Grupo.Text
    If DLL_BD.BDSIS_TBGRU.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura do grupo no banco de dados.")
        Exit Sub
    End If
    TXT_Grupo.Text = DLL_BD.BDSIS_TBGRU_CPGRU.Value
    TXT_Descricao.Text = DLL_BD.BDSIS_TBGRU_CPDES.Value
    TXT_Valor.Text = DLL_BD.BDSIS_TBGRU_CPVAL.Value
    TXT_Tipo.Text = DLL_BD.BDSIS_TBGRU_CPTIP.Value
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Grupo_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Novo.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.SelLength = Len(TXT_Descricao.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Valor.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Grupo_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Grupo.SelLength = Len(TXT_Grupo.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Grupo_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Grupo_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    If ModoEdicao = False Then
        For I = 0 To LT_Grupo.ListCount - 1
            If LT_Grupo.List(I) = TXT_Grupo.Text Then
                MsgBox ("Esse grupo já existe. Digite novamente.")
                TXT_Grupo.SetFocus
                Exit Sub
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Tipo_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Tipo.SelLength = Len(TXT_Tipo.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Tipo_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Salvar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Valor_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Valor.SelLength = Len(TXT_Valor.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Valor_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Tipo.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub AtivaTelaEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Valor = True Then
        BT_Novo.Enabled = False
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        BT_Salvar.Enabled = True
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
        LB_ListaGrupo.Enabled = False
        LT_Grupo.Enabled = False
        FR_1.Enabled = True
        LB_Grupo.Enabled = True
        TXT_Grupo.Enabled = True
        LB_Descricao.Enabled = True
        TXT_Descricao.Enabled = True
        LB_Valor.Enabled = True
        TXT_Valor.Enabled = True
        LB_Tipo.Enabled = True
        TXT_Tipo.Enabled = True
    Else
        BT_Novo.Enabled = True
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
        BT_Salvar.Enabled = False
        BT_Apagar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Voltar.Enabled = True
        LB_ListaGrupo.Enabled = True
        LT_Grupo.Enabled = True
        FR_1.Enabled = False
        LB_Grupo.Enabled = False
        TXT_Grupo.Enabled = False
        LB_Descricao.Enabled = False
        TXT_Descricao.Enabled = False
        LB_Valor.Enabled = False
        TXT_Valor.Enabled = False
        LB_Tipo.Enabled = False
        TXT_Tipo.Enabled = False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_Cfg_Grupos.MousePointer = vbHourglass
        Tela_Cfg_Grupos.Enabled = False
    Else
        Tela_Cfg_Grupos.MousePointer = vbDefault
        Tela_Cfg_Grupos.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub CarregaBSEP(Texto As String)
    On Error GoTo ERRO_SISCOVAL
    BS_1.SimpleText = Texto
    BP_1.Value = BP_1.Value + 1
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ResetaBP(Max As Integer)
    On Error GoTo ERRO_SISCOVAL
    BP_1.Max = Max
    BP_1.Value = 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ResetaBSEP()
    On Error GoTo ERRO_SISCOVAL
    BP_1.Value = 0
    BS_1.SimpleText = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
