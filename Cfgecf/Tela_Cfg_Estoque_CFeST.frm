VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Cfg_Estoque_CFeST 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações das C.F. e S.T. dos ítens de estoque"
   ClientHeight    =   4170
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6495
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar BP_1 
      Height          =   240
      Left            =   4560
      TabIndex        =   22
      Top             =   3936
      Width           =   1932
      _ExtentX        =   3413
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar BS_1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   3915
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FR_1 
      Caption         =   "Configurações das Classificações Fiscais e Situações Tributárias"
      Height          =   2892
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   6252
      Begin VB.Frame FR_3 
         Caption         =   "C.F. e S.T.:"
         Height          =   2292
         Left            =   4320
         TabIndex        =   13
         Top             =   360
         Width           =   1692
         Begin VB.TextBox TXT_CF 
            Height          =   288
            Left            =   120
            MaxLength       =   20
            TabIndex        =   17
            ToolTipText     =   "Nome do grupo."
            Top             =   840
            Width           =   1452
         End
         Begin VB.TextBox TXT_ST 
            Height          =   288
            Left            =   120
            MaxLength       =   30
            TabIndex        =   16
            ToolTipText     =   "Valor deste grupo."
            Top             =   1800
            Width           =   1452
         End
         Begin VB.ComboBox CB_GrupoCF 
            Height          =   288
            ItemData        =   "Tela_Cfg_Estoque_CFeST.frx":0000
            Left            =   120
            List            =   "Tela_Cfg_Estoque_CFeST.frx":0002
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            ToolTipText     =   "Grupo de internos desta figura."
            Top             =   480
            Width           =   1452
         End
         Begin VB.ComboBox CB_GrupoST 
            Height          =   288
            ItemData        =   "Tela_Cfg_Estoque_CFeST.frx":0004
            Left            =   120
            List            =   "Tela_Cfg_Estoque_CFeST.frx":0006
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            ToolTipText     =   "Grupo de internos desta figura."
            Top             =   1440
            Width           =   1452
         End
         Begin VB.Label LB_GrupoCF 
            AutoSize        =   -1  'True
            Caption         =   "Grupo da C.F.:"
            Height          =   192
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label LB_GrupoST 
            AutoSize        =   -1  'True
            Caption         =   "Grupo da S.T.:"
            Height          =   192
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1032
         End
      End
      Begin VB.ListBox LT_Figuras 
         Height          =   1230
         ItemData        =   "Tela_Cfg_Estoque_CFeST.frx":0008
         Left            =   240
         List            =   "Tela_Cfg_Estoque_CFeST.frx":000A
         Sorted          =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Lista de grupos que já existem."
         Top             =   1200
         Width           =   1812
      End
      Begin VB.Frame FR_2 
         Caption         =   "Exibir por:"
         Height          =   732
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1812
         Begin VB.OptionButton RB_Indice 
            Caption         =   "Índices de figuras"
            Height          =   192
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1572
         End
         Begin VB.OptionButton RB_Figuras 
            Caption         =   "Figuras"
            Height          =   192
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1572
         End
      End
      Begin VB.ListBox LT_Material 
         Height          =   1815
         ItemData        =   "Tela_Cfg_Estoque_CFeST.frx":000C
         Left            =   2280
         List            =   "Tela_Cfg_Estoque_CFeST.frx":000E
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Lista de grupos que já existem."
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label LB_Material 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   192
         Left            =   2280
         TabIndex        =   20
         Top             =   480
         Width           =   420
      End
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   5640
      Picture         =   "Tela_Cfg_Estoque_CFeST.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Volta à tela principal."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   732
      Left            =   4560
      Picture         =   "Tela_Cfg_Estoque_CFeST.frx":0452
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancela operação."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "Apa&gar"
      Height          =   732
      Left            =   3720
      Picture         =   "Tela_Cfg_Estoque_CFeST.frx":075C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Limpa todos os campos."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Deletar 
      Caption         =   "&Deletar"
      Height          =   732
      Left            =   1800
      Picture         =   "Tela_Cfg_Estoque_CFeST.frx":0B9E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Deleta."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   732
      Left            =   2880
      Picture         =   "Tela_Cfg_Estoque_CFeST.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salva informações."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Editar 
      Caption         =   "&Editar"
      Height          =   732
      Left            =   960
      Picture         =   "Tela_Cfg_Estoque_CFeST.frx":1422
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Edita."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Novo 
      Caption         =   "&Novo"
      Height          =   732
      Left            =   120
      Picture         =   "Tela_Cfg_Estoque_CFeST.frx":1864
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Novo."
      Top             =   3120
      Width           =   732
   End
End
Attribute VB_Name = "Tela_Cfg_Estoque_CFeST"
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
Const NOMEAPLIC As String = "Configurações das C.F. e S.T. dos ítens de estoque"
Dim I, J As Integer
Dim RespMsg, Resp
Dim ModoEdicao As Boolean
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaFR3
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (True)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (True)
    ModoEdicao = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    If RB_Indice.Value = False And RB_Figuras.Value = False Then
        MsgBox ("Selecione primeiro exibir por índice ou figura...")
        Exit Sub
    ElseIf LT_Figuras.ListIndex = -1 Then
        MsgBox ("Selecione primeira um ítem...")
        LT_Figuras.SetFocus
        Exit Sub
    ElseIf LT_Material.SelCount = 0 Then
        MsgBox ("Selecione pelo menos um material na lista...")
        LT_Material.SetFocus
        Exit Sub
    End If
    AtivaTelaEmEdicao (True)
    ModoEdicao = False
    LimpaFR3
    CB_GrupoCF.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    If RB_Indice.Value = True Then
        If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
            TelaEmEspera (True)
            ResetaBP (DLL_BD.BDSIS_TBEFG.RecordCount + 1)
            DLL_BD.BDSIS_TBEFG.MoveFirst
            Do While Not DLL_BD.BDSIS_TBEFG.EOF
                If DLL_BD.BDSIS_TBEFG_CPIFG.Value = LT_Figuras.Text Then
                    CarregaBSEP ("Salvando CF e ST...")
                    For I = 0 To LT_Material.SelCount - 1
                        If LT_Material.Selected(I) = True Then
                            DLL_BD.BDSIS_TBCFS.Seek "=", DLL_BD.BDSIS_TBEFG_CPFIG.Value, LT_Material.List(I)
                            If DLL_BD.BDSIS_TBCFS.NoMatch Then
                                DLL_BD.BDSIS_TBCFS.AddNew
                                DLL_BD.BDSIS_TBCFS_CPFIG.Value = DLL_BD.BDSIS_TBEFG_CPFIG.Value
                                DLL_BD.BDSIS_TBCFS_CPGMT.Value = LT_Material.List(I)
                            Else
                                DLL_BD.BDSIS_TBCFS.Edit
                            End If
                            DLL_BD.BDSIS_TBCFS_CPGCF.Value = CB_GrupoCF.Text
                            DLL_BD.BDSIS_TBCFS_CPGST.Value = CB_GrupoST.Text
                            DLL_BD.BDSIS_TBCFS.Update
                        End If
                    Next I
                End If
                DLL_BD.BDSIS_TBEFG.MoveNext
            Loop
        ElseIf RB_Figuras.Value = True Then
                    
                    
        End If
        ResetaBSEP
        
    End If
    DLL_FUNCS.RegistraEvento "Salvar - C.F. e S.T.", ""
ERRO_SISCOVAL:
    AtivaTelaEmEdicao (False)
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_Estoque_CFeST
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_GrupoCF_Click()
    On Error GoTo ERRO_SISCOVAL
    If CB_GrupoCF.ListIndex = -1 Then Exit Sub
    TXT_CF.Text = DLL_FUNCS.ProcuraGrupo(CB_GrupoCF.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_GrupoCF_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then CB_GrupoST.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_GrupoST_Click()
    On Error GoTo ERRO_SISCOVAL
    If CB_GrupoST.ListIndex = -1 Then Exit Sub
    TXT_ST.Text = DLL_FUNCS.ProcuraGrupo(CB_GrupoST.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_GrupoST_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then BT_Salvar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (12)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela Grupos...")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Índice...")
    If DLL_BD.AbreTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Figuras...")
    If DLL_BD.AbreTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - CF e ST...")
    If DLL_BD.AbreTabela_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Índice...")
    If DLL_BD.AbreCampos_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Figuras...")
    If DLL_BD.AbreCampos_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - CF e ST...")
    If DLL_BD.AbreCampos_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
        
    DLL_CARGA.CarregaTexto ("Carregando CF e ST...")
    CB_GrupoCF.Clear
    CB_GrupoST.Clear
    With DLL_BD.BDSIS_TBGRU
        If .RecordCount > 0 Then
            .MoveFirst
            Do While Not .EOF
                If DLL_BD.BDSIS_TBGRU_CPTIP = "CF" Then
                    CB_GrupoCF.AddItem (DLL_BD.BDSIS_TBGRU_CPGRU.Value)
                ElseIf DLL_BD.BDSIS_TBGRU_CPTIP = "ST" Then
                    CB_GrupoST.AddItem (DLL_BD.BDSIS_TBGRU_CPGRU.Value)
                End If
                .MoveNext
            Loop
        End If
    End With
    
    AtivaTelaEmEdicao (False)
    RB_Figuras.Value = False
    RB_Indice.Value = False
    DLL_FUNCS.RegistraEvento "Abrir Configurações C.F. e S.T.", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_Estoque_CFeST
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Figuras_Click()
    On Error GoTo ERRO_SISCOVAL
    CB_GrupoCF.ListIndex = -1
    CB_GrupoST.ListIndex = -1
    TXT_CF.Text = ""
    TXT_ST.Text = ""
    Dim cA As String
    If RB_Indice.Value = True Then
        If LT_Figuras.ListIndex = -1 Then Exit Sub
        DLL_BD.BDSIS_TBEID.Seek "=", LT_Figuras.Text
        If DLL_BD.BDSIS_TBEID.NoMatch Then
            MsgBox ("Ocorreu algum erro durante a procura do índice da figura.")
            Exit Sub
        End If
        TelaEmEspera (True)
        ResetaBP (1)
        CarregaBSEP ("Carregando lista de materiais...")
        'Carrega lista de materiais
        LT_Material.Clear
        cA = ""
        For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
            If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
                cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
            ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
                LT_Material.AddItem (cA)
                cA = ""
            End If
        Next I
        LT_Material.AddItem (cA)
        ResetaBSEP
        TelaEmEspera (False)
    ElseIf RB_Figuras.Value = True Then
        If LT_Figuras.ListIndex = -1 Then Exit Sub
        DLL_BD.BDSIS_TBEFG.Seek "=", LT_Figuras.Text
        If DLL_BD.BDSIS_TBEFG.NoMatch Then
            MsgBox ("Ocorreu algum erro durante a procura da figura.")
            BT_Apagar.Value = True
            Exit Sub
        End If
        DLL_BD.BDSIS_TBEID.Seek "=", DLL_BD.BDSIS_TBEFG_CPIFG.Value
        If DLL_BD.BDSIS_TBEID.NoMatch Then
            MsgBox ("Ocorreu algum erro durante a procura do índice da figura.")
            Exit Sub
        End If
        TelaEmEspera (True)
        ResetaBP (1)
        CarregaBSEP ("Carregando lista de materiais...")
        'Carrega lista de materiais
        LT_Material.Clear
        cA = ""
        For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
            If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
                cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
            ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
                LT_Material.AddItem (cA)
                cA = ""
            End If
        Next I
        LT_Material.AddItem (cA)
        ResetaBSEP
    End If
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Figuras_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then LT_Material.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Material_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaFR3
    If LT_Figuras.ListIndex = -1 Then Exit Sub
    If LT_Material.SelCount = 1 And _
       RB_Figuras.Value = True Then
        DLL_BD.BDSIS_TBCFS.Seek "=", DLL_BD.BDSIS_TBEFG_CPFIG.Value, LT_Material.Text
        If DLL_BD.BDSIS_TBCFS.NoMatch Then
            Exit Sub
        End If
        For I = 0 To CB_GrupoCF.ListCount - 1
            If CB_GrupoCF.List(I) = DLL_BD.BDSIS_TBCFS_CPGCF.Value Then
                CB_GrupoCF.ListIndex = I
                Exit For
            End If
        Next I
        For I = 0 To CB_GrupoST.ListCount - 1
            If CB_GrupoST.List(I) = DLL_BD.BDSIS_TBCFS_CPGST.Value Then
                CB_GrupoST.ListIndex = I
                Exit For
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Material_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then BT_Novo.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Figuras_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    ResetaBP (DLL_BD.BDSIS_TBEFG.RecordCount + 1)
    CarregaBSEP ("Carregando figuras...")
    'Carrega lista de figuras
    LT_Figuras.Clear
    LT_Material.Clear
    If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEFG.MoveFirst
        Do While Not DLL_BD.BDSIS_TBEFG.EOF
            LT_Figuras.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value)
            DLL_BD.BDSIS_TBEFG.MoveNext
        Loop
    Else
        MsgBox ("Não existe nenhuma figura ainda cadastrado.")
    End If
    ResetaBSEP
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Figuras_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then LT_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Indice_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    ResetaBP (DLL_BD.BDSIS_TBEID.RecordCount + 1)
    CarregaBSEP ("Carregando índices de figura...")
    'Carrega lista de índices de figura
    LT_Figuras.Clear
    LT_Material.Clear
    If DLL_BD.BDSIS_TBEID.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEID.MoveFirst
        Do While Not DLL_BD.BDSIS_TBEID.EOF
            LT_Figuras.AddItem (DLL_BD.BDSIS_TBEID_CPIFI.Value)
            DLL_BD.BDSIS_TBEID.MoveNext
        Loop
    Else
        MsgBox ("Não existe nenhum índice de figura ainda cadastrado.")
    End If
    ResetaBSEP
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Indice_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then LT_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub ST_1_Click(PreviousTab As Integer)
    On Error GoTo ERRO_SISCOVAL
    If BT_Voltar.Enabled = True Then AtivaTelaEmEdicao (False)
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
        FR_2.Enabled = False
        RB_Figuras.Enabled = False
        RB_Indice.Enabled = False
        LT_Figuras.Enabled = False
        LB_Material.Enabled = False
        LT_Material.Enabled = False
        FR_3.Enabled = True
        LB_GrupoCF.Enabled = True
        CB_GrupoCF.Enabled = True
        TXT_CF.Enabled = True
        LB_GrupoST.Enabled = True
        CB_GrupoST.Enabled = True
        TXT_ST.Enabled = True
    Else
        BT_Novo.Enabled = True
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        BT_Salvar.Enabled = False
        BT_Apagar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Voltar.Enabled = True
        FR_2.Enabled = True
        RB_Figuras.Enabled = True
        RB_Indice.Enabled = True
        LT_Figuras.Enabled = True
        LB_Material.Enabled = True
        LT_Material.Enabled = True
        FR_3.Enabled = False
        LB_GrupoCF.Enabled = False
        CB_GrupoCF.Enabled = False
        TXT_CF.Enabled = False
        LB_GrupoST.Enabled = False
        CB_GrupoST.Enabled = False
        TXT_ST.Enabled = False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_Cfg_Estoque_CFeST.MousePointer = vbHourglass
        Tela_Cfg_Estoque_CFeST.Enabled = False
    Else
        Tela_Cfg_Estoque_CFeST.MousePointer = vbDefault
        Tela_Cfg_Estoque_CFeST.Enabled = True
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
Private Static Sub LimpaFR3()
    On Error GoTo ERRO_SISCOVAL
    CB_GrupoCF.ListIndex = -1
    CB_GrupoST.ListIndex = -1
    TXT_CF.Text = ""
    TXT_ST.Text = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
