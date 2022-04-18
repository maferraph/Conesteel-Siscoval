VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Cfg_Estoque_Declaracoes 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações das declarações de ítens de estoque para nota fiscal"
   ClientHeight    =   4170
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5295
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar BP_1 
      Height          =   240
      Left            =   3840
      TabIndex        =   15
      Top             =   3948
      Width           =   1452
      _ExtentX        =   2566
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar BS_1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   14
      Top             =   3915
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FR_3 
      Height          =   3012
      Left            =   2160
      TabIndex        =   11
      Top             =   0
      Width           =   3012
      Begin VB.TextBox TXT_DecOutros 
         Height          =   1092
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Declarações fiscais para demais estados."
         Top             =   1800
         Width           =   2652
      End
      Begin VB.TextBox TXT_DecSP 
         Height          =   1092
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   5
         ToolTipText     =   "Declaração fiscal para São Paulo."
         Top             =   360
         Width           =   2652
      End
      Begin VB.Label LB_DecOutros 
         AutoSize        =   -1  'True
         Caption         =   "Demais declarações:"
         Height          =   192
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   1548
      End
      Begin VB.Label LB_DecSP 
         AutoSize        =   -1  'True
         Caption         =   "Declaração em S.P.:"
         Height          =   192
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1476
      End
   End
   Begin VB.Frame FR_2 
      Height          =   3012
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   1932
      Begin VB.ListBox LT_CF 
         Height          =   1620
         ItemData        =   "Tela_Cfg_Estoque_Declaracoes.frx":0000
         Left            =   120
         List            =   "Tela_Cfg_Estoque_Declaracoes.frx":0007
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Lista de classificações fiscais cadastradas."
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Frame FR_2_2 
         Caption         =   "Tipo do Grupo"
         Height          =   612
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1692
         Begin VB.ComboBox CB_TG 
            Height          =   288
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   9
            TabStop         =   0   'False
            ToolTipText     =   "Só altere esta lista se você tem certeza que os dados exibidos na lista abaixo não são os esperados (neste caso, CF)."
            Top             =   240
            Width           =   1452
         End
      End
      Begin VB.Label LB_CF 
         AutoSize        =   -1  'True
         Caption         =   "Lista de CF:"
         Height          =   192
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   840
      End
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   4440
      Picture         =   "Tela_Cfg_Estoque_Declaracoes.frx":0012
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Volta à tela principal."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   732
      Left            =   3000
      Picture         =   "Tela_Cfg_Estoque_Declaracoes.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancela operação."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   732
      Left            =   1560
      Picture         =   "Tela_Cfg_Estoque_Declaracoes.frx":075E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salva informações."
      Top             =   3120
      Width           =   732
   End
   Begin VB.CommandButton BT_Editar 
      Caption         =   "&Editar"
      Height          =   732
      Left            =   120
      Picture         =   "Tela_Cfg_Estoque_Declaracoes.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Edita."
      Top             =   3120
      Width           =   732
   End
End
Attribute VB_Name = "Tela_Cfg_Estoque_Declaracoes"
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
Const NOMEAPLIC As String = "Configurações das declarações de ítens de estoque para nota fiscal"
Dim I, J As Integer
Dim RespMsg, cResp
Dim ModoEdicao As Boolean
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (False)
    LT_CF.ListIndex = -1
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_CF.ListIndex = -1 Then
        MsgBox ("Você deve primeiro selecionar uma classificação fiscal na lista...")
        LT_CF.SetFocus
        Exit Sub
    End If
    AtivaTelaEmEdicao (True)
    TXT_DecSP.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    ResetaBP (1)
    DLL_BD.BDSIS_TBEDC.Seek "=", LT_CF.Text
    CarregaBSEP ("Salvando declarações...")
    If DLL_BD.BDSIS_TBEDC.NoMatch Then
        DLL_BD.BDSIS_TBEDC.AddNew
        DLL_BD.BDSIS_TBEDC_CPCCF.Value = LT_CF.Text
    Else
        DLL_BD.BDSIS_TBEDC.Edit
    End If
    DLL_BD.BDSIS_TBEDC_CPDSP.Value = TXT_DecSP.Text
    DLL_BD.BDSIS_TBEDC_CPDOU.Value = TXT_DecOutros.Text
    DLL_BD.BDSIS_TBEDC.Update
    DLL_FUNCS.RegistraEvento "Salvar - Declarações", LT_CF.Text
    AtivaTelaEmEdicao (False)
    ResetaBSEP
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_Estoque_Declaracoes
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_TG_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    LT_CF.Clear
    DLL_BD.BDSIS_TBGRU.MoveFirst
    Do While Not DLL_BD.BDSIS_TBGRU.EOF
        If DLL_BD.BDSIS_TBGRU_CPTIP.Value = CB_TG.Text Then
            LT_CF.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
        End If
        DLL_BD.BDSIS_TBGRU.MoveNext
    Loop
    cResp = SalvaCB_TG(CB_TG.Text, Tela_Cfg_Estoque_Declaracoes.Name, CB_TG.Name)
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_TG_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then LT_CF.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (7)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela Grupos...")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Declarações...")
    If DLL_BD.AbreTabela_EstoqueDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abre campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Declarações...")
    If DLL_BD.AbreCampos_EstoqueDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega lista de grupos
    DLL_CARGA.CarregaTexto ("Carregando lista de grupos...")
    cResp = CarregaCB_TG(Tela_Cfg_Estoque_Declaracoes.CB_TG)
    cResp = LeCB_TG(Tela_Cfg_Estoque_Declaracoes.CB_TG, Tela_Cfg_Estoque_Declaracoes.Name)
                        
    DLL_FUNCS.RegistraEvento "Abrir Configurações das Declarações", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    AtivaTelaEmEdicao (False)
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_Estoque_Declaracoes
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_CF_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_CF.ListIndex = -1 Then
        TXT_DecSP.Text = ""
        TXT_DecOutros.Text = ""
        Exit Sub
    End If
    DLL_BD.BDSIS_TBEDC.Seek "=", LT_CF.Text
    If DLL_BD.BDSIS_TBEDC.NoMatch Then
        MsgBox ("Não existe declaração para esta classificação fiscal.")
        TXT_DecSP.Text = ""
        TXT_DecOutros.Text = ""
        Exit Sub
    End If
    TXT_DecSP.Text = DLL_BD.BDSIS_TBEDC_CPDSP.Value
    TXT_DecOutros.Text = DLL_BD.BDSIS_TBEDC_CPDOU.Value
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_CF_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Editar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DecOutros_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_DecOutros.SelLength = Len(TXT_DecOutros.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DecOutros_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Salvar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DecSP_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_DecSP.SelLength = Len(TXT_DecSP.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DecSP_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_DecOutros.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub AtivaTelaEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Valor = True Then
        FR_2.Enabled = False
        FR_2_2.Enabled = False
        CB_TG.Enabled = False
        LB_CF.Enabled = False
        LT_CF.Enabled = False
        FR_3.Enabled = True
        LB_DecSP.Enabled = True
        TXT_DecSP.Enabled = True
        LB_DecOutros.Enabled = True
        TXT_DecOutros.Enabled = True
        BT_Editar.Enabled = False
        BT_Salvar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
    Else
        FR_2.Enabled = True
        FR_2_2.Enabled = True
        CB_TG.Enabled = True
        LB_CF.Enabled = True
        LT_CF.Enabled = True
        FR_3.Enabled = False
        LB_DecSP.Enabled = False
        TXT_DecSP.Enabled = False
        LB_DecOutros.Enabled = False
        TXT_DecOutros.Enabled = False
        BT_Editar.Enabled = True
        BT_Salvar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Voltar.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_Cfg_Estoque_Declaracoes.MousePointer = vbHourglass
        Tela_Cfg_Estoque_Declaracoes.Enabled = False
    Else
        Tela_Cfg_Estoque_Declaracoes.MousePointer = vbDefault
        Tela_Cfg_Estoque_Declaracoes.Enabled = True
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
Private Static Function CarregaCB_TG(Combo As ComboBox)
    On Error GoTo ERRO_SISCOVAL
    Set DLL_BD.BDSIS_TBGRU = DLL_BD.BDSIS.OpenRecordset("Grupos")
    Set DLL_BD.BDSIS_TBGRU_CPGRU = DLL_BD.BDSIS_TBGRU.Fields("Grupo")
    Set DLL_BD.BDSIS_TBGRU_CPDES = DLL_BD.BDSIS_TBGRU.Fields("Descrição")
    Set DLL_BD.BDSIS_TBGRU_CPVAL = DLL_BD.BDSIS_TBGRU.Fields("Valor")
    Set DLL_BD.BDSIS_TBGRU_CPTIP = DLL_BD.BDSIS_TBGRU.Fields("Tipo")
    DLL_BD.BDSIS_TBGRU.Index = "Grupo"
    'Carrega CB_TG
    DLL_BD.BDSIS_TBGRU.MoveFirst
    Combo.AddItem (DLL_BD.BDSIS_TBGRU_CPTIP.Value)
    Do While Not DLL_BD.BDSIS_TBGRU.EOF
        For I = 0 To Combo.ListCount - 1
            If Combo.List(I) <> DLL_BD.BDSIS_TBGRU_CPTIP.Value And _
               I = Combo.ListCount - 1 Then
                Combo.AddItem (DLL_BD.BDSIS_TBGRU_CPTIP.Value)
            End If
        Next I
        DLL_BD.BDSIS_TBGRU.MoveNext
    Loop
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Function LeCB_TG(Combo As ComboBox, Tela As String)
    On Error GoTo ERRO_SISCOVAL
    Set DLL_BD.BDSIS_TBCTG = DLL_BD.BDSIS.OpenRecordset("Configurações Tela-Grupo")
    Set DLL_BD.BDSIS_TBCTG_CPTEL = DLL_BD.BDSIS_TBCTG.Fields("Tela")
    Set DLL_BD.BDSIS_TBCTG_CPCOM = DLL_BD.BDSIS_TBCTG.Fields("Combo")
    Set DLL_BD.BDSIS_TBCTG_CPVAL = DLL_BD.BDSIS_TBCTG.Fields("Valor")
    DLL_BD.BDSIS_TBCTG.Index = "Tela"
    DLL_BD.BDSIS_TBCTG.Seek "=", Tela, Combo.Name
    If DLL_BD.BDSIS_TBCTG.NoMatch Then
        Combo.ListIndex = 0
    Else
        For I = 0 To Combo.ListCount - 1
            If Combo.List(I) = DLL_BD.BDSIS_TBCTG_CPVAL.Value Then
                Combo.ListIndex = I
                Exit For
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Function SalvaCB_TG(Valor As String, Tela As String, Combo As String)
    On Error GoTo ERRO_SISCOVAL
    Set DLL_BD.BDSIS_TBCTG = DLL_BD.BDSIS.OpenRecordset("Configurações Tela-Grupo")
    Set DLL_BD.BDSIS_TBCTG_CPTEL = DLL_BD.BDSIS_TBCTG.Fields("Tela")
    Set DLL_BD.BDSIS_TBCTG_CPCOM = DLL_BD.BDSIS_TBCTG.Fields("Combo")
    Set DLL_BD.BDSIS_TBCTG_CPVAL = DLL_BD.BDSIS_TBCTG.Fields("Valor")
    DLL_BD.BDSIS_TBCTG.Index = "Tela"
    DLL_BD.BDSIS_TBCTG.Seek "=", Tela, Combo
    If DLL_BD.BDSIS_TBCTG.NoMatch Then
        DLL_BD.BDSIS_TBCTG.AddNew
        DLL_BD.BDSIS_TBCTG_CPTEL.Value = Tela
        DLL_BD.BDSIS_TBCTG_CPCOM.Value = Combo
    Else
        DLL_BD.BDSIS_TBCTG.Edit
    End If
    DLL_BD.BDSIS_TBCTG_CPVAL.Value = Valor
    DLL_BD.BDSIS_TBCTG.Update
    DLL_BD.BDSIS_TBCTG.Close
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
