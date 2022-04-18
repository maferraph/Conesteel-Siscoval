VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Cfg_Estoque_Aliquotas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações de Alíquotas"
   ClientHeight    =   3945
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5670
   ControlBox      =   0   'False
   Icon            =   "Tela_Cfg_Estoque_Aliquotas.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar BP_1 
      Height          =   240
      Left            =   3840
      TabIndex        =   14
      Top             =   3708
      Width           =   1812
      _ExtentX        =   3201
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar BS_1 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   13
      Top             =   3696
      Width           =   5664
      _ExtentX        =   10001
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid FG_1 
      Height          =   1812
      Left            =   2040
      TabIndex        =   3
      ToolTipText     =   $"Tela_Cfg_Estoque_Aliquotas.frx":030A
      Top             =   960
      Width           =   3492
      _ExtentX        =   6165
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   4
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      ScrollBars      =   2
   End
   Begin VB.Frame FR_1 
      Caption         =   "Tipo do Grupo"
      Height          =   612
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   1692
      Begin VB.ComboBox CB_TG 
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Só altere esta lista se você tem certeza que os dados exibidos na lista abaixo não são os esperados (neste caso, Estados)."
         Top             =   240
         Width           =   1452
      End
   End
   Begin VB.ListBox LT_Estados 
      Height          =   1620
      ItemData        =   "Tela_Cfg_Estoque_Aliquotas.frx":0393
      Left            =   120
      List            =   "Tela_Cfg_Estoque_Aliquotas.frx":039A
      Sorted          =   -1  'True
      TabIndex        =   2
      ToolTipText     =   "Lista de estados."
      Top             =   960
      Width           =   1692
   End
   Begin VB.Frame FR_2 
      Caption         =   "Selecione qual alíquota deseja configurar:"
      Height          =   612
      Left            =   2040
      TabIndex        =   8
      Top             =   0
      Width           =   3492
      Begin VB.OptionButton RB_ICMS 
         Caption         =   "I.C.M.S."
         Height          =   192
         Left            =   1920
         TabIndex        =   1
         ToolTipText     =   "Configurar alíquotas de I.C.M.S."
         Top             =   240
         Width           =   1092
      End
      Begin VB.OptionButton RB_IPI 
         Caption         =   "I.P.I."
         Height          =   192
         Left            =   480
         TabIndex        =   0
         ToolTipText     =   "Configurar alíquotas de I.P.I."
         Top             =   240
         Width           =   852
      End
   End
   Begin VB.CommandButton BT_Editar 
      Caption         =   "&Editar"
      Height          =   732
      Left            =   120
      Picture         =   "Tela_Cfg_Estoque_Aliquotas.frx":03AA
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Edita."
      Top             =   2880
      Width           =   732
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   732
      Left            =   1800
      Picture         =   "Tela_Cfg_Estoque_Aliquotas.frx":07EC
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salva informações."
      Top             =   2880
      Width           =   732
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   732
      Left            =   3360
      Picture         =   "Tela_Cfg_Estoque_Aliquotas.frx":0C2E
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancela operação."
      Top             =   2880
      Width           =   732
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   4800
      Picture         =   "Tela_Cfg_Estoque_Aliquotas.frx":0F38
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Volta à tela principal."
      Top             =   2880
      Width           =   732
   End
   Begin VB.Label LB_Aliquota 
      AutoSize        =   -1  'True
      Caption         =   "Configurações das alíquotas:"
      Height          =   192
      Left            =   2040
      TabIndex        =   11
      Top             =   720
      Width           =   2088
   End
   Begin VB.Label LB_Lista 
      AutoSize        =   -1  'True
      Caption         =   "Lista de Estados:"
      Height          =   192
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   1236
   End
End
Attribute VB_Name = "Tela_Cfg_Estoque_Aliquotas"
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
Const NOMEAPLIC As String = "Configurações de Alíquotas"
Dim I, J As Integer
Dim RespMsg, Resp
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Estados.ListIndex = -1 Then
        MsgBox ("Você deve primeiro selecionar um estado na lista...")
        LT_Estados.SetFocus
        Exit Sub
    End If
    AtivaTelaEmEdicao (True)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    ResetaBP (FG_1.Rows)
    If RB_IPI.Value = True Then
        For I = 1 To (FG_1.Rows - 1)
            DLL_BD.BDSIS_TBEAL.Seek "=", "IPI", LT_Estados.Text, FG_1.TextMatrix(I, 0)
            If Not DLL_BD.BDSIS_TBEAL.NoMatch Then
                DLL_BD.BDSIS_TBEAL.Edit
            Else
                DLL_BD.BDSIS_TBEAL.AddNew
                DLL_BD.BDSIS_TBEAL_CPALI.Value = "IPI"
                DLL_BD.BDSIS_TBEAL_CPEST.Value = LT_Estados.Text
                DLL_BD.BDSIS_TBEAL_CPCCF.Value = FG_1.TextMatrix(I, 0)
            End If
            DLL_BD.BDSIS_TBEAL_CPPOR.Value = FG_1.TextMatrix(I, 1)
            DLL_BD.BDSIS_TBEAL_CPRBC.Value = 0
            DLL_BD.BDSIS_TBEAL.Update
            CarregaBSEP ("Salvando alíquotas...")
        Next I
    ElseIf RB_ICMS.Value = True Then
        For I = 1 To (FG_1.Rows - 1)
            DLL_BD.BDSIS_TBEAL.Seek "=", "ICMS", LT_Estados.Text, FG_1.TextMatrix(I, 0)
            If Not DLL_BD.BDSIS_TBEAL.NoMatch Then
                DLL_BD.BDSIS_TBEAL.Edit
            Else
                DLL_BD.BDSIS_TBEAL.AddNew
                DLL_BD.BDSIS_TBEAL_CPALI.Value = "ICMS"
                DLL_BD.BDSIS_TBEAL_CPEST.Value = LT_Estados.Text
                DLL_BD.BDSIS_TBEAL_CPCCF.Value = FG_1.TextMatrix(I, 0)
            End If
            DLL_BD.BDSIS_TBEAL_CPPOR.Value = FG_1.TextMatrix(I, 1)
            DLL_BD.BDSIS_TBEAL_CPRBC.Value = FG_1.TextMatrix(I, 2)
            DLL_BD.BDSIS_TBEAL.Update
            CarregaBSEP ("Salvando alíquotas...")
        Next I
    End If
    AtivaTelaEmEdicao (False)
    DLL_FUNCS.RegistraEvento "Salvar - Alíquotas de Estoque", ""
    ResetaBSEP
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_Estoque_Aliquotas
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_TG_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    LT_Estados.Clear
    DLL_BD.BDSIS_TBGRU.MoveFirst
    Do While Not DLL_BD.BDSIS_TBGRU.EOF
        If DLL_BD.BDSIS_TBGRU_CPTIP.Value = CB_TG.Text Then
            LT_Estados.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
        End If
        DLL_BD.BDSIS_TBGRU.MoveNext
    Loop
    RespMsg = SalvaCB_TG(CB_TG.Text, Tela_Cfg_Estoque_Aliquotas.Name, CB_TG.Name)
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_TG_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then RB_IPI.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub FG_1_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then
        BT_Salvar.SetFocus
    ElseIf KeyAscii = Asc("0") Or _
           KeyAscii = Asc("1") Or _
           KeyAscii = Asc("2") Or _
           KeyAscii = Asc("3") Or _
           KeyAscii = Asc("4") Or _
           KeyAscii = Asc("5") Or _
           KeyAscii = Asc("6") Or _
           KeyAscii = Asc("7") Or _
           KeyAscii = Asc("8") Or _
           KeyAscii = Asc("9") Or _
           KeyAscii = Asc(",") Then
        FG_1.TextMatrix(FG_1.Row, FG_1.Col) = FG_1.TextMatrix(FG_1.Row, FG_1.Col) & Chr(KeyAscii)
    ElseIf KeyAscii = vbKeyBack Then
        If Len(FG_1.TextMatrix(FG_1.Row, FG_1.Col)) <= 1 Then
            FG_1.TextMatrix(FG_1.Row, FG_1.Col) = ""
        Else
            FG_1.TextMatrix(FG_1.Row, FG_1.Col) = Mid(FG_1.TextMatrix(FG_1.Row, FG_1.Col), 1, Len(FG_1.TextMatrix(FG_1.Row, FG_1.Col)) - 1)
        End If
    ElseIf KeyAscii = vbKeyUp Then
        If FG_1.Row > 1 Then FG_1.Row = FG_1.Row - 1
    ElseIf KeyAscii = vbKeyDown Then
        If FG_1.Row < FG_1.Rows Then FG_1.Row = FG_1.Row + 1
    ElseIf KeyAscii = vbKeyLeft Then
        If FG_1.Col > 1 Then FG_1.Col = FG_1.Col - 1
    ElseIf KeyAscii = vbKeyRight Then
        If FG_1.Col < FG_1.Cols Then FG_1.Col = FG_1.Col + 1
    Else
        KeyAscii = vbKeyEscape
    End If
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
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Alíquotas...")
    If DLL_BD.AbreTabela_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abre campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Alíquotas...")
    If DLL_BD.AbreCampos_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega lista de grupos
    DLL_CARGA.CarregaTexto ("Carregando lista de grupos...")
    Dim cResp
    cResp = CarregaCB_TG(Tela_Cfg_Estoque_Aliquotas.CB_TG)
    cResp = LeCB_TG(Tela_Cfg_Estoque_Aliquotas.CB_TG, Tela_Cfg_Estoque_Aliquotas.Name)
    For I = 0 To CB_TG.ListCount - 1
        If CB_TG.List(I) = "EST" Then
            CB_TG.ListIndex = I
            Exit For
        End If
    Next I
    RB_IPI.Value = True
                    
    DLL_FUNCS.RegistraEvento "Abrir Configurações de Alíquotas", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    AtivaTelaEmEdicao (False)
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_Estoque_Aliquotas
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Estados_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Estados.ListIndex = -1 Then Exit Sub
    TelaEmEspera (True)
    ResetaBP (FG_1.Rows)
    If RB_IPI.Value = True Then
        For I = 1 To (FG_1.Rows - 1)
            DLL_BD.BDSIS_TBEAL.Seek "=", "IPI", LT_Estados.Text, FG_1.TextMatrix(I, 0)
            If Not DLL_BD.BDSIS_TBEAL.NoMatch Then
                FG_1.TextMatrix(I, 1) = DLL_BD.BDSIS_TBEAL_CPPOR.Value
            Else
                FG_1.TextMatrix(I, 1) = "0"
            End If
            CarregaBSEP ("Carregando alíquotas...")
        Next I
    ElseIf RB_ICMS.Value = True Then
        For I = 1 To (FG_1.Rows - 1)
            DLL_BD.BDSIS_TBEAL.Seek "=", "ICMS", LT_Estados.Text, FG_1.TextMatrix(I, 0)
            If Not DLL_BD.BDSIS_TBEAL.NoMatch Then
                FG_1.TextMatrix(I, 1) = DLL_BD.BDSIS_TBEAL_CPPOR.Value
                FG_1.TextMatrix(I, 2) = DLL_BD.BDSIS_TBEAL_CPRBC.Value
            Else
                FG_1.TextMatrix(I, 1) = "0"
                FG_1.TextMatrix(I, 2) = "0"
            End If
            CarregaBSEP ("Carregando alíquotas...")
        Next I
    End If
    ResetaBSEP
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Estados_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Salvar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_ICMS_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    LT_Estados.ListIndex = -1
    FG_1.Clear
    FG_1.FixedCols = 1
    FG_1.FixedRows = 1
    FG_1.Cols = 3
    FG_1.Rows = 1
    FG_1.ColWidth(0) = 400
    FG_1.ColWidth(1) = 1400
    FG_1.ColWidth(2) = 1400
    FG_1.TextArray(0) = "C.F."
    FG_1.TextArray(1) = "Porcentagem"
    FG_1.TextArray(2) = "Red. Base Calc."
    
    'Carrega C.F.
    DLL_BD.BDSIS_TBGRU.MoveFirst
    Do While Not DLL_BD.BDSIS_TBGRU.EOF
        If DLL_BD.BDSIS_TBGRU_CPTIP.Value = "CF" Then
            FG_1.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
        End If
        DLL_BD.BDSIS_TBGRU.MoveNext
    Loop
    For I = 0 To 2
        FG_1.ColAlignment(I) = flexAlignCenterCenter
    Next I
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_ICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then LT_Estados.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_IPI_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    LT_Estados.ListIndex = -1
    FG_1.Clear
    FG_1.FixedCols = 1
    FG_1.FixedRows = 1
    FG_1.Cols = 2
    FG_1.Rows = 1
    FG_1.ColWidth(0) = 400
    FG_1.ColWidth(1) = 1400
    FG_1.TextArray(0) = "C.F."
    FG_1.TextArray(1) = "Porcentagem"
    'Carrega C.F.
    DLL_BD.BDSIS_TBGRU.MoveFirst
    Do While Not DLL_BD.BDSIS_TBGRU.EOF
        If DLL_BD.BDSIS_TBGRU_CPTIP.Value = "CF" Then
            FG_1.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
        End If
        DLL_BD.BDSIS_TBGRU.MoveNext
    Loop
    For I = 0 To 1
        FG_1.ColAlignment(I) = flexAlignCenterCenter
    Next I
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_IPI_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then LT_Estados.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub AtivaTelaEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Valor = True Then
        FR_1.Enabled = False
        CB_TG.Enabled = False
        LB_Lista.Enabled = False
        LT_Estados.Enabled = False
        FR_2.Enabled = False
        RB_IPI.Enabled = False
        RB_ICMS.Enabled = False
        LB_Aliquota.Enabled = True
        FG_1.Enabled = True
        BT_Editar.Enabled = False
        BT_Voltar.Enabled = False
        BT_Salvar.Enabled = True
        BT_Cancelar.Enabled = True
    Else
        FR_1.Enabled = True
        CB_TG.Enabled = True
        LB_Lista.Enabled = True
        LT_Estados.Enabled = True
        FR_2.Enabled = True
        RB_IPI.Enabled = True
        RB_ICMS.Enabled = True
        LB_Aliquota.Enabled = False
        FG_1.Enabled = False
        BT_Editar.Enabled = True
        BT_Voltar.Enabled = True
        BT_Salvar.Enabled = False
        BT_Cancelar.Enabled = False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_Cfg_Estoque_Aliquotas.MousePointer = vbHourglass
        Tela_Cfg_Estoque_Aliquotas.Enabled = False
    Else
        Tela_Cfg_Estoque_Aliquotas.MousePointer = vbDefault
        Tela_Cfg_Estoque_Aliquotas.Enabled = True
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
