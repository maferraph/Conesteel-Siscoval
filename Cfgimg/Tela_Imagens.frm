VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Tela_Imagens 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurações de Imagens"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   2880
      Top             =   3720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FR 
      Caption         =   "Exibir por:"
      Height          =   1095
      Index           =   2
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton RB_Todas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Clique aqui para exibir todas imagens"
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton RB_Imagem 
         Caption         =   "Tipo Imagem"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Clique aqui para exibir as imagens pelo tipo"
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox CB_Exibir 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Selecione o tipo de imagem que deseja exibir"
         Top             =   720
         Width           =   1335
      End
   End
   Begin VB.Frame FR 
      Caption         =   "Imagem:"
      Height          =   2295
      Index           =   1
      Left            =   0
      TabIndex        =   23
      Top             =   1080
      Width           =   1575
      Begin VB.ListBox LT_Indice 
         Height          =   2010
         ItemData        =   "Tela_Imagens.frx":0000
         Left            =   120
         List            =   "Tela_Imagens.frx":0007
         TabIndex        =   4
         ToolTipText     =   "Lista das imagens"
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame FR 
      Height          =   3375
      Index           =   0
      Left            =   1680
      TabIndex        =   18
      Top             =   0
      Width           =   6735
      Begin VB.PictureBox IMG 
         Height          =   3015
         Left            =   3600
         ScaleHeight     =   2955
         ScaleWidth      =   2955
         TabIndex        =   27
         ToolTipText     =   "Imagem"
         Top             =   240
         Width           =   3015
      End
      Begin VB.TextBox TXT_Path 
         BackColor       =   &H8000000F&
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Text            =   "Tela_Imagens.frx":0016
         ToolTipText     =   "Caminho da foto para inserir no sistema"
         Top             =   2760
         Width           =   3375
      End
      Begin VB.CommandButton BT_RemoverFoto 
         Caption         =   "Remover foto"
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Apagar a foto ao lado"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.CommandButton BT_ProcurarFoto 
         Caption         =   "Procurar foto"
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Clique aqui para procurar a foto para inserir no sistema"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox TXT_Indice 
         Enabled         =   0   'False
         Height          =   330
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Índice da imagem"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXT_Nome 
         Height          =   330
         Left            =   120
         MaxLength       =   20
         TabIndex        =   7
         ToolTipText     =   "Nome da imagem"
         Top             =   840
         Width           =   3375
      End
      Begin VB.ComboBox CB_Tipo 
         Height          =   315
         Left            =   1560
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Selecione o tipo da imagem"
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "A foto deverá ser previamente editada com as seguintes características: Tamanho de 300 X 300 DPI, Resolução de 150 DPI"
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   1440
         Width           =   3375
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Imagem:"
         Height          =   195
         Index           =   0
         Left            =   3600
         TabIndex        =   22
         Top             =   0
         Width           =   600
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Índice da Imagem:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   21
         Top             =   0
         Width           =   1305
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Tipo da Imagem:"
         Height          =   195
         Index           =   2
         Left            =   1560
         TabIndex        =   20
         Top             =   0
         Width           =   1185
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Imagem:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   1290
      End
   End
   Begin VB.CommandButton BT_Deletar 
      Caption         =   "&Deletar"
      Height          =   855
      Left            =   1800
      Picture         =   "Tela_Imagens.frx":001C
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Deletar foto"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5160
      Picture         =   "Tela_Imagens.frx":045E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancela edição"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   7440
      Picture         =   "Tela_Imagens.frx":0768
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   3480
      Width           =   855
   End
   Begin MSComctlLib.ProgressBar BP 
      Height          =   255
      Left            =   6000
      TabIndex        =   16
      Top             =   4440
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar BS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   4425
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   855
      Left            =   4320
      Picture         =   "Tela_Imagens.frx":0BAA
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salvar dados"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton BT_Editar 
      Caption         =   "&Editar"
      Height          =   855
      Left            =   960
      Picture         =   "Tela_Imagens.frx":0FEC
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Editar foto"
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton BT_Novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "Tela_Imagens.frx":142E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nova foto"
      Top             =   3480
      Width           =   855
   End
   Begin RichTextLib.RichTextBox RTB 
      Height          =   495
      Left            =   3480
      TabIndex        =   26
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Tela_Imagens.frx":1738
   End
End
Attribute VB_Name = "Tela_Imagens"
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
Const NOMEAPLIC As String = "Configurações de Imagens"
Dim I As Integer, J As Integer, ModoEdicao As Boolean
Dim RespMsg

Private Sub BT_Cancelar_Click()
    LimpaCampos
    TelaEmEdicao False
End Sub
Private Sub BT_Deletar_Click()
    If LT_Indice.ListIndex = -1 Then
        MsgBox ("Selecione primeiro um índice na lista acima.")
        LT_Indice.SetFocus
        Exit Sub
    End If
    RespMsg = MsgBox("A exclusão de uma imagem pode afetar outras tabelas do banco de dados; só faça isso se você tem certeza absoluta de proceder esta operação. Você deseja realmente deletar uma imagem ?", vbInformation + vbYesNo + vbDefaultButton2, "Deletar imagem")
    If RespMsg = vbYes Then
        TelaEmEspera True
        With DLL_BD
            BS.SimpleText = "Deletando ítem..."
            .BDSIS_TBIMA.Delete
            LT_Indice.RemoveItem LT_Indice.ListIndex
            BS.SimpleText = ""
            LimpaCampos
        End With
        TelaEmEspera False
    End If
End Sub
Private Sub BT_Editar_Click()
    If LT_Indice.ListIndex = -1 Then
        MsgBox ("Selecione primeiro um índice na lista acima.")
        LT_Indice.SetFocus
        Exit Sub
    End If
    TelaEmEspera True
    TelaEmEdicao True
    ModoEdicao = True
    TelaEmEspera False
    CB_Tipo.SetFocus
End Sub
Private Sub BT_Novo_Click()
    TelaEmEspera True
    TelaEmEdicao True
    ModoEdicao = False
    LimpaCampos
    TelaEmEspera False
    CB_Tipo.SetFocus
End Sub
Private Sub BT_ProcurarFoto_Click()
    On Error GoTo ERRO_SISCOVAL
    CD.DialogTitle = "Indique o caminho da imagem"
    CD.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames
    CD.Filter = "Imagens (*.jpg;*.gif)|*.jpg;*.gif"
    CD.ShowOpen
    TelaEmEspera True
    If CD.FileName <> "" Then
        RTB.LoadFile CD.FileName, rtfText
        TXT_Path.Text = CD.FileName
        IMG.Picture = LoadPicture(TXT_Path.Text)
    End If
    TelaEmEspera False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_ProcurarFoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BT_Salvar.SetFocus
End Sub
Private Sub BT_RemoverFoto_Click()
    TXT_Path.Text = ""
    RTB.Text = ""
    IMG.Picture = LoadPicture()
End Sub
Private Sub BT_RemoverFoto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BT_ProcurarFoto.SetFocus
End Sub
Private Sub BT_Salvar_Click()
    If CB_Tipo.ListIndex = -1 Then
        MsgBox "É necessário digitar todos os dados.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Tipo.SetFocus
        Exit Sub
    ElseIf TXT_Nome.Text = "" Then
        MsgBox "É necessário digitar todos os dados.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_Nome.SetFocus
        Exit Sub
    ElseIf RTB.Text = "" Or TXT_Path.Text = "" Then
        MsgBox "É necessário digitar todos os dados.", vbInformation + vbOKOnly, NOMEAPLIC
        BT_ProcurarFoto.SetFocus
        Exit Sub
    End If
    TelaEmEspera True
    With DLL_BD
        BS.SimpleText = "Aguarde... salvando informações."
        If ModoEdicao = False Then 'Novo
            'confirma se não existe o mesmo nome
            .BDSIS_TBIMA.Seek "=", TXT_Nome.Text
            If Not .BDSIS_TBIMA.NoMatch Then
                MsgBox "Este nome digitado já existe no banco de dados - para incluir uma nova imagem, digite outro nome.", vbExclamation + vbOKOnly, NOMEAPLIC
                TelaEmEspera False
                TXT_Nome.SetFocus
                Exit Sub
            End If
            .BDSIS_TBIMA.AddNew
        ElseIf ModoEdicao = True Then 'edicao
            .BDSIS_TBIMA.Edit
        End If
        If ModoEdicao = False Then 'Novo
            TXT_Indice.Text = .BDSIS_TBIMA_CPIND.Value
            LT_Indice.AddItem TXT_Nome.Text
        End If
        .BDSIS_TBIMA_CPNOM.Value = TXT_Nome.Text
        .BDSIS_TBIMA_CPTIP.Value = PegaGrupo(CB_Tipo.Text, "IMG")
        .BDSIS_TBIMA_CPIMA.Value = RTB.Text
        .BDSIS_TBIMA.Update
        BS.SimpleText = ""
    End With
    BT_Cancelar_Click
    TelaEmEspera False
End Sub
Private Sub BT_Voltar_Click()
    Unload Me
End Sub
Private Sub CB_Exibir_Click()
    TelaEmEspera True
    LT_Indice.Clear
    With DLL_BD
        If .BDSIS_TBIMA.RecordCount > 0 Then
            ResetaBSEP
            ResetaBP .BDSIS_TBIMA.RecordCount + 10
            .BDSIS_TBIMA.MoveFirst
            Do While Not .BDSIS_TBIMA.EOF
                CarregaBSEP "Aguarde... carregando lista de índices das imagens."
                If .BDSIS_TBIMA_CPTIP.Value = CB_Exibir.Text Then LT_Indice.AddItem .BDSIS_TBIMA_CPIND.Value
                .BDSIS_TBIMA.MoveNext
            Loop
            ResetaBSEP
        End If
    End With
    TelaEmEspera False
End Sub
Private Sub CB_Exibir_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_Indice.SetFocus
End Sub
Private Sub CB_Tipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_Nome.SetFocus
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
    DLL_CARGA.CarregaTexto ("Abrindo tabela Configuração de Imagens...")
    If DLL_BD.AbreTabela_Imagens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Grupos...")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abre campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Configuração de Imagens...")
    If DLL_BD.AbreCampos_Imagens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega lista de grupos
    DLL_CARGA.CarregaTexto ("Carregando lista de grupos...")
    With DLL_BD
        CB_Exibir.Clear
        CB_Tipo.Clear
        .BDSIS_TBGRU.MoveFirst
        Do While Not .BDSIS_TBGRU.EOF
            If .BDSIS_TBGRU_CPTIP.Value = "IMG" Then
                CB_Exibir.AddItem .BDSIS_TBGRU_CPVAL.Value
                CB_Tipo.AddItem .BDSIS_TBGRU_CPVAL.Value
            End If
            .BDSIS_TBGRU.MoveNext
        Loop
    End With
    
    DLL_CARGA.CarregaTexto ("Finalizando...")
    BT_Cancelar_Click
    DLL_FUNCS.RegistraEvento "Abrir Configurações de Imagens", ""
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Me
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Imagens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Indice_Click()
    CB_Tipo.ListIndex = -1
    TXT_Nome.Text = ""
    TXT_Path.Text = ""
    TXT_Indice.Text = ""
    RTB.Text = ""
    IMG.Picture = LoadPicture()
    If LT_Indice.ListIndex = -1 Then Exit Sub
    TelaEmEspera True
    With DLL_BD
        .BDSIS_TBIMA.Seek "=", LT_Indice.Text
        If Not .BDSIS_TBIMA.NoMatch Then
            TXT_Indice.Text = .BDSIS_TBIMA_CPIND.Value
            TXT_Nome.Text = .BDSIS_TBIMA_CPNOM.Value
            CB_Tipo.Text = PegaValorGrupo(.BDSIS_TBIMA_CPTIP.Value)
            RTB.Text = .BDSIS_TBIMA_CPIMA.Value
            'le imagem
            TXT_Path.Text = Trim(DLL_FUNCS.DiretorioTemporario) & "scvimg.jpg"
            RTB.SaveFile TXT_Path.Text, rtfText
            IMG.Picture = LoadPicture(TXT_Path.Text)
        End If
    End With
    TelaEmEspera False
End Sub
Private Sub LT_Indice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BT_Voltar.SetFocus
End Sub
Private Sub RB_Imagem_Click()
    CB_Exibir.Enabled = True
    LT_Indice.Clear
    CB_Exibir.SetFocus
End Sub
Private Sub RB_Imagem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_Exibir.SetFocus
End Sub
Private Sub RB_Todas_Click()
    TelaEmEspera True
    CB_Exibir.Enabled = False
    LT_Indice.Clear
    With DLL_BD
        If .BDSIS_TBIMA.RecordCount > 0 Then
            ResetaBSEP
            ResetaBP .BDSIS_TBIMA.RecordCount + 10
            .BDSIS_TBIMA.MoveFirst
            Do While Not .BDSIS_TBIMA.EOF
                CarregaBSEP "Aguarde... carregando lista de imagens."
                LT_Indice.AddItem .BDSIS_TBIMA_CPNOM.Value
                .BDSIS_TBIMA.MoveNext
            Loop
            ResetaBSEP
        End If
    End With
    TelaEmEspera False
End Sub
Private Sub RB_Todas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_Indice.SetFocus
End Sub
Private Sub TXT_Nome_GotFocus()
    TXT_Nome.SelLength = Len(TXT_Nome.Text)
End Sub
Private Sub TXT_Nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BT_ProcurarFoto.SetFocus
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Me.MousePointer = vbHourglass
        Me.Enabled = False
    Else
        Me.MousePointer = vbDefault
        Me.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub CarregaBSEP(Texto As String)
    On Error GoTo ERRO_SISCOVAL
    BS.SimpleText = Texto
    BP.Value = BP.Value + 1
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ResetaBP(Max As Integer)
    On Error GoTo ERRO_SISCOVAL
    BP.Max = Max
    BP.Value = 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ResetaBSEP()
    On Error GoTo ERRO_SISCOVAL
    BP.Value = 0
    BS.SimpleText = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub LimpaCampos()
    RB_Todas.Value = False
    RB_Imagem.Value = False
    LT_Indice.Clear
    CB_Tipo.ListIndex = -1
    CB_Exibir.ListIndex = -1
    TXT_Nome.Text = ""
    TXT_Path.Text = ""
    TXT_Indice.Text = ""
    RTB.Text = ""
    IMG.Picture = LoadPicture()
End Sub
Private Static Sub TelaEmEdicao(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        BT_Novo.Enabled = False
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        BT_Salvar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
        BT_ProcurarFoto.Enabled = True
        BT_RemoverFoto.Enabled = True
        RB_Todas.Enabled = False
        RB_Imagem.Enabled = False
        LT_Indice.Enabled = False
        CB_Exibir.Enabled = False
        CB_Tipo.Enabled = True
        TXT_Nome.Enabled = True
    Else
        BT_Novo.Enabled = True
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
        BT_Salvar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Voltar.Enabled = True
        BT_ProcurarFoto.Enabled = False
        BT_RemoverFoto.Enabled = False
        RB_Todas.Enabled = True
        RB_Imagem.Enabled = True
        LT_Indice.Enabled = True
        CB_Exibir.Enabled = True
        CB_Tipo.Enabled = False
        TXT_Nome.Enabled = False
    End If
    TXT_Indice.Enabled = False
    RTB.Visible = False
    TXT_Path.Enabled = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function PegaGrupo(Valor As String, Tipo As String) As String
    PegaGrupo = ""
    With DLL_BD
        If .BDSIS_TBGRU.RecordCount < 1 Then Exit Function
        .BDSIS_TBGRU.MoveFirst
        Do While Not .BDSIS_TBGRU.EOF
            If .BDSIS_TBGRU_CPTIP.Value = Tipo Then
                If .BDSIS_TBGRU_CPVAL.Value = Valor Then
                    PegaGrupo = .BDSIS_TBGRU_CPGRU.Value
                    Exit Function
                End If
            End If
            .BDSIS_TBGRU.MoveNext
        Loop
    End With
End Function
Private Static Function PegaValorGrupo(Grupo As String) As String
    PegaValorGrupo = ""
    With DLL_BD
        If .BDSIS_TBGRU.RecordCount < 1 Then Exit Function
        .BDSIS_TBGRU.Seek "=", Grupo
        If Not .BDSIS_TBGRU.NoMatch Then
            PegaValorGrupo = .BDSIS_TBGRU_CPVAL.Value
            Exit Function
        End If
    End With
End Function

