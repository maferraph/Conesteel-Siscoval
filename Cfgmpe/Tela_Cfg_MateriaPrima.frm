VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Cfg_MateriaPrima 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações de Matéria-Prima"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7935
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "MP"
      Height          =   495
      Left            =   6240
      TabIndex        =   39
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   5160
      Picture         =   "Tela_Cfg_MateriaPrima.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancela operação"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "Apa&gar"
      Height          =   855
      Left            =   4320
      Picture         =   "Tela_Cfg_MateriaPrima.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Apaga campos"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   855
      Left            =   3480
      Picture         =   "Tela_Cfg_MateriaPrima.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprime configuração"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "Salvar"
      Height          =   855
      Left            =   2640
      Picture         =   "Tela_Cfg_MateriaPrima.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salva dados"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton BT_ConfigMP 
      Caption         =   "Config"
      Height          =   855
      Left            =   1800
      Picture         =   "Tela_Cfg_MateriaPrima.frx":0E98
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Configuração de materiais das peças"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton BT_Consulta 
      Caption         =   "Consulta"
      Height          =   855
      Left            =   960
      Picture         =   "Tela_Cfg_MateriaPrima.frx":1762
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Nova configuração de matéria-prima"
      Top             =   4200
      Width           =   855
   End
   Begin VB.Frame FR_MAT 
      Caption         =   "Configurações de Matéria-Prima:"
      Height          =   4095
      Left            =   6360
      TabIndex        =   26
      Top             =   5040
      Width           =   7935
      Begin VB.Frame FR_MAT2 
         Height          =   2295
         Left            =   4080
         TabIndex        =   30
         Top             =   840
         Width           =   3495
         Begin VB.ComboBox CB_MATMP 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Para o material da peça acabada ao lado, especificar nesta lista o material da matéria-prima"
            Top             =   1080
            Width           =   3255
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Selecione o material da matéria-prima:"
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   2670
         End
      End
      Begin VB.Frame FR_MAT1 
         Height          =   2295
         Left            =   360
         TabIndex        =   27
         Top             =   840
         Width           =   3495
         Begin VB.ComboBox CB_MP 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   "Selecione o item que é matéria-prima que você deseja configurar"
            Top             =   720
            Width           =   3255
         End
         Begin VB.ComboBox CB_MATPECA 
            Height          =   315
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Material da peça acabada"
            Top             =   1560
            Width           =   3255
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Selecione a matéria-prima:"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   480
            Width           =   1860
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Selecione o material da peça acabada:"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   1320
            Width           =   2775
         End
      End
   End
   Begin TabDlg.SSTab ST 
      Height          =   2775
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   7725
      _ExtentX        =   13626
      _ExtentY        =   4895
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Lista de Figuras"
      TabPicture(0)   =   "Tela_Cfg_MateriaPrima.frx":1A6C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FG_FIG"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LT"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Lista de Matéria-Prima"
      TabPicture(1)   =   "Tela_Cfg_MateriaPrima.frx":1A88
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FG_MP"
      Tab(1).ControlCount=   1
      Begin VB.ListBox LT 
         Height          =   255
         ItemData        =   "Tela_Cfg_MateriaPrima.frx":1AA4
         Left            =   240
         List            =   "Tela_Cfg_MateriaPrima.frx":1AA6
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid FG_FIG 
         Height          =   2295
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Lista de matéria-prima"
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   350
         AllowBigSelection=   0   'False
         SelectionMode   =   1
      End
      Begin MSFlexGridLib.MSFlexGrid FG_MP 
         Height          =   2295
         Left            =   -74880
         TabIndex        =   20
         ToolTipText     =   "Lista de matéria-prima"
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   4048
         _Version        =   393216
         FixedCols       =   0
         RowHeightMin    =   350
         AllowBigSelection=   0   'False
         SelectionMode   =   1
      End
   End
   Begin VB.Frame FR 
      Height          =   975
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Width           =   7695
      Begin VB.CommandButton BT_Alterar 
         Caption         =   "Alterar Item"
         Height          =   315
         Left            =   3840
         TabIndex        =   15
         ToolTipText     =   "Altera ítem da linha selecionada"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BT_Remover 
         Caption         =   "Remover Item"
         Height          =   315
         Left            =   2280
         TabIndex        =   14
         ToolTipText     =   "Remove ítens de linhas selecionadas"
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TXT_Quantidade 
         Height          =   315
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Quantidade de peças"
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CB_Figura 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         ToolTipText     =   "Peça"
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CB_Material 
         Height          =   315
         Left            =   6120
         TabIndex        =   12
         ToolTipText     =   "Material da matéria-prima"
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TXT_Nome 
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         ToolTipText     =   "Nome da matéria-prima"
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox CB_Bitola 
         Height          =   315
         Left            =   4560
         TabIndex        =   11
         ToolTipText     =   "Bitola da matéria-prima"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton BT_Importar 
         Caption         =   "Importar"
         Height          =   315
         Left            =   5400
         TabIndex        =   16
         ToolTipText     =   "Importa configurações de outras peças"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton BT_Adicionar 
         Caption         =   "Adicionar Item"
         Height          =   315
         Left            =   720
         TabIndex        =   13
         ToolTipText     =   "Adiciona ítens de matéria-prima na lista abaixo"
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label LB_M 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   195
         Left            =   6120
         TabIndex        =   37
         Top             =   0
         Width           =   600
      End
      Begin VB.Label LB_B 
         AutoSize        =   -1  'True
         Caption         =   "Bitola:"
         Height          =   195
         Left            =   4560
         TabIndex        =   36
         Top             =   0
         Width           =   435
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Peça:"
         Height          =   195
         Left            =   2640
         TabIndex        =   35
         Top             =   0
         Width           =   1110
      End
      Begin VB.Label LB_P 
         AutoSize        =   -1  'True
         Caption         =   "Figura:"
         Height          =   195
         Left            =   1080
         TabIndex        =   34
         Top             =   0
         Width           =   480
      End
      Begin VB.Label LB_Q 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   0
         Width           =   870
      End
   End
   Begin MSComctlLib.ProgressBar BP 
      Height          =   255
      Left            =   5470
      TabIndex        =   25
      Top             =   5160
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
      TabIndex        =   24
      Top             =   5130
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   6960
      Picture         =   "Tela_Cfg_MateriaPrima.frx":1AA8
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton BT_Novo 
      Caption         =   "&Novo"
      Height          =   855
      Left            =   120
      Picture         =   "Tela_Cfg_MateriaPrima.frx":1EEA
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Nova configuração de matéria-prima"
      Top             =   4200
      Width           =   855
   End
   Begin VB.CheckBox CK_M 
      Caption         =   "Usar configurações de materiais previamente selecionados para configurar a matéria-prima"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1080
      Width           =   7695
   End
End
Attribute VB_Name = "Tela_Cfg_MateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** VARIÁVEIS DLL's ****************
Public DLL_BD As Scvbd.Classe_Scvbd
Public DLL_CARGA As Scvcarr.Classe_Scvcarr
Public DLL_FUNCS As Scvfunc.Classe_Scvfunc

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Configurações de Matéria-Prima"
Dim I As Integer, J As Integer, K As Integer, RespMsg, ESTIND As String
Public MAT As String
Dim MATPRI As MP
Private Type MP
    QUA As String
    PEC As String
    NOM As String
    BIT As String
    MAT As String
End Type

Private Sub BT_Adicionar_Click()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        MsgBox "É necessário especificar a figura para adicioná-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        MsgBox "É necessário especificar a bitola para adicioná-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Material.Text = "" And CK_M.Value = 0 Then
        MsgBox "É necessário especificar o material para adicioná-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Material.SetFocus
        Exit Sub
    End If
    If ST.Tab = 0 Then 'Figuras
        FG_FIG.AddItem (FG_FIG.Rows - 1)
        FG_FIG.TextMatrix(FG_FIG.Rows - 1, 0) = CB_Figura.Text
        FG_FIG.TextMatrix(FG_FIG.Rows - 1, 1) = CB_Bitola.Text
        If CK_M.Value = 0 Then FG_FIG.TextMatrix(FG_FIG.Rows - 1, 2) = CB_Material.Text
    Else 'MP
        If TXT_Quantidade.Text = "" Then
            MsgBox "É necessário especificar a quantidade para adicioná-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT_Quantidade.SetFocus
            Exit Sub
        ElseIf TXT_Nome.Text = "" Then
            MsgBox "É necessário especificar o nome da peça para adicioná-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT_Nome.SetFocus
            Exit Sub
        End If
        FG_MP.AddItem (FG_MP.Rows - 1)
        FG_MP.TextMatrix(FG_MP.Rows - 1, 0) = TXT_Quantidade.Text
        FG_MP.TextMatrix(FG_MP.Rows - 1, 1) = CB_Figura.Text
        FG_MP.TextMatrix(FG_MP.Rows - 1, 2) = TXT_Nome.Text
        FG_MP.TextMatrix(FG_MP.Rows - 1, 3) = CB_Bitola.Text
        If CK_M.Value = 0 Then FG_MP.TextMatrix(FG_MP.Rows - 1, 4) = CB_Material.Text
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Alterar_Click()
    If ST.Tab = 0 Then If FG_FIG.RowSel < 1 Then Exit Sub
    If ST.Tab = 1 Then If FG_MP.RowSel < 1 Then Exit Sub
    If CB_Figura.Text = "" Then
        MsgBox "É necessário especificar a figura para alterá-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        MsgBox "É necessário especificar a bitola para alterá-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Material.Text = "" And CK_M.Value = 0 Then
        MsgBox "É necessário especificar o material para alterá-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Material.SetFocus
        Exit Sub
    End If
    If ST.Tab = 0 Then 'Figuras
        FG_FIG.TextMatrix(FG_FIG.RowSel, 0) = CB_Figura.Text
        FG_FIG.TextMatrix(FG_FIG.RowSel, 1) = CB_Bitola.Text
        If CK_M.Value = 0 Then FG_FIG.TextMatrix(FG_FIG.RowSel, 2) = CB_Material.Text
    Else 'MP
        If TXT_Quantidade.Text = "" Then
            MsgBox "É necessário especificar a quantidade para adicioná-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT_Quantidade.SetFocus
            Exit Sub
        ElseIf TXT_Nome.Text = "" Then
            MsgBox "É necessário especificar o nome da peça para adicioná-la na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT_Nome.SetFocus
            Exit Sub
        End If
        FG_MP.TextMatrix(FG_MP.RowSel, 0) = TXT_Quantidade.Text
        FG_MP.TextMatrix(FG_MP.RowSel, 1) = CB_Figura.Text
        FG_MP.TextMatrix(FG_MP.RowSel, 2) = TXT_Nome.Text
        FG_MP.TextMatrix(FG_MP.RowSel, 3) = CB_Bitola.Text
        If CK_M.Value = 0 Then FG_MP.TextMatrix(FG_MP.RowSel, 4) = CB_Material.Text
    End If
End Sub
Private Sub BT_Apagar_Click()
    If FR_MAT.Visible = True Then 'Config
        CB_MP.ListIndex = -1
        CB_MATPECA.ListIndex = -1
        CB_MATMP.ListIndex = -1
        BT_Salvar.Enabled = False
    Else
        CK_M.Value = 0
        TXT_Quantidade.Text = ""
        CB_Figura.Text = ""
        TXT_Nome.Text = ""
        CarregaFG (0)
        CarregaFG (1)
    End If
    BP.Value = 0
    BS.SimpleText = ""
End Sub
Private Sub BT_Cancelar_Click()
    BT_Apagar_Click
    ModoTela (0)
End Sub
Private Sub BT_ConfigMP_Click()
    ModoTela (3)
    CB_MP.SetFocus
End Sub
Private Sub BT_Consulta_Click()
    If CK_M.Value = 1 Then
        MsgBox "Não é possível fazer a consulta sem o material.", vbInformation + vbOKOnly, "Falta dados"
        CK_M.SetFocus
        Exit Sub
    End If
    If CB_Figura.Text = "" Then
        MsgBox "É necessário especificar a figura para continuar a consulta.", vbInformation + vbOKOnly, "Falta dados"
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        MsgBox "É necessário especificar a bitola para continuar a consulta.", vbInformation + vbOKOnly, "Falta dados"
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Material.Text = "" Then
        MsgBox "É necessário especificar o material para continuar a consulta.", vbInformation + vbOKOnly, "Falta dados"
        CB_Material.SetFocus
        Exit Sub
    End If
    BS.SimpleText = "Aguarde... procurando informações sobre este ítem."
    TelaEmEspera True
    ModoTela (2)
    'procura configuração desta peça
    DLL_BD.BDSIS_TBEST.Seek "=", CB_Figura.Text, CB_Bitola.Text, CB_Material.Text
    If Not DLL_BD.BDSIS_TBEST.NoMatch Then
        If IsEmpty(DLL_BD.BDSIS_TBEST_CPINQ.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINP.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINN.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINB.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINM.Value) Then
            MsgBox "Um ou mais dados sobre a matéria-prima deste ítem podem estar faltando, configure-a antes de consultar.", vbInformation + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        CarregaFG (1)
        ProcuraMP DLL_BD.BDSIS_TBEST_CPINQ.Value, DLL_BD.BDSIS_TBEST_CPINP.Value, DLL_BD.BDSIS_TBEST_CPINN.Value, DLL_BD.BDSIS_TBEST_CPINB.Value, DLL_BD.BDSIS_TBEST_CPINM.Value
        DivideMP "QUA", MATPRI.QUA, 1
        DivideMP "PEC", MATPRI.PEC, 1
        DivideMP "NOM", MATPRI.NOM, 1
        DivideMP "BIT", MATPRI.BIT, 1
        DivideMP "MAT", MATPRI.MAT, 1
        CarregaFG (0)
        FG_FIG.AddItem (1)
        FG_FIG.TextMatrix(1, 0) = CB_Figura.Text
        FG_FIG.TextMatrix(1, 1) = CB_Bitola.Text
        FG_FIG.TextMatrix(1, 2) = CB_Material.Text
    Else
        MsgBox "Não foi possível localizar a ficha de estoque - verifique os dados digitados.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    BS.SimpleText = ""
    TelaEmEspera False
End Sub
Private Sub BT_Importar_Click()
    If CB_Figura.Text = "" Then
        MsgBox "É necessário especificar a figura para importar dados.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        MsgBox "É necessário especificar a bitola para importar dados.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Material.Text = "" And CK_M.Value = 0 Then
        MsgBox "É necessário especificar o material para importar dados.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Material.SetFocus
        Exit Sub
    End If
    BS.SimpleText = "Aguarde... procurando informações sobre este ítem."
    TelaEmEspera True
    'procura configuração desta peça
    If CK_M.Value = 0 Then
        DLL_BD.BDSIS_TBEST.Seek "=", CB_Figura.Text, CB_Bitola.Text, CB_Material.Text
    Else
        DLL_BD.BDSIS_TBEST.Seek "=", CB_Figura.Text, CB_Bitola.Text, "A-105"
    End If
    If Not DLL_BD.BDSIS_TBEST.NoMatch Then
        If IsEmpty(DLL_BD.BDSIS_TBEST_CPINQ.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINP.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINN.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINB.Value) = True Or IsEmpty(DLL_BD.BDSIS_TBEST_CPINM.Value) Then
            MsgBox "Um ou mais dados sobre a matéria-prima deste ítem podem estar faltando, configure-a antes de consultar.", vbInformation + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        ProcuraMP DLL_BD.BDSIS_TBEST_CPINQ.Value, DLL_BD.BDSIS_TBEST_CPINP.Value, DLL_BD.BDSIS_TBEST_CPINN.Value, DLL_BD.BDSIS_TBEST_CPINB.Value, DLL_BD.BDSIS_TBEST_CPINM.Value
        Dim nA As Integer
        RespMsg = MsgBox("As informações sobre matéria-prima do ítem importado foram encontradas. Você tem certeza que deseja importá-las ?", vbQuestion + vbYesNo + vbDefaultButton1, "Importar configurações")
        If RespMsg <> vbYes Then GoTo SAIDA
        nA = FG_MP.Rows
        DivideMP "QUA", MATPRI.QUA, nA
        DivideMP "PEC", MATPRI.PEC, nA
        DivideMP "NOM", MATPRI.NOM, nA
        DivideMP "BIT", MATPRI.BIT, nA
        If CK_M.Value = 0 Then DivideMP "MAT", MATPRI.MAT, 1
    Else
        MsgBox "Não foi possível localizar a ficha de estoque - verifique os dados digitados.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
SAIDA:
    BS.SimpleText = ""
    TelaEmEspera False
End Sub
Private Sub BT_Imprimir_Click()
    If FG_FIG.Rows < 1 Then
        MsgBox "Não existe nenhuma peça incluída na lista de figuras para imprimir o relatório de matéria-prima.", vbInformation + vbOKOnly, NOMEAPLIC
        ST.Tab = 0
        Exit Sub
    ElseIf FG_MP.Rows < 1 Then
        MsgBox "Não existe nenhuma peça incluída na lista de matéria-prima para imprimir o relatório de matéria-prima.", vbInformation + vbOKOnly, NOMEAPLIC
        ST.Tab = 1
        Exit Sub
    End If
    RespMsg = MsgBox("Você tem certeza que deseja imprimir o relatório de matéria-prima das peças configuradas nas listas de figura e matéria-prima ?", vbQuestion + vbYesNo + vbDefaultButton1, "Imprimir relatório")
    If RespMsg = vbYes Then
        TelaEmEspera True
        DLL_FUNCS.SelecionaImpressora (DLL_FUNCS.NomeImpressora("Tela_Cfg_MateriaPrima_Relatorio"))
        'limpa campos
        For I = 0 To 9
            Tela_Cfg_MateriaPrima_Relatorio.LB_Figura(I).Caption = ""
            Tela_Cfg_MateriaPrima_Relatorio.LB_Bitola(I).Caption = ""
            Tela_Cfg_MateriaPrima_Relatorio.LB_Material(I).Caption = ""
            Tela_Cfg_MateriaPrima_Relatorio.LB_Descricao(I).Caption = ""
        Next I
        For I = 0 To 23
            Tela_Cfg_MateriaPrima_Relatorio.LB_QUA(I).Caption = ""
            Tela_Cfg_MateriaPrima_Relatorio.LB_FIG(I).Caption = ""
            Tela_Cfg_MateriaPrima_Relatorio.LB_NOM(I).Caption = ""
            Tela_Cfg_MateriaPrima_Relatorio.LB_BIT(I).Caption = ""
            Tela_Cfg_MateriaPrima_Relatorio.LB_MAT(I).Caption = ""
        Next I
        'carrega lista de figuras
        For I = 1 To FG_FIG.Rows - 1
            Tela_Cfg_MateriaPrima_Relatorio.LB_Figura(I - 1).Caption = FG_FIG.TextMatrix(I, 0)
            Tela_Cfg_MateriaPrima_Relatorio.LB_Bitola(I - 1).Caption = FG_FIG.TextMatrix(I, 1)
            Tela_Cfg_MateriaPrima_Relatorio.LB_Material(I - 1).Caption = FG_FIG.TextMatrix(I, 2)
            'If FG_FIG.Cols = 3 Then
            '    DLL_BD.BDSIS_TBEST.Seek "=", FG_FIG.TextMatrix(I, 0), FG_FIG.TextMatrix(I, 1), FG_FIG.TextMatrix(I, 2)
            '    If Not DLL_BD.BDSIS_TBEST.NoMatch Then
        Next I
        'carrega lista de matéria-prima
        For I = 1 To FG_MP.Rows - 1
            Tela_Cfg_MateriaPrima_Relatorio.LB_QUA(I - 1).Caption = FG_MP.TextMatrix(I, 0)
            Tela_Cfg_MateriaPrima_Relatorio.LB_FIG(I - 1).Caption = FG_MP.TextMatrix(I, 1)
            Tela_Cfg_MateriaPrima_Relatorio.LB_NOM(I - 1).Caption = FG_MP.TextMatrix(I, 2)
            Tela_Cfg_MateriaPrima_Relatorio.LB_BIT(I - 1).Caption = FG_MP.TextMatrix(I, 3)
            If CK_M.Value = 0 Then Tela_Cfg_MateriaPrima_Relatorio.LB_MAT(I - 1).Caption = FG_MP.TextMatrix(I, 4)
        Next I
        Tela_Cfg_MateriaPrima_Relatorio.LB_Relatorio.Caption = "Relatório emitido em " & Format(Date, "dd/mm/yyyy") & " às " & Format(Time, "hh:mm:ss") & "."
        'Imprime relatório
        Tela_Cfg_MateriaPrima_Relatorio.PrintForm
        DLL_FUNCS.SelecionaImpressora (DLL_FUNCS.NomeImpressora("PADRÃO"))
        DLL_FUNCS.RegistraEvento "Imprimir - Relatório Matéria-prima", ""
        TelaEmEspera False
    End If
End Sub
Private Sub BT_Novo_Click()
    ModoTela (1)
    CB_Figura.SetFocus
End Sub
Private Sub BT_Remover_Click()
    If ST.Tab = 0 Then 'Figuras
        If FG_FIG.RowSel > 1 Then
            FG_FIG.RemoveItem (FG_FIG.RowSel)
        ElseIf FG_FIG.RowSel = 1 Then
            If CK_M.Value = 0 Then
                CarregaFG (0)
            Else
                CarregaFG (2)
            End If
        Else
            MsgBox "Não existe linha ou não foi selecionada uma na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        End If
    Else 'MP
        If FG_MP.RowSel > 1 Then
            FG_MP.RemoveItem (FG_MP.RowSel)
        ElseIf FG_MP.RowSel = 1 Then
            If CK_M.Value = 0 Then
                CarregaFG (1)
            Else
                CarregaFG (3)
            End If
        Else
            MsgBox "Não existe linha ou não foi selecionada uma na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
        End If
    End If
End Sub
Private Sub BT_Salvar_Click()
    If FR_MAT.Visible = True Then 'Config
        If CB_MP.Text = "" Then
            MsgBox "É necessário preencher todos os campos.", vbInformation + vbOKOnly, NOMEAPLIC
            CB_MP.SetFocus
            Exit Sub
        ElseIf CB_MATPECA.Text = "" Then
            MsgBox "É necessário preencher todos os campos.", vbInformation + vbOKOnly, NOMEAPLIC
            CB_MATPECA.SetFocus
            Exit Sub
        ElseIf CB_MATMP.Text = "" Then
            MsgBox "É necessário preencher todos os campos.", vbInformation + vbOKOnly, NOMEAPLIC
            CB_MATMP.SetFocus
            Exit Sub
        End If
        BS.SimpleText = "Salvando informações sobre a matéria-prima..."
        DLL_BD.BDSIS_TBMPR.Seek "=", CB_MP.Text, CB_MATPECA.Text
        If DLL_BD.BDSIS_TBMPR.NoMatch Then
            DLL_BD.BDSIS_TBMPR.AddNew
        Else
            DLL_BD.BDSIS_TBMPR.Edit
        End If
        DLL_BD.BDSIS_TBMPR_CPFIG.Value = CB_MP.Text
        DLL_BD.BDSIS_TBMPR_CPMFG.Value = CB_MATPECA.Text
        DLL_BD.BDSIS_TBMPR_CPMMP.Value = CB_MATMP.Text
        DLL_BD.BDSIS_TBMPR.Update
    Else 'salva MP
        If FG_FIG.Rows < 1 Then
            MsgBox "Não existe nenhum ítem na lista de figuras ainda adicionado para salvar configurações das fichas de estoque.", vbInformation + vbOKOnly, "Falta dados"
            ST.Tab = 0
            Exit Sub
        ElseIf FG_MP.Rows < 1 Then
            MsgBox "Não existe nenhum ítem na lista de matéria-prima ainda adicionado para salvar configurações das fichas de estoque.", vbInformation + vbOKOnly, "Falta dados"
            ST.Tab = 1
            Exit Sub
        End If
        
        RespMsg = MsgBox("Esta operação irá verificar todo banco de dados de estoque e se em cada ítem, a figura/bitola/material forem iguais aos dados das listas dos mesmos, serão alteradas as informações sobre a matéria-prima deste ítem de estoque conforme a relação que você escolheu. Sendo esta operação muito demorada e podendo afetar muitos registros, você tem certeza que deseja continuar ?", vbQuestion + vbYesNo + vbDefaultButton1, "Configuração de Matéria-Prima")
        
        If RespMsg = vbYes Then
            TelaEmEspera True
            Dim cQ As String, cP As String, cN As String, cB As String, cM As String
            Dim nQ As Integer, nP As Integer, nN As Integer, nB As Integer, nM As Integer
            'analisa MP
            BS.SimpleText = "Montando lista de matéria-prima..."
            cQ = FG_MP.TextMatrix(1, 0)
            cP = FG_MP.TextMatrix(1, 1)
            cN = FG_MP.TextMatrix(1, 2)
            cB = FG_MP.TextMatrix(1, 3)
            If CK_M.Value = 0 Then cM = FG_MP.TextMatrix(1, 4)
            If (FG_MP.Rows - 1) > 1 Then
                For I = 2 To (FG_MP.Rows - 1)
                    cQ = cQ & ";" & FG_MP.TextMatrix(I, 0)
                    cP = cP & ";" & FG_MP.TextMatrix(I, 1)
                    cN = cN & ";" & FG_MP.TextMatrix(I, 2)
                    cB = cB & ";" & FG_MP.TextMatrix(I, 3)
                    If CK_M.Value = 0 Then cM = cM & ";" & FG_MP.TextMatrix(I, 4)
                Next I
            End If
            'procura se o índice de peças existe
            Dim aMP As Variant
            aMP = ProcuraIndicesMP(cQ, cP, cN, cB, cM)
            
            'verifica banco de dados de estoque e alterar informações sobre matéria-prima
            BS.SimpleText = "Alterando informações sobre matéria-prima do estoque..."
            BP.Value = 0
            If CK_M.Value = 0 Then 'foi selecionado o material
                BP.Max = FG_FIG.Rows
                For I = 1 To FG_FIG.Rows - 1
                    BS.SimpleText = "Procurando fichas de estoque de:" & Trim(FG_FIG.TextMatrix(I, 1))
                    DLL_BD.BDSIS_TBEST.Seek "=", FG_FIG.TextMatrix(I, 0), FG_FIG.TextMatrix(I, 1), FG_FIG.TextMatrix(I, 2)
                    If Not DLL_BD.BDSIS_TBEST.NoMatch Then
                        BS.SimpleText = "Alterando informações sobre matéria-prima da ficha de estoque..."
                        DLL_BD.BDSIS_TBEST.Edit
                        DLL_BD.BDSIS_TBEST_CPINQ.Value = nQ
                        DLL_BD.BDSIS_TBEST_CPINP.Value = nP
                        DLL_BD.BDSIS_TBEST_CPINN.Value = nN
                        DLL_BD.BDSIS_TBEST_CPINB.Value = nB
                        DLL_BD.BDSIS_TBEST_CPINM.Value = nM
                        DLL_BD.BDSIS_TBEST.Update
                    End If
                    BP.Value = BP.Value + 1
                Next I
            Else 'procura por todos materiais de cada figura selecionada
                'procura dados sobre a figura
                'a procura sera feita somente com a primeiro figura da lista de figuras, pois as figuras da lista
                'devem ter o mesmo indice de figura, portanto a mesma relação de materiais;
                'sendo assim, a configuração de materiais da MP será a mesma para todas figuras da lista
                BS.SimpleText = "Procurando informações sobre as figuras..."
                DLL_BD.BDSIS_TBEFG.Seek "=", FG_FIG.TextMatrix(1, 0)
                DLL_BD.BDSIS_TBEID.Seek "=", DLL_BD.BDSIS_TBEFG_CPIFG.Value
                If DLL_BD.BDSIS_TBEFG.NoMatch And DLL_BD.BDSIS_TBEID.NoMatch Then
                    MsgBox "Ocorreu algum erro durante a procura do índice da figura - não é possível prosseguir com a operação.", vbOKOnly + vbInformation, NOMEAPLIC
                    Exit Sub
                End If
                Dim cA As String
                cA = ""
                LT.Clear
                'monta lista de materiais da figura
                BS.SimpleText = "Montando lista de materiais..."
                For J = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
                    If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, J, 1) <> ";" Then
                        cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, J, 1)
                    ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, J, 1) = ";" Then
                        LT.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                        cA = ""
                    End If
                Next J
                LT.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                FG_MP.Cols = 5
                FG_MP.ColAlignment(4) = flexAlignLeftCenter
                FG_MP.ColWidth(4) = 1500
                FG_MP.TextArray(4) = "Material"
                'verifica lista de materiais da figura e depois procura pelo material de cada MP
                BP.Max = LT.ListCount
                For J = 0 To LT.ListCount - 1
                    BS.SimpleText = "Procurando materiais de cada matéria-prima..."
                    For K = 1 To FG_MP.Rows - 1
                        'procura informações de materiais sobre MP
                        DLL_BD.BDSIS_TBMPR.Seek "=", FG_MP.TextMatrix(K, 1), LT.List(J)
                        If Not DLL_BD.BDSIS_TBMPR.NoMatch Then
                            FG_MP.TextMatrix(K, 4) = DLL_BD.BDSIS_TBMPR_CPMMP.Value
                        End If
                    Next K
                    'verifica se todos materiais de MP estão incluídos
                    BS.SimpleText = "Verificando se cada matéria-prima está com o material configurado..."
                    For K = 1 To FG_MP.Rows - 1
                        If FG_MP.TextMatrix(K, 4) = "" Then
                            Tela_Cfg_MateriaPrima_Materiais.LB_F.Caption = LT.List(J)
                            Tela_Cfg_MateriaPrima_Materiais.LB_MP.Caption = FG_MP.TextMatrix(K, 1)
                            Tela_Cfg_MateriaPrima_Materiais.CB_MAT.ListIndex = -1
                            MAT = ""
                            Do While MAT = ""
                                Tela_Cfg_MateriaPrima_Materiais.Show vbModal
                            Loop
                            'salva material da MP
                            DLL_BD.BDSIS_TBMPR.Seek "=", FG_MP.TextMatrix(K, 1), LT.List(J)
                            If DLL_BD.BDSIS_TBMPR.NoMatch Then
                                DLL_BD.BDSIS_TBMPR.AddNew
                            Else
                                DLL_BD.BDSIS_TBMPR.Edit
                            End If
                            DLL_BD.BDSIS_TBMPR_CPFIG.Value = FG_MP.TextMatrix(K, 1)
                            DLL_BD.BDSIS_TBMPR_CPMFG.Value = LT.List(J)
                            DLL_BD.BDSIS_TBMPR_CPMMP.Value = MAT
                            DLL_BD.BDSIS_TBMPR.Update
                            'altera lista da MP
                            FG_MP.TextMatrix(K, 4) = MAT
                        End If
                    Next K
                    BS.SimpleText = "Montando lista de materiais..."
                    'monta lista de materiais da MP selecionados
                    cM = FG_MP.TextMatrix(1, 4)
                    If (FG_MP.Rows - 1) > 1 Then
                        For I = 2 To (FG_MP.Rows - 1)
                            cM = cM & ";" & FG_MP.TextMatrix(I, 4)
                        Next I
                    End If
                    'verifica se a lista já existe cadastrada
                    BS.SimpleText = "Verificando índice de materiais..."
                    DLL_BD.BDSIS_TBMPM.Seek "=", cM
                    If DLL_BD.BDSIS_TBMPM.NoMatch Then
                        DLL_BD.BDSIS_TBMPM.AddNew
                        DLL_BD.BDSIS_TBMPM_CPMAT.Value = cM
                        nM = DLL_BD.BDSIS_TBMPM_CPINM.Value
                        DLL_BD.BDSIS_TBMPM.Update
                    Else
                        nM = DLL_BD.BDSIS_TBMPM_CPINM.Value
                    End If
                    'salva dados do estoque
                    For K = 1 To FG_FIG.Rows - 1
                        BS.SimpleText = "Procurando fichas de estoque selecionadas..."
                        DLL_BD.BDSIS_TBEST.Seek "=", FG_FIG.TextMatrix(K, 0), FG_FIG.TextMatrix(K, 1), LT.List(J)
                        If Not DLL_BD.BDSIS_TBEST.NoMatch Then
                            BS.SimpleText = "Alterando informações sobre matéria-prima da ficha de estoque..."
                            DLL_BD.BDSIS_TBEST.Edit
                            DLL_BD.BDSIS_TBEST_CPINQ.Value = nQ
                            DLL_BD.BDSIS_TBEST_CPINP.Value = nP
                            DLL_BD.BDSIS_TBEST_CPINN.Value = nN
                            DLL_BD.BDSIS_TBEST_CPINB.Value = nB
                            DLL_BD.BDSIS_TBEST_CPINM.Value = nM
                            DLL_BD.BDSIS_TBEST.Update
                        End If
                    Next K
                    BP.Value = BP.Value + 1
                Next J 'da lista de MAT
            End If 'do CK_M
            TelaEmEspera False
        End If 'do vbYes
    End If 'salvar
    BT_Cancelar_Click
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_MateriaPrima
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        MsgBox "Selecione primeiro uma figura.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Figura.SetFocus
    End If
    CB_Bitola.SelLength = Len(CB_Bitola.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CB_Bitola.Text <> "" And CK_M.Value = 0 Then CB_Material.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    CB_Bitola.Text = UCase(CB_Bitola.Text)
    If CB_Bitola.Text <> "" Then
        For I = 0 To CB_Bitola.ListCount - 1
            If CB_Bitola.Text = CB_Bitola.List(I) Then
                Exit For
            ElseIf CB_Bitola.Text <> CB_Bitola.List(I) And I = CB_Bitola.ListCount - 1 Then
                MsgBox "Essa bitola digitada não existe - consulte esta lista.", vbOKOnly + vbInformation, NOMEAPLIC
                CB_Bitola.SetFocus
                Exit Sub
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_Click()
    CarregaFIGBITMAT
End Sub
Private Sub CB_Figura_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    CB_Figura.SelLength = Len(CB_Figura.Text)
    CB_Material.Text = ""
    CB_Bitola.Text = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CB_Figura.Text <> "" Then
        CB_Bitola.SetFocus
    ElseIf KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        CarregaFIGBITMAT
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_LostFocus()
    CarregaFIGBITMAT
End Sub
Private Sub CB_Material_Change()
    On Error GoTo ERRO_SISCOVAL
    CB_Material.SelLength = Len(CB_Material.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma figura.", vbOKOnly + vbInformation, NOMEAPLIC)
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma bitola.", vbOKOnly + vbInformation, NOMEAPLIC)
        CB_Bitola.SetFocus
        Exit Sub
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CB_Material.Text <> "" Then
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    CB_Material.Text = UCase(CB_Material.Text)
    If CB_Material.Text <> "" Then
        For I = 0 To CB_Material.ListCount - 1
            If CB_Material.Text = CB_Material.List(I) Then
                Exit For
            ElseIf CB_Material.Text <> CB_Material.List(I) And I = CB_Material.ListCount - 1 Then
                MsgBox "Esse material digitado não existe - consulte esta lista.", vbOKOnly + vbInformation, NOMEAPLIC
                CB_Material.SetFocus
                Exit Sub
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_MATMP_Click()
    BT_Salvar.Enabled = True
End Sub
Private Sub CB_MATPECA_Change()
    'ProcuraMaterialMP
End Sub
Private Sub CB_MATPECA_Click()
    'ProcuraMaterialMP
End Sub
Private Sub CB_MP_Change()
    'ProcuraMaterialMP
End Sub
Private Sub CB_MP_Click()
    'ProcuraMaterialMP
End Sub
Private Sub CK_M_Click()
    If CK_M.Value = 0 Then
        CarregaFG (0)
        CarregaFG (1)
    Else
        CarregaFG (2)
        CarregaFG (3)
    End If
End Sub

Private Sub Command1_Click()
CONFIGURAMP
End Sub
Private Sub FG_FIG_Click()
    If FG_FIG.Rows > 0 Then
        TXT_Quantidade.Text = ""
        CB_Figura.Text = FG_FIG.TextMatrix(FG_FIG.RowSel, 0)
        TXT_Nome.Text = ""
        If CK_M.Value = 0 Then
            CarregaFIGBITMAT FG_FIG.TextMatrix(FG_FIG.RowSel, 1), FG_FIG.TextMatrix(FG_FIG.RowSel, 2)
        Else
            CarregaFIGBITMAT FG_FIG.TextMatrix(FG_FIG.RowSel, 1)
        End If
    End If
End Sub
Private Sub FG_FIG_SelChange()
    FG_FIG_Click
End Sub
Private Sub FG_MP_Click()
    If FG_MP.Rows > 0 Then
        TXT_Quantidade.Text = FG_MP.TextMatrix(FG_MP.RowSel, 0)
        CB_Figura.Text = FG_MP.TextMatrix(FG_MP.RowSel, 1)
        TXT_Nome.Text = FG_MP.TextMatrix(FG_MP.RowSel, 2)
        If CK_M.Value = 0 Then
            CarregaFIGBITMAT FG_MP.TextMatrix(FG_MP.RowSel, 3), FG_MP.TextMatrix(FG_MP.RowSel, 4)
        Else
            CarregaFIGBITMAT FG_MP.TextMatrix(FG_MP.RowSel, 3)
        End If
    End If
End Sub
Private Sub FG_MP_SelChange()
    FG_MP_Click
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (26)
    DLL_CARGA.ResetaBP
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela Grupos...")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque...")
    If DLL_BD.AbreTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Índice...")
    If DLL_BD.AbreTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Figuras...")
    If DLL_BD.AbreTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Quantidades...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeQuantidades(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Pecas...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDePecas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Nomes...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeNomes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Bitolas...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeBitolas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Materiais...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Relação de Materiais...")
    If DLL_BD.AbreTabela_MateriaPrimaRelacaoMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    ' Abre Campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque...")
    If DLL_BD.AbreCampos_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Índice...")
    If DLL_BD.AbreCampos_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Figuras...")
    If DLL_BD.AbreCampos_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Quantidades...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeQuantidades(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Pecas...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDePecas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Nomes...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeNomes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Bitolas...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeBitolas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Materiais...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Relação de Materiais...")
    If DLL_BD.AbreCampos_MateriaPrimaRelacaoMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Carrega figuras
    CB_Figura.Clear
    CB_MP.Clear
    DLL_CARGA.CarregaTexto ("Carregando listas de figuras...")
    If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEFG.MoveFirst
        While Not DLL_BD.BDSIS_TBEFG.EOF
            If DLL_BD.BDSIS_TBEFG_CPFIG.Value <> "" Then
                CB_Figura.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value)
                CB_MP.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value)
            End If
            DLL_BD.BDSIS_TBEFG.MoveNext
        Wend
    End If
    'Carrega materiais
    CB_MATPECA.Clear
    CB_MATMP.Clear
    Tela_Cfg_MateriaPrima_Materiais.CB_MAT.Clear
    DLL_CARGA.CarregaTexto ("Carregando listas de materiais...")
    If DLL_BD.BDSIS_TBGRU.RecordCount > 0 Then
        DLL_BD.BDSIS_TBGRU.MoveFirst
        Do While Not DLL_BD.BDSIS_TBGRU.EOF
            If DLL_BD.BDSIS_TBGRU_CPTIP.Value = "MAT" Then
                CB_MATPECA.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
                CB_MATMP.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
                Tela_Cfg_MateriaPrima_Materiais.CB_MAT.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
            End If
            DLL_BD.BDSIS_TBGRU.MoveNext
        Loop
    End If
    'Monta tela
    DLL_CARGA.CarregaTexto ("Organizando tela...")
    BP.Value = 0
    BS.SimpleText = ""
    FR_MAT.Left = 0
    FR_MAT.Top = 0
    ModoTela (0)
    CarregaFG (0)
    CarregaFG (1)
    Tela_Cfg_MateriaPrima_Materiais.Hide
    ESTIND = ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    DLL_CARGA.Exibe (False)
    DLL_FUNCS.RegistraEvento "Abrir Configurações de Matéria-Prima", ""
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
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeQuantidades(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDePecas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeNomes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeBitolas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaRelacaoMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
    Set DLL_FUNCS = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub

Private Sub ST_Click(PreviousTab As Integer)
    If ST.Tab = 0 Then
        LB_Q.Enabled = False
        TXT_Quantidade.Enabled = False
        LB_N.Enabled = False
        TXT_Nome.Enabled = False
        BT_Importar.Enabled = False
    Else
        LB_Q.Enabled = True
        TXT_Quantidade.Enabled = True
        LB_N.Enabled = True
        TXT_Nome.Enabled = True
        BT_Importar.Enabled = True
    End If
End Sub
Private Sub TXT_Nome_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_Bitola.SetFocus
End Sub
Private Sub TXT_Quantidade_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CB_Figura.SetFocus
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Public Static Sub TelaEmEspera(Estado As Boolean)
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
Private Static Sub CarregaFG(FG As Integer)
    If FG = 0 Then 'Com material na FIG
        FG_FIG.Clear
        FG_FIG.FixedCols = 0
        FG_FIG.Cols = 3
        FG_FIG.Rows = 1
        FG_FIG.ColAlignment(0) = flexAlignLeftCenter
        FG_FIG.ColAlignment(1) = flexAlignLeftCenter
        FG_FIG.ColAlignment(2) = flexAlignLeftCenter
        FG_FIG.ColWidth(0) = 2000
        FG_FIG.ColWidth(1) = 2000
        FG_FIG.ColWidth(2) = 2000
        FG_FIG.TextArray(0) = "Figura"
        FG_FIG.TextArray(1) = "Bitola"
        FG_FIG.TextArray(2) = "Material"
    ElseIf FG = 1 Then 'Com material na MP
        FG_MP.Clear
        FG_MP.FixedCols = 0
        FG_MP.Cols = 5
        FG_MP.Rows = 1
        FG_MP.ColAlignment(0) = flexAlignLeftCenter
        FG_MP.ColAlignment(1) = flexAlignLeftCenter
        FG_MP.ColAlignment(2) = flexAlignLeftCenter
        FG_MP.ColAlignment(3) = flexAlignLeftCenter
        FG_MP.ColAlignment(4) = flexAlignLeftCenter
        FG_MP.ColWidth(0) = 1100
        FG_MP.ColWidth(1) = 1600
        FG_MP.ColWidth(2) = 1600
        FG_MP.ColWidth(3) = 1500
        FG_MP.ColWidth(4) = 1500
        FG_MP.TextArray(0) = "Quantidade"
        FG_MP.TextArray(1) = "Peça"
        FG_MP.TextArray(2) = "Nome da Peça"
        FG_MP.TextArray(3) = "Bitola"
        FG_MP.TextArray(4) = "Material"
    ElseIf FG = 2 Then 'sem material na FIG
        FG_FIG.Clear
        FG_FIG.FixedCols = 0
        FG_FIG.Cols = 2
        FG_FIG.Rows = 1
        FG_FIG.ColAlignment(0) = flexAlignLeftCenter
        FG_FIG.ColAlignment(1) = flexAlignLeftCenter
        FG_FIG.ColWidth(0) = 2000
        FG_FIG.ColWidth(1) = 2000
        FG_FIG.TextArray(0) = "Figura"
        FG_FIG.TextArray(1) = "Bitola"
    ElseIf FG = 3 Then 'sem material na MP
        FG_MP.Clear
        FG_MP.FixedCols = 0
        FG_MP.Cols = 4
        FG_MP.Rows = 1
        FG_MP.ColAlignment(0) = flexAlignLeftCenter
        FG_MP.ColAlignment(1) = flexAlignLeftCenter
        FG_MP.ColAlignment(2) = flexAlignLeftCenter
        FG_MP.ColAlignment(3) = flexAlignLeftCenter
        FG_MP.ColWidth(0) = 1200
        FG_MP.ColWidth(1) = 1500
        FG_MP.ColWidth(2) = 1500
        FG_MP.ColWidth(3) = 1500
        FG_MP.TextArray(0) = "Quantidade"
        FG_MP.TextArray(1) = "Peça"
        FG_MP.TextArray(2) = "Nome da Peça"
        FG_MP.TextArray(3) = "Bitola"
    End If
    If FG = 0 Or FG = 1 Then
        LB_M.Enabled = True
        CB_Material.Enabled = True
    ElseIf FG = 2 Or FG = 3 Then
        LB_M.Enabled = False
        CB_Material.Enabled = False
    End If
End Sub
Private Static Sub CarregaBitolaMaterial()
    On Error GoTo ERRO_SISCOVAL
    Dim cA As String
    If CB_Figura.Text = "" Then Exit Sub
    CB_Figura.Text = UCase(CB_Figura.Text)
    For I = 0 To CB_Figura.ListCount - 1
        If CB_Figura.Text = CB_Figura.List(I) Then
            Exit For
        ElseIf CB_Figura.Text <> CB_Figura.List(I) And I = CB_Figura.ListCount - 1 Then
            RespMsg = MsgBox("Essa figura digitada não existe - consulte esta lista.", vbOKOnly, NOMEAPLIC)
            CB_Figura.SetFocus
            Exit Sub
        End If
    Next I
    'procura figura
    DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figura.Text
    If DLL_BD.BDSIS_TBEFG.NoMatch Then
        RespMsg = MsgBox("Ocorreu algum erro durante a procura do índice da figura.", vbOKOnly, NOMEAPLIC)
        Exit Sub
    End If
    'Como é tabelas relacionadas, a procura acima ja acha o indice de figura
    'Montando lista de materiais
    cA = ""
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
            CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    'Montando lista de bitolas
    cA = ""
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGBI.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) = ";" Then
            CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    'seleciona material A-105
    CB_Material.ListIndex = 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ModoTela(Modo As Integer)
    If Modo = 0 Then 'Normal
        FR.Enabled = True
        LB_Q.Enabled = False
        LB_P.Enabled = True
        LB_N.Enabled = False
        LB_B.Enabled = True
        LB_M.Enabled = True
        TXT_Quantidade.Enabled = False
        CB_Figura.Enabled = True
        TXT_Nome.Enabled = False
        CB_Bitola.Enabled = True
        CB_Material.Enabled = True
        BT_Adicionar.Enabled = False
        BT_Remover.Enabled = False
        BT_Alterar.Enabled = False
        BT_Importar.Enabled = False
        CK_M.Enabled = True
        ST.Tab = 0
        ST.TabEnabled(0) = False
        ST.TabEnabled(1) = False
        FG_FIG.Enabled = False
        FG_MP.Enabled = False
        BT_Novo.Enabled = True
        BT_Consulta.Enabled = True
        BT_ConfigMP.Enabled = True
        BT_Salvar.Enabled = False
        BT_Imprimir.Enabled = False
        BT_Apagar.Enabled = False
        BT_Cancelar.Enabled = False
        BT_Voltar.Enabled = True
        FR_MAT.Visible = False
    ElseIf Modo = 1 Then 'Novo
        FR.Enabled = True
        LB_Q.Enabled = False
        LB_P.Enabled = True
        LB_N.Enabled = False
        LB_B.Enabled = True
        If CK_M.Value = 0 Then LB_M.Enabled = True
        TXT_Quantidade.Enabled = False
        CB_Figura.Enabled = True
        TXT_Nome.Enabled = False
        CB_Bitola.Enabled = True
        If CK_M.Value = 0 Then CB_Material.Enabled = True
        BT_Adicionar.Enabled = True
        BT_Remover.Enabled = True
        BT_Alterar.Enabled = True
        BT_Importar.Enabled = True
        CK_M.Enabled = False
        ST.Tab = 0
        ST.TabEnabled(0) = True
        ST.TabEnabled(1) = True
        FG_FIG.Enabled = True
        FG_MP.Enabled = True
        BT_Novo.Enabled = False
        BT_Consulta.Enabled = False
        BT_ConfigMP.Enabled = False
        BT_Salvar.Enabled = True
        BT_Imprimir.Enabled = True
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
        FR_MAT.Visible = False
    ElseIf Modo = 2 Then 'Consulta
        FR.Enabled = True
        LB_Q.Enabled = True
        LB_P.Enabled = True
        LB_N.Enabled = True
        LB_B.Enabled = True
        LB_M.Enabled = True
        TXT_Quantidade.Enabled = True
        CB_Figura.Enabled = True
        TXT_Nome.Enabled = True
        CB_Bitola.Enabled = True
        CB_Material.Enabled = True
        BT_Adicionar.Enabled = True
        BT_Remover.Enabled = True
        BT_Alterar.Enabled = True
        BT_Importar.Enabled = True
        CK_M.Enabled = False
        ST.Tab = 1
        ST.TabEnabled(0) = False
        ST.TabEnabled(1) = True
        FG_FIG.Enabled = False
        FG_MP.Enabled = True
        BT_Novo.Enabled = False
        BT_Consulta.Enabled = False
        BT_ConfigMP.Enabled = False
        BT_Salvar.Enabled = True
        BT_Imprimir.Enabled = True
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
        FR_MAT.Visible = False
    ElseIf Modo = 3 Then 'ConfigMP
        BT_Novo.Enabled = False
        BT_Consulta.Enabled = False
        BT_ConfigMP.Enabled = False
        BT_Salvar.Enabled = True
        BT_Imprimir.Enabled = False
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
        BT_Voltar.Enabled = False
        FR_MAT.Visible = True
    End If
End Sub
Private Static Sub ModoFrame(Modo As Integer)
    If Modo = 0 Then 'Consulta
        FR_MAT.Visible = False
    ElseIf Modo = 1 Then 'Config
        FR_MAT.Visible = True
    End If
End Sub
Private Static Function ProcuraMaterialMP(aFIG As Variant, sMat As Variant) As Variant
    If UBound(aFIG) < 1 And sMat = "" Then Exit Function
    Dim aFIGTMP As Variant
    aFIGTMP = Array()
    For I = 0 To UBound(aFIG)
        DLL_BD.BDSIS_TBMPR.Seek "=", aFIG(I), sMat
        ReDim Preserve aFIGTMP(UBound(aFIGTMP) + 1)
        If DLL_BD.BDSIS_TBMPR.NoMatch Then
            RespMsg = InputBox("Não foi possível localizar as configurações da seguinte peça:" & vbCr & vbCr & aFIG(I) & " em " & sMat & vbCr & vbCr & "Digite o material que você deseja usar.", "Falta material")
            aFIGTMP(UBound(aFIGTMP)) = RespMsg
        Else
            'achou a cfg do material
            aFIGTMP(UBound(aFIGTMP)) = DLL_BD.BDSIS_TBMPR_CPMMP.Value
        End If
    Next I
    ProcuraMaterialMP = aFIGTMP
End Function
Private Static Sub ProcuraMP(QUA As String, PEC As String, NOM As String, BIT As String, MAT As String)
    MATPRI.QUA = ""
    MATPRI.PEC = ""
    MATPRI.NOM = ""
    MATPRI.BIT = ""
    MATPRI.MAT = ""
    'altera para novos indices
    DLL_BD.BDSIS_TBMPQ.Index = "Índice de Quantidades"
    DLL_BD.BDSIS_TBMPP.Index = "Índice de Peças"
    DLL_BD.BDSIS_TBMPN.Index = "Índice de Nomes"
    DLL_BD.BDSIS_TBMPB.Index = "Índice de Bitolas"
    DLL_BD.BDSIS_TBMPM.Index = "Índice de Materiais"
    'procura pecas
    DLL_BD.BDSIS_TBMPQ.Seek "=", QUA
    If Not DLL_BD.BDSIS_TBMPQ.NoMatch Then MATPRI.QUA = DLL_BD.BDSIS_TBMPQ_CPQUA.Value
    DLL_BD.BDSIS_TBMPP.Seek "=", PEC
    If Not DLL_BD.BDSIS_TBMPP.NoMatch Then MATPRI.PEC = DLL_BD.BDSIS_TBMPP_CPPEC.Value
    DLL_BD.BDSIS_TBMPN.Seek "=", NOM
    If Not DLL_BD.BDSIS_TBMPN.NoMatch Then MATPRI.NOM = DLL_BD.BDSIS_TBMPN_CPNOM.Value
    DLL_BD.BDSIS_TBMPB.Seek "=", BIT
    If Not DLL_BD.BDSIS_TBMPB.NoMatch Then MATPRI.BIT = DLL_BD.BDSIS_TBMPB_CPBIT.Value
    DLL_BD.BDSIS_TBMPM.Seek "=", MAT
    If Not DLL_BD.BDSIS_TBMPM.NoMatch Then MATPRI.MAT = DLL_BD.BDSIS_TBMPM_CPMAT.Value
    'altera para velhos indices
    DLL_BD.BDSIS_TBMPQ.Index = "Quantidades"
    DLL_BD.BDSIS_TBMPP.Index = "Peças"
    'DLL_BD.BDSIS_TBMPN.Index = "Nomes"
    DLL_BD.BDSIS_TBMPB.Index = "Bitolas"
    DLL_BD.BDSIS_TBMPM.Index = "Materiais"
End Sub
Private Static Sub DivideMP(Tipo As String, Valor As String, LinIni As Integer)
    If Tipo = "" Or Valor = "" Then Exit Sub
    'comeca dividir
    Dim cA As String, nA As Integer
    cA = ""
    nA = LinIni
    For I = 1 To Len(Valor)
        If Mid(Valor, I, 1) <> ";" Then
            cA = cA & Mid(Valor, I, 1)
        ElseIf Mid(Valor, I, 1) = ";" Then
            'insere dados
            If Tipo = "QUA" Then 'quantidade
                FG_MP.AddItem (1)
                FG_MP.TextMatrix(FG_MP.Rows - 1, 0) = cA
            ElseIf Tipo = "PEC" Then 'peca
                FG_MP.TextMatrix(nA, 1) = cA
            ElseIf Tipo = "NOM" Then 'nome
                FG_MP.TextMatrix(nA, 2) = cA
            ElseIf Tipo = "BIT" Then 'bitola
                FG_MP.TextMatrix(nA, 3) = cA
            ElseIf Tipo = "MAT" Then 'material
                FG_MP.TextMatrix(nA, 4) = cA
            End If
            cA = ""
            If Tipo <> "QUA" Then nA = nA + 1
        End If
    Next I
    'insere o primeiro ou o ultimo
    If Tipo = "QUA" Then 'quantidade
        FG_MP.AddItem (1)
        FG_MP.TextMatrix(FG_MP.Rows - 1, 0) = cA
    ElseIf Tipo = "PEC" Then 'peca
        FG_MP.TextMatrix(FG_MP.Rows - 1, 1) = cA
    ElseIf Tipo = "NOM" Then 'nome
        FG_MP.TextMatrix(FG_MP.Rows - 1, 2) = cA
    ElseIf Tipo = "BIT" Then 'bitola
        FG_MP.TextMatrix(FG_MP.Rows - 1, 3) = cA
    ElseIf Tipo = "MAT" Then 'material
        FG_MP.TextMatrix(FG_MP.Rows - 1, 4) = cA
    End If
End Sub
Private Static Sub CarregaFIGBITMAT(Optional ColocaBIT As String, Optional ColocaMAT As String)
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then Exit Sub
    CB_Figura.Text = UCase(CB_Figura.Text)
    For I = 0 To CB_Figura.ListCount - 1
        If CB_Figura.Text = CB_Figura.List(I) Then
            Exit For
        ElseIf CB_Figura.Text <> CB_Figura.List(I) And I = CB_Figura.ListCount - 1 Then
            MsgBox "Essa figura digitada não existe - consulte esta lista.", vbOKOnly + vbInformation, NOMEAPLIC
            CB_Figura.SetFocus
            Exit Sub
        End If
    Next I
    'procura figura
    DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figura.Text
    've se o indice de figura da nova consulta é igual a figura anteriormente consultada
    If DLL_BD.BDSIS_TBEFG_CPIFG.Value = ESTIND And CB_Bitola.ListCount > 1 Then Exit Sub
    ESTIND = DLL_BD.BDSIS_TBEFG_CPIFG.Value
    DLL_BD.BDSIS_TBEID.Seek "=", DLL_BD.BDSIS_TBEFG_CPIFG.Value
    If DLL_BD.BDSIS_TBEFG.NoMatch And DLL_BD.BDSIS_TBEID.NoMatch Then
        MsgBox "Ocorreu algum erro durante a procura do índice da figura.", vbOKOnly + vbInformation, NOMEAPLIC
        Exit Sub
    End If
    
    'pega nome peça
    TXT_Nome.Text = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & Trim(DLL_BD.BDSIS_TBEID_CPTNO.Value) & " " & Trim(DLL_BD.BDSIS_TBEFG_CPCOM.Value)
    
    'Como são tabelas relacionadas, a procura acima ja acha o indice de figura
    Dim cA As String
    'Montando lista de bitolas
    cA = ""
    CB_Bitola.Clear
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGBI.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) = ";" Then
            CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    If IsEmpty(ColocaBIT) = False Then CB_Bitola.Text = ColocaBIT
    'Montando lista de materiais
    cA = ""
    CB_Material.Clear
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
            CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    'seleciona material A-105
    CB_Material.ListIndex = 0
    If IsEmpty(ColocaMAT) = False Then CB_Material.Text = ColocaMAT
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Public Static Function ProcuraIndicesMP(sQua As String, sFIG As String, sNOM As String, sBit As String, sMat As String) As Variant
    Dim nQ As Integer, nP As Integer, nN As Integer, nB As Integer, nM As Integer
    nQ = 0
    nP = 0
    nN = 0
    nB = 0
    nM = 0
    'procura se o índice de peças existe
    DLL_BD.BDSIS_TBMPQ.Seek "=", sQua 'Quantidades
    If DLL_BD.BDSIS_TBMPQ.NoMatch Then
        DLL_BD.BDSIS_TBMPQ.AddNew
        DLL_BD.BDSIS_TBMPQ_CPQUA.Value = sQua
        nQ = DLL_BD.BDSIS_TBMPQ_CPINQ.Value
        DLL_BD.BDSIS_TBMPQ.Update
    Else
        nQ = DLL_BD.BDSIS_TBMPQ_CPINQ.Value
    End If
    DLL_BD.BDSIS_TBMPP.Seek "=", sFIG 'Peças
    If DLL_BD.BDSIS_TBMPP.NoMatch Then
        DLL_BD.BDSIS_TBMPP.AddNew
        DLL_BD.BDSIS_TBMPP_CPPEC.Value = sFIG
        nP = DLL_BD.BDSIS_TBMPP_CPINP.Value
        DLL_BD.BDSIS_TBMPP.Update
    Else
        nP = DLL_BD.BDSIS_TBMPP_CPINP.Value
    End If
    nN = 0
    If DLL_BD.BDSIS_TBMPN.RecordCount > 0 Then 'Nomes
        DLL_BD.BDSIS_TBMPN.MoveFirst
        Do While Not DLL_BD.BDSIS_TBMPN.EOF
            If DLL_BD.BDSIS_TBMPN_CPNOM.Value = sNOM And _
               Len(DLL_BD.BDSIS_TBMPN_CPNOM.Value) = Len(sNOM) Then
                nN = DLL_BD.BDSIS_TBMPN_CPINN.Value
                Exit Do
            End If
            DLL_BD.BDSIS_TBMPN.MoveNext
        Loop
    End If
    If DLL_BD.BDSIS_TBMPN.RecordCount = 0 Or nN = 0 Then
        DLL_BD.BDSIS_TBMPN.AddNew
        DLL_BD.BDSIS_TBMPN_CPNOM.Value = sNOM
        nN = DLL_BD.BDSIS_TBMPN_CPINN.Value
        DLL_BD.BDSIS_TBMPN.Update
    End If
    DLL_BD.BDSIS_TBMPB.Seek "=", sBit 'Bitolas
    If DLL_BD.BDSIS_TBMPB.NoMatch Then
        DLL_BD.BDSIS_TBMPB.AddNew
        DLL_BD.BDSIS_TBMPB_CPBIT.Value = sBit
        nB = DLL_BD.BDSIS_TBMPB_CPINB.Value
        DLL_BD.BDSIS_TBMPB.Update
    Else
        nB = DLL_BD.BDSIS_TBMPB_CPINB.Value
    End If
    DLL_BD.BDSIS_TBMPM.Seek "=", sMat 'Materiais
    If DLL_BD.BDSIS_TBMPM.NoMatch Then
        DLL_BD.BDSIS_TBMPM.AddNew
        DLL_BD.BDSIS_TBMPM_CPMAT.Value = sMat
        nM = DLL_BD.BDSIS_TBMPM_CPINM.Value
        DLL_BD.BDSIS_TBMPM.Update
    Else
        nM = DLL_BD.BDSIS_TBMPM_CPINM.Value
    End If
    ProcuraIndicesMP = Array(nQ, nP, nN, nB, nM)
End Function




'********************************************************
' FUNCAO AUXILIAR PARA CONFIGURAR TODO ESTOQUE DE UMA VEZ
'********************************************************
Private Static Sub CONFIGURAMP()
    Dim sFIG As String, cA As String
    sFIG = InputBox("Digite o indice de figura q vc deseja configurar:", "Indice de Figura")
    If sFIG = "0" Then 'zera as MP nao configuradas
        BP.Max = DLL_BD.BDSIS_TBEST.RecordCount
        DLL_BD.BDSIS_TBEST.MoveFirst
        Do While Not DLL_BD.BDSIS_TBEST.EOF
            If DLL_BD.BDSIS_TBEST_CPINQ.Value = "" Then
                DLL_BD.BDSIS_TBEST.Edit
                DLL_BD.BDSIS_TBEST_CPINQ.Value = 0
                DLL_BD.BDSIS_TBEST.Update
            ElseIf DLL_BD.BDSIS_TBEST_CPINP.Value = "" Then
                DLL_BD.BDSIS_TBEST.Edit
                DLL_BD.BDSIS_TBEST_CPINP.Value = 0
                DLL_BD.BDSIS_TBEST.Update
            ElseIf DLL_BD.BDSIS_TBEST_CPINN.Value = "" Then
                DLL_BD.BDSIS_TBEST.Edit
                DLL_BD.BDSIS_TBEST_CPINN.Value = 0
                DLL_BD.BDSIS_TBEST.Update
            ElseIf DLL_BD.BDSIS_TBEST_CPINB.Value = "" Then
                DLL_BD.BDSIS_TBEST.Edit
                DLL_BD.BDSIS_TBEST_CPINB.Value = 0
                DLL_BD.BDSIS_TBEST.Update
            ElseIf DLL_BD.BDSIS_TBEST_CPINM.Value = "" Then
                DLL_BD.BDSIS_TBEST.Edit
                DLL_BD.BDSIS_TBEST_CPINM.Value = 0
                DLL_BD.BDSIS_TBEST.Update
            End If
            DLL_BD.BDSIS_TBEST.MoveNext
            If BP.Value < BP.Max Then BP.Value = BP.Value + 1
        Loop
        Exit Sub
    End If
    If sFIG = "CP" Then
        ConfigCP
        Exit Sub
    ElseIf sFIG = "PA" Then
        ConfigPA
        Exit Sub
    End If
    If sFIG = "" Or Not IsNumeric(sFIG) Then
        GoTo CONFIGURAMP
    ElseIf sFIG <> "1" And sFIG <> "2" And sFIG <> "3" And sFIG <> "4" And sFIG <> "5" And sFIG <> "7" And sFIG <> "9" And _
       sFIG <> "10" And sFIG <> "12" And sFIG <> "13" And sFIG <> "14" And sFIG <> "15" And sFIG <> "20" And _
       sFIG <> "100" And sFIG <> "120" Then
        MsgBox "O índice de figura que você digitou não existe ou não foi configurado.", vbInformation + vbOKOnly, "Erro"
        GoTo CONFIGURAMP
    End If
    DLL_BD.BDSIS_TBEID.Seek "=", sFIG
    If DLL_BD.BDSIS_TBEID.NoMatch Then
        MsgBox "Este índice de figura não existe!", vbOKOnly + vbInformation, NOMEAPLIC
        GoTo CONFIGURAMP
    Else
        'procura todas figuras deste indice
        BS.SimpleText = "Procurando figuras deste índice..."
        BP.Max = DLL_BD.BDSIS_TBEFG.RecordCount
        BP.Value = 0
        DLL_BD.BDSIS_TBEFG.MoveFirst
        Dim aFIGURAS As Variant, aBITOLAS As Variant, aMATERIAIS As Variant
        aFIGURAS = Array()
        Do While Not DLL_BD.BDSIS_TBEFG.EOF
            If DLL_BD.BDSIS_TBEFG_CPIFG.Value = sFIG Then
                ReDim Preserve aFIGURAS(UBound(aFIGURAS) + 1)
                aFIGURAS(UBound(aFIGURAS)) = DLL_BD.BDSIS_TBEFG_CPFIG.Value
            End If
            DLL_BD.BDSIS_TBEFG.MoveNext
            BP.Value = BP.Value + 1
        Loop
        BS.SimpleText = ""
        BP.Value = 0
    End If
    RespMsg = MsgBox("Foram encontradas " & Val(UBound(aFIGURAS) + 1) & " do índice de figuras " & sFIG & " - Você tem certeza que deseja configurar os ítens de estoque de todas as figuras deste índice ?", vbYesNo + vbQuestion + vbDefaultButton1, "Config de MP")
    If RespMsg = vbYes Then
        TelaEmEspera True
        'procura indice
        DLL_BD.BDSIS_TBEID.Seek "=", sFIG
        If DLL_BD.BDSIS_TBEID.NoMatch Then
            MsgBox "Ocorreu algum erro durante a procura do índice da figura.", vbOKOnly, NOMEAPLIC
            TelaEmEspera False
            GoTo CONFIGURAMP
        End If
        'Montando vetor de bitolas
        aBITOLAS = Array()
        cA = ""
        For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGBI.Value)
            If Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) <> ";" Then
                cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1)
            ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) = ";" Then
                ReDim Preserve aBITOLAS(UBound(aBITOLAS) + 1)
                aBITOLAS(UBound(aBITOLAS)) = DLL_FUNCS.ProcuraGrupo(cA)
                cA = ""
            End If
        Next I
        ReDim Preserve aBITOLAS(UBound(aBITOLAS) + 1)
        aBITOLAS(UBound(aBITOLAS)) = DLL_FUNCS.ProcuraGrupo(cA)
        'Montando vetor de materiais
        aMATERIAIS = Array()
        cA = ""
        For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
            If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
                cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
            ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
                ReDim Preserve aMATERIAIS(UBound(aMATERIAIS) + 1)
                aMATERIAIS(UBound(aMATERIAIS)) = DLL_FUNCS.ProcuraGrupo(cA)
                cA = ""
            End If
        Next I
        ReDim Preserve aMATERIAIS(UBound(aMATERIAIS) + 1)
        aMATERIAIS(UBound(aMATERIAIS)) = DLL_FUNCS.ProcuraGrupo(cA)
        
        'procura informacoes sobre materia-prima
        Dim aMP As Variant, nA As Integer, nB As Integer, nC As Integer, nIII As Long, lProc As Boolean
        Dim aQUA As Variant, aFIG As Variant, aNOM As Variant, aBIT As Variant, aMAT As Variant
        Dim iQUA As String, iFIG As String, iNOM As String, iBIT As String, iMAT As String
        BP.Max = (Val(UBound(aFIGURAS) + 1) * Val(UBound(aBITOLAS) + 1) * Val(UBound(aMATERIAIS) + 1))
        BP.Value = 0
        nIII = 0
        BS.SimpleText = "Aguarde... configurando estoque."
        For nA = 0 To UBound(aFIGURAS)
            For nB = 0 To UBound(aBITOLAS)
                For nC = 0 To UBound(aMATERIAIS)
        
                    'pega informaçoes sobre a peca
                    If sFIG = "100" Then 'VALVULA GAVETA
                        aMP = MP_Gaveta(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = True
                    ElseIf sFIG = "120" Then 'VALVULA GLOBO
                        aMP = MP_Globo(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = True
                    ElseIf sFIG = "1" Then 'BUCHA
                        aMP = MP_BUCHA(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "2" Then 'CAPS
                        aMP = MP_CAPS(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "3" Then 'COTOVELO 90º
                        aMP = MP_COT90(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "4" Then 'COTOVELO 90º M/F
                        aMP = MP_COTMF(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "5" Then 'COTOVELO 45º
                        aMP = MP_COT45(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "7" Then 'LUVA
                        aMP = MP_LUVA(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "9" Then 'MEIA-LUVA
                        aMP = MP_MEIALUVA(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "10" Then 'NIPLE
                        aMP = MP_NIPLE(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "12" Then 'PLUG REDONDO
                        aMP = MP_PRED(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "13" Then 'PLUG QUADRADO
                        aMP = MP_PQUA(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "14" Then 'PLUG SEXTAVADO
                        aMP = MP_PSEX(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "15" Then 'TEE 90º
                        aMP = MP_TEE90(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    ElseIf sFIG = "20" Then 'CRUZETA
                        aMP = MP_CRUZETA(aFIGURAS(nA), aBITOLAS(nB))
                        lProc = False
                    Else
                        MsgBox "Ainda não foi configurado função para: " & sFIG, vbOKOnly + vbInformation, NOMEAPLIC
                        GoTo CONFIGURAMP
                    End If
                    
                    If Not IsArray(aMP) Then
                        MsgBox "Não foi possível localizar dados da ficha:" & vbCr & vbCr & aFIGURAS(nA) & " de " & aBITOLAS(nB) & " em " & aMATERIAIS(nC), vbOKOnly + vbInformation, NOMEAPLIC
                    Else
                        aQUA = aMP(0)
                        aFIG = aMP(1)
                        aNOM = aMP(2)
                        aBIT = aMP(3)
                        If lProc = True Then
                            aMAT = ProcuraMaterialMP(aFIG, aMATERIAIS(nC))
                        Else
                            aMAT = Array(aMATERIAIS(nC))
                        End If
                        'verifica tamanhos das variaveis
                        If UBound(aQUA) = UBound(aFIG) And _
                           UBound(aFIG) = UBound(aNOM) And _
                           UBound(aNOM) = UBound(aBIT) And _
                           UBound(aBIT) = UBound(aMAT) Then
                            iQUA = ""
                            iFIG = ""
                            iNOM = ""
                            iBIT = ""
                            iMAT = ""
                            For I = 0 To UBound(aQUA)
                                If I = 0 Then
                                    iQUA = aQUA(I)
                                    iFIG = aFIG(I)
                                    iNOM = aNOM(I)
                                    iBIT = aBIT(I)
                                    iMAT = aMAT(I)
                                Else
                                    iQUA = iQUA & ";" & aQUA(I)
                                    iFIG = iFIG & ";" & aFIG(I)
                                    iNOM = iNOM & ";" & aNOM(I)
                                    iBIT = iBIT & ";" & aBIT(I)
                                    iMAT = iMAT & ";" & aMAT(I)
                                End If
                            Next I
                            aMP = ProcuraIndicesMP(iQUA, iFIG, iNOM, iBIT, iMAT)
                            'salva informacoes nas fichas
                            DLL_BD.BDSIS_TBEST.Seek "=", aFIGURAS(nA), aBITOLAS(nB), aMATERIAIS(nC)
                            If Not DLL_BD.BDSIS_TBEST.NoMatch Then
                                DLL_BD.BDSIS_TBEST.Edit
                                DLL_BD.BDSIS_TBEST_CPINQ.Value = aMP(0)
                                DLL_BD.BDSIS_TBEST_CPINP.Value = aMP(1)
                                DLL_BD.BDSIS_TBEST_CPINN.Value = aMP(2)
                                DLL_BD.BDSIS_TBEST_CPINB.Value = aMP(3)
                                DLL_BD.BDSIS_TBEST_CPINM.Value = aMP(4)
                                DLL_BD.BDSIS_TBEST.Update
                            End If
                            If BP.Value < BP.Max Then BP.Value = BP.Value + 1
                            nIII = nIII + 1
                            BS.SimpleText = "Configurando ficha " & nIII & "/" & BP.Max & ": " & aFIGURAS(nA) & " de " & aBITOLAS(nB) & " em " & aMATERIAIS(nC)
                        Else
                            MsgBox "Não foi possível configurar os dados da ficha:" & vbCr & vbCr & aFIGURAS(nA) & " de " & aBITOLAS(nB) & " em " & aMATERIAIS(nC), vbOKOnly + vbInformation, NOMEAPLIC
                        End If
                    End If
                Next nC
            Next nB
        Next nA
    End If
CONFIGURAMP:
        BP.Value = 0
        BS.SimpleText = ""
        TelaEmEspera False
End Sub
