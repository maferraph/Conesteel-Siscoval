VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Tela_Cfg_NotaFiscal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações sobre a nota fiscal"
   ClientHeight    =   5070
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   8040
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Editar 
      Caption         =   "&Editar"
      Height          =   732
      Left            =   960
      Picture         =   "Tela_Cfg_NotaFiscal.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Editar os dados existentes."
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "Apa&gar"
      Height          =   732
      Left            =   2160
      Picture         =   "Tela_Cfg_NotaFiscal.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Apaga todos os campos."
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   732
      Left            =   3000
      Picture         =   "Tela_Cfg_NotaFiscal.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salva os dados."
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   732
      Left            =   3840
      Picture         =   "Tela_Cfg_NotaFiscal.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancela edição."
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   7200
      Picture         =   "Tela_Cfg_NotaFiscal.frx":0FD0
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Volta à Tela Principal."
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton BT_Incluir 
      Caption         =   "&Incluir"
      Height          =   732
      Left            =   120
      Picture         =   "Tela_Cfg_NotaFiscal.frx":1412
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Novo cadastro de empresa"
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton BT_Teste 
      Caption         =   "&Teste"
      Height          =   732
      Left            =   6360
      Picture         =   "Tela_Cfg_NotaFiscal.frx":1854
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimi nota fiscal"
      Top             =   4200
      Width           =   732
   End
   Begin VB.CommandButton BT_Img 
      Caption         =   "I&magem"
      Height          =   732
      Left            =   5520
      Picture         =   "Tela_Cfg_NotaFiscal.frx":1C96
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Imprimi nota fiscal"
      Top             =   4200
      Width           =   732
   End
   Begin TabDlg.SSTab ST_1 
      Height          =   3972
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Edição de Alíquotas."
      Top             =   120
      Width           =   7812
      _ExtentX        =   13785
      _ExtentY        =   7011
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   420
      TabCaption(0)   =   "Nota Fiscal"
      TabPicture(0)   =   "Tela_Cfg_NotaFiscal.frx":20D8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LB_XM"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LB_YM"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LB_Y1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LB_Y2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LB_X2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LB_X1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TXT_Lin"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FR_1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TXT_Col"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "CB_Item"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Ítens"
      TabPicture(1)   =   "Tela_Cfg_NotaFiscal.frx":20F4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LT_Item"
      Tab(1).Control(1)=   "TXT_Nome"
      Tab(1).Control(2)=   "TXT_Descricao"
      Tab(1).Control(3)=   "TXT_Indice"
      Tab(1).Control(4)=   "TXT_X"
      Tab(1).Control(5)=   "TXT_Y"
      Tab(1).Control(6)=   "FR_2"
      Tab(1).Control(7)=   "TXT_Exemplo"
      Tab(1).Control(8)=   "Label1"
      Tab(1).Control(9)=   "Label2"
      Tab(1).Control(10)=   "Label3"
      Tab(1).Control(11)=   "Label4"
      Tab(1).Control(12)=   "Label5"
      Tab(1).Control(13)=   "Label9"
      Tab(1).ControlCount=   14
      Begin VB.ListBox LT_Item 
         Height          =   2985
         Left            =   -74640
         Sorted          =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   2172
      End
      Begin VB.TextBox TXT_Nome 
         Height          =   288
         Left            =   -72000
         TabIndex        =   28
         Top             =   720
         Width           =   4332
      End
      Begin VB.TextBox TXT_Descricao 
         Height          =   288
         Left            =   -72000
         MaxLength       =   25
         TabIndex        =   27
         Top             =   1320
         Width           =   4332
      End
      Begin VB.TextBox TXT_Indice 
         Height          =   288
         Left            =   -72000
         TabIndex        =   26
         Top             =   3480
         Width           =   1452
      End
      Begin VB.TextBox TXT_X 
         Height          =   288
         Left            =   -70440
         TabIndex        =   25
         Top             =   3480
         Width           =   1332
      End
      Begin VB.TextBox TXT_Y 
         Height          =   288
         Left            =   -69000
         TabIndex        =   24
         Top             =   3480
         Width           =   1332
      End
      Begin VB.ComboBox CB_Item 
         Height          =   315
         ItemData        =   "Tela_Cfg_NotaFiscal.frx":2110
         Left            =   240
         List            =   "Tela_Cfg_NotaFiscal.frx":2112
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   600
         Width           =   2772
      End
      Begin VB.TextBox TXT_Col 
         Height          =   288
         Left            =   3240
         TabIndex        =   22
         Top             =   600
         Width           =   2052
      End
      Begin VB.Frame FR_1 
         Height          =   2412
         Left            =   600
         TabIndex        =   20
         Top             =   1320
         Width           =   6612
         Begin VB.Label LB_NF 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BlaBla"
            Height          =   192
            Left            =   3360
            TabIndex        =   21
            Top             =   1080
            Width           =   480
         End
         Begin VB.Image IMG_1 
            Height          =   2292
            Left            =   0
            Top             =   120
            Width           =   6612
         End
      End
      Begin VB.Frame FR_2 
         Height          =   852
         Left            =   -72000
         TabIndex        =   11
         Top             =   2280
         Width           =   4332
         Begin VB.TextBox TXT_Fonte 
            Height          =   288
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   1452
         End
         Begin VB.TextBox TXT_Tam 
            Height          =   288
            Left            =   1680
            TabIndex        =   16
            Top             =   480
            Width           =   492
         End
         Begin VB.CheckBox CK_Negrito 
            Caption         =   "Negrito"
            Height          =   192
            Left            =   2400
            TabIndex        =   15
            Top             =   120
            Width           =   972
         End
         Begin VB.CheckBox CK_Italico 
            Caption         =   "Itálico"
            Height          =   192
            Left            =   2400
            TabIndex        =   14
            Top             =   360
            Width           =   972
         End
         Begin VB.CheckBox CK_Sublinhado 
            Caption         =   "Sublinhado"
            Height          =   192
            Left            =   2400
            TabIndex        =   13
            Top             =   600
            Width           =   1092
         End
         Begin VB.CommandButton BT_Fonte 
            Height          =   492
            Left            =   3600
            Picture         =   "Tela_Cfg_NotaFiscal.frx":2114
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Novo cadastro de empresa"
            Top             =   240
            Width           =   612
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fonte"
            Height          =   192
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   408
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tam"
            Height          =   192
            Left            =   1680
            TabIndex        =   18
            Top             =   240
            Width           =   336
         End
      End
      Begin VB.TextBox TXT_Exemplo 
         Height          =   372
         Left            =   -72000
         MaxLength       =   50
         TabIndex        =   10
         Top             =   1920
         Width           =   4332
      End
      Begin VB.TextBox TXT_Lin 
         Height          =   288
         Left            =   5520
         TabIndex        =   9
         Top             =   600
         Width           =   2052
      End
      Begin VB.Label LB_X1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Left            =   600
         TabIndex        =   44
         Top             =   1080
         Width           =   84
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   192
         Left            =   -72000
         TabIndex        =   43
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   192
         Left            =   -72000
         TabIndex        =   42
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Índice:"
         Height          =   192
         Left            =   -72000
         TabIndex        =   41
         Top             =   3240
         Width           =   468
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Coordenada X:"
         Height          =   192
         Left            =   -70440
         TabIndex        =   40
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Coordenada Y:"
         Height          =   192
         Left            =   -69000
         TabIndex        =   39
         Top             =   3240
         Width           =   1092
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Coordenada - Linha:"
         Height          =   192
         Left            =   5520
         TabIndex        =   38
         Top             =   360
         Width           =   1452
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Coordenada - Coluna:"
         Height          =   192
         Left            =   3240
         TabIndex        =   37
         Top             =   360
         Width           =   1572
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Ítem:"
         Height          =   192
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   336
      End
      Begin VB.Label LB_X2 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Left            =   6720
         TabIndex        =   35
         Top             =   1080
         Width           =   84
      End
      Begin VB.Label LB_Y2 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Left            =   120
         TabIndex        =   34
         Top             =   3480
         Width           =   84
      End
      Begin VB.Label LB_Y1 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Left            =   120
         TabIndex        =   33
         Top             =   1440
         Width           =   84
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Dados de exemplo de impressão da N.F.:"
         Height          =   192
         Left            =   -72000
         TabIndex        =   32
         Top             =   1680
         Width           =   3000
      End
      Begin VB.Label LB_YM 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   84
      End
      Begin VB.Label LB_XM 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   192
         Left            =   3720
         TabIndex        =   30
         Top             =   1080
         Width           =   84
      End
   End
   Begin MSComDlg.CommonDialog CD_1 
      Left            =   1680
      Top             =   4560
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
End
Attribute VB_Name = "Tela_Cfg_NotaFiscal"
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
Const NOMEAPLIC As String = "Configurações sobre a nota fiscal"
Dim I, J As Integer, DirTmp As String, RespMsg, cResp, ArquivoNF As String
Dim ModoEdicao, ModoReposionado As Boolean
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.Text = ""
    TXT_Exemplo.Text = ""
    TXT_Indice.Text = ""
    TXT_X.Text = ""
    TXT_Y.Text = ""
    TXT_Fonte.Text = ""
    TXT_Tam.Text = ""
    CK_Negrito.Value = 0
    CK_Italico.Value = 0
    CK_Sublinhado.Value = 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    BotoesEmEdicao (False)
    BT_Apagar.Value = True
    TXT_Nome.Text = ""
    LT_Item.ListIndex = -1
    BT_Voltar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    BotoesEmEdicao (True)
    TXT_Descricao.SetFocus
    ModoEdicao = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Fonte_Click()
    On Error GoTo ERRO_SISCOVAL
    'Carrega o Dlg de fontes com dados iniciais
    CD_1.FontName = TXT_Exemplo.FontName
    CD_1.FontSize = TXT_Exemplo.FontSize
    CD_1.FontBold = TXT_Exemplo.FontBold
    CD_1.FontItalic = TXT_Exemplo.FontItalic
    CD_1.FontUnderline = TXT_Exemplo.FontUnderline
    CD_1.DialogTitle = "Selecione a fonte para o item da nota fiscal"
    CD_1.Flags = cdlCFBoth + cdlCFEffects
    CD_1.ShowFont
    'Repassa dados do Dlg de fontes para o programa
    TXT_Fonte.Text = CD_1.FontName
    TXT_Tam.Text = CD_1.FontSize
    If CD_1.FontBold = True Then
        CK_Negrito.Value = 1
    Else
        CK_Negrito.Value = 0
    End If
    If CD_1.FontItalic = True Then
        CK_Italico.Value = 1
    Else
        CK_Italico.Value = 0
    End If
    If CD_1.FontUnderline = True Then
        CK_Sublinhado.Value = 1
    Else
        CK_Sublinhado.Value = 0
    End If
    MudaFonte
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Img_Click()
    On Error GoTo ERRO_SISCOVAL
    CD_1.DialogTitle = "Indique o caminho da imagem da nota fiscal"
    'CD.Flags =
    CD_1.ShowOpen
    If CD_1.FileName <> "" Then
        'salva imagem
        ArquivoNF = ""
        DLL_BD.BDSIS_TBINF.MoveFirst
        DLL_BD.BDSIS_TBINF.Edit
        Open CD_1.FileName For Input As #1 'abre arquivo
        ArquivoNF = Input(LOF(1), 1)
        Close #1
        DLL_BD.BDSIS_TBINF_CPIMG.Value = ArquivoNF
        DLL_BD.BDSIS_TBINF.Update
        'le imagem
        DirTmp = DLL_FUNCS.DiretorioTemporario
        Open Trim(DirTmp) & "scvinf.jpg" For Output As #1
        Write #1, ArquivoNF
        Close #1
        IMG_1.Picture = LoadPicture(Trim(DirTmp) & "scvinf.jpg")
        Kill (Trim(DirTmp) & "scvinf.jpg")
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Incluir_Click()
    On Error GoTo ERRO_SISCOVAL
    Dim Peca As String
    Peca = InputBox("Digite o nome do item que você deseja incluir:", "Incluir ítens")
    If Peca = "" Then Exit Sub
    Peca = Trim(UCase(Peca))
    If Len(Peca) > 20 Then
        MsgBox ("O nome do ítem deve ter no máximo 20 caracteres.")
        Exit Sub
    End If
    
    BT_Apagar.Value = True 'Limpa campos
    DLL_BD.BDSIS_TBCNF.Seek "=", Peca
    If Not DLL_BD.BDSIS_TBCNF.NoMatch Then
        MsgBox ("Esse nome já existe... digite outro.")
        Exit Sub
    End If
    Dim Num As Integer
    If DLL_BD.BDSIS_TBCNF.RecordCount = 0 Then
        Num = 0
    Else
        Num = (DLL_BD.BDSIS_TBCNF.RecordCount) + 1
    End If
    
    TXT_Nome.Text = Peca
    TXT_Indice.Text = Num
    TXT_Fonte.Text = TXT_Exemplo.FontName
    TXT_Tam.Text = TXT_Exemplo.FontSize
    TXT_X.Text = 0
    TXT_Y.Text = 0
    BotoesEmEdicao (True)
    ModoEdicao = False
    TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    'Teste para salvar
    If TXT_Nome.Text = "" Then
        MsgBox ("Não foi digitado um nome para o ítem.")
        TXT_Nome.SetFocus
        Exit Sub
    ElseIf TXT_Descricao.Text = "" Then
        MsgBox ("Não foi digitado uma descrição para o ítem.")
        TXT_Descricao.SetFocus
        Exit Sub
    ElseIf TXT_Exemplo.Text = "" Then
        MsgBox ("Não foi digitado um exemplo para o ítem.")
        TXT_Exemplo.SetFocus
        Exit Sub
    ElseIf TXT_Fonte.Text = "" Then
        MsgBox ("Não foi digitado uma fonte para o ítem.")
        TXT_Fonte.SetFocus
        Exit Sub
    ElseIf TXT_Tam.Text = "" Then
        MsgBox ("Não foi digitado um tamanho de fonte para o ítem.")
        TXT_Tam.SetFocus
        Exit Sub
    ElseIf TXT_Indice.Text = "" Then
        MsgBox ("Não foi digitado um índice para o ítem.")
        TXT_Indice.SetFocus
        Exit Sub
    ElseIf TXT_X.Text = "" Then
        MsgBox ("Não foi digitado uma coordenada X para o ítem.")
        TXT_X.SetFocus
        Exit Sub
    ElseIf TXT_Y.Text = "" Then
        MsgBox ("Não foi digitado uma coordenada Y para o ítem.")
        TXT_Y.SetFocus
        Exit Sub
    End If
    
    If ModoEdicao = False Then
        DLL_BD.BDSIS_TBCNF.AddNew
    Else
        DLL_BD.BDSIS_TBCNF.Edit
    End If
    DLL_BD.BDSIS_TBCNF_CPITE.Value = Trim(TXT_Nome.Text)
    DLL_BD.BDSIS_TBCNF_CPDES.Value = Trim(TXT_Descricao.Text)
    DLL_BD.BDSIS_TBCNF_CPEXE.Value = Trim(TXT_Exemplo.Text)
    DLL_BD.BDSIS_TBCNF_CPIND.Value = TXT_Indice.Text
    DLL_BD.BDSIS_TBCNF_CPCOL.Value = TXT_X.Text
    DLL_BD.BDSIS_TBCNF_CPLIN.Value = TXT_Y.Text
    DLL_BD.BDSIS_TBCNF_CPFON.Value = TXT_Fonte.Text
    DLL_BD.BDSIS_TBCNF_CPTAM.Value = TXT_Tam.Text
    If CK_Negrito.Value = 1 Then DLL_BD.BDSIS_TBCNF_CPNEG.Value = True
    If CK_Italico.Value = 1 Then DLL_BD.BDSIS_TBCNF_CPITA.Value = True
    If CK_Sublinhado.Value = 1 Then DLL_BD.BDSIS_TBCNF_CPSUB.Value = True
    DLL_BD.BDSIS_TBCNF.Update
    DLL_FUNCS.RegistraEvento "Salvar - Configurações de Notas Fiscais", TXT_Nome.Text
    If ModoEdicao = False Then
        LT_Item.AddItem (Trim(TXT_Nome.Text))
        CB_Item.AddItem (Trim(TXT_Nome.Text))
    End If
    BotoesEmEdicao (False)
    BT_Apagar.Value = True
    TXT_Nome.Text = ""
    BT_Incluir.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Teste_Click()
    On Error GoTo ERRO_SISCOVAL
    'Aqui imprime uma nota de teste
    Dim cResp
    cResp = MsgBox("Você deseja imprimir em uma folha em branco os dados de nota fiscal para teste de posições dos ítens ?", vbInformation + vbYesNo + vbDefaultButton1, "Impressão de teste de Nota Fiscal")
    If cResp = vbNo Then
        Exit Sub
    End If
        
    DLL_CARGA.Max (4)
    DLL_CARGA.CarregaTexto ("Iniciando impressão da nota fiscal de teste...")
    For I = 0 To Tela_Cfg_NotaFiscal_IT.Controls.Count - 1
        Tela_Cfg_NotaFiscal_IT.Controls(I).Visible = False
    Next I
    DLL_CARGA.CarregaTexto ("Carregando valores dos textos de exemplo para impressão...")
    
    Dim nNumLB
    'Aqui le o banco de dados da configuracao da nota fiscal e carrega valores
    DLL_BD.BDSIS_TBCNF.MoveFirst
    Do While Not DLL_BD.BDSIS_TBCNF.EOF
        nNumLB = DLL_BD.BDSIS_TBCNF_CPIND.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).Visible = True
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).Caption = DLL_BD.BDSIS_TBCNF_CPEXE.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).Left = DLL_BD.BDSIS_TBCNF_CPCOL.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).Top = DLL_BD.BDSIS_TBCNF_CPLIN.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).FontName = DLL_BD.BDSIS_TBCNF_CPFON.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).FontSize = DLL_BD.BDSIS_TBCNF_CPTAM.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).FontBold = DLL_BD.BDSIS_TBCNF_CPNEG.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).FontItalic = DLL_BD.BDSIS_TBCNF_CPITA.Value
        Tela_Cfg_NotaFiscal_IT.LB_NF(nNumLB).FontUnderline = DLL_BD.BDSIS_TBCNF_CPSUB.Value
        DLL_BD.BDSIS_TBCNF.MoveNext
    Loop
    
    DLL_CARGA.CarregaTexto ("Preparando impressão...")
        Dim TP
        TP = Printer.PaperSize
        Printer.PaperSize = vbPRPSLegal  'Tamanho da nota fiscal
        Tela_Cfg_NotaFiscal_IT.PrintForm
        Printer.PaperSize = TP
    DLL_CARGA.CarregaTexto ("Finalizando.")
    DLL_CARGA.Exibe (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_NotaFiscal
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Item_Click()
    On Error GoTo ERRO_SISCOVAL
    ReposicionaNF (CB_Item.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Item_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Col.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Italico_Click()
    On Error GoTo ERRO_SISCOVAL
    MudaFonte
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Italico_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        CK_Sublinhado.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Negrito_Click()
    On Error GoTo ERRO_SISCOVAL
    MudaFonte
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Negrito_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        CK_Italico.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Sublinhado_Click()
    On Error GoTo ERRO_SISCOVAL
    MudaFonte
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Sublinhado_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Indice.SetFocus
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
    DLL_CARGA.Max (8)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela Configurações de Nota Fiscal...")
    If DLL_BD.AbreTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Configurações de Nota Fiscal - Imagem...")
    If DLL_BD.AbreTabela_ConfiguracoesNotaFiscalImagem(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Abrindo campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Configurações da Nota Fiscal...")
    If DLL_BD.AbreCampos_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Configurações de Nota Fiscal - Imagem...")
    If DLL_BD.AbreCampos_ConfiguracoesNotaFiscalImagem(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega a imagem da nota fiscal
    DLL_CARGA.CarregaTexto ("Carregando a imagem da nota fiscal...")
    If LeImagemNF = True Then
        BT_Img.Visible = False
    Else
        BT_Img.Visible = True
    End If

    'Carrega ítens
    DLL_CARGA.CarregaTexto ("Carregando lista de ítens")
    CB_Item.Clear
    LT_Item.Clear
    If DLL_BD.BDSIS_TBCNF.RecordCount <> 0 Then
        DLL_BD.BDSIS_TBCNF.MoveFirst
        Do While Not DLL_BD.BDSIS_TBCNF.EOF
            CB_Item.AddItem (DLL_BD.BDSIS_TBCNF_CPITE.Value)
            LT_Item.AddItem (DLL_BD.BDSIS_TBCNF_CPITE.Value)
            DLL_BD.BDSIS_TBCNF.MoveNext
        Loop
    End If
    
    DLL_CARGA.CarregaTexto ("Finalizando...")
    BT_Incluir.Enabled = False
    BT_Editar.Enabled = False
    BotoesEmEdicao (False)
    
    LB_X1.Caption = 0
    LB_Y1.Caption = 0
    LB_X2.Caption = FR_1.Width
    LB_Y2.Caption = FR_1.Height
    LB_XM.Caption = Int(FR_1.Width / 2)
    LB_YM.Caption = Int(FR_1.Height / 2)
    LB_NF.Visible = False
    TXT_Fonte.Text = "Arial"
    TXT_Tam.Text = "8"
    DLL_CARGA.Exibe (False)
    
    If BT_Img.Visible = True Then 'Se houve algum erro durante a leitura da imagem
        RespMsg = MsgBox("Ocorreu algum erro durante a abertura da imagem da nota fiscal. Verifique se a imagem e renicie esta tela.", vbCritical + vbOKOnly, "Erro na imagem")
        ST_1.Enabled = False
        BT_Incluir.Enabled = False
        BT_Editar.Enabled = False
        BT_Teste.Enabled = False
        BT_Img.Value = True
        Exit Sub
    End If
    
    DLL_FUNCS.RegistraEvento "Abrir Configurações de Notas Fiscais", ""
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_NotaFiscal
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_ConfiguracoesNotaFiscalImagem(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
    Set DLL_FUNCS = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Item_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Item.ListIndex = -1 Then Exit Sub
    DLL_BD.BDSIS_TBCNF.Seek "=", LT_Item.Text
    If DLL_BD.BDSIS_TBCNF.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura deste item no banco de dados.")
        LT_Item.ListIndex = -1
        Exit Sub
    End If
    TXT_Nome.Text = DLL_BD.BDSIS_TBCNF_CPITE.Value
    TXT_Descricao.Text = DLL_BD.BDSIS_TBCNF_CPDES.Value
    TXT_Exemplo.Text = DLL_BD.BDSIS_TBCNF_CPEXE.Value
    TXT_Indice.Text = DLL_BD.BDSIS_TBCNF_CPIND.Value
    TXT_X.Text = DLL_BD.BDSIS_TBCNF_CPCOL.Value
    TXT_Y.Text = DLL_BD.BDSIS_TBCNF_CPLIN.Value
    TXT_Fonte.Text = DLL_BD.BDSIS_TBCNF_CPFON.Value
    TXT_Tam.Text = DLL_BD.BDSIS_TBCNF_CPTAM.Value
    If DLL_BD.BDSIS_TBCNF_CPNEG.Value = True Then
        CK_Negrito.Value = 1
    Else
        CK_Negrito.Value = 0
    End If
    If DLL_BD.BDSIS_TBCNF_CPITA.Value = True Then
        CK_Italico.Value = 1
    Else
        CK_Italico.Value = 0
    End If
    If DLL_BD.BDSIS_TBCNF_CPSUB.Value = True Then
        CK_Sublinhado.Value = 1
    Else
        CK_Sublinhado.Value = 0
    End If
    'Caracteristicas de fonte da caixa de texto
    TXT_Exemplo.FontName = TXT_Fonte.Text
    TXT_Exemplo.FontSize = TXT_Tam.Text
    TXT_Exemplo.FontBold = CK_Negrito.Value
    TXT_Exemplo.FontItalic = CK_Italico.Value
    TXT_Exemplo.FontUnderline = CK_Sublinhado.Value
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub ST_1_Click(PreviousTab As Integer)
    On Error GoTo ERRO_SISCOVAL
    If ST_1.Tab = 0 Then
        BT_Incluir.Enabled = False
        BT_Editar.Enabled = False
    Else
        BT_Incluir.Enabled = True
        BT_Editar.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Col_Change()
    On Error GoTo ERRO_SISCOVAL
    If TestaXY(TXT_Col.Text, "X") = True Then
        If ModoReposionado = True Then cResp = ReposicionaItem(Val(TXT_Col.Text), Val(TXT_Lin.Text))
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Col_Click()
    On Error GoTo ERRO_SISCOVAL
    If TestaXY(TXT_Col.Text, "X") = True Then
        If ModoReposionado = True Then cResp = ReposicionaItem(Val(TXT_Col.Text), Val(TXT_Lin.Text))
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Col_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Lin.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.SelLength = Len(TXT_Descricao.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Exemplo.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Exemplo_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Exemplo.SelLength = Len(TXT_Exemplo.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Exemplo_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        BT_Fonte.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Fonte_Change()
    On Error GoTo ERRO_SISCOVAL
    MudaFonte
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Fonte_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Fonte.SelLength = Len(TXT_Fonte.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Fonte_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Tam.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Indice_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Indice.SelLength = Len(TXT_Indice.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Lin_Change()
    On Error GoTo ERRO_SISCOVAL
    If TestaXY(TXT_Lin.Text, "Y") = True Then
        If ModoReposionado = True Then cResp = ReposicionaItem(Val(TXT_Col.Text), Val(TXT_Lin.Text))
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Lin_Click()
    On Error GoTo ERRO_SISCOVAL
    If TestaXY(TXT_Lin.Text, "Y") = True Then
        If ModoReposionado = True Then cResp = ReposicionaItem(Val(TXT_Col.Text), Val(TXT_Lin.Text))
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Lin_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        CB_Item.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Tam_Change()
    On Error GoTo ERRO_SISCOVAL
    MudaFonte
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Tam_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Tam.SelLength = Len(TXT_Tam.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Tam_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        CK_Negrito.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_X_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_X.SelLength = Len(TXT_X.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_X_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Y.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Y_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Y.SelLength = Len(TXT_Y.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Y_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        BT_Salvar.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Function LeImagemNF() As Boolean
    On Error GoTo Erro_ImagemNF
    If DLL_BD.BDSIS_TBINF.RecordCount = 0 Then
        CD_1.DialogTitle = "Indique o caminho da foto"
        CD_1.Filter = "Todos arquivos|*.*"
        CD_1.Flags = cdlOFNExplorer + cdlOFNHideReadOnly + cdlOFNLongNames
        CD_1.InitDir = "C:"
        CD_1.ShowOpen
        If CD_1.FileName <> "" Then
            DLL_BD.BDSIS_TBINF.AddNew
            ArquivoNF = ""
            Open CD_1.FileName For Input As #1 'abre arquivo
            ArquivoNF = Input(LOF(1), 1)
            Close #1
            DLL_BD.BDSIS_TBINF_CPIMG.Value = ArquivoNF
            DLL_BD.BDSIS_TBINF.Update
        End If
    End If
    'Pega diretorio temporario do windows
    DirTmp = DLL_FUNCS.DiretorioTemporario
    'Abre imagem do banco de dados
    DLL_BD.BDSIS_TBINF.MoveFirst
    ArquivoNF = DLL_BD.BDSIS_TBINF_CPIMG.Value
    Open DirTmp & "\scvinf.000" For Output As #1
    Write #1, ArquivoNF
    Close #1
    IMG_1.Picture = LoadPicture(DirTmp & "\scvinf.000")
    Kill (DirTmp & "\scvinf.000")
    LeImagemNF = True
    Exit Function
Erro_ImagemNF:
    MsgBox ("Não foi possível ler a imagem da nota fiscal.")
    LeImagemNF = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Function BotoesEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Valor = True Then
        BT_Cancelar.Enabled = True
        BT_Apagar.Enabled = True
        BT_Salvar.Enabled = True
        BT_Voltar.Enabled = False
        BT_Teste.Enabled = False
        LT_Item.Enabled = False
        BT_Incluir.Enabled = False
        BT_Editar.Enabled = False
        ST_1.TabEnabled(0) = False
        TXT_Descricao.Enabled = True
        TXT_Exemplo.Enabled = True
        TXT_X.Enabled = True
        TXT_Y.Enabled = True
        TXT_Fonte.Enabled = True
        TXT_Tam.Enabled = True
        CK_Negrito.Enabled = True
        CK_Italico.Enabled = True
        CK_Sublinhado.Enabled = True
        BT_Fonte.Enabled = True
        FR_2.Enabled = True
    Else
        BT_Cancelar.Enabled = False
        BT_Apagar.Enabled = False
        BT_Salvar.Enabled = False
        BT_Voltar.Enabled = True
        BT_Teste.Enabled = True
        LT_Item.Enabled = True
        If ST_1.Tab = 1 Then
            BT_Incluir.Enabled = True
            BT_Editar.Enabled = True
        End If
        TXT_Descricao.Enabled = False
        TXT_Exemplo.Enabled = False
        TXT_X.Enabled = False
        TXT_Y.Enabled = False
        ST_1.TabEnabled(0) = True
        TXT_Fonte.Enabled = False
        TXT_Tam.Enabled = False
        CK_Negrito.Enabled = False
        CK_Italico.Enabled = False
        CK_Sublinhado.Enabled = False
        BT_Fonte.Enabled = False
        FR_2.Enabled = False
    End If
    TXT_Nome.Enabled = False
    TXT_Indice.Enabled = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Sub ReposicionaNF(Item As String)
    On Error GoTo ERRO_SISCOVAL
    BT_Apagar.Value = True
    TXT_Nome.Text = ""
    
    DLL_BD.BDSIS_TBCNF.Seek "=", Item
    If DLL_BD.BDSIS_TBCNF.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura deste item no banco de dados para a exibição na tela.")
        Exit Sub
    End If
    TXT_Nome.Text = DLL_BD.BDSIS_TBCNF_CPITE.Value
    TXT_Descricao.Text = DLL_BD.BDSIS_TBCNF_CPDES.Value
    TXT_Exemplo.Text = DLL_BD.BDSIS_TBCNF_CPEXE.Value
    TXT_Indice.Text = Str(DLL_BD.BDSIS_TBCNF_CPIND.Value)
    ModoReposionado = False
    TXT_Col.Text = DLL_BD.BDSIS_TBCNF_CPCOL.Value
    TXT_Lin.Text = DLL_BD.BDSIS_TBCNF_CPLIN.Value
    ModoReposionado = True
    ReposicionaItem Val(TXT_Col.Text), Val(TXT_Lin.Text)
        
    LB_NF.Enabled = True
    LB_NF.Visible = True
    LB_NF.Caption = TXT_Exemplo.Text
    LB_NF.ToolTipText = TXT_Descricao.Text
    LB_NF.FontName = DLL_BD.BDSIS_TBCNF_CPFON.Value
    LB_NF.FontSize = DLL_BD.BDSIS_TBCNF_CPTAM.Value
    LB_NF.FontBold = DLL_BD.BDSIS_TBCNF_CPNEG.Value
    LB_NF.FontItalic = DLL_BD.BDSIS_TBCNF_CPITA.Value
    LB_NF.FontUnderline = DLL_BD.BDSIS_TBCNF_CPSUB.Value
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function ReposicionaItem(nCol As Integer, nLin As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Esta funcao resposiciona o label dentro do frame,
    'dependendo da sua posicao (lin,col), irá movimentar
    'a imagem da nota fiscal dentro do frame
    
    If nLin < (FR_1.Height / 2) Then 'nLin é menor que metade do frame
        IMG_1.Top = 0
        LB_NF.Top = nLin
    ElseIf nLin > (FR_1.Height / 2) And (IMG_1.Height - nLin) >= FR_1.Height / 2 Then 'nLin é maior que metade do frame e exibe a nota toda
        LB_NF.Top = (FR_1.Height / 2) - (LB_NF.Height / 2)
        IMG_1.Top = (FR_1.Height / 2) - nLin - (LB_NF.Height / 2)
    ElseIf nLin > (FR_1.Height / 2) And (IMG_1.Height - nLin) < FR_1.Height / 2 Then 'nLin maior e exibe o fim da nota
        LB_NF.Top = nLin - (IMG_1.Height - FR_1.Height)
        IMG_1.Top = FR_1.Height - IMG_1.Height
    End If
    
    If nCol < (FR_1.Width / 2) Then
        IMG_1.Left = 0
        LB_NF.Left = nCol
    ElseIf nCol > (FR_1.Width / 2) And (IMG_1.Width - nCol) >= FR_1.Width / 2 Then
        LB_NF.Left = (FR_1.Width / 2) - (LB_NF.Width / 2)
        IMG_1.Left = (FR_1.Width / 2) - nCol - (LB_NF.Width / 2)
    ElseIf nCol > (FR_1.Width / 2) And (IMG_1.Width - nCol) < FR_1.Width / 2 Then
        LB_NF.Left = nCol - (IMG_1.Width - FR_1.Width)
        IMG_1.Left = FR_1.Width - IMG_1.Width
    End If
    
    'Marca a posicao da regua
    LB_X1.Caption = Int(Abs(IMG_1.Left))
    LB_Y1.Caption = Int(Abs(IMG_1.Top))
    LB_X2.Caption = Int(Abs(IMG_1.Left)) + FR_1.Width
    LB_Y2.Caption = Int(Abs(IMG_1.Top)) + FR_1.Height
    LB_XM.Caption = ((Int(Abs(IMG_1.Left)) + FR_1.Width) + Int(Abs(IMG_1.Left))) / 2
    LB_YM.Caption = ((Int(Abs(IMG_1.Top)) + FR_1.Height) + Int(Abs(IMG_1.Top))) / 2

    ' Grava os novos dados
    DLL_BD.BDSIS_TBCNF.Seek "=", CB_Item.Text
    If DLL_BD.BDSIS_TBCNF.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura deste item no banco de dados para a exibição na tela.")
        Exit Function
    End If
    DLL_BD.BDSIS_TBCNF.Edit
    DLL_BD.BDSIS_TBCNF_CPCOL.Value = Val(TXT_Col.Text)
    DLL_BD.BDSIS_TBCNF_CPLIN.Value = Val(TXT_Lin.Text)
    DLL_BD.BDSIS_TBCNF.Update
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Function TestaXY(Valor As String, Tipo As String) As Boolean
    On Error GoTo ERRO_SISCOVAL
    TestaXY = False
    
    If Tipo = "X" Then 'É Coluna
        If TXT_Col.Text = "" Then Exit Function
        If Not IsNumeric(TXT_Col.Text) Then
            MsgBox ("Não foi digitado um número válido na caixa de texto de linha.")
            TXT_Col.Text = ""
            TXT_Col.SetFocus
            Exit Function
        End If
        If Val(TXT_Col.Text) > IMG_1.Width Then
            MsgBox ("Foi digitado um número que ultrapassará as margens da nota fiscal. O número para linha deve estar entre 0 e " & Str(IMG_1.Width) & ".")
            TXT_Col.SelLength = Len(Trim(TXT_Col.Text))
            Exit Function
        ElseIf Val(TXT_Col.Text) < 0 Then
            MsgBox ("Foi digitado um número que ultrapassará as margens da nota fiscal. O número para linha deve estar entre 0 e " & Str(IMG_1.Width) & ".")
            TXT_Col.SelLength = Len(Trim(TXT_Col.Text))
            Exit Function
        End If
        If Val(TXT_Col.Text) > (IMG_1.Width - LB_NF.Width) Then
            MsgBox ("Esse valor deixará o texto para fora das margens da nota fiscal.")
            TXT_Col.SetFocus
            Exit Function
        End If
    Else ' É linha
        If TXT_Lin.Text = "" Then Exit Function
        If Not IsNumeric(TXT_Lin.Text) Then
            MsgBox ("Não foi digitado um número válido na caixa de texto de linha.")
            TXT_Lin.Text = ""
            TXT_Lin.SetFocus
            Exit Function
        End If
        If Val(TXT_Lin.Text) > IMG_1.Height Then
            MsgBox ("Foi digitado um número que ultrapassará as margens da nota fiscal. O número para linha deve estar entre 0 e " & Str(IMG_1.Height) & ".")
            TXT_Lin.SelLength = Len(Trim(TXT_Lin.Text))
            Exit Function
        ElseIf Val(TXT_Lin.Text) < 0 Then
            MsgBox ("Foi digitado um número que ultrapassará as margens da nota fiscal. O número para linha deve estar entre 0 e " & Str(IMG_1.Height) & ".")
            TXT_Lin.SelLength = Len(Trim(TXT_Lin.Text))
            Exit Function
        End If
        If Val(TXT_Lin.Text) > (IMG_1.Height - LB_NF.Height) Then
            MsgBox ("Esse valor deixará o texto para fora das margens da nota fiscal.")
            TXT_Lin.SetFocus
            Exit Function
        End If
    End If
    TestaXY = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Sub MudaFonte()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Fonte.Text <> "" Then TXT_Exemplo.FontName = TXT_Fonte.Text
    If TXT_Tam.Text <> "" And Val(TXT_Tam.Text) > 0 Then TXT_Exemplo.FontSize = Val(TXT_Tam.Text)
    TXT_Exemplo.FontBold = CK_Negrito.Value
    TXT_Exemplo.FontItalic = CK_Italico.Value
    TXT_Exemplo.FontUnderline = CK_Sublinhado.Value
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


