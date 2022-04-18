VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Tela_Empresas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cadastros das Empresas"
   ClientHeight    =   4815
   ClientLeft      =   30
   ClientTop       =   165
   ClientWidth     =   7695
   ControlBox      =   0   'False
   Icon            =   "Tela_Empresas.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Cadastro 
      Caption         =   "Ca&dastro"
      Enabled         =   0   'False
      Height          =   732
      Left            =   3360
      Picture         =   "Tela_Empresas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3960
      Visible         =   0   'False
      Width           =   732
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   732
      Left            =   5880
      Picture         =   "Tela_Empresas.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Cancela edição."
      Top             =   3960
      Width           =   732
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "Apa&gar"
      Height          =   732
      Left            =   5160
      Picture         =   "Tela_Empresas.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3960
      Width           =   732
   End
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   732
      Left            =   2280
      Picture         =   "Tela_Empresas.frx":0D60
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Imprimir ficha de cadastro de uma empresa."
      Top             =   3960
      Width           =   732
   End
   Begin VB.CommandButton BT_Deletar 
      Caption         =   "&Deletar"
      Height          =   732
      Left            =   1560
      Picture         =   "Tela_Empresas.frx":13CA
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Apagar um cadastro de uma empresa"
      Top             =   3960
      Width           =   732
   End
   Begin VB.CommandButton BT_Editar 
      Caption         =   "&Editar"
      Height          =   732
      Left            =   840
      Picture         =   "Tela_Empresas.frx":180C
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Editar os cadastros de empresas existentes."
      Top             =   3960
      Width           =   732
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados sobre a Empresa:"
      Height          =   3852
      Left            =   2280
      TabIndex        =   25
      Top             =   0
      Width           =   5292
      Begin VB.ComboBox CB_Trans 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Tela_Empresas.frx":1C4E
         Left            =   3360
         List            =   "Tela_Empresas.frx":1C50
         Sorted          =   -1  'True
         TabIndex        =   44
         ToolTipText     =   "Se esta empresa usar alguma transportadora em especial, selecione nesta lista."
         Top             =   3470
         Width           =   1815
      End
      Begin VB.TextBox TXT_Apelido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   120
         MaxLength       =   15
         TabIndex        =   43
         Top             =   480
         Width           =   1332
      End
      Begin VB.TextBox TXT_Fax 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   3240
         MaxLength       =   16
         TabIndex        =   13
         Top             =   2880
         Width           =   1932
      End
      Begin VB.TextBox TXT_Observacao 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   15
         Top             =   3480
         Width           =   1575
      End
      Begin VB.TextBox TXT_Comentario 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   120
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox TXT_Empresa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   2
         Top             =   480
         Width           =   3612
      End
      Begin VB.TextBox TXT_Endereco 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1680
         Width           =   3132
      End
      Begin VB.TextBox TXT_Cidade 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   9
         Top             =   2280
         Width           =   1692
      End
      Begin VB.ComboBox CB_Estado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         ItemData        =   "Tela_Empresas.frx":1C52
         Left            =   4320
         List            =   "Tela_Empresas.frx":1C54
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2280
         Width           =   852
      End
      Begin VB.TextBox TXT_Fone 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1200
         MaxLength       =   16
         TabIndex        =   12
         Top             =   2880
         Width           =   1932
      End
      Begin VB.TextBox TXT_InsEst 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   4
         Top             =   1080
         Width           =   1692
      End
      Begin VB.ComboBox CB_Tipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         ItemData        =   "Tela_Empresas.frx":1C56
         Left            =   3720
         List            =   "Tela_Empresas.frx":1C58
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   1452
      End
      Begin VB.TextBox TXT_PracaPagamento 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   120
         MaxLength       =   100
         TabIndex        =   8
         ToolTipText     =   "Se a praça de pagamento for a mesmo do endereço, deixe este campo em branco."
         Top             =   2280
         Width           =   2292
      End
      Begin VB.TextBox TXT_Bairro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   3360
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1680
         Width           =   1812
      End
      Begin MSMask.MaskEdBox TXT_CGC 
         Height          =   288
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1692
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         ForeColor       =   -2147483630
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   18
         Mask            =   "99.999.999/9999-99"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_CEP 
         Height          =   288
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         ForeColor       =   -2147483630
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   9
         Mask            =   "99999-999"
         PromptChar      =   "_"
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Transportadora:"
         Height          =   195
         Left            =   3360
         TabIndex        =   45
         Top             =   3240
         Width           =   1125
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   192
         Left            =   3240
         TabIndex        =   40
         Top             =   2640
         Width           =   300
      End
      Begin VB.Label LB_Observacao 
         AutoSize        =   -1  'True
         Caption         =   "Observações:"
         Height          =   195
         Left            =   1680
         TabIndex        =   39
         Top             =   3240
         Width           =   1020
      End
      Begin VB.Label LB_Comentarios 
         AutoSize        =   -1  'True
         Caption         =   "Comentários de N.F.:"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   3240
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   192
         Left            =   1560
         TabIndex        =   37
         Top             =   240
         Width           =   696
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   192
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   744
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   192
         Left            =   3360
         TabIndex        =   35
         Top             =   1440
         Width           =   468
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   192
         Left            =   2520
         TabIndex        =   34
         Top             =   2040
         Width           =   564
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   192
         Left            =   4320
         TabIndex        =   33
         Top             =   2040
         Width           =   552
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   192
         Left            =   120
         TabIndex        =   32
         Top             =   2640
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fone:"
         Height          =   192
         Left            =   1200
         TabIndex        =   31
         Top             =   2640
         Width           =   408
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   1920
         TabIndex        =   30
         Top             =   840
         Width           =   1356
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "C.N.P.J.:"
         Enabled         =   0   'False
         Height          =   192
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   192
         Left            =   3720
         TabIndex        =   28
         Top             =   840
         Width           =   372
      End
      Begin VB.Label LB_Praca 
         AutoSize        =   -1  'True
         Caption         =   "Praça de Pagamento:"
         Height          =   192
         Left            =   120
         TabIndex        =   27
         Top             =   2040
         Width           =   1572
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia:"
         Height          =   192
         Left            =   120
         TabIndex        =   26
         Top             =   240
         Width           =   1140
      End
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   6840
      Picture         =   "Tela_Empresas.frx":1C5A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3960
      Width           =   732
   End
   Begin VB.CommandButton BT_Novo 
      Caption         =   "&Novo"
      Height          =   732
      Left            =   120
      Picture         =   "Tela_Empresas.frx":209C
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Novo cadastro de empresa"
      Top             =   3960
      Width           =   732
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   732
      Left            =   4440
      Picture         =   "Tela_Empresas.frx":2706
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3960
      Width           =   732
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresas"
      Height          =   3852
      Left            =   120
      TabIndex        =   24
      Top             =   0
      Width           =   2052
      Begin VB.ComboBox CB_Exibir 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         ItemData        =   "Tela_Empresas.frx":2B48
         Left            =   120
         List            =   "Tela_Empresas.frx":2B4A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1812
      End
      Begin VB.ListBox LT_Apelido 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2940
         ItemData        =   "Tela_Empresas.frx":2B4C
         Left            =   120
         List            =   "Tela_Empresas.frx":2B4E
         TabIndex        =   1
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label LB_Exibir 
         AutoSize        =   -1  'True
         Height          =   192
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   36
      End
   End
End
Attribute VB_Name = "Tela_Empresas"
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
Dim ModoEdicao As Boolean
Dim RespMsg, NomeFantasia
Dim I As Integer
Const NOMEAPLIC As String = "Cadastro de Empresas"
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaTextoTelaEmpresa
    TXT_Empresa.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cadastro_Click()
    'DLL_FUNCS.RegistraEvento "Consulta Cadastro de Empresa", ""
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaTextoTelaEmpresa
    AtivaTelaEmpresa (False)
    AtivaBotoesEmEdicao (False)
    LT_Apelido.ListIndex = -1
    ModoEdicao = False
    Tela_Empresas.Refresh
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Apelido.ListIndex = -1 Then
        MsgBox "Selecione alguma empresa na lista.", vbInformation + vbOKOnly, NOMEAPLIC
        LT_Apelido.SetFocus
        Exit Sub
    End If
    RespMsg = MsgBox("Deseja deletar esta empresa do Banco de Dados ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Deletar Empresa")
    DLL_BD.BDSIS_TBEMP.Seek "=", LT_Apelido.Text
    If DLL_BD.BDSIS_TBEMP.NoMatch Then
        MsgBox "Erro ao procurar a empresa no banco de dados.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    If RespMsg = vbYes Then
        DLL_BD.BDSIS_TBEMP.Delete
        DLL_FUNCS.RegistraEvento "Deletar - Cadastro de Empresas", LT_Apelido.Text
        LT_Apelido.RemoveItem (LT_Apelido.ListIndex)
        LimpaTextoTelaEmpresa
    Else
        Tela_Empresas.Refresh
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Apelido.ListIndex = -1 Then
        RespMsg = MsgBox("Selecione alguma empresa na lista.")
        LT_Apelido.SetFocus
        Exit Sub
    End If
    DLL_BD.BDSIS_TBEMP.Seek "=", UCase(LT_Apelido.Text)
    If Not DLL_BD.BDSIS_TBEMP.NoMatch Then
        ModoEdicao = True
        AtivaTelaEmpresa (True)
        AtivaBotoesEmEdicao (True)
        TXT_Apelido.Text = UCase(LT_Apelido)
        TXT_Apelido.Enabled = False
        TXT_Empresa.SetFocus
    Else
        Beep
        MsgBox "Erro ao procurar esta empresa no banco de dados.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Imprimir_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Apelido.Text = "" Then
        RespMsg = MsgBox("Selecione alguma empresa na lista.")
        LT_Apelido.SetFocus
        Exit Sub
    End If
    'Tela_Empresas_IT.LB_Apelido.Caption = DLL_BD.BDSIS_TBEMP_CPAPE
    'Tela_Empresas_IT.LB_Empresa.Caption = DLL_BD.BDSIS_TBEMP_CPEMP
    'Tela_Empresas_IT.LB_CGC.Caption = DLL_BD.BDSIS_TBEMP_CPCGC
    'Tela_Empresas_IT.LB_InsEst.Caption = DLL_BD.BDSIS_TBEMP_CPINE
    'Tela_Empresas_IT.LB_Endereco.Caption = DLL_BD.BDSIS_TBEMP_CPEND
    'Tela_Empresas_IT.LB_Bairro.Caption = DLL_BD.BDSIS_TBEMP_CPBAI
    'Tela_Empresas_IT.LB_Cidade.Caption = DLL_BD.BDSIS_TBEMP_CPCID
    'Tela_Empresas_IT.LB_Estado.Caption = DLL_BD.BDSIS_TBEMP_CPEST
    'Tela_Empresas_IT.LB_Cep.Caption = DLL_BD.BDSIS_TBEMP_CPCEP
    'Tela_Empresas_IT.LB_Fone.Caption = DLL_BD.BDSIS_TBEMP_CPFON
    'Tela_Empresas_IT.LB_Tipo.Caption = DLL_BD.BDSIS_TBEMP_CPTIP
    'RespMsg = MsgBox("Deseja imprimir cadastro desta empresa ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Imprimir cadastro")
    'If RespMsg = vbYes Then
    '    Tela_Empresas_IT.PrintForm
    'End If
    DLL_FUNCS.RegistraEvento "Imprimir Cadastro de Empresa", ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    LT_Apelido.ListIndex = -1
    LimpaTextoTelaEmpresa
    AtivaTelaEmpresa (True)
    AtivaBotoesEmEdicao (True)
    ModoEdicao = False
    TXT_Apelido.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Apelido.Text = "" Then
        MsgBox "Campo nome fantasia deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_Apelido.SetFocus
        Exit Sub
    ElseIf TXT_Empresa.Text = "" Then
        MsgBox "Campo empresas deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_Empresa.SetFocus
        Exit Sub
    ElseIf TXT_CGC.Text = "" Or TXT_CGC.Text = "__.___.___/____-__" Then
        MsgBox "Campo CGC deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_CGC.SetFocus
        Exit Sub
    ElseIf TXT_InsEst.Text = "" Then
        MsgBox "Campo Inscrição Estadual deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_InsEst.SetFocus
        Exit Sub
    ElseIf CB_Tipo.ListIndex = -1 Then
        MsgBox "Campo tipo deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Tipo.SetFocus
        Exit Sub
    ElseIf TXT_Endereco.Text = "" Then
        MsgBox "Este campo deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_Endereco.SetFocus
        Exit Sub
    ElseIf TXT_Cidade.Text = "" Then
        MsgBox "Campo cidade deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_Cidade.SetFocus
        Exit Sub
    ElseIf CB_Estado.ListIndex = -1 Then
        MsgBox "Campo estado deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        CB_Estado.SetFocus
        Exit Sub
    ElseIf TXT_CEP.Text = "" Or TXT_CEP.Text = "     -   " Then
        MsgBox "Este campo deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_CEP.SetFocus
        Exit Sub
    ElseIf TXT_Fone.Text = "" Then
        MsgBox "Este campo deve ser preenchido.", vbInformation + vbOKOnly, NOMEAPLIC
        TXT_Fone.SetFocus
        Exit Sub
    End If
    
    If ModoEdicao = True Then
        DLL_BD.BDSIS_TBEMP.Edit
    Else
        LT_Apelido.AddItem (TXT_Apelido.Text)
        DLL_BD.BDSIS_TBEMP.AddNew
    End If
                       
    DLL_BD.BDSIS_TBEMP_CPAPE.Value = TXT_Apelido.Text
    DLL_BD.BDSIS_TBEMP_CPEMP.Value = TXT_Empresa.Text
    DLL_BD.BDSIS_TBEMP_CPCGC.Value = TXT_CGC.Text
    DLL_BD.BDSIS_TBEMP_CPINE.Value = TXT_InsEst.Text
    DLL_BD.BDSIS_TBEMP_CPEND.Value = TXT_Endereco.Text
    If TXT_Bairro.Text <> "" Then DLL_BD.BDSIS_TBEMP_CPBAI.Value = TXT_Bairro.Text
    If TXT_PracaPagamento.Text <> "" Then DLL_BD.BDSIS_TBEMP_CPPRA.Value = TXT_PracaPagamento.Text
    DLL_BD.BDSIS_TBEMP_CPCID.Value = TXT_Cidade.Text
    DLL_BD.BDSIS_TBEMP_CPEST.Value = Trim(CB_Estado.Text)
    DLL_BD.BDSIS_TBEMP_CPCEP.Value = TXT_CEP.Text
    DLL_BD.BDSIS_TBEMP_CPFON.Value = TXT_Fone.Text
    If CB_Tipo.Text <> "" Then DLL_BD.BDSIS_TBEMP_CPTIP.Value = (CB_Tipo.Text)
    If TXT_Comentario.Text <> "" Then DLL_BD.BDSIS_TBEMP_CPCOM.Value = TXT_Comentario.Text
    If TXT_Observacao.Text <> "" Then DLL_BD.BDSIS_TBEMP_CPOBS.Value = TXT_Observacao.Text
    If CB_Trans.Text <> "" Then DLL_BD.BDSIS_TBEMP_CPTRA.Value = CB_Trans.Text
    DLL_BD.BDSIS_TBEMP.Update
    LimpaTextoTelaEmpresa
    AtivaTelaEmpresa (False)
    AtivaBotoesEmEdicao (False)
    LT_Apelido.ListIndex = -1
    ModoEdicao = False
    BT_Voltar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Estado_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_CEP.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Exibir_Click()
    On Error GoTo ERRO_SISCOVAL
    LT_Apelido.Clear
    DLL_BD.BDSIS_TBEMP.MoveFirst
    If CB_Exibir.List(CB_Exibir.ListIndex) = "Todos" Then 'Todos
        While Not DLL_BD.BDSIS_TBEMP.EOF
            If DLL_BD.BDSIS_TBEMP_CPAPE.Value <> "" Then
                LT_Apelido.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            DLL_BD.BDSIS_TBEMP.MoveNext
        Wend
    Else
        While Not DLL_BD.BDSIS_TBEMP.EOF
            If DLL_BD.BDSIS_TBEMP_CPTIP = CB_Exibir.List(CB_Exibir.ListIndex) Then
                LT_Apelido.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            DLL_BD.BDSIS_TBEMP.MoveNext
        Wend
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Tipo_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Endereco.SetFocus
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
    DLL_CARGA.CarregaTexto ("Abrindo tabela Empresas...")
    If DLL_BD.AbreTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Empresas...")
    If DLL_BD.AbreCampos_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    AtivaTelaEmpresa (False)
    AtivaBotoesEmEdicao (False)
    LimpaTextoTelaEmpresa
    
    'Carregando combo de tipo de empresas
    DLL_CARGA.CarregaTexto ("Carregando lista de tipos de empresas e estados...")
    CB_Exibir.Clear
    DLL_BD.BDSIS_TBGRU.MoveFirst
    Do While Not DLL_BD.BDSIS_TBGRU.EOF
        If DLL_BD.BDSIS_TBGRU_CPTIP.Value = "EMP" Then
            CB_Exibir.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
            CB_Tipo.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
        ElseIf DLL_BD.BDSIS_TBGRU_CPTIP.Value = "EST" Then
            CB_Estado.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
        End If
        DLL_BD.BDSIS_TBGRU.MoveNext
    Loop
    'seleciona todas empresas
    For I = 0 To CB_Exibir.ListCount - 1
        If CB_Exibir.List(I) = "Todos" Then
            CB_Exibir.ListIndex = I
            Exit For
        End If
    Next I
    'carrega transportadoras
    CB_Trans.Clear
    DLL_BD.BDSIS_TBEMP.MoveFirst
    While Not DLL_BD.BDSIS_TBEMP.EOF
        If DLL_BD.BDSIS_TBEMP_CPTIP = "Transportadora" Then
            CB_Trans.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
        End If
        DLL_BD.BDSIS_TBEMP.MoveNext
    Wend
    If EMPRESA <> "" Then LT_Apelido.Text = EMPRESA
    DLL_FUNCS.RegistraEvento "Abrir Cadastro de Empresas", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Empresas
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Empresas
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Terminate()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Empresas
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Apelido_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Apelido.ListIndex = -1 Then
        Exit Sub
    End If
    LimpaTextoTelaEmpresa
    EMPRESA = LT_Apelido.Text
    DLL_BD.BDSIS_TBEMP.Seek "=", LT_Apelido.Text
    If DLL_BD.BDSIS_TBEMP.NoMatch Then
        RespMsg = MsgBox("Ocorreu erro durante a procura do nome fantasia da empresa.")
        Exit Sub
    Else
        TXT_Apelido.Text = DLL_BD.BDSIS_TBEMP_CPAPE.Value
        If DLL_BD.BDSIS_TBEMP_CPEMP <> "" Then
            TXT_Empresa.Text = DLL_BD.BDSIS_TBEMP_CPEMP
        Else
            TXT_Empresa.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCGC <> "" Then
            TXT_CGC.Text = DLL_BD.BDSIS_TBEMP_CPCGC
        Else
            TXT_CGC.Text = "__.___.___/____-__"
        End If
        If DLL_BD.BDSIS_TBEMP_CPINE <> "" Then
            TXT_InsEst.Text = DLL_BD.BDSIS_TBEMP_CPINE
        Else
            TXT_InsEst.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPEND <> "" Then
            TXT_Endereco.Text = DLL_BD.BDSIS_TBEMP_CPEND
        Else
            TXT_Endereco.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPBAI <> "" Then
            TXT_Bairro.Text = DLL_BD.BDSIS_TBEMP_CPBAI
        Else
            TXT_Bairro.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPPRA <> "" Then
            TXT_PracaPagamento.Text = DLL_BD.BDSIS_TBEMP_CPPRA
        Else
            TXT_PracaPagamento.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCID <> "" Then
            TXT_Cidade.Text = DLL_BD.BDSIS_TBEMP_CPCID
        Else
            TXT_Cidade.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCEP <> "" Then
            TXT_CEP.Text = DLL_BD.BDSIS_TBEMP_CPCEP
        Else
            TXT_CEP.Text = "_____-___"
        End If
        If DLL_BD.BDSIS_TBEMP_CPFON <> "" Then
            TXT_Fone.Text = DLL_BD.BDSIS_TBEMP_CPFON
        Else
            TXT_Fone.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCOM.Value <> "" Then
            TXT_Comentario.Text = DLL_BD.BDSIS_TBEMP_CPCOM.Value
        Else
            TXT_Comentario.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPOBS <> "" Then
            TXT_Observacao.Text = DLL_BD.BDSIS_TBEMP_CPOBS
        Else
            TXT_Observacao.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPTRA <> "" Then
            CB_Trans.Text = DLL_BD.BDSIS_TBEMP_CPTRA.Value
        Else
            CB_Trans.Text = ""
        End If
        CB_Estado.ListIndex = -1
        CB_Tipo.ListIndex = -1
        Do While Not DLL_BD.BDSIS_TBEMP.NoMatch
            For I = 0 To CB_Estado.ListCount - 1
                If DLL_BD.BDSIS_TBEMP_CPEST.Value = CB_Estado.List(I) Then
                    CB_Estado.ListIndex = I
                    Exit For
                End If
            Next
            For I = 0 To CB_Tipo.ListCount - 1
                If DLL_BD.BDSIS_TBEMP_CPTIP.Value = CB_Tipo.List(I) Then
                    CB_Tipo.ListIndex = I
                    Exit For
                End If
            Next
            DLL_BD.BDSIS_TBEMP.MoveNext
            If CB_Estado.ListCount >= 0 And _
                CB_Tipo.ListCount >= 0 Then
                Exit Do
            End If
        Loop
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Apelido_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then TXT_Empresa.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Apelido_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Apelido.Text = "" Then Exit Sub
    DLL_BD.BDSIS_TBEMP.Seek "=", TXT_Apelido.Text
    If DLL_BD.BDSIS_TBEMP.NoMatch Then
        TXT_Empresa.SetFocus
    ElseIf Not DLL_BD.BDSIS_TBEMP.NoMatch And ModoEdicao = False Then
        Beep
        MsgBox ("Já existe esse nome fantasia cadastrado... tente novamente.")
        TXT_Apelido.SetFocus
        Exit Sub
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Bairro_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Bairro.SelLength = Len(Trim(TXT_Bairro.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Bairro_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_PracaPagamento.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CEP_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_CEP.SelLength = Len(Trim(TXT_CEP.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CEP_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Fone.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CGC_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_CGC.SelLength = Len(Trim(TXT_CGC.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CGC_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_InsEst.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Cidade_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Cidade.SelLength = Len(Trim(TXT_Cidade.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Cidade_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Estado.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Comentario_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Comentario.SelLength = Len(Trim(TXT_Comentario.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Comentario_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Observacao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Empresa_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Empresa.SelLength = Len(Trim(TXT_Empresa.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Empresa_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_CGC.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Endereco_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Endereco.SelLength = Len(Trim(TXT_Endereco.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Endereco_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Bairro.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Fax_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Comentario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Fone_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Fone.SelLength = Len(Trim(TXT_Fone.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Fone_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Fax.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_InsEst_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_InsEst.SelLength = Len(Trim(TXT_InsEst.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_InsEst_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Tipo.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Observacao_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Observacao.SelLength = Len(Trim(TXT_Observacao.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Observacao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Salvar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PracaPagamento_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PracaPagamento.SelLength = Len(Trim(TXT_PracaPagamento.Text))
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PracaPagamento_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Cidade.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub AtivaTelaEmpresa(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    ' Valor = True -> Habilita todos controles
    Frame2.Enabled = Valor
    Label1.Enabled = Valor
    Label2.Enabled = Valor
    Label3.Enabled = Valor
    Label4.Enabled = Valor
    Label5.Enabled = Valor
    Label6.Enabled = Valor
    Label7.Enabled = Valor
    Label8.Enabled = Valor
    Label9.Enabled = Valor
    Label10.Enabled = Valor
    Label11.Enabled = Valor
    Label12.Enabled = Valor
    Label13.Enabled = Valor
    TXT_Apelido.Enabled = Valor
    LB_Praca.Enabled = Valor
    LB_Comentarios.Enabled = Valor
    LB_Observacao.Enabled = Valor
    TXT_Empresa.Enabled = Valor
    If Valor = False Then
        TXT_CGC.ForeColor = &H80000011
        TXT_CEP.ForeColor = &H80000011
    Else
        TXT_CGC.ForeColor = &H80000012
        TXT_CEP.ForeColor = &H80000012
    End If
    TXT_InsEst.Enabled = Valor
    TXT_Endereco.Enabled = Valor
    TXT_PracaPagamento.Enabled = Valor
    TXT_Bairro.Enabled = Valor
    TXT_Cidade.Enabled = Valor
    TXT_Fone.Enabled = Valor
    TXT_Observacao.Enabled = Valor
    TXT_Comentario.Enabled = Valor
    CB_Estado.Enabled = Valor
    CB_Tipo.Enabled = Valor
    CB_Trans.Enabled = Valor
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub LimpaTextoTelaEmpresa()
    On Error GoTo ERRO_SISCOVAL
    TXT_Apelido.Text = ""
    TXT_Empresa.Text = ""
    TXT_CGC.Text = "__.___.___/____-__"
    TXT_InsEst.Text = ""
    TXT_Endereco.Text = ""
    TXT_Bairro.Text = ""
    TXT_PracaPagamento.Text = ""
    TXT_Cidade.Text = ""
    TXT_CEP.Text = "_____-___"
    TXT_Fone.Text = ""
    TXT_Fax.Text = ""
    TXT_Comentario.Text = ""
    TXT_Observacao.Text = ""
    CB_Estado.ListIndex = -1
    CB_Tipo.ListIndex = -1
    CB_Trans.ListIndex = -1
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub AtivaBotoesEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    ' Valor = True -> Habilita todos controles
    If Valor = True Then
        BT_Novo.Enabled = False
        BT_Editar.Enabled = False
        BT_Deletar.Enabled = False
        BT_Imprimir.Enabled = False
        BT_Voltar.Enabled = False
        BT_Salvar.Enabled = True
        BT_Apagar.Enabled = True
        BT_Cancelar.Enabled = True
        LT_Apelido.Enabled = False
        Frame1.Enabled = False
        CB_Exibir.Enabled = False
    Else
        BT_Novo.Enabled = True
        BT_Editar.Enabled = True
        BT_Deletar.Enabled = True
        BT_Imprimir.Enabled = False
        BT_Voltar.Enabled = True
        BT_Salvar.Enabled = False
        BT_Apagar.Enabled = False
        BT_Cancelar.Enabled = False
        LT_Apelido.Enabled = True
        Frame1.Enabled = True
        CB_Exibir.Enabled = True
    End If
    BT_Cadastro.Enabled = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_Empresas.MousePointer = vbHourglass
        Tela_Empresas.Enabled = False
    Else
        Tela_Empresas.MousePointer = vbDefault
        Tela_Empresas.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
