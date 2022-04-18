VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form Tela_Cfg_Usuarios 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações de Usuários do Sistema"
   ClientHeight    =   4320
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   7215
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Tela 
      BorderStyle     =   0  'None
      Height          =   4212
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6972
      Begin SysInfoLib.SysInfo SI 
         Left            =   5640
         Top             =   3600
         _ExtentX        =   794
         _ExtentY        =   794
         _Version        =   393216
      End
      Begin VB.CommandButton BT_Novo 
         Caption         =   "&Novo"
         Height          =   732
         Left            =   0
         Picture         =   "Tela_Cfg_Usuarios.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Novo cadastro de empresa"
         Top             =   3360
         Width           =   732
      End
      Begin VB.CommandButton BT_Deletar 
         Caption         =   "&Deletar"
         Height          =   732
         Left            =   1680
         Picture         =   "Tela_Cfg_Usuarios.frx":066A
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Apagar um cadastro de uma empresa"
         Top             =   3360
         Width           =   732
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   6240
         Picture         =   "Tela_Cfg_Usuarios.frx":0AAC
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Volta à Tela Principal."
         Top             =   3360
         Width           =   732
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   732
         Left            =   4800
         Picture         =   "Tela_Cfg_Usuarios.frx":0EEE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancela edição."
         Top             =   3360
         Width           =   732
      End
      Begin VB.CommandButton BT_Salvar 
         Caption         =   "&Salvar"
         Height          =   732
         Left            =   3120
         Picture         =   "Tela_Cfg_Usuarios.frx":11F8
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Salva os dados."
         Top             =   3360
         Width           =   732
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "Apa&gar"
         Height          =   732
         Left            =   3960
         Picture         =   "Tela_Cfg_Usuarios.frx":163A
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Apaga todos os campos."
         Top             =   3360
         Width           =   732
      End
      Begin VB.CommandButton BT_Editar 
         Caption         =   "&Editar"
         Height          =   732
         Left            =   840
         Picture         =   "Tela_Cfg_Usuarios.frx":1A7C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Editar os dados existentes."
         Top             =   3360
         Width           =   732
      End
      Begin TabDlg.SSTab ST_1 
         Height          =   3252
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   6972
         _ExtentX        =   12303
         _ExtentY        =   5741
         _Version        =   393216
         Tab             =   1
         TabHeight       =   420
         TabCaption(0)   =   "Usuários"
         TabPicture(0)   =   "Tela_Cfg_Usuarios.frx":1EBE
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "CK_Adm"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "TXT_Senha2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "TXT_Senha1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "TXT_Logon"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "TXT_Nome"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "LT_Usuarios"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label5"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label4"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label3"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label2"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label1"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).ControlCount=   11
         TabCaption(1)   =   "Permissões"
         TabPicture(1)   =   "Tela_Cfg_Usuarios.frx":1EDA
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label12"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Label13"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "LT_Menus"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "LT_UsuariosMenus"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Computadores"
         TabPicture(2)   =   "Tela_Cfg_Usuarios.frx":1EF6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label10"
         Tab(2).Control(1)=   "Label20"
         Tab(2).Control(2)=   "LT_Computadores"
         Tab(2).Control(3)=   "FR"
         Tab(2).Control(4)=   "TXT_Computador"
         Tab(2).ControlCount=   5
         Begin VB.TextBox TXT_Computador 
            Enabled         =   0   'False
            Height          =   288
            Left            =   -74760
            TabIndex        =   51
            Top             =   2760
            Width           =   1812
         End
         Begin VB.Frame FR 
            Caption         =   "Informações sobre este computador"
            Height          =   2772
            Left            =   -72720
            TabIndex        =   26
            ToolTipText     =   "Informações sobre este computador"
            Top             =   360
            Width           =   4452
            Begin VB.TextBox TXT_TamC 
               Enabled         =   0   'False
               Height          =   288
               Left            =   1560
               TabIndex        =   53
               Top             =   2400
               Width           =   1332
            End
            Begin VB.TextBox TXT_RAM 
               Enabled         =   0   'False
               Height          =   288
               Left            =   120
               TabIndex        =   48
               Top             =   2400
               Width           =   1332
            End
            Begin VB.TextBox TXT_SNC 
               Enabled         =   0   'False
               Height          =   288
               Left            =   3000
               TabIndex        =   47
               Top             =   2400
               Width           =   1332
            End
            Begin VB.TextBox TXT_SisTipProc 
               Enabled         =   0   'False
               Height          =   288
               Left            =   3000
               TabIndex        =   43
               Top             =   1800
               Width           =   1332
            End
            Begin VB.TextBox TXT_SisOEM 
               Enabled         =   0   'False
               Height          =   288
               Left            =   1560
               TabIndex        =   42
               Top             =   1800
               Width           =   1332
            End
            Begin VB.TextBox TXT_SisMascProc 
               Enabled         =   0   'False
               Height          =   288
               Left            =   120
               TabIndex        =   41
               Top             =   1800
               Width           =   1332
            End
            Begin VB.TextBox TXT_OSVersao 
               Enabled         =   0   'False
               Height          =   288
               Left            =   2280
               TabIndex        =   36
               Top             =   1200
               Width           =   972
            End
            Begin VB.TextBox TXT_OSCSD 
               Enabled         =   0   'False
               Height          =   288
               Left            =   1200
               TabIndex        =   35
               Top             =   1200
               Width           =   972
            End
            Begin VB.TextBox TXT_OSBuild 
               Enabled         =   0   'False
               Height          =   288
               Left            =   120
               TabIndex        =   34
               Top             =   1200
               Width           =   972
            End
            Begin VB.TextBox TXT_OSPla 
               Enabled         =   0   'False
               Height          =   288
               Left            =   3360
               TabIndex        =   33
               Top             =   1200
               Width           =   972
            End
            Begin VB.TextBox TXT_DirWin 
               Enabled         =   0   'False
               Height          =   288
               Left            =   120
               TabIndex        =   29
               Top             =   600
               Width           =   1332
            End
            Begin VB.TextBox TXT_DirSys 
               Enabled         =   0   'False
               Height          =   288
               Left            =   1560
               TabIndex        =   28
               Top             =   600
               Width           =   1332
            End
            Begin VB.TextBox TXT_DirTemp 
               Enabled         =   0   'False
               Height          =   288
               Left            =   3000
               TabIndex        =   27
               Top             =   600
               Width           =   1332
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Tamanho C:"
               Height          =   192
               Left            =   1560
               TabIndex        =   54
               Top             =   2160
               Width           =   876
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "RAM Total"
               Height          =   192
               Left            =   120
               TabIndex        =   50
               Top             =   2160
               Width           =   768
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Serial Number C:"
               Height          =   192
               Left            =   3000
               TabIndex        =   49
               Top             =   2160
               Width           =   1212
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Sis - Tip.Proc."
               Height          =   192
               Left            =   3000
               TabIndex        =   46
               Top             =   1560
               Width           =   996
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Sis - OEM"
               Height          =   192
               Left            =   1560
               TabIndex        =   45
               Top             =   1560
               Width           =   708
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Sis - Masc.Proc."
               Height          =   192
               Left            =   120
               TabIndex        =   44
               Top             =   1560
               Width           =   1152
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "OS - CSD"
               Height          =   192
               Left            =   1200
               TabIndex        =   40
               Top             =   960
               Width           =   684
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "OS - Build"
               Height          =   192
               Left            =   120
               TabIndex        =   39
               Top             =   960
               Width           =   708
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "OS - Versão"
               Height          =   192
               Left            =   2280
               TabIndex        =   38
               Top             =   960
               Width           =   876
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               Caption         =   "OS - Plataf."
               Height          =   192
               Left            =   3360
               TabIndex        =   37
               Top             =   960
               Width           =   792
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Dir - Windows"
               Height          =   192
               Left            =   120
               TabIndex        =   32
               Top             =   360
               Width           =   984
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Dir - System"
               Height          =   192
               Left            =   1560
               TabIndex        =   31
               Top             =   360
               Width           =   864
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Dir - Temporário"
               Height          =   192
               Left            =   3000
               TabIndex        =   30
               Top             =   360
               Width           =   1176
            End
         End
         Begin VB.CheckBox CK_Adm 
            Caption         =   "Administrador"
            Height          =   252
            Left            =   -69720
            TabIndex        =   10
            Top             =   960
            Width           =   1452
         End
         Begin VB.ListBox LT_UsuariosMenus 
            Height          =   2010
            ItemData        =   "Tela_Cfg_Usuarios.frx":1F12
            Left            =   600
            List            =   "Tela_Cfg_Usuarios.frx":1F14
            Sorted          =   -1  'True
            TabIndex        =   14
            Top             =   720
            Width           =   1812
         End
         Begin VB.ListBox LT_Menus 
            Height          =   2085
            ItemData        =   "Tela_Cfg_Usuarios.frx":1F16
            Left            =   2760
            List            =   "Tela_Cfg_Usuarios.frx":1F18
            Sorted          =   -1  'True
            Style           =   1  'Checkbox
            TabIndex        =   15
            Top             =   720
            Width           =   3612
         End
         Begin VB.ListBox LT_Computadores 
            Height          =   1620
            ItemData        =   "Tela_Cfg_Usuarios.frx":1F1A
            Left            =   -74760
            List            =   "Tela_Cfg_Usuarios.frx":1F1C
            Sorted          =   -1  'True
            TabIndex        =   16
            Top             =   600
            Width           =   1812
         End
         Begin VB.TextBox TXT_Senha2 
            Height          =   288
            IMEMode         =   3  'DISABLE
            Left            =   -70080
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   13
            Top             =   2400
            Width           =   1452
         End
         Begin VB.TextBox TXT_Senha1 
            Height          =   288
            IMEMode         =   3  'DISABLE
            Left            =   -72000
            MaxLength       =   10
            PasswordChar    =   "*"
            TabIndex        =   12
            Top             =   2400
            Width           =   1452
         End
         Begin VB.TextBox TXT_Logon 
            Height          =   288
            Left            =   -72000
            MaxLength       =   10
            TabIndex        =   9
            Top             =   960
            Width           =   1932
         End
         Begin VB.TextBox TXT_Nome 
            Height          =   288
            Left            =   -72000
            MaxLength       =   50
            TabIndex        =   11
            Top             =   1680
            Width           =   3372
         End
         Begin VB.ListBox LT_Usuarios 
            Height          =   2010
            ItemData        =   "Tela_Cfg_Usuarios.frx":1F1E
            Left            =   -74400
            List            =   "Tela_Cfg_Usuarios.frx":1F20
            Sorted          =   -1  'True
            TabIndex        =   8
            Top             =   720
            Width           =   1812
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Computador"
            Height          =   192
            Left            =   -74760
            TabIndex        =   52
            Top             =   2520
            Width           =   888
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Usuários Cadastrados:"
            Height          =   192
            Left            =   600
            TabIndex        =   25
            Top             =   480
            Width           =   1656
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Menus 'checados' aparecerão:"
            Height          =   192
            Left            =   2760
            TabIndex        =   24
            Top             =   480
            Width           =   2232
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Computadores:"
            Height          =   192
            Left            =   -74760
            TabIndex        =   23
            Top             =   360
            Width           =   1104
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Confirme a senha:"
            Height          =   192
            Left            =   -70080
            TabIndex        =   22
            Top             =   2160
            Width           =   1284
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Digite a senha:"
            Height          =   192
            Left            =   -72000
            TabIndex        =   21
            Top             =   2160
            Width           =   1068
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Logon:"
            Height          =   192
            Left            =   -72000
            TabIndex        =   20
            Top             =   720
            Width           =   492
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nome:"
            Height          =   192
            Left            =   -72000
            TabIndex        =   19
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Usuários Cadastrados:"
            Height          =   192
            Left            =   -74400
            TabIndex        =   18
            Top             =   480
            Width           =   1656
         End
      End
   End
End
Attribute VB_Name = "Tela_Cfg_Usuarios"
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
Const NOMEAPLIC As String = "Configurações de Usuários do Sistema"
Dim LDMenu As String
Dim I, J As Integer
Dim RespMsg, Resp
Dim ModoEdicao As Boolean
Public MeusMenus As Variant
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaCampos
    If ST_1.Tab = 0 Then 'Usuários
        TXT_Logon.SetFocus
    ElseIf ST_1.Tab = 1 Then 'Permissões
        If LT_Menus.ListCount > 0 Then
            For I = 0 To LT_Menus.ListCount - 1
                LT_Menus.Selected(I) = False
            Next I
        End If
        LT_Menus.ListIndex = -1
    ElseIf ST_1.Tab = 2 Then 'Usuários - Computadores
        TXT_Computador.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaCampos
    AtivaCampos (False)
    BT_Voltar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    If ST_1.Tab = 0 Then 'Usuários
        If LT_Usuarios.ListIndex = -1 Then
            MsgBox ("Não foi selecionado nenhum usuários na lista.")
            LT_Usuarios.SetFocus
        End If
        LT_Usuarios.RemoveItem (LT_Usuarios.ListIndex)
        DLL_BD.BDSIS_TBUSU.Delete
    ElseIf ST_1.Tab = 1 Then 'Permissões
        MsgBox ("Nesta tela de permissões você deve somente conceder ou não as permissões de acesso aos programas. Para excluir uma conta usa a aba Usuários.")
        ST_1.SetFocus
        Exit Sub
    ElseIf ST_1.Tab = 2 Then 'Usuários - Computadores
        If LT_Computadores.ListIndex = -1 Then
            MsgBox ("Não foi selecionado nenhum computador na lista.")
            LT_Computadores.SetFocus
        End If
        LT_Computadores.RemoveItem (LT_Computadores.ListIndex)
        DLL_BD.BDSIS_TBUCO.Delete
    End If
    LimpaCampos
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    If ST_1.Tab = 0 Then 'Usuários
        If LT_Usuarios.ListIndex = -1 Then
            MsgBox ("Não foi selecionado nenhum usuários na lista.")
            LT_Usuarios.SetFocus
            Exit Sub
        End If
        AtivaCampos (True)
        ModoEdicao = True
        TXT_Logon.SetFocus
    ElseIf ST_1.Tab = 1 Then 'Permissões
        If LT_UsuariosMenus.ListIndex = -1 Then
            MsgBox ("Não foi selecionado nenhum usuários na lista.")
            LT_UsuariosMenus.SetFocus
            Exit Sub
        End If
        AtivaCampos (True)
        ModoEdicao = True
        LT_Menus.SetFocus
    End If
    ModoEdicao = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    LimpaCampos
    AtivaCampos (True)
    ModoEdicao = False
    If ST_1.Tab = 0 Then 'Usuários
        TXT_Logon.SetFocus
    ElseIf ST_1.Tab = 1 Then 'Permissões
        

    ElseIf ST_1.Tab = 2 Then 'Computadores
        LimpaCampos
        If CarregaComputadores = False Then
            BT_Cancelar.Value = True
            Exit Sub
        End If
        'Procura se já existe o computador
        DLL_BD.BDSIS_TBUCO.Seek "=", TXT_Computador.Text
        If Not DLL_BD.BDSIS_TBUCO.NoMatch Then
            RespMsg = MsgBox("Este computador já está cadastrado. Se existe algum informação errada sobre ele, apague-o e pressione o botão Novo, que o sistema automaticamente pegará as informações atualizadas.", vbInformation + vbOKOnly, "Computador já existe.")
            LimpaCampos
            BT_Cancelar.Value = True
            Exit Sub
        End If
        RespMsg = MsgBox("Você deseja cadastrar o computador " & Trim(TXT_Computador.Text) & " no sistema Siscoval ?", vbQuestion + vbYesNo + vbDefaultButton1, "Cadastrar computador")
        If RespMsg = vbYes Then
            BT_Salvar.Value = True
        Else
            BT_Cancelar.Value = True
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    'On Error GoTo ERRO_SISCOVAL
    ' ********** ABA USUARIOS **********
    If ST_1.Tab = 0 Then
        If TXT_Logon.Text = "" Then
            MsgBox ("É necessário digitar o nome do usuário do sistema.")
            TXT_Logon.SetFocus
            Exit Sub
        ElseIf TXT_Logon.Text <> "" And ModoEdicao = False Then
            For I = 0 To LT_Usuarios.ListCount - 1
                If LT_Usuarios.List(I) = TXT_Logon.Text Then
                    MsgBox ("Este usuário já existe... escolha outro nome.")
                    TXT_Logon.SetFocus
                    Exit Sub
                End If
            Next I
        End If
        If TXT_Nome.Text = "" Then
            MsgBox ("É necessário digitar o nome completo do usuário do sistema.")
            TXT_Nome.SetFocus
            Exit Sub
        ElseIf TXT_Senha1.Text = "" Then
            MsgBox ("É necessário digitar a senha do usuário do sistema.")
            TXT_Senha1.SetFocus
            Exit Sub
        ElseIf TXT_Senha2.Text = "" Then
            MsgBox ("É necessário confirmar a senha do usuário do sistema.")
            TXT_Senha2.SetFocus
            Exit Sub
        ElseIf TXT_Senha1.Text <> TXT_Senha2.Text Then
            MsgBox ("Você confirmou uma senha errada. Digite novamente.")
            TXT_Senha2.SetFocus
            Exit Sub
        End If
        Dim NomeVelho As String
        NomeVelho = LT_Usuarios.Text
        If ModoEdicao = True Then
            DLL_BD.BDSIS_TBUSU.Edit
        Else
            DLL_BD.BDSIS_TBUSU.AddNew
        End If
        DLL_BD.BDSIS_TBUSU_CPUSU.Value = TXT_Logon.Text
        DLL_BD.BDSIS_TBUSU_CPNOM.Value = TXT_Nome.Text
        DLL_BD.BDSIS_TBUSU_CPSEN.Value = TXT_Senha1.Text
        DLL_BD.BDSIS_TBUSU.Update
        DLL_FUNCS.RegistraEvento "Salvar - Configurações de Usuários - Usuário", TXT_Logon.Text
        If ModoEdicao = True And NomeVelho <> TXT_Logon.Text Then
            LT_Usuarios.AddItem (TXT_Logon.Text)
            For I = 0 To LT_Usuarios.ListCount - 1
                If LT_Usuarios.List(I) = NomeVelho Then
                    LT_Usuarios.RemoveItem (I)
                    LT_UsuariosMenus.RemoveItem (I)
                    Exit For
                End If
            Next I
        ElseIf ModoEdicao = False Then
            LT_Usuarios.AddItem (TXT_Logon.Text)
            LT_UsuariosMenus.AddItem (TXT_Logon.Text)
        End If
        LimpaCampos
        AtivaCampos (False)
        LT_Usuarios.SetFocus
    ' ********** ABA PERMISSOES **********
    ElseIf ST_1.Tab = 1 Then
        Dim sMenus As String
        sMenus = ""
        For I = 0 To (LT_Menus.ListCount - 1)
            If LT_Menus.Selected(I) = True Then
                sMenus = sMenus & "1"
            Else
                sMenus = sMenus & "0"
            End If
        Next I
        DLL_BD.BDSIS_TBUSU.Seek "=", LT_UsuariosMenus.Text
        If DLL_BD.BDSIS_TBUSU.NoMatch Then
            MsgBox "Não foi possível localizar dados sobre o usuário - tente mais tarde.", vbCritical + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        DLL_BD.BDSIS_TBUSU.Edit
        DLL_BD.BDSIS_TBUSU_CPPER.Value = sMenus
        DLL_BD.BDSIS_TBUSU.Update
        DLL_FUNCS.RegistraEvento "Salvar - Configurações de Usuários - Permissões", LT_Usuarios.List(LT_Usuarios.ListIndex)
        sMenus = ""
        LimpaCampos
        AtivaCampos (False)
        LT_UsuariosMenus.SetFocus
    ' ********** ABA COMPUTADORES **********
    ElseIf ST_1.Tab = 2 Then
        If TXT_Computador.Text = "" Then
            MsgBox ("É necessário digitar o nome do computador do sistema.")
            TXT_Computador.SetFocus
            Exit Sub
        End If
        DLL_BD.BDSIS_TBUCO.Seek "=", TXT_Computador.Text
        If DLL_BD.BDSIS_TBUCO.NoMatch Then
            DLL_BD.BDSIS_TBUCO.AddNew
        Else
            DLL_BD.BDSIS_TBUCO.Edit
        End If
        DLL_BD.BDSIS_TBUCO_CPCOM = TXT_Computador.Text
        DLL_BD.BDSIS_TBUCO_CPDWI = TXT_DirWin.Text
        DLL_BD.BDSIS_TBUCO_CPDSY = TXT_DirSys.Text
        DLL_BD.BDSIS_TBUCO_CPDTP = TXT_DirTemp.Text
        DLL_BD.BDSIS_TBUCO_CPOBU = TXT_OSBuild.Text
        DLL_BD.BDSIS_TBUCO_CPOCS = TXT_OSCSD.Text
        DLL_BD.BDSIS_TBUCO_CPOVE = TXT_OSVersao.Text
        DLL_BD.BDSIS_TBUCO_CPOPL = TXT_OSPla.Text
        DLL_BD.BDSIS_TBUCO_CPSMP = TXT_SisMascProc.Text
        DLL_BD.BDSIS_TBUCO_CPSOE = TXT_SisOEM.Text
        DLL_BD.BDSIS_TBUCO_CPSTP = TXT_SisTipProc.Text
        DLL_BD.BDSIS_TBUCO_CPRAM = TXT_RAM.Text
        DLL_BD.BDSIS_TBUCO_CPCTA = TXT_TamC.Text
        DLL_BD.BDSIS_TBUCO_CPCSN = TXT_SNC.Text
        DLL_BD.BDSIS_TBUCO.Update
        DLL_FUNCS.RegistraEvento "Salvar - Configurações de Usuários - Computadores", TXT_Computador.Text
        
        LT_Computadores.AddItem (TXT_Computador.Text)
        LimpaCampos
        AtivaCampos (False)
        LT_Computadores.SetFocus
    End If
'ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_Usuarios
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
    DLL_CARGA.Max (11)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Usuários")
    If DLL_BD.AbreTabela_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Usuários - Computadores")
    If DLL_BD.AbreTabela_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Usuários - Menus")
    If DLL_BD.AbreTabela_UsuariosMenus(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Usuários")
    If DLL_BD.AbreCampos_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Usuários - Computadores")
    If DLL_BD.AbreCampos_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Usuários - Menus")
    If DLL_BD.AbreCampos_UsuariosMenus(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega lista de usuários
    DLL_CARGA.CarregaTexto ("Carregando usuários")
    LT_Usuarios.Clear
    LT_UsuariosMenus.Clear
    If DLL_BD.BDSIS_TBUSU.RecordCount > 0 Then
        DLL_BD.BDSIS_TBUSU.MoveFirst
        Do While Not DLL_BD.BDSIS_TBUSU.EOF
            LT_Usuarios.AddItem (DLL_BD.BDSIS_TBUSU_CPUSU.Value)
            LT_UsuariosMenus.AddItem (DLL_BD.BDSIS_TBUSU_CPUSU.Value)
            DLL_BD.BDSIS_TBUSU.MoveNext
        Loop
    End If
    
    'Carrega lista de permissões
    DLL_CARGA.CarregaTexto ("Carregando menus")
    LT_Menus.Clear
    
    If DLL_BD.BDSIS_TBUME.RecordCount > 0 Then
        DLL_BD.BDSIS_TBUME.MoveFirst
        Do While Not DLL_BD.BDSIS_TBUME.EOF
            LT_Menus.AddItem (DLL_BD.BDSIS_TBUME_CPMEN.Value)
            DLL_BD.BDSIS_TBUME.MoveNext
        Loop
    End If
    
    'Carrega lista de computadores
    DLL_CARGA.CarregaTexto ("Carregando Usuários - Computadores")
    LT_Computadores.Clear
    If DLL_BD.BDSIS_TBUCO.RecordCount > 0 Then
        DLL_BD.BDSIS_TBUCO.MoveFirst
        Do While Not DLL_BD.BDSIS_TBUCO.EOF
            LT_Computadores.AddItem (DLL_BD.BDSIS_TBUCO_CPCOM.Value)
            DLL_BD.BDSIS_TBUCO.MoveNext
        Loop
    End If
    
    DLL_FUNCS.RegistraEvento "Abrir Configurações de Usuários", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    AtivaCampos (False)
    LimpaCampos
    ST_1.Tab = 0
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_Usuarios
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_UsuariosMenus(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Computadores_Click()
    If LT_Computadores.ListIndex = -1 Then Exit Sub
    DLL_BD.BDSIS_TBUCO.Seek "=", LT_Computadores.Text
    If DLL_BD.BDSIS_TBUCO.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura do computador no banco de dados. Tente novamente.")
        LT_Computadores.SetFocus
        Exit Sub
    End If
    TXT_Computador.Text = DLL_BD.BDSIS_TBUCO_CPCOM
    TXT_DirWin.Text = DLL_BD.BDSIS_TBUCO_CPDWI
    TXT_DirSys.Text = DLL_BD.BDSIS_TBUCO_CPDSY
    TXT_DirTemp.Text = DLL_BD.BDSIS_TBUCO_CPDTP
    TXT_OSBuild.Text = DLL_BD.BDSIS_TBUCO_CPOBU
    TXT_OSCSD.Text = DLL_BD.BDSIS_TBUCO_CPOCS
    TXT_OSVersao.Text = DLL_BD.BDSIS_TBUCO_CPOVE
    TXT_OSPla.Text = DLL_BD.BDSIS_TBUCO_CPOPL
    TXT_SisMascProc.Text = DLL_BD.BDSIS_TBUCO_CPSMP
    TXT_SisOEM.Text = DLL_BD.BDSIS_TBUCO_CPSOE
    TXT_SisTipProc.Text = DLL_BD.BDSIS_TBUCO_CPSTP
    TXT_RAM.Text = DLL_BD.BDSIS_TBUCO_CPRAM
    TXT_TamC.Text = DLL_BD.BDSIS_TBUCO_CPCTA
    TXT_SNC.Text = DLL_BD.BDSIS_TBUCO_CPCSN
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub

Private Sub LT_Usuarios_Click()
    On Error GoTo ERRO_SISCOVAL
    DLL_BD.BDSIS_TBUSU.Seek "=", LT_Usuarios.Text
    If DLL_BD.BDSIS_TBUSU.NoMatch Then
        MsgBox ("Algum erro ocorreu procurando o usuário na banco de dados.")
        Exit Sub
    End If
    TXT_Logon.Text = DLL_BD.BDSIS_TBUSU_CPUSU.Value
    TXT_Nome.Text = DLL_BD.BDSIS_TBUSU_CPNOM.Value
    TXT_Senha1.Text = DLL_BD.BDSIS_TBUSU_CPSEN.Value
    TXT_Senha2.Text = DLL_BD.BDSIS_TBUSU_CPSEN.Value
    If DLL_BD.BDSIS_TBUSU_CPADM.Value = True Then
        CK_Adm.Value = 1
    Else
        CK_Adm.Value = 0
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_UsuariosMenus_Click()
    On Error GoTo ERRO_SISCOVAL
    DLL_BD.BDSIS_TBUSU.Seek "=", LT_UsuariosMenus.Text
    If DLL_BD.BDSIS_TBUSU.NoMatch Then
        MsgBox ("Algum erro ocorreu procurando o usuário na banco de dados.")
        Exit Sub
    End If
    If LT_Menus.ListCount <> Len(DLL_BD.BDSIS_TBUSU_CPPER.Value) Then
        LimpaCampos
        MsgBox ("Conflitos de valores que estão gravados no banco de dados. Edite novamente.")
        BT_Editar.SetFocus
        Exit Sub
    End If
    For I = 0 To LT_Menus.ListCount - 1
        If Mid(DLL_BD.BDSIS_TBUSU_CPPER.Value, I + 1, 1) = "1" Then
            LT_Menus.Selected(I) = True
        Else
            LT_Menus.Selected(I) = False
        End If
    Next I
    LT_Menus.ListIndex = -1
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub ST_1_Click(PreviousTab As Integer)
    On Error GoTo ERRO_SISCOVAL
    AtivaCampos (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Logon_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Logon.SelLength = Len(TXT_Logon.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Logon_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Nome.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Nome_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Nome.SelLength = Len(TXT_Nome.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Nome_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Senha1.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Senha1_GotFocus()
    TXT_Senha1.SelLength = Len(TXT_Senha1.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Senha1_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Senha2.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Senha2_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Senha2.SelLength = Len(TXT_Senha2.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Senha2_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then BT_Salvar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub AtivaCampos(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    'Usuários
    TXT_Logon.Enabled = Valor
    TXT_Nome.Enabled = Valor
    TXT_Senha1.Enabled = Valor
    TXT_Senha2.Enabled = Valor
    CK_Adm.Enabled = Valor
    TXT_Computador.Enabled = Valor
    
    If ST_1.Tab = 0 Then 'Usuários
        If Valor = True Then
            LT_Usuarios.Enabled = False
            ST_1.TabEnabled(1) = False
            ST_1.TabEnabled(2) = False
        Else
            LT_Usuarios.Enabled = True
            ST_1.TabEnabled(1) = True
            ST_1.TabEnabled(2) = True
        End If
    ElseIf ST_1.Tab = 1 Then 'Permissões
        If Valor = True Then
            LT_UsuariosMenus.Enabled = False
            LT_Menus.Enabled = True
            ST_1.TabEnabled(0) = False
            ST_1.TabEnabled(2) = False
        Else
            LT_UsuariosMenus.Enabled = True
            LT_Menus.Enabled = True
            ST_1.TabEnabled(0) = True
            ST_1.TabEnabled(2) = True
        End If
    ElseIf ST_1.Tab = 2 Then 'Computadores
        
        If Valor = True Then
            ST_1.TabEnabled(0) = False
            ST_1.TabEnabled(1) = False
        Else
            ST_1.TabEnabled(0) = True
            ST_1.TabEnabled(1) = True
        End If
    End If
    'Ativa e desativa botoes
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
    If ST_1.Tab = 1 Then 'Permissões
        BT_Novo.Enabled = False
        BT_Deletar.Enabled = False
        BT_Apagar.Enabled = True
    ElseIf ST_1.Tab = 2 Then 'Computadores
        BT_Editar.Enabled = False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub LimpaCampos()
    On Error GoTo ERRO_SISCOVAL
    If ST_1.Tab = 0 Then 'Usuários
        TXT_Logon.Text = ""
        TXT_Nome.Text = ""
        TXT_Senha1.Text = ""
        TXT_Senha2.Text = ""
        CK_Adm.Value = 0
    ElseIf ST_1.Tab = 1 Then 'Permissões
        For I = 0 To LT_Menus.ListCount - 1
            LT_Menus.Selected(I) = False
        Next I
    ElseIf ST_1.Tab = 2 Then 'Computadores
        TXT_Computador.Text = ""
        TXT_DirWin.Text = ""
        TXT_DirSys.Text = ""
        TXT_DirTemp.Text = ""
        TXT_OSBuild.Text = ""
        TXT_OSCSD.Text = ""
        TXT_OSVersao.Text = ""
        TXT_OSPla.Text = ""
        TXT_SisMascProc.Text = ""
        TXT_SisOEM.Text = ""
        TXT_SisTipProc.Text = ""
        TXT_RAM.Text = ""
        TXT_TamC.Text = ""
        TXT_SNC.Text = ""
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function CarregaComputadores() As Boolean
    On Error GoTo ERRO_DLL:
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    TXT_Computador.Text = DLL_FUNCS.PegaNomeComputador
    TXT_DirWin.Text = DLL_FUNCS.DiretorioWindows
    TXT_DirSys.Text = DLL_FUNCS.DiretorioSystem
    TXT_DirTemp.Text = DLL_FUNCS.DiretorioTemporario
    TXT_OSBuild.Text = DLL_FUNCS.OS_Build
    TXT_OSCSD.Text = DLL_FUNCS.OS_CSDVersao
    TXT_OSVersao.Text = DLL_FUNCS.OS_NumVersao
    TXT_OSPla.Text = DLL_FUNCS.OS_Plataforma
    TXT_SisMascProc.Text = DLL_FUNCS.SIS_MascaraProcessador
    TXT_SisOEM.Text = DLL_FUNCS.SIS_OEMID
    TXT_SisTipProc.Text = DLL_FUNCS.SIS_TipoProcessador
    TXT_RAM.Text = DLL_FUNCS.MEM_TotRAM
    TXT_TamC.Text = DLL_FUNCS.PegaTamanhoC
    TXT_SNC.Text = DLL_FUNCS.PegaSerialC
    Set DLL_FUNCS = Nothing
    CarregaComputadores = True
    Exit Function
ERRO_DLL:
    MsgBox ("Não foi possível acessar a biblioteca para carregar dados sobre este computador.")
    CarregaComputadores = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
