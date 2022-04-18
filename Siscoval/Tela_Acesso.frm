VERSION 5.00
Begin VB.Form Tela_Acesso 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acesso de usuário"
   ClientHeight    =   1590
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   2805
   ControlBox      =   0   'False
   Icon            =   "Tela_Acesso.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_Tela 
      Caption         =   "Digite o nome e senha do usuário:"
      Height          =   1572
      Left            =   0
      TabIndex        =   4
      ToolTipText     =   "Digite o nome e senha do usuário para efetuar o logon."
      Top             =   0
      Width           =   2772
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   372
         Left            =   1440
         TabIndex        =   3
         ToolTipText     =   "Cancela Logon."
         Top             =   1080
         Width           =   972
      End
      Begin VB.CommandButton BT_Iniciar 
         Caption         =   "&Iniciar"
         Height          =   372
         Left            =   360
         TabIndex        =   2
         ToolTipText     =   "Entra no sistema Siscoval."
         Top             =   1080
         Width           =   972
      End
      Begin VB.TextBox TXT_Senha 
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   720
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Digite a senha do usuário."
         Top             =   720
         Width           =   1932
      End
      Begin VB.TextBox TXT_Nome 
         Height          =   288
         Left            =   720
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   "Digite o nome do usuário."
         Top             =   360
         Width           =   1932
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Senha:"
         Height          =   192
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   504
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Nome:"
         Height          =   192
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "Tela_Acesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NOMEAPLIC As String = "Acesso de Usuário"

Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    EncerraSistema
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Iniciar_Click()
    On Error GoTo ERRO_SISCOVAL
    NumeroTentativasLogon = NumeroTentativasLogon + 1
    VerificaTentativas
    If TXT_Nome.Text = "" Then
        RespMsg = MsgBox("É necessário digitar o nome e senha do usuário do sistema Siscoval para o programa efetuar o logon.", vbInformation + vbOKOnly, "Nome ou senha não digitado")
        TXT_Nome.SetFocus
        Exit Sub
    ElseIf TXT_Senha.Text = "" Then
        RespMsg = MsgBox("É necessário digitar o nome e senha do usuário do sistema Siscoval para o programa efetuar o logon.", vbInformation + vbOKOnly, "Nome ou senha não digitado")
        TXT_Senha.SetFocus
        Exit Sub
    End If
    Usuario = TXT_Nome.Text
    DLL_BD.BDSIS_TBUSU.Seek "=", Usuario
    If DLL_BD.BDSIS_TBUSU.NoMatch Then
        RespMsg = MsgBox("Este usuário ou senha não conferem. Tente novamente.", vbInformation + vbOKOnly, "Erro de acesso")
        TXT_Nome.SetFocus
        Exit Sub
    Else
        If DLL_BD.BDSIS_TBUSU_CPUSU.Value = TXT_Nome.Text And _
           DLL_BD.BDSIS_TBUSU_CPSEN.Value = TXT_Senha.Text Then 'Aqui efetuou o logon
            'Registra usuario
            DLL_FUNCS.RegistraUsuario TXT_Nome.Text, DLL_FUNCS.PegaNomeComputador
            DLL_FUNCS.RegistraEvento "Logon de Sistema", TXT_Nome.Text
            Unload Tela_Acesso
            Tela_Entrada.Show vbModal
        Else
            RespMsg = MsgBox("Este usuário ou senha não conferem. Tente novamente.", vbInformation + vbOKOnly, "Erro de acesso")
            TXT_Nome.SetFocus
            Exit Sub
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    'Abre banco de dados
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    If DLL_BD.AbreTabela_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreTabela_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    If DLL_BD.AbreCampos_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreCampos_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    NumeroTentativasLogon = 0
    Usuario = ""
    'pega nome do ultimo usuario
    TXT_Nome.Text = GetSetting("Siscoval", "Logon", "Nome")
'******************************************************
'TXT_Nome.Text = "Admin"
'TXT_Senha.Text = "gangazumba"
'******************************************************
    Exit Sub

ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("(Acesso) Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    EncerraSistema
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Nome_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Nome.SelLength = Len(TXT_Nome.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Nome_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Senha.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Senha_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Senha.SelLength = Len(TXT_Senha.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Senha_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Iniciar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub EncerraSistema()
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
    'Encerra sistema
    Shell_NotifyIcon NIM_DELETE, IconeTela
    End
End Sub
Private Static Sub VerificaTentativas()
    On Error GoTo ERRO_SISCOVAL
    If NumeroTentativasLogon >= 3 Then
        RespMsg = MsgBox("Você alcançou o número máximo de tentativas. Tente mais tarde.", vbCritical + vbOKOnly, "Falha de acesso")
        EncerraSistema
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ProtegeSistema()
    On Error GoTo ERRO_SISCOVAL
    'Encerra o sistema
    EncerraSistema
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
