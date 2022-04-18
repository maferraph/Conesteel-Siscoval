VERSION 5.00
Begin VB.Form Tela_MudarSenha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mudar a Senha"
   ClientHeight    =   3240
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   3030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame F2 
      Caption         =   "Digite e confirme a nova senha:"
      Height          =   1212
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Digite o nome e senha do usuário para efetuar o logon."
      Top             =   1440
      Width           =   2772
      Begin VB.TextBox TXT_NS2 
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Redigite para confirmar a nova senha"
         Top             =   720
         Width           =   1812
      End
      Begin VB.TextBox TXT_NS1 
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Nova senha"
         Top             =   360
         Width           =   1812
      End
      Begin VB.Label L4 
         AutoSize        =   -1  'True
         Caption         =   "Confirme:"
         Height          =   192
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   672
      End
      Begin VB.Label L3 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   192
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   504
      End
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Cancelar"
      Height          =   372
      Left            =   1920
      TabIndex        =   5
      ToolTipText     =   "Cancela alteração de senha."
      Top             =   2760
      Width           =   972
   End
   Begin VB.CommandButton BT_Alterar 
      Caption         =   "&Alterar"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Altera a senha atual"
      Top             =   2760
      Width           =   972
   End
   Begin VB.Frame F1 
      Caption         =   "Digite o nome e senha atuais:"
      Height          =   1212
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Digite o nome e senha do usuário para efetuar o logon."
      Top             =   120
      Width           =   2772
      Begin VB.TextBox TXT_Nome 
         Height          =   288
         Left            =   840
         MaxLength       =   15
         TabIndex        =   0
         ToolTipText     =   "Nome do usuário"
         Top             =   360
         Width           =   1812
      End
      Begin VB.TextBox TXT_Senha 
         Height          =   288
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Senha atual"
         Top             =   720
         Width           =   1812
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         Caption         =   "Nome:"
         Height          =   192
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   480
      End
      Begin VB.Label L2 
         AutoSize        =   -1  'True
         Caption         =   "Senha:"
         Height          =   192
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   504
      End
   End
End
Attribute VB_Name = "Tela_MudarSenha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NOMEAPLIC As String = "Mudar Senhas"

Private Sub BT_Alterar_Click()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Nome.Text = "" Then
        MsgBox "É necessário digitar o nome do usuário.", vbInformation + vbOKOnly, "Nome não digitado"
        TXT_Nome.SetFocus
        Exit Sub
    ElseIf TXT_Senha.Text = "" Then
        MsgBox "É necessário digitar a senha atual do usuário.", vbInformation + vbOKOnly, "Senha atual não digitada"
        TXT_Senha.SetFocus
        Exit Sub
    ElseIf TXT_NS1.Text = "" Then
        MsgBox "É necessário digitar a nova senha do usuário.", vbInformation + vbOKOnly, "Nova senha não digitada"
        TXT_NS1.SetFocus
        Exit Sub
    ElseIf TXT_NS2.Text = "" Then
        MsgBox "É necessário confirmar a nova senha do usuário.", vbInformation + vbOKOnly, "Nova senha não digitada"
        TXT_NS2.SetFocus
        Exit Sub
    ElseIf TXT_NS1.Text <> TXT_NS2.Text Then
        MsgBox "A nova senha digitada é diferente da confirmação da nova senha. Digite novamente.", vbInformation + vbOKOnly, "Nova senha não confere"
        TXT_NS1.Text = ""
        TXT_NS2.Text = ""
        TXT_NS1.SetFocus
        Exit Sub
    End If
    DLL_BD.BDSIS_TBUSU.Seek "=", TXT_Nome.Text
    If DLL_BD.BDSIS_TBUSU.NoMatch Then
        MsgBox "Este usuário ou senha não conferem. Tente novamente.", vbInformation + vbOKOnly, "Erro de acesso"
        TXT_Nome.SetFocus
        Exit Sub
    Else
        If DLL_BD.BDSIS_TBUSU_CPSEN.Value = TXT_Senha.Text Then
            DLL_BD.BDSIS_TBUSU.Edit
            DLL_BD.BDSIS_TBUSU_CPSEN.Value = TXT_NS1.Text
            DLL_BD.BDSIS_TBUSU.Update
            DLL_FUNCS.RegistraEvento "Alteração de Senhas", TXT_Nome.Text
            BT_Cancelar.Value = True
            Exit Sub
        Else
            MsgBox "A senha atual não é válida. Digite novamente.", vbInformation + vbOKOnly, "Erro de acesso"
            TXT_Senha.SetFocus
            Exit Sub
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_MudarSenha
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    TXT_Nome.Text = Usuario
    TXT_Senha.Text = ""
    TXT_NS1.Text = ""
    TXT_NS2.Text = ""
    DLL_FUNCS.RegistraEvento "Abrir Alteração de Senhas", ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Nome_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Senha.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NS1_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_NS2.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NS2_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Alterar.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Senha_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_NS1.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
