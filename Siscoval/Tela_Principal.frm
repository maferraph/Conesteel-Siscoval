VERSION 5.00
Begin VB.Form Tela_Principal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Siscoval"
   ClientHeight    =   645
   ClientLeft      =   30
   ClientTop       =   570
   ClientWidth     =   3645
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   3645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox ICONE_ABRIR 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1200
      Picture         =   "Tela_Principal.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      ToolTipText     =   "Registrar usuário"
      Top             =   240
      Width           =   240
   End
   Begin VB.FileListBox LA 
      Height          =   285
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.PictureBox ICONE_INICIAR 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   240
      Picture         =   "Tela_Principal.frx":0102
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   3
      ToolTipText     =   "Iniciar o Sistema Siscoval"
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox ICONE_SISCOVAL 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   720
      Picture         =   "Tela_Principal.frx":040C
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      ToolTipText     =   "Sistema Siscoval"
      Top             =   120
      Width           =   480
   End
   Begin VB.PictureBox ICONE_SAIR 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1560
      Picture         =   "Tela_Principal.frx":0716
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      ToolTipText     =   "Encerra o Sistema Siscoval"
      Top             =   240
      Width           =   240
   End
   Begin VB.PictureBox ICONE_LOGON 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1920
      Picture         =   "Tela_Principal.frx":0818
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      ToolTipText     =   "Registrar usuário"
      Top             =   240
      Width           =   240
   End
   Begin VB.Menu MENU_LOGON 
      Caption         =   "MENU_LOGON"
      Begin VB.Menu Menu_Iniciar 
         Caption         =   "&Iniciar"
      End
      Begin VB.Menu lixo_1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Cancelar 
         Caption         =   "&Sair"
      End
   End
   Begin VB.Menu MENU_ICONE 
      Caption         =   "MENU_ICONE"
      Begin VB.Menu Menu_Abrir 
         Caption         =   "&Abrir"
      End
      Begin VB.Menu lixo_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Sair 
         Caption         =   "&Sair"
      End
   End
End
Attribute VB_Name = "Tela_Principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** VARIÁVEIS DLL's ****************
Dim DLL_BD As Scvbd.Classe_Scvbd
Dim DLL_SCVFUNC  As Scvfunc.Classe_Scvfunc

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Sistema Siscoval"
Dim I As Integer
Private Sub Form_Load()
    On Error GoTo ERRO_INICIO
    Screen.MousePointer = vbHourglass
    
    If App.PrevInstance Then End
    Tela_Principal.Visible = False
    lTempo = False
    'procura o bd aberto
'    If VerificaBD = False Then End
    'Insere figuras no menu logon
    Dim HWND_MENU As Long, lRet As Long, HWND_ITEMMENU As Long, HWND_SUBMENU As Long
    HWND_MENU = GetMenu(Tela_Principal.hwnd) 'Pega handle do menu da tela
    HWND_ITEMMENU = GetSubMenu(HWND_MENU, 0)
    lRet = SetMenuItemBitmaps(HWND_ITEMMENU, 0, MF_BYPOSITION, ICONE_LOGON.Picture, ICONE_LOGON.Picture)
    lRet = SetMenuItemBitmaps(HWND_ITEMMENU, 2, MF_BYPOSITION, ICONE_SAIR.Picture, ICONE_SAIR.Picture)
    HWND_ITEMMENU = GetSubMenu(HWND_MENU, 1)
    lRet = SetMenuItemBitmaps(HWND_ITEMMENU, 0, MF_BYPOSITION, ICONE_ABRIR.Picture, ICONE_ABRIR.Picture)
    lRet = SetMenuItemBitmaps(HWND_ITEMMENU, 2, MF_BYPOSITION, ICONE_SAIR.Picture, ICONE_SAIR.Picture)
    
    'Desenha icone na barra de tarefas do windows
    With IconeTela
        .cbSize = Len(IconeTela)
        .hwnd = ICONE_INICIAR.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = ICONE_INICIAR.Picture
        .szTip = ICONE_INICIAR.ToolTipText & vbNullChar
        .Tela = "LOGON"
    End With
    Shell_NotifyIcon NIM_ADD, IconeTela
    Screen.MousePointer = vbNormal
    Exit Sub
ERRO_INICIO:
    MsgBox "Ocorreu algum erro na inicialização do Siscoval.", vbCritical + vbOKOnly, NOMEAPLIC
    End
End Sub
Private Sub ICONE_INICIAR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ERRO_SISCOVAL
    If IconeTela.Tela = "LOGON" Then
        IconeLogon (X)
    ElseIf IconeTela.Tela = "SISCOVAL" Then
        IconeSiscoval (X)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub ICONE_SISCOVAL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ERRO_SISCOVAL
    If IconeTela.Tela = "LOGON" Then
        IconeLogon (X)
    ElseIf IconeTela.Tela = "SISCOVAL" Then
        IconeSiscoval (X)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Abrir_Click()
    On Error GoTo ERRO_SISCOVAL
    'remove icone da barra de tarefas
    Shell_NotifyIcon NIM_DELETE, IconeTela
    'Abre tela siscoval
    Tela_Siscoval.Top = 0
    Tela_Siscoval.Left = 0
    Tela_Siscoval.Width = Screen.Width
    Tela_Siscoval.Height = 1300
    Tela_Siscoval.Visible = True
    lTempo = True
    DLL_FUNCS.RegistraEvento "Inicialização de Sistema", ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    'Encerra sistema
    Shell_NotifyIcon NIM_DELETE, IconeTela
    End
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Iniciar_Click()
    On Error GoTo ERRO_SISCOVAL
    'verifica licensa
    'If VerificaRegistro = False Then End
    'If VerificaComputador = False Then End
    Tela_Acesso.Show vbModal
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Sair_Click()
    On Error GoTo ERRO_SISCOVAL
    'Encerra sistema
    Shell_NotifyIcon NIM_DELETE, IconeTela
    End
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Sub IconeLogon(X)
    On Error GoTo ERRO_SISCOVAL
    'abre menu
    If ICONE_INICIAR.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_RBUTTONUP 'Botão direito pressionado
            PopupMenu Tela_Principal.MENU_LOGON
        Case WM_LBUTTONUP 'Botão esquerdo pressionado
            PopupMenu Tela_Principal.MENU_LOGON
    End Select
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub IconeSiscoval(X)
    On Error GoTo ERRO_SISCOVAL
    If ICONE_SISCOVAL.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_RBUTTONUP 'Botão direito pressionado
            PopupMenu Tela_Principal.MENU_ICONE
        Case WM_LBUTTONUP 'Botão esquerdo pressionado
            PopupMenu Tela_Principal.MENU_ICONE
    End Select
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function VerificaRegistro() As Boolean
    On Error GoTo ERRO_REGISTRO
    'Verifica diretorio do EXE
    LA.Path = GetSetting("Siscoval", "Diretório", "Dir_EXE")
    If LA.ListCount > 0 Then
        For I = 0 To LA.ListCount - 1
            If VBA.UCase(LA.List(I)) = VBA.UCase("Siscoval.exe") Then
                Exit For
            ElseIf VBA.UCase(LA.List(I)) <> VBA.UCase("siscoval.exe") And I = LA.ListCount - 1 Then
                GoTo ERRO_REGISTRO
            End If
        Next I
        Else: GoTo ERRO_REGISTRO
    End If
    'Verifica diretorio do banco de dados do Siscoval
    LA.Path = GetSetting("Siscoval", "Diretório", "Dir_BD")
    If LA.ListCount > 0 Then
        For I = 0 To LA.ListCount - 1
            If VBA.UCase(LA.List(I)) = VBA.UCase("Siscoval.scv") Then
                Exit For
            ElseIf VBA.UCase(LA.List(I)) <> VBA.UCase("Siscoval.scv") And I = LA.ListCount - 1 Then
                GoTo ERRO_REGISTRO
            End If
        Next I
        Else: GoTo ERRO_REGISTRO
    End If
    
    VerificaRegistro = True
    Exit Function
ERRO_REGISTRO:
    VerificaRegistro = False
    RespMsg = MsgBox("Erro de inicialização do Siscoval.", vbOKOnly + vbCritical, "Carga")
End Function
Private Static Function VerificaComputador() As Boolean
    On Error GoTo ERRO_SISCOVAL
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_SCVFUNC = New Scvfunc.Classe_Scvfunc
    'Abre BD
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreTabela_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreCampos_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Procura dados sobre o computador
    DLL_BD.BDSIS_TBUCO.Seek "=", DLL_SCVFUNC.PegaNomeComputador
    If DLL_BD.BDSIS_TBUCO.NoMatch Then
        GoTo ERRO_ACESSO_BANCODADOS
    Else
        'Verifica dados
        If DLL_BD.BDSIS_TBUCO_CPDWI = DLL_SCVFUNC.DiretorioWindows And _
           DLL_BD.BDSIS_TBUCO_CPDSY = DLL_SCVFUNC.DiretorioSystem And _
           DLL_BD.BDSIS_TBUCO_CPDTP = DLL_SCVFUNC.DiretorioTemporario And _
           DLL_BD.BDSIS_TBUCO_CPOBU = DLL_SCVFUNC.OS_Build And _
           DLL_BD.BDSIS_TBUCO_CPOCS = DLL_SCVFUNC.OS_CSDVersao And _
           DLL_BD.BDSIS_TBUCO_CPOVE = DLL_SCVFUNC.OS_NumVersao And _
           DLL_BD.BDSIS_TBUCO_CPOPL = DLL_SCVFUNC.OS_Plataforma And _
           DLL_BD.BDSIS_TBUCO_CPSMP = DLL_SCVFUNC.SIS_MascaraProcessador And _
           DLL_BD.BDSIS_TBUCO_CPSOE = DLL_SCVFUNC.SIS_OEMID And _
           DLL_BD.BDSIS_TBUCO_CPSTP = DLL_SCVFUNC.SIS_TipoProcessador And _
           DLL_BD.BDSIS_TBUCO_CPRAM = DLL_SCVFUNC.MEM_TotRAM And _
           DLL_BD.BDSIS_TBUCO_CPCTA = DLL_SCVFUNC.PegaTamanhoC And _
           DLL_BD.BDSIS_TBUCO_CPCSN = DLL_SCVFUNC.PegaSerialC Then
            VerificaComputador = True
        Else
            GoTo ERRO_ACESSO_BANCODADOS
       End If
    End If
    'Fecha banco de dados
    If DLL_BD.FechaTabela_UsuariosComputadores(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    Set DLL_SCVFUNC = Nothing
    Set DLL_BD = Nothing
    Exit Function
ERRO_ACESSO_BANCODADOS:
    'Computador não permitido
    VerificaComputador = False
    RespMsg = MsgBox("Este computador não está licensiado.", vbOKOnly + vbCritical, "Falta licensa")
    Dim DB, DE, DW
    DE = GetSetting("Siscoval", "Diretório", "Dir_EXE")
    Kill (DE & "\Siscoval.exe")
    DB = GetSetting("Siscoval", "Diretório", "Dir_BD")
    DW = DLL_SCVFUNC.DiretorioSystem
    'Move arquico do banco de dados
    Dim XXX, ZZZ As String
    'ZZZ = "xcopy " & Trim(DB) & "\Siscoval.scv " & Trim(DW) & " /A /Y /U"
    'XXX = Shell(ZZZ, vbHide)
    'Kill (Trim(DB) & "\Siscoval.scv")
    'ZZZ = DLL_SCVFUNC.MoveArquivo(VBA.Trim(DB) & "\Siscoval.scv", VBA.Trim(DW))
    'MsgBox ZZZ
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Function VerificaBD() As Boolean
    VerificaBD = False
    On Error GoTo ERRO_SISCOVAL
    Set DLL_BD = New Scvbd.Classe_Scvbd
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_SISCOVAL
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    Set DLL_BD = Nothing
    VerificaBD = True
    Exit Function
ERRO_SISCOVAL:
    MsgBox "Não foi possível abrir o banco de dados do Siscoval. Verifique se o computador onde está o banco de dados está ligado e/ou funcionando.", vbCritical + vbOKOnly, NOMEAPLIC
    VerificaBD = False
End Function
