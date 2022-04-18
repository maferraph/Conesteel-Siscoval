VERSION 5.00
Begin VB.Form Tela_Entrada 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000E&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   30
   ClientTop       =   30
   ClientWidth     =   10035
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "Tela_Entrada.frx":0000
   ScaleHeight     =   5775
   ScaleWidth      =   10035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LT_Menus 
      Height          =   285
      ItemData        =   "Tela_Entrada.frx":BFF82
      Left            =   720
      List            =   "Tela_Entrada.frx":BFF84
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer TM 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label LB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LB"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1800
      TabIndex        =   0
      Top             =   4440
      Width           =   6375
   End
End
Attribute VB_Name = "Tela_Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** VARIÁVEIS DLL's ****************
Dim DLL_BD As Scvbd.Classe_Scvbd

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Entrada"
Dim I As Integer, nOp As Integer
Dim RespMsg
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Na tela de entrada ficará todos procedimentos
    'necessários para verificação do sistema.
    Screen.MousePointer = vbHourglass
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    'Abre bancos de dados
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    If DLL_BD.AbreTabela_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    If DLL_BD.AbreCampos_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    LT_Menus.Visible = False
    TM.Interval = 1
    nOp = 1
    LB.Caption = ""
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    EncerraSistema
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Screen.MousePointer = vbNormal
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TM_Timer()
    On Error GoTo ERRO_SISCOVAL
    IniciaSistema (nOp)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub IniciaSistema(Operacao As Integer)
    On Error GoTo ERRO_INICIASISTEMA
    TM.Interval = 500
    If Operacao = 1 Then 'Insere icone na barra do windows
        LB.Caption = "Inserindo ícone na barra de tarefas do windows..."
        IconeTela.Tela = "SISCOVAL"
        IconeTela.hIcon = Tela_Principal.ICONE_SISCOVAL.Picture
        IconeTela.szTip = Tela_Principal.ICONE_SISCOVAL.ToolTipText & vbNullChar
        Shell_NotifyIcon NIM_MODIFY, IconeTela
    ElseIf Operacao = 2 Then 'Monta menu pop-up
        LB.Caption = "Montando menus do sistema..."
        If MontaMenu = False Then GoTo ERRO_INICIASISTEMA
    ElseIf Operacao = 3 Then 'Monta menu para o usuario
        LB.Caption = "Cadastrando usuário no sistema..."
        If MontaMenuUsuario = False Then GoTo ERRO_INICIASISTEMA
    ElseIf Operacao = 4 Then 'Finalizando
        LB.Caption = "Finalizando..."
    Else 'Fecha esta tela
        Unload Tela_Entrada
        Tela_Siscoval.Visible = False
        Tela_Principal.Visible = False
        Screen.MousePointer = vbNormal
    End If
    nOp = nOp + 1
    Exit Sub
ERRO_INICIASISTEMA:
    RespMsg = MsgBox("Ocorreu algum erro durante o carregamento do sistema.", vbCritical + vbOKOnly, "Erro de abertura")
    EncerraSistema
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub EncerraSistema()
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Usuarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    'Encerra sistema
    Shell_NotifyIcon NIM_DELETE, IconeTela
    Screen.MousePointer = vbNormal
    End
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function MontaMenu() As Boolean
    On Error GoTo ERRO_SISCOVAL
    MontaMenu = False
    Tela_Siscoval.Visible = False
    Dim HWND_MENU As Long, lRet As Long, HWND_ITEMMENU As Long, HWND_SUBMENU As Long
    HWND_MENU = GetMenu(Tela_Siscoval.hwnd)
    HWND_ITEMMENU = GetSubMenu(HWND_MENU, 0)
    lRet = SetMenuItemBitmaps(HWND_ITEMMENU, 7, MF_BYPOSITION, Tela_Principal.ICONE_ABRIR.Picture, Tela_Principal.ICONE_ABRIR.Picture)
    lRet = SetMenuItemBitmaps(HWND_ITEMMENU, 9, MF_BYPOSITION, Tela_Principal.ICONE_SAIR.Picture, Tela_Principal.ICONE_SAIR.Picture)
    
    MontaMenu = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Function MontaMenuUsuario() As Boolean
    On Error GoTo ERRO_MONTAMENUUSUARIO
    DLL_BD.BDSIS_TBUSU.Seek "=", Usuario
    If DLL_BD.BDSIS_TBUSU.NoMatch Then GoTo ERRO_MONTAMENUUSUARIO
    Dim CCC As Control
    If DLL_BD.BDSIS_TBUSU_CPADM.Value = True Then 'Se for administrador
        For Each CCC In Tela_Siscoval
            If TypeOf CCC Is Menu Then
                CCC.Enabled = True
            End If
        Next CCC
    Else
        Dim sVM As String
        sVM = DLL_BD.BDSIS_TBUSU_CPPER.Value
        'carrega lista de menus
        LT_Menus.Clear
        For Each CCC In Tela_Siscoval
            If TypeOf CCC Is Menu Then
                LT_Menus.AddItem CCC.Name
            End If
        Next CCC
        For I = 0 To (LT_Menus.ListCount - 1)
            If VBA.Mid(sVM, (I + 1), 1) = 1 Then
                LT_Menus.Selected(I) = True
            Else
                LT_Menus.Selected(I) = False
            End If
        Next I
        'Carrega menus
        For I = 0 To (LT_Menus.ListIndex - 1)
            If LT_Menus.Selected(I) = True Then
                CVM VBA.Trim(LT_Menus.List(I)), True
            Else
                CVM VBA.Trim(LT_Menus.List(I)), False
            End If
        Next I
        With Tela_Siscoval
            .BF.Buttons(4).Enabled = .Menu_Estoque_ConsultaRápida.Enabled 'BT Consulta
            .BF.Buttons(5).Enabled = .Menu_Escritorio_Cotacoes.Enabled 'BT Cotação
            .BF.Buttons(6).Enabled = .Menu_Escritorio_Pedidos.Enabled 'BT Pedidos
            .BF.Buttons(7).Enabled = .Menu_Escritorio_NotaFiscal_Assistente.Enabled 'BT NF
            .BF.Buttons(8).Enabled = .Menu_Escritorio_Certificado.Enabled 'BT CQ
            .BF.Buttons(9).Enabled = .Menu_Escritorio_CadastrosEmpresas.Enabled 'BT Empresas
        End With
    End If
    
    MontaMenuUsuario = True
    Exit Function
ERRO_MONTAMENUUSUARIO:
    MsgBox Err.Description
    MontaMenuUsuario = False
End Function
Private Static Sub CVM(NomeMenu As String, Habilitado As Boolean)
    Dim MeuControle As Control
    For Each MeuControle In Tela_Siscoval.Controls
        If TypeOf MeuControle Is Menu Then
            If MeuControle.Name = NomeMenu Then
                If MeuControle.Caption <> "-" Then
                    MeuControle.Enabled = Habilitado
                End If
            End If
        End If
    Next MeuControle
End Sub
