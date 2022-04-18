VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Cfg_Estoque_PrecoPeso 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configurações dos Preços e Pesos dos Ítens de Estoque"
   ClientHeight    =   4215
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   4695
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Apaga 
      Caption         =   "Apaga lista"
      Height          =   372
      Left            =   3720
      TabIndex        =   10
      Top             =   1800
      Width           =   852
   End
   Begin VB.CheckBox CK_Incluir 
      Caption         =   "Incluir preços de peças equivalentes"
      Height          =   192
      Left            =   1800
      TabIndex        =   11
      Top             =   2280
      Width           =   2892
   End
   Begin VB.CommandButton BT_Remove 
      Caption         =   "Remove da lista"
      Height          =   372
      Left            =   2760
      TabIndex        =   9
      Top             =   1800
      Width           =   852
   End
   Begin VB.CommandButton BT_Inclui 
      Caption         =   "Inclui na lista"
      Height          =   372
      Left            =   1800
      TabIndex        =   8
      Top             =   1800
      Width           =   852
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   960
      Picture         =   "Tela_Cfg_Estoque_PrecoPeso.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Volta à tela principal."
      Top             =   3360
      Width           =   732
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "&Salvar"
      Height          =   732
      Left            =   120
      Picture         =   "Tela_Cfg_Estoque_PrecoPeso.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salva informações."
      Top             =   3360
      Width           =   732
   End
   Begin VB.ListBox LT_Figuras 
      Height          =   1230
      Left            =   1800
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   2760
      Width           =   2772
   End
   Begin VB.ComboBox CB_Bitolas 
      Height          =   288
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1440
      Width           =   2772
   End
   Begin VB.ComboBox CB_Materiais 
      Height          =   288
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   2772
   End
   Begin VB.ComboBox CB_Figuras 
      Height          =   288
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   2772
   End
   Begin VB.Frame FR_1 
      Height          =   1092
      Left            =   120
      TabIndex        =   17
      Top             =   0
      Width           =   1572
      Begin MSComctlLib.ProgressBar BP_1 
         Height          =   132
         Left            =   240
         TabIndex        =   23
         Top             =   840
         Width           =   1092
         _ExtentX        =   1931
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.OptionButton RB_Preco 
         Caption         =   "Preço"
         Height          =   252
         Left            =   360
         TabIndex        =   0
         Top             =   240
         Width           =   732
      End
      Begin VB.OptionButton RB_Peso 
         Caption         =   "Peso"
         Height          =   252
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   732
      End
   End
   Begin VB.Frame FR_2 
      Caption         =   "Exibir figuras por:"
      Height          =   1092
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   1572
      Begin VB.OptionButton RB_Indice 
         Caption         =   "Por índice"
         Height          =   252
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1092
      End
      Begin VB.ComboBox CB_Indice 
         Height          =   288
         ItemData        =   "Tela_Cfg_Estoque_PrecoPeso.frx":0884
         Left            =   120
         List            =   "Tela_Cfg_Estoque_PrecoPeso.frx":0886
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Lista de índices de figuras."
         Top             =   720
         Width           =   1332
      End
      Begin VB.OptionButton RB_Todas 
         Caption         =   "Todas"
         Height          =   252
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   732
      End
   End
   Begin VB.Frame FR_3 
      Caption         =   "Preço"
      Height          =   852
      Left            =   120
      TabIndex        =   21
      Top             =   2280
      Width           =   1572
      Begin MSMask.MaskEdBox TXT_1 
         Height          =   372
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   "###,##0.00"
         PromptChar      =   "_"
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Lista de fichas que serão alteradas"
      Height          =   192
      Left            =   1800
      TabIndex        =   22
      Top             =   2520
      Width           =   2508
   End
   Begin VB.Label LB_Bitolas 
      AutoSize        =   -1  'True
      Caption         =   "Bitolas:"
      Height          =   192
      Left            =   1800
      TabIndex        =   20
      Top             =   1200
      Width           =   528
   End
   Begin VB.Label LB_Materiais 
      AutoSize        =   -1  'True
      Caption         =   "Materiais:"
      Height          =   192
      Left            =   1800
      TabIndex        =   19
      Top             =   600
      Width           =   696
   End
   Begin VB.Label LB_Figuras 
      AutoSize        =   -1  'True
      Caption         =   "Figuras:"
      Height          =   192
      Left            =   1800
      TabIndex        =   18
      Top             =   0
      Width           =   576
   End
End
Attribute VB_Name = "Tela_Cfg_Estoque_PrecoPeso"
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
Const NOMEAPLIC As String = "Configurações dos Preços e Pesos dos Ítens de Estoque"
Dim I, J As Integer
Dim RespMsg, cResp
Dim ModoEdicao, lTeste As Boolean
Private Sub BT_Apaga_Click()
    On Error GoTo ERRO_SISCOVAL
    LT_Figuras.Clear
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Apaga_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Inclui_Click()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figuras.ListIndex = -1 Then
        MsgBox ("Não foi selecionada nenhuma figura.")
        CB_Figuras.SetFocus
        Exit Sub
    ElseIf CB_Materiais.ListIndex = -1 Then
        MsgBox ("Não foi selecionado nenhum material.")
        CB_Materiais.SetFocus
        Exit Sub
    ElseIf CB_Bitolas.ListIndex = -1 Then
        MsgBox ("Não foi selecionada nenhuma bitola.")
        CB_Bitolas.SetFocus
        Exit Sub
    End If
    PrecoPesoEmEspera (True)
    If CK_Incluir.Value = 0 Then
        LT_Figuras.AddItem (CB_Figuras.Text & " em " & CB_Materiais.Text & " de " & CB_Bitolas.Text)
    Else
        Dim cIndice, cClasse As String
        If RB_Preco.Value = True Then
            'Procura grupo de classe desta figura para incluir
            'na lista de figuras peças de mesmo índice de figura
            'e mesma classe de pressão
            DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figuras.Text
            If DLL_BD.BDSIS_TBEFG.NoMatch Then
                MsgBox ("Ocorreu algum erro durante a procura da figura no banco de dados de estoque.")
                PrecoPesoEmEspera (False)
                Exit Sub
            End If
            cIndice = DLL_BD.BDSIS_TBEFG_CPIFG.Value
            cClasse = DLL_BD.BDSIS_TBEFG_CPGCL.Value
            For I = 0 To CB_Figuras.ListCount - 1
                BP_1.Max = CB_Figuras.ListCount
                BP_1.Value = 0
                DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figuras.List(I)
                If DLL_BD.BDSIS_TBEFG.NoMatch Then
                    Exit For
                Else
                    If DLL_BD.BDSIS_TBEFG_CPIFG.Value = cIndice And _
                       DLL_BD.BDSIS_TBEFG_CPGCL.Value = cClasse Then
                        LT_Figuras.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value & " em " & CB_Materiais.Text & " de " & CB_Bitolas.Text)
                    End If
                End If
            Next I
        ElseIf RB_Peso.Value = True Then
            'Procura grupo de classe desta figura para incluir
            'na lista de figuras peças de mesmo índice de figura
            'e mesma classe de pressão e mesmo material
            DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figuras.Text
            If DLL_BD.BDSIS_TBEFG.NoMatch Then
                MsgBox ("Ocorreu algum erro durante a procura da figura no banco de dados de estoque.")
                PrecoPesoEmEspera (False)
                Exit Sub
            End If
            cIndice = DLL_BD.BDSIS_TBEFG_CPIFG.Value
            cClasse = DLL_BD.BDSIS_TBEFG_CPGCL.Value
            For I = 0 To CB_Figuras.ListCount - 1
                BP_1.Max = CB_Figuras.ListCount
                BP_1.Value = 0
                DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figuras.List(I)
                If DLL_BD.BDSIS_TBEFG.NoMatch Then
                    Exit For
                Else
                    If DLL_BD.BDSIS_TBEFG_CPIFG.Value = cIndice And _
                       DLL_BD.BDSIS_TBEFG_CPGCL.Value = cClasse Then
                        For J = 0 To CB_Materiais.ListCount - 1
                            LT_Figuras.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value & " em " & CB_Materiais.List(J) & " de " & CB_Bitolas.Text)
                        Next J
                    End If
                End If
            Next I
        End If
    End If
    BP_1.Value = 0
    PrecoPesoEmEspera (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Remove_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Figuras.ListIndex = -1 Then
        MsgBox ("Você precisa primeiro selecionar uma figura na lista abaixo para removê-la.")
        LT_Figuras.SetFocus
        Exit Sub
    End If
    LT_Figuras.RemoveItem (LT_Figuras.ListIndex)
    LT_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Remove_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then LT_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Salvar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Figuras.ListCount = 0 Then
        MsgBox ("Não existe nenhuma figura ainda para ser editada.")
        BT_Salvar.Enabled = False
        TXT_1.Text = ""
        lTeste = True
        LT_Figuras.SetFocus
        Exit Sub
    End If
    PrecoPesoEmEspera (True)
    Dim cVal, cFig, cBit, cMat As String
    cFig = ""
    cBit = ""
    cMat = ""
    Dim nPos As Integer
    nPos = 1
    For I = 0 To LT_Figuras.ListCount - 1
        BP_1.Max = LT_Figuras.ListCount
        BP_1.Value = 0
        For J = 1 To Len(LT_Figuras.List(I))
            If Mid(LT_Figuras.List(I), J, 2) = "em" Then
                cFig = Mid(LT_Figuras.List(I), nPos, J - 2)
                nPos = J + 3
            End If
            If Mid(LT_Figuras.List(I), J, 2) = "de" Then
                cMat = Mid(LT_Figuras.List(I), nPos, (J - 1) - nPos)
                nPos = J + 3
                cBit = Mid(LT_Figuras.List(I), nPos, Len(LT_Figuras.List(I)))
            End If
            If cFig <> "" And cBit <> "" And cMat <> "" Then
                DLL_BD.BDSIS_TBEST.Seek "=", cFig, cBit, cMat
                If DLL_BD.BDSIS_TBEST.NoMatch Then
                    MsgBox ("Não foi encontrada a ficha de estoque para gravar os dados.")
                    Exit Sub
                End If
                DLL_BD.BDSIS_TBEST.Edit
                If RB_Preco.Value = True Then
                    DLL_BD.BDSIS_TBEST_CPVUN.Value = TXT_1.Text
                ElseIf RB_Peso.Value = True Then
                    DLL_BD.BDSIS_TBEST_CPPUN.Value = TXT_1.Text
                End If
                DLL_BD.BDSIS_TBEST.Update
                Exit For
            End If
        Next J
        nPos = 1
        cFig = ""
        cBit = ""
        cMat = ""
        BP_1.Value = BP_1.Value + 1
    Next I
    DLL_FUNCS.RegistraEvento "Salvar - Preços e Pesos de Estoque", cFig & " " & cBit & " " & cMat
    BP_1.Value = 0
    TXT_1.Text = ""
    CB_Figuras.ListIndex = -1
    CB_Bitolas.Clear
    CB_Materiais.Clear
    LT_Figuras.Clear
    PrecoPesoEmEspera (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Cfg_Estoque_PrecoPeso
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitolas_Click()
    On Error GoTo ERRO_SISCOVAL
    BT_Salvar.Enabled = False
    If CB_Figuras.ListIndex <> -1 And CB_Materiais.ListIndex <> -1 And CB_Bitolas.ListIndex <> -1 Then
        cResp = ProcuraFigura(CB_Figuras.Text, CB_Materiais.Text, CB_Bitolas.Text)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitolas_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CK_Incluir.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figuras_Click()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figuras.ListIndex = -1 Then Exit Sub
    PrecoPesoEmEspera (True)
    DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figuras.Text
    If DLL_BD.BDSIS_TBEFG.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura da figura.")
        PrecoPesoEmEspera (False)
        Exit Sub
    End If
    Dim cProc As String
    cProc = DLL_BD.BDSIS_TBEFG_CPIFG.Value
    DLL_BD.BDSIS_TBEID.Seek "=", cProc
    If DLL_BD.BDSIS_TBEID.NoMatch Then
        MsgBox ("Ocorreu algum erro durante a procura do índice da figura.")
        PrecoPesoEmEspera (False)
        Exit Sub
    End If
    CB_Materiais.Clear
    CB_Bitolas.Clear
    Dim cA As String
    cA = ""
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
            CB_Materiais.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Materiais.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    cA = ""
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGBI.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) = ";" Then
            CB_Bitolas.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Bitolas.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    TXT_1.Text = ""
    If CB_Figuras.ListIndex <> -1 And CB_Materiais.ListIndex <> -1 And CB_Bitolas.ListIndex <> -1 Then
        cResp = ProcuraFigura(CB_Figuras.Text, CB_Materiais.Text, CB_Bitolas.Text)
    End If
    BT_Salvar.Enabled = False
    PrecoPesoEmEspera (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figuras_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Materiais.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Indice_Click()
    On Error GoTo ERRO_SISCOVAL
    If CB_Indice.ListIndex = -1 Then Exit Sub
    PrecoPesoEmEspera (True)
    CB_Figuras.Clear
    If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
        BP_1.Max = DLL_BD.BDSIS_TBEFG.RecordCount
        BP_1.Value = 0
        DLL_BD.BDSIS_TBEFG.MoveFirst
        Do While Not DLL_BD.BDSIS_TBEFG.EOF
            If DLL_BD.BDSIS_TBEFG_CPIFG.Value = CB_Indice.Text Then
                CB_Figuras.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value)
            End If
            DLL_BD.BDSIS_TBEFG.MoveNext
            BP_1.Value = BP_1.Value + 1
        Loop
        BP_1.Value = 0
    End If
    CB_Materiais.Clear
    CB_Bitolas.Clear
    PrecoPesoEmEspera (False)
    CB_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Indice_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Materiais_Click()
    On Error GoTo ERRO_SISCOVAL
    BT_Salvar.Enabled = False
    If CB_Figuras.ListIndex <> -1 And CB_Materiais.ListIndex <> -1 And CB_Bitolas.ListIndex <> -1 Then
        cResp = ProcuraFigura(CB_Figuras.Text, CB_Materiais.Text, CB_Bitolas.Text)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Incluir_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Inclui.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc

    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (10)
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
    ' Abre Campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque...")
    If DLL_BD.AbreCampos_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Índice...")
    If DLL_BD.AbreCampos_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Figuras...")
    If DLL_BD.AbreCampos_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS

    DLL_CARGA.CarregaTexto ("Finalizando...")
    CB_Figuras.Clear
    CB_Materiais.Clear
    CB_Bitolas.Clear
    RB_Preco.Value = True
    BP_1.Value = 0
    BT_Salvar.Enabled = False
    DLL_FUNCS.RegistraEvento "Abrir Configurações de Preços e Pesos", ""
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cfg_Estoque_PrecoPeso
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
    Set DLL_FUNCS = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Figuras_Click()
    On Error GoTo ERRO_SISCOVAL
    Dim cVal, cFig, cBit, cMat As String
    cFig = ""
    cBit = ""
    cMat = ""
    Dim nPos As Integer
    nPos = 1
    For J = 1 To Len(LT_Figuras.Text)
        If Mid(LT_Figuras.Text, J, 2) = "em" Then
            cFig = Mid(LT_Figuras.Text, nPos, J - 2)
            nPos = J + 3
        End If
        If Mid(LT_Figuras.Text, J, 2) = "de" Then
            cMat = Mid(LT_Figuras.Text, nPos, (J - 1) - nPos)
            nPos = J + 3
            cBit = Mid(LT_Figuras.Text, nPos, Len(LT_Figuras.Text))
        End If
        If cFig <> "" And cBit <> "" And cMat <> "" Then
            DLL_BD.BDSIS_TBEST.Seek "=", cFig, cBit, cMat
            If DLL_BD.BDSIS_TBEST.NoMatch Then
                MsgBox ("Não foi encontrada a ficha de estoque para gravar os dados.")
                Exit Sub
            End If
            If RB_Preco.Value = True Then
                TXT_1.Text = DLL_BD.BDSIS_TBEST_CPVUN.Value
            ElseIf RB_Peso.Value = True Then
                TXT_1.Text = DLL_BD.BDSIS_TBEST_CPPUN.Value
            End If
        End If
    Next J
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Figuras_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_1.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Indice_Click()
    On Error GoTo ERRO_SISCOVAL
    PrecoPesoEmEspera (True)
    BP_1.Value = 0
    CB_Indice.Enabled = True
    CB_Indice.Clear
    CB_Figuras.Clear
    CB_Materiais.Clear
    CB_Bitolas.Clear
    If DLL_BD.BDSIS_TBEID.RecordCount > 0 Then
        BP_1.Max = DLL_BD.BDSIS_TBEID.RecordCount
        DLL_BD.BDSIS_TBEID.MoveFirst
        Do While Not DLL_BD.BDSIS_TBEID.EOF
            CB_Indice.AddItem (DLL_BD.BDSIS_TBEID_CPIFI.Value)
            DLL_BD.BDSIS_TBEID.MoveNext
            BP_1.Value = BP_1.Value + 1
        Loop
    End If
    BP_1.Value = 0
    CB_Indice.ListIndex = -1
    PrecoPesoEmEspera (False)
    BT_Salvar.Enabled = False
    CB_Indice.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Indice_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Indice.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Peso_Click()
    On Error GoTo ERRO_SISCOVAL
    FR_3.Caption = "Peso"
    BT_Salvar.Enabled = False
    CK_Incluir.Caption = "Incluir pesos de peças equivalentes"
    If CB_Figuras.ListIndex <> -1 And CB_Materiais.ListIndex <> -1 And CB_Bitolas.ListIndex <> -1 Then
        cResp = ProcuraFigura(CB_Figuras.Text, CB_Materiais.Text, CB_Bitolas.Text)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Peso_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then RB_Todas.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Preco_Click()
    On Error GoTo ERRO_SISCOVAL
    FR_3.Caption = "Preço"
    BT_Salvar.Enabled = False
    CK_Incluir.Caption = "Incluir preços de peças equivalentes"
    If CB_Figuras.ListIndex <> -1 And CB_Materiais.ListIndex <> -1 And CB_Bitolas.ListIndex <> -1 Then
        cResp = ProcuraFigura(CB_Figuras.Text, CB_Materiais.Text, CB_Bitolas.Text)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Preco_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then RB_Todas.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Todas_Click()
    On Error GoTo ERRO_SISCOVAL
    PrecoPesoEmEspera (True)
    BP_1.Max = DLL_BD.BDSIS_TBEFG.RecordCount
    BP_1.Value = 0
    CB_Indice.Enabled = False
    CB_Figuras.Clear
    CB_Materiais.Clear
    CB_Bitolas.Clear
    'Carrega lista de figuras
    If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEFG.MoveFirst
        Do While Not DLL_BD.BDSIS_TBEFG.EOF
            CB_Figuras.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value)
            DLL_BD.BDSIS_TBEFG.MoveNext
            BP_1.Value = BP_1.Value + 1
        Loop
    Else
        MsgBox ("Não Existe Nenhuma ficha aberta - execute o assistente de configuração de estoque primeiro.")
        BT_Voltar.Value = True
        Exit Sub
    End If
    BP_1.Value = 0
    BT_Salvar.Enabled = False
    PrecoPesoEmEspera (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Todas_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Figuras.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_1_Change()
    On Error GoTo ERRO_SISCOVAL
    If lTeste = True Then
        lTeste = False
        BT_Salvar.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub PrecoPesoEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_Cfg_Estoque_PrecoPeso.MousePointer = vbHourglass
        Tela_Cfg_Estoque_PrecoPeso.Enabled = False
    Else
        Tela_Cfg_Estoque_PrecoPeso.MousePointer = vbDefault
        Tela_Cfg_Estoque_PrecoPeso.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function ProcuraFigura(Fig As String, Mat As String, Bit As String)
    On Error GoTo ERRO_SISCOVAL
    DLL_BD.BDSIS_TBEST.Seek "=", Fig, Bit, Mat
    If DLL_BD.BDSIS_TBEST.NoMatch Then
        MsgBox ("Erro ao procurar ficha no banco de dados de estoque.")
    End If
    If RB_Preco.Value = True Then
        TXT_1.Text = DLL_BD.BDSIS_TBEST_CPVUN.Value
    ElseIf RB_Peso.Value = True Then
        TXT_1.Text = DLL_BD.BDSIS_TBEST_CPPUN.Value
    End If
    BT_Salvar.Enabled = False
    lTeste = True
    ProcuraFigura = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Sub TXT_1_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_1.SelLength = Len(TXT_1.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_1_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then
        BT_Salvar.SetFocus
        Exit Sub
    End If
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
