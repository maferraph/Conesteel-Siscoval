VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Tela_Condicional 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Escolha o critério para listagem das Notas Fiscais"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_VOL 
      Caption         =   "Voltar à Tela Principal"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      ToolTipText     =   "Volta à tela de notas fiscais emitidas"
      Top             =   1920
      Width           =   1935
   End
   Begin VB.CommandButton BT_PRO 
      Caption         =   "Procurar Notas Fiscais"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      ToolTipText     =   "Iniciar procura das notas fiscais"
      Top             =   1920
      Width           =   1935
   End
   Begin MSMask.MaskEdBox TXT1 
      Height          =   315
      Left            =   2520
      TabIndex        =   7
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin VB.ComboBox CB_VA 
      Height          =   315
      ItemData        =   "Tela_Condicional.frx":0000
      Left            =   2520
      List            =   "Tela_Condicional.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   11
      ToolTipText     =   "Selecione um critério de procura de nota fiscal para valores"
      Top             =   1440
      Width           =   735
   End
   Begin VB.ComboBox CB_EMP 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "A procura será filtrada para a empresa nesta lista"
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CheckBox CK_EMP 
      Caption         =   "Filtrar procura por Empresa"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Se desejar que a procura das notas fiscais seja feita somente com a empresa selecionada na lista abaixo"
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Critério:"
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   2175
      Begin VB.OptionButton RB_VA 
         Caption         =   "Valor"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Procura nota fiscal por valores"
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton RB_DA 
         Caption         =   "Data"
         Height          =   255
         Left            =   1320
         TabIndex        =   4
         ToolTipText     =   "Procura nota fiscal por data de emissão"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton RB_PE 
         Caption         =   "Período"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Procura nota fiscal por data inicial até data final"
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton RB_AN 
         Caption         =   "Ano"
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         ToolTipText     =   "Procura notas fiscais por ano"
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton RB_ME 
         Caption         =   "Mês"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         ToolTipText     =   "Procura notas fisicais por Mês/Ano"
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton RB_NU 
         Caption         =   "Número"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Procura por número de nota fiscal"
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSMask.MaskEdBox TXT2 
      Height          =   315
      Left            =   2520
      TabIndex        =   8
      Top             =   840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      PromptChar      =   "_"
   End
   Begin MSComctlLib.ProgressBar BP 
      Height          =   255
      Left            =   3360
      TabIndex        =   17
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label LB3 
      AutoSize        =   -1  'True
      Caption         =   "Valor:"
      Height          =   195
      Left            =   2520
      TabIndex        =   15
      Top             =   1200
      Width           =   405
   End
   Begin VB.Label LB2 
      AutoSize        =   -1  'True
      Caption         =   "***"
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   600
      Width           =   180
   End
   Begin VB.Label LB1 
      AutoSize        =   -1  'True
      Caption         =   "Digite o número da N.F.:"
      Height          =   195
      Left            =   2520
      TabIndex        =   14
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "Tela_Condicional"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const NOMEAPLIC As String = "Critério de Procura de N.F."
Private Sub BT_PRO_Click()
    On Error GoTo ERRO_SISCOVAL
    'Testa critérios de procura
    If RB_NU.Value = True Then
        If TXT1.Text = "" Or IsNumeric(TXT1.Text) = False Then
            MsgBox "O número da Nota Fiscal não foi digitado ou é inválido.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT1.SetFocus
            Exit Sub
        End If
        Tela_NFEmitidas.CRITERIO = "Por número: " & Trim(TXT1.Text)
    ElseIf RB_PE.Value = True Then
        If TXT1.Text = "__/__/____" Or IsDate(TXT1.Text) = False Then
            MsgBox "A data inicial das notas fiscais não foi digitada ou é inválida.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT1.SetFocus
            Exit Sub
        ElseIf TXT2.Text = "__/__/____" Or IsDate(TXT2.Text) = False Then
            MsgBox "A data final das notas fiscais não foi digitada ou é inválida.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT2.SetFocus
            Exit Sub
        End If
        Tela_NFEmitidas.CRITERIO = "Por período - de: " & Trim(TXT1.Text) & " até: " & Trim(TXT2.Text)
    ElseIf RB_VA.Value = True Then
        If TXT1.Text = "" Or IsNumeric(TXT1.Text) = False Then
            MsgBox "O valor da Nota Fiscal não foi digitado ou é inválido.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT1.SetFocus
            Exit Sub
        ElseIf CB_VA.ListIndex = -1 Then
            MsgBox "O critério de valor da Nota Fiscal não foi escolhido.", vbInformation + vbOKOnly, NOMEAPLIC
            CB_VA.SetFocus
            Exit Sub
        End If
        Tela_NFEmitidas.CRITERIO = "Por valores " & Trim(CB_VA.Text) & " " & Trim(TXT1.Text)
    ElseIf RB_DA.Value = True Then
        If TXT1.Text = "__/__/____" Or IsDate(TXT1.Text) = False Then
            MsgBox "A data das notas fiscais não foi digitada ou é inválida.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT1.SetFocus
            Exit Sub
        End If
        Tela_NFEmitidas.CRITERIO = "Por data: " & Trim(TXT1.Text)
    ElseIf RB_ME.Value = True Then
        If TXT1.Text = "__/____" Then
            MsgBox "A data mês/ano das notas fiscais não foi digitada ou é inválida.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT1.SetFocus
            Exit Sub
        End If
        Tela_NFEmitidas.CRITERIO = "Por data (Mês/Ano): " & Trim(TXT1.Text)
    ElseIf RB_AN.Value = True Then
        If TXT1.Text = "____" Then
            MsgBox "O ano das notas fiscais não foi digitada ou é inválida.", vbInformation + vbOKOnly, NOMEAPLIC
            TXT1.SetFocus
            Exit Sub
        End If
        Tela_NFEmitidas.CRITERIO = "Por data (Ano): " & Trim(TXT1.Text)
    End If
    If CK_EMP.Value = 1 Then Tela_NFEmitidas.CRITERIO = Tela_NFEmitidas.CRITERIO & " c/filtro por empresa"
    
    'Começa a pesquisa
    With Tela_NFEmitidas
        .DLL_BD.BDSIS_TBNTF.MoveFirst
        BP.Max = .DLL_BD.BDSIS_TBNTF.RecordCount + 5
        BP.Value = 0
        .LT_NF.Clear
        If RB_NU.Value = True Then 'Por número
            .DLL_BD.BDSIS_TBNTF.Seek "=", TXT1.Text
            If .DLL_BD.BDSIS_TBNTF.NoMatch Then
                MsgBox "Não foi encontrada a N.F. " & TXT1.Text, vbInformation + vbOKOnly, NOMEAPLIC
                Exit Sub
            End If
            .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
        ElseIf RB_PE.Value = True Then 'Por período
            Do While Not .DLL_BD.BDSIS_TBNTF.EOF
                If CK_EMP.Value = 1 And CB_EMP.ListIndex >= 0 Then
                    If DateValue(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) >= DateValue(TXT1.Text) And DateValue(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) <= DateValue(TXT2.Text) And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                Else
                    If DateValue(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) >= DateValue(TXT1.Text) And DateValue(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) <= DateValue(TXT2.Text) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                End If
                BP.Value = BP.Value + 1
                .DLL_BD.BDSIS_TBNTF.MoveNext
            Loop
        ElseIf RB_VA.Value = True Then 'Por valor
            Do While Not .DLL_BD.BDSIS_TBNTF.EOF
                If CK_EMP.Value = 1 And CB_EMP.ListIndex >= 0 Then
                    If CB_VA.Text = "<=" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) <= CDbl(TXT1.Text) And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = "<" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) < CDbl(TXT1.Text) And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = "=" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) = CDbl(TXT1.Text) And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = ">" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) > CDbl(TXT1.Text) And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = ">=" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) >= CDbl(TXT1.Text) And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                Else
                    If CB_VA.Text = "<=" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) <= CDbl(TXT1.Text) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = "<" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) < CDbl(TXT1.Text) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = "=" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) = CDbl(TXT1.Text) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = ">" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) > CDbl(TXT1.Text) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                    If CB_VA.Text = ">=" And CDbl(.DLL_BD.BDSIS_TBNTF_CPVAL.Value) >= CDbl(TXT1.Text) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                End If
                BP.Value = BP.Value + 1
                .DLL_BD.BDSIS_TBNTF.MoveNext
            Loop
        ElseIf RB_DA.Value = True Then 'Por Data
            Do While Not .DLL_BD.BDSIS_TBNTF.EOF
                If CK_EMP.Value = 1 And CB_EMP.ListIndex >= 0 Then
                    If DateValue(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) = DateValue(TXT1.Text) And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                Else
                    If DateValue(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) = DateValue(TXT1.Text) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                End If
                BP.Value = BP.Value + 1
                .DLL_BD.BDSIS_TBNTF.MoveNext
            Loop
        ElseIf RB_ME.Value = True Then 'Por Mês/Ano
            Do While Not .DLL_BD.BDSIS_TBNTF.EOF
                If CK_EMP.Value = 1 And CB_EMP.ListIndex >= 0 Then
                    If Format(Trim(Month(.DLL_BD.BDSIS_TBNTF_CPDEM.Value)) & "/" & Trim(Year(.DLL_BD.BDSIS_TBNTF_CPDEM.Value)), "mm/yyyy") = Format(TXT1.Text, "mm/yyyy") And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                Else
                    If Format(Trim(Month(.DLL_BD.BDSIS_TBNTF_CPDEM.Value)) & "/" & Trim(Year(.DLL_BD.BDSIS_TBNTF_CPDEM.Value)), "mm/yyyy") = Format(TXT1.Text, "mm/yyyy") Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                End If
                BP.Value = BP.Value + 1
                .DLL_BD.BDSIS_TBNTF.MoveNext
            Loop
        ElseIf RB_AN.Value = True Then 'Por Ano
            Do While Not .DLL_BD.BDSIS_TBNTF.EOF
                If CK_EMP.Value = 1 And CB_EMP.ListIndex >= 0 Then
                    If Year(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) = TXT1.Text And .DLL_BD.BDSIS_TBNTF_CPEMP.Value = CB_EMP.List(CB_EMP.ListIndex) Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                Else
                    If Year(.DLL_BD.BDSIS_TBNTF_CPDEM.Value) = TXT1.Text Then .LT_NF.AddItem (.DLL_BD.BDSIS_TBNTF_CPNNF.Value)
                End If
                BP.Value = BP.Value + 1
                .DLL_BD.BDSIS_TBNTF.MoveNext
            Loop
        End If
        If .LT_NF.ListCount = 0 Then MsgBox "Não foi possível encontrar nenhuma nota fiscal dentro dos critérios acima estipulados.", vbCritical + vbOKOnly, NOMEAPLIC
        Tela_NFEmitidas.RB_Condicional.Value = False
        Unload Me
    End With
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_VOL_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Condicional
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_EMP_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_PRO.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_VA_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_PRO.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EMP_Click()
    On Error GoTo ERRO_SISCOVAL
    If CK_EMP.Value = 1 Then
        CB_EMP.Enabled = True
    Else
        CB_EMP.Enabled = False
    End If
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EMP_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_PRO.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    For I = 0 To Tela_NFEmitidas.CB_Empresas.ListCount - 1
        CB_EMP.AddItem (Tela_NFEmitidas.CB_Empresas.List(I))
    Next I
    CB_EMP.ListIndex = -1
    CK_EMP.Value = 0
    CB_EMP.Enabled = False
    CB_VA.ListIndex = -1
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_AN_Click()
    On Error GoTo ERRO_SISCOVAL
    LB1.Enabled = True
    LB1.Caption = "Digite o ano da N.F.:"
    TXT1.Enabled = True
    TXT1.Format = ""
    TXT1.Mask = "####"
    TXT1.MaxLength = 5
    LB2.Enabled = False
    LB2.Caption = "-"
    TXT2.Enabled = False
    TXT2.Format = ""
    TXT2.Mask = ""
    TXT2.Text = ""
    TXT2.MaxLength = 10
    LB3.Enabled = False
    CB_VA.Enabled = False
    TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_AN_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_DA_Click()
    On Error GoTo ERRO_SISCOVAL
    LB1.Enabled = True
    LB1.Caption = "Digite a data da N.F.:"
    TXT1.Enabled = True
    TXT1.Format = "dd/mm/yyyy"
    TXT1.Mask = "##/##/####"
    TXT1.MaxLength = 10
    LB2.Enabled = False
    LB2.Caption = "-"
    TXT2.Enabled = False
    TXT2.Format = ""
    TXT2.Mask = ""
    TXT2.Text = ""
    TXT2.MaxLength = 10
    LB3.Enabled = False
    CB_VA.Enabled = False
    TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_DA_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_ME_Click()
    On Error GoTo ERRO_SISCOVAL
    LB1.Enabled = True
    LB1.Caption = "Digite o mês/ano da N.F.:"
    TXT1.Enabled = True
    TXT1.Format = "mm/yyyy"
    TXT1.Mask = "##/####"
    TXT1.MaxLength = 8
    LB2.Enabled = False
    LB2.Caption = "-"
    TXT2.Enabled = False
    TXT2.Format = ""
    TXT2.Mask = ""
    TXT2.Text = ""
    TXT2.MaxLength = 10
    LB3.Enabled = False
    CB_VA.Enabled = False
    TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_ME_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_NU_Click()
    On Error GoTo ERRO_SISCOVAL
    LB1.Enabled = True
    LB1.Caption = "Digite o nº da N.F."
    TXT1.Enabled = True
    TXT1.Format = ""
    TXT1.Mask = ""
    TXT1.Text = ""
    TXT1.MaxLength = 10
    LB2.Enabled = False
    LB2.Caption = "-"
    TXT2.Enabled = False
    TXT2.Format = ""
    TXT2.Mask = ""
    TXT2.Text = ""
    TXT2.MaxLength = 10
    LB3.Enabled = False
    CB_VA.Enabled = False
    TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_NU_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_PE_Click()
    On Error GoTo ERRO_SISCOVAL
    LB1.Enabled = True
    LB1.Caption = "Digite a data inicial"
    TXT1.Enabled = True
    TXT1.Format = "dd/mm/yyyy"
    TXT1.Mask = "##/##/####"
    TXT1.MaxLength = 10
    LB2.Enabled = True
    LB2.Caption = "Digite a data final"
    TXT2.Enabled = True
    TXT2.Format = "dd/mm/yyyy"
    TXT2.Mask = "##/##/####"
    TXT2.MaxLength = 10
    LB3.Enabled = False
    CB_VA.Enabled = False
    TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_PE_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_VA_Click()
    On Error GoTo ERRO_SISCOVAL
    LB1.Enabled = True
    LB1.Caption = "Digite o valor da N.F.:"
    TXT1.Enabled = True
    TXT1.Format = ""
    TXT1.Mask = ""
    TXT1.Text = ""
    TXT1.MaxLength = 15
    LB2.Enabled = False
    LB2.Caption = "-"
    TXT2.Enabled = False
    TXT2.Format = ""
    TXT2.Mask = ""
    TXT2.Text = ""
    TXT2.MaxLength = 10
    LB3.Enabled = True
    CB_VA.Enabled = True
    TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_VA_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT1.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT1_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = Tela_NFEmitidas.DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then BT_PRO.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT2_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = Tela_NFEmitidas.DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then BT_PRO.SetFocus
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_Condicional.MousePointer = vbHourglass
        Tela_Condicional.Enabled = False
    Else
        Tela_Condicional.MousePointer = vbDefault
        Tela_Condicional.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If Tela_NFEmitidas.DLL_FUNCS.MensagemErro(Tela_NFEmitidas.DLL_FUNCS.PegaUsuario, Tela_NFEmitidas.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
