VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Tela_Movimento 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contas Diversas"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "Voltar à tela de contas sem fazer alterações"
      Height          =   855
      Left            =   4080
      Picture         =   "Tela_Modelo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Volta à tela de contas sem fazer alterações"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.CommandButton BT_Salvar 
      Caption         =   "Salvar dados digitados"
      Height          =   855
      Left            =   120
      Picture         =   "Tela_Modelo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salva os dados digitados para esta conta"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Frame FR 
      Height          =   2415
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   7335
      Begin VB.TextBox TXT_Observacoes 
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         ToolTipText     =   "Digite aqui alguma observação se existir"
         Top             =   1800
         Width           =   4215
      End
      Begin VB.TextBox TXT_NossoNum 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "Digite o nosso número do banco (no boleto) se existir"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox TXT_SeuNum 
         Height          =   285
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Digite o seu número do banco (no boleto) se existir"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox CB_Banco 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Selecione uma conta de banco para este movimento"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TXT_Historico 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4680
         TabIndex        =   2
         ToolTipText     =   "Digite o histórico desta conta"
         Top             =   360
         Width           =   2535
      End
      Begin VB.ComboBox CB_Movimento 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Selecione o tipo do movimento"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox TXT_NumDoc 
         Height          =   285
         Left            =   3720
         TabIndex        =   5
         ToolTipText     =   "Digite o número do documento desta conta"
         Top             =   1080
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TXT_DataEmissao 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Digite a data de emissão desta conta"
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "&&/&&/&&&&"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_DataVencimento 
         Height          =   285
         Left            =   5400
         TabIndex        =   6
         ToolTipText     =   "Digite a data de vencimento desta conta"
         Top             =   1080
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "&&/&&/&&&&"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_Valor 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         ToolTipText     =   "Digite o valor desta conta"
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   20
         Format          =   "$###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Observações:"
         Height          =   195
         Index           =   10
         Left            =   3000
         TabIndex        =   22
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Nosso Número:"
         Height          =   195
         Index           =   9
         Left            =   1560
         TabIndex        =   21
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Seu Número:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   20
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Conta de Banco:"
         Height          =   195
         Index           =   7
         Left            =   2400
         TabIndex        =   19
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Histórico:"
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   18
         Top             =   120
         Width           =   660
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Movimento:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   1410
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento:"
         Height          =   195
         Index           =   3
         Left            =   3720
         TabIndex        =   16
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Data Emissão:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1020
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Valor:"
         Height          =   195
         Index           =   5
         Left            =   2040
         TabIndex        =   14
         Top             =   840
         Width           =   405
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Data Vencimento:"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   13
         Top             =   840
         Width           =   1275
      End
   End
End
Attribute VB_Name = "Tela_Movimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private bEdicao As Boolean
Private NOMEAPLIC As String

Private Sub BT_Salvar_Click()
    'verifica dados digitados
    Dim sMsg As String
    sMsg = "Alguns campos são obrigatórios o preenchimento."
    If CB_Movimento.ListIndex = -1 Then
        MsgBox sMsg, vbExclamation + vbOKOnly, NOMEAPLIC
        CB_Movimento.SetFocus
        Exit Sub
    ElseIf TXT_Historico.Text = "" Then
        MsgBox sMsg, vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_Historico.SetFocus
        Exit Sub
    ElseIf TXT_DataEmissao.Text = "__/__/____" Then
        MsgBox sMsg, vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_DataEmissao.SetFocus
        Exit Sub
    ElseIf IsDate(TXT_DataEmissao.Text) = False Then
        MsgBox "Data inválida.", vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_DataEmissao.SetFocus
        Exit Sub
    ElseIf TXT_Valor.Text = "" Then
        MsgBox sMsg, vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_Valor.SetFocus
        Exit Sub
    ElseIf TXT_DataVencimento.Text = "__/__/____" Then
        MsgBox sMsg, vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_DataVencimento.SetFocus
        Exit Sub
    ElseIf IsDate(TXT_DataVencimento.Text) = False Then
        MsgBox "Data inválida.", vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_DataVencimento.SetFocus
        Exit Sub
    End If
    TelaEmEspera True
    With Tela_Contas
        .BS.SimpleText = "Aguarde... salvando informações digitadas."
        .BP.Max = 3
        .BP.Value = 0
        If bEdicao = False Then 'Novo Movimento
            .DLL_BD.BDSIS_TBCPR.AddNew
        Else 'Edição do Movimento
            .DLL_BD.BDSIS_TBCPR.Edit
        End If
        .BP.Value = .BP.Value + 1
        'grava informações
        .DLL_BD.BDSIS_TBCPR_CPEMI.Value = TXT_DataEmissao.Text
        .DLL_BD.BDSIS_TBCPR_CPVEN.Value = TXT_DataVencimento.Text
        If Me.Caption = "Contas Diversas à Pagar" Then
            .DLL_BD.BDSIS_TBCPR_CPMOV.Value = "P"
        ElseIf Me.Caption = "Contas Diversas à Receber" Then
            .DLL_BD.BDSIS_TBCPR_CPMOV.Value = "R"
        End If
        .DLL_BD.BDSIS_TBCPR_CPVAL.Value = TXT_Valor.Text
        .DLL_BD.BDSIS_TBCPR_CPORI.Value = TXT_Historico.Text
        .DLL_BD.BDSIS_TBCPR_CPNDO.Value = PegaDadosDigitados(TXT_NumDoc)
        .DLL_BD.BDSIS_TBCPR_CPNNU.Value = PegaDadosDigitados(TXT_NossoNum)
        .DLL_BD.BDSIS_TBCPR_CPSNU.Value = PegaDadosDigitados(TXT_SeuNum)
        .DLL_BD.BDSIS_TBCPR_CPTIP.Value = CB_Movimento.List(CB_Movimento.ListIndex)
        .DLL_BD.BDSIS_TBCPR_CPBAN.Value = PegaDadosDigitados(CB_Banco)
        .DLL_BD.BDSIS_TBCPR_CPCAR.Value = False
        .DLL_BD.BDSIS_TBCPR_CPCOB.Value = False
        .DLL_BD.BDSIS_TBCPR_CPOBS.Value = PegaDadosDigitados(TXT_Observacoes)
        .BP.Value = .BP.Value + 1
        .DLL_BD.BDSIS_TBCPR.Update
        .BP.Value = .BP.Value + 1
        .BS.SimpleText = ""
        .BP.Value = 0
    End With
    ApagaCampos
    TelaEmEspera False
End Sub
Private Sub BT_Voltar_Click()
    Unload Me
End Sub
Private Sub CB_Banco_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_Historico.SetFocus
End Sub
Private Sub CB_Movimento_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_Banco.SetFocus
End Sub
Private Sub Form_Load()
    With Tela_Contas.DLL_BD
        'Carregando combo de tipos de contas
        CB_Movimento.Clear
        Tela_Contas.BS.SimpleText = "Aguarde... carregando lista de movimentos."
        If .BDSIS_TBGRU.RecordCount > 0 Then
            Tela_Contas.BP.Max = .BDSIS_TBGRU.RecordCount + 1
            Tela_Contas.BP.Value = 0
            .BDSIS_TBGRU.MoveFirst
            While Not .BDSIS_TBGRU.EOF
                If .BDSIS_TBGRU_CPTIP.Value = "MCN" Then CB_Movimento.AddItem .BDSIS_TBGRU_CPVAL.Value
                .BDSIS_TBGRU.MoveNext
                Tela_Contas.BP.Value = Tela_Contas.BP.Value + 1
            Wend
        End If
        'Carregando combo de tipos de contas
        CB_Banco.Clear
        Tela_Contas.BS.SimpleText = "Aguarde... carregando lista de bancos."
        If .BDSIS_TBBAN.RecordCount > 0 Then
            Tela_Contas.BP.Max = .BDSIS_TBBAN.RecordCount + 1
            Tela_Contas.BP.Value = 0
            .BDSIS_TBBAN.MoveFirst
            While Not .BDSIS_TBBAN.EOF
                CB_Banco.AddItem .BDSIS_TBBAN_CPNMC.Value
                .BDSIS_TBBAN.MoveNext
                Tela_Contas.BP.Value = Tela_Contas.BP.Value + 1
            Wend
        End If
    End With
    Tela_Contas.BP.Value = 0
    Tela_Contas.BS.SimpleText = ""
    ApagaCampos
End Sub
Private Sub TXT_DataEmissao_GotFocus()
    TXT_DataEmissao.SelLength = Len(TXT_DataEmissao.Text)
End Sub
Private Sub TXT_DataEmissao_KeyPress(KeyAscii As Integer)
    KeyAscii = Tela_Contas.DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_Valor.SetFocus
End Sub
Private Sub TXT_DataVencimento_GotFocus()
    TXT_DataVencimento.SelLength = Len(TXT_DataVencimento.Text)
End Sub
Private Sub TXT_DataVencimento_KeyPress(KeyAscii As Integer)
    KeyAscii = Tela_Contas.DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_SeuNum.SetFocus
End Sub
Private Sub TXT_Historico_GotFocus()
    TXT_Historico.SelLength = Len(TXT_Historico.Text)
End Sub
Private Sub TXT_Historico_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_DataEmissao.SetFocus
End Sub
Private Sub TXT_NossoNum_GotFocus()
    TXT_NossoNum.SelLength = Len(TXT_NossoNum.Text)
End Sub
Private Sub TXT_NossoNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_Observacoes.SetFocus
End Sub
Private Sub TXT_NumDoc_GotFocus()
    TXT_NumDoc.SelLength = Len(TXT_NumDoc.Text)
End Sub
Private Sub TXT_NumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_DataVencimento.SetFocus
End Sub
Private Sub TXT_NumNF_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_Valor.SetFocus
End Sub
Private Sub TXT_Observacoes_GotFocus()
    TXT_Observacoes.SelLength = Len(TXT_Observacoes.Text)
End Sub
Private Sub TXT_Observacoes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BT_Salvar.SetFocus
End Sub
Private Sub TXT_SeuNum_GotFocus()
    TXT_SeuNum.SelLength = Len(TXT_SeuNum.Text)
End Sub
Private Sub TXT_SeuNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_NossoNum.SetFocus
End Sub
Private Sub TXT_Valor_GotFocus()
    TXT_Valor.SelLength = Len(TXT_Valor.Text)
End Sub
Private Sub TXT_Valor_KeyPress(KeyAscii As Integer)
    KeyAscii = Tela_Contas.DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_NumDoc.SetFocus
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub ApagaCampos()
    CB_Movimento.ListIndex = -1
    CB_Banco.ListIndex = -1
    TXT_Historico.Text = ""
    TXT_DataEmissao.Text = "__/__/____"
    TXT_Valor.Text = ""
    TXT_NumDoc.Text = ""
    TXT_DataVencimento.Text = "__/__/____"
    TXT_SeuNum.Text = ""
    TXT_NossoNum.Text = ""
    TXT_Observacoes.Text = ""
End Sub
Public Static Sub Duplicata_Nova(sPouR As String)
    CB_Movimento.Enabled = True
    CB_Banco.Enabled = True
    TXT_Historico.Enabled = True
    TXT_DataEmissao.Enabled = True
    TXT_Valor.Enabled = True
    TXT_NumDoc.Enabled = True
    TXT_DataVencimento.Enabled = True
    TXT_SeuNum.Enabled = True
    TXT_NossoNum.Enabled = True
    TXT_Observacoes.Enabled = True
    If sPouR = "P" Then
        NOMEAPLIC = "Duplicatas à Pagar"
        Me.Caption = NOMEAPLIC
    ElseIf sPouR = "R" Then
        NOMEAPLIC = "Duplicatas à Receber"
        Me.Caption = NOMEAPLIC
    End If
    bEdicao = False
End Sub
Public Static Sub Duplicata_Edicao(sPouR As String)
    CB_Movimento.Enabled = False
    CB_Banco.Enabled = True
    TXT_Historico.Enabled = True
    TXT_DataEmissao.Enabled = True
    TXT_Valor.Enabled = True
    TXT_NumDoc.Enabled = True
    TXT_DataVencimento.Enabled = True
    TXT_SeuNum.Enabled = True
    TXT_NossoNum.Enabled = True
    TXT_Observacoes.Enabled = True
    If sPouR = "P" Then
        NOMEAPLIC = "Duplicatas à Pagar"
        Me.Caption = NOMEAPLIC
    ElseIf sPouR = "R" Then
        NOMEAPLIC = "Duplicatas à Receber"
        Me.Caption = NOMEAPLIC
    End If
    bEdicao = True
End Sub
Public Static Sub Conta_Nova(sPouR As String)
    CB_Movimento.Enabled = True
    CB_Banco.Enabled = True
    TXT_Historico.Enabled = True
    TXT_DataEmissao.Enabled = True
    TXT_Valor.Enabled = True
    TXT_NumDoc.Enabled = True
    TXT_DataVencimento.Enabled = True
    TXT_SeuNum.Enabled = True
    TXT_NossoNum.Enabled = True
    TXT_Observacoes.Enabled = True
    If sPouR = "P" Then
        NOMEAPLIC = "Contas Diversas à Pagar"
        Me.Caption = NOMEAPLIC
    ElseIf sPouR = "R" Then
        NOMEAPLIC = "Contas Diversas à Receber"
        Me.Caption = NOMEAPLIC
    End If
    bEdicao = False
End Sub
Public Static Sub Conta_Edicao(sPouR As String)
    CB_Movimento.Enabled = False
    CB_Banco.Enabled = True
    TXT_Historico.Enabled = True
    TXT_DataEmissao.Enabled = True
    TXT_Valor.Enabled = True
    TXT_NumDoc.Enabled = True
    TXT_DataVencimento.Enabled = True
    TXT_SeuNum.Enabled = True
    TXT_NossoNum.Enabled = True
    TXT_Observacoes.Enabled = True
    If sPouR = "P" Then
        NOMEAPLIC = "Contas Diversas à Pagar"
        Me.Caption = NOMEAPLIC
    ElseIf sPouR = "R" Then
        NOMEAPLIC = "Contas Diversas à Receber"
        Me.Caption = NOMEAPLIC
    End If
    bEdicao = True
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    If Estado = True Then
        Me.MousePointer = vbHourglass
        Me.Enabled = False
    Else
        Me.MousePointer = vbDefault
        Me.Enabled = True
    End If
End Sub
Private Static Function PegaDadosDigitados(ByRef Controle As Control) As String
    If Controle.Text = "" Then
        PegaDadosDigitados = ""
    Else
        PegaDadosDigitados = Controle.Text
    End If
End Function
