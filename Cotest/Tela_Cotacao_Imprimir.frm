VERSION 5.00
Begin VB.Form Tela_Cotacao_Imprimir 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione os dados abaixo para imprimir a Cotação de Preços:"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6870
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   5880
      Picture         =   "Tela_Cotacao_Imprimir.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Volta à Tela de Cotações"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton BT_Email 
      Caption         =   "&E-mail"
      Height          =   855
      Left            =   4800
      Picture         =   "Tela_Cotacao_Imprimir.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Manda um e-mail para o cliente"
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton BT_Fax 
      Caption         =   "&Fax"
      Height          =   855
      Left            =   3720
      Picture         =   "Tela_Cotacao_Imprimir.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Manda um fax para o cliente"
      Top             =   2640
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   855
      Left            =   2760
      Picture         =   "Tela_Cotacao_Imprimir.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprimir Cotação"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox TXT_Vias 
      Height          =   285
      Left            =   720
      TabIndex        =   9
      ToolTipText     =   "Digite aqui o número de vias que você deseja imprimir"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame FR 
      Height          =   2535
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   6855
      Begin VB.TextBox TXT_Email 
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         ToolTipText     =   "Digite aqui o E-mail do Contato"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox TXT_Fax 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         ToolTipText     =   "Digite aqui o número do Fax do Contato"
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox TXT_Dados 
         Height          =   885
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Digite aqui Dados Adicionais sobre a Cotação"
         Top             =   1560
         Width           =   4575
      End
      Begin VB.ComboBox CB_Descricao 
         Height          =   315
         ItemData        =   "Tela_Cotacao_Imprimir.frx":0FD0
         Left            =   4560
         List            =   "Tela_Cotacao_Imprimir.frx":0FDD
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Digite aqui o Departamento ou Cargo do Vendedor"
         Top             =   360
         Width           =   2175
      End
      Begin VB.ComboBox CB_Vendedor 
         Height          =   315
         ItemData        =   "Tela_Cotacao_Imprimir.frx":1017
         Left            =   2160
         List            =   "Tela_Cotacao_Imprimir.frx":102A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Digite aqui o nome do Vendedor"
         Top             =   360
         Width           =   2295
      End
      Begin VB.ComboBox CB_Frete 
         Height          =   315
         ItemData        =   "Tela_Cotacao_Imprimir.frx":10A5
         Left            =   120
         List            =   "Tela_Cotacao_Imprimir.frx":10AF
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Digite aqui o tipo do Frete"
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox TXT_Ramal 
         Height          =   285
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Digite aqui o Ramal do Contato"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox TXT_Contato 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Digite aqui o nome do Contato"
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox CB_Depto 
         Height          =   315
         ItemData        =   "Tela_Cotacao_Imprimir.frx":10CF
         Left            =   120
         List            =   "Tela_Cotacao_Imprimir.frx":10D9
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Digite aqui o nome do Departamento"
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Departamento:"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número do Fax:"
         Height          =   195
         Left            =   2160
         TabIndex        =   23
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Dados Adicionais:"
         Height          =   195
         Left            =   2160
         TabIndex        =   21
         Top             =   1320
         Width           =   1275
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   4560
         TabIndex        =   20
         Top             =   120
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Vendedor:"
         Height          =   195
         Left            =   2160
         TabIndex        =   19
         Top             =   120
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Frete:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1920
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ramal:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
         Height          =   195
         Left            =   4560
         TabIndex        =   15
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Número de Vias:"
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   2880
      Width           =   1170
   End
End
Attribute VB_Name = "Tela_Cotacao_Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Imprimir Cotação de Estoque"
Dim I As Integer, nLin As Integer, RespMsg

Private Sub BT_Imprimir_Click()
    If TXT_Vias.Text = "" Then Exit Sub
    If IsNumeric(TXT_Vias.Text) = False Then Exit Sub
    If Val(TXT_Vias.Text) < 1 Then Exit Sub
    RespMsg = MsgBox("Você deseja imprimir " & Str(TXT_Vias.Text) & " via(s) da cotação de preço ?", vbQuestion + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then MontaCotacao Val(TXT_Vias.Text), 0
    Unload Me
End Sub
Private Sub BT_Voltar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    CB_Depto.Text = "Vendas"
    TXT_Contato.Text = Tela_Cotacao.TXT_Contato.Text
    TXT_Ramal.Text = ""
    CB_Frete.ListIndex = 0
    CB_Vendedor.Text = "Marcos Toyoki Hayashi"
    CB_Descricao.ListIndex = 0
    TXT_Fax.Text = Tela_Cotacao.sFax
    TXT_Email.Text = Tela_Cotacao.sEmail
    TXT_Vias.Text = "1"
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Sub MontaCotacao(NumVias As Integer, Destino As Integer)
    If NumVias < 1 Then Exit Sub
    'Destino 0: Impressora, Destino 1: Fax, Destino 2: E-mail
    Dim lTeste As Boolean, lFuncTeste As Boolean
    lTeste = True
    lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_LimpaItens
    If lFuncTeste = False Then lTeste = False
    'monta cotacao
    With Tela_Cotacao.DLL_BD
        .BDSIS_TBEMP.Seek "=", Tela_Cotacao.TXT_Empresa.Text 'procura dados sobre a empresa
        If .BDSIS_TBEMP.NoMatch Then
            MsgBox "Não foi possível localizar os dados sobre a empresa desta cotação. Tente mais tarde.", vbCritical + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        Dim sEmpresa As String, sCGC As String, sINE As String, sFone As String, sFax As String, sEndereco As String, sBairro As String, sCEP As String, sCidade As String, sEsta As String
        sEmpresa = ""
        sCGC = ""
        sINE = ""
        sEndereco = ""
        sBairro = ""
        sCEP = ""
        sCidade = ""
        sEsta = ""
        sFone = ""
        sFax = ""
        If IsNull(.BDSIS_TBEMP_CPEMP.Value) = False Then sEmpresa = .BDSIS_TBEMP_CPEMP.Value
        If IsNull(.BDSIS_TBEMP_CPCGC.Value) = False Then sCGC = .BDSIS_TBEMP_CPCGC.Value
        If IsNull(.BDSIS_TBEMP_CPINE.Value) = False Then sINE = .BDSIS_TBEMP_CPINE.Value
        If IsNull(.BDSIS_TBEMP_CPEND.Value) = False Then sEndereco = .BDSIS_TBEMP_CPEND.Value
        If IsNull(.BDSIS_TBEMP_CPBAI.Value) = False Then sBairro = .BDSIS_TBEMP_CPBAI.Value
        If IsNull(.BDSIS_TBEMP_CPCEP.Value) = False Then sCEP = .BDSIS_TBEMP_CPCEP.Value
        If IsNull(.BDSIS_TBEMP_CPCID.Value) = False Then sCidade = .BDSIS_TBEMP_CPCID.Value
        If IsNull(.BDSIS_TBEMP_CPEST.Value) = False Then sEsta = .BDSIS_TBEMP_CPEST.Value
        If IsNull(.BDSIS_TBEMP_CPFON.Value) = False Then sFone = .BDSIS_TBEMP_CPFON.Value
        If IsNull(.BDSIS_TBEMP_CPFAX.Value) = False Then sFax = .BDSIS_TBEMP_CPFAX.Value
        'cabeçalho
        lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_Cabecalho(Tela_Cotacao.LT_NumCot.Text, Tela_Cotacao.TXT_Data.Text, sEmpresa, sCGC, sINE, sEndereco, sBairro, sCEP, sCidade, sEsta, sFone, sFax, CB_Depto.Text, TXT_Contato.Text, TXT_Ramal.Text)
        If lFuncTeste = False Then lTeste = False
        'itens
        Dim nInd As Integer
        nLin = 0
        nInd = 1
        With Tela_Cotacao.FG
            For I = 1 To (.Rows - 1)
                If (Len(Trim(.TextMatrix(I, 5))) + Len(Trim(.TextMatrix(I, 6)))) > 70 Then
                    lFuncTeste = MontaCotacao_DivideDescricao(Trim(Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6))), nInd)
                    If lFuncTeste = False Then lTeste = False
                Else
                    lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_Itens(nLin, Tela_Cotacao.DLL_FUNCS.PegaNumeroItem(nInd), Tela_Cotacao.DLL_FUNCS.PegaUnidade(CDbl(.TextMatrix(I, 1)), 0), .TextMatrix(I, 2), Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6)), Format(Trim(.TextMatrix(I, 13)), "###,###,###,##0.00"), .TextMatrix(I, 16), .TextMatrix(I, 15), .TextMatrix(I, 14))
                    If lFuncTeste = False Then lTeste = False
                End If
                nLin = nLin + 1
                nInd = nInd + 1
            Next I
        End With
        'rodape
        lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_Rodape(PegaCondPagto, Tela_Cotacao.TXT_Transportadora.Text, CB_Frete.Text, CB_Vendedor.Text, CB_Descricao.Text, TXT_Dados.Text)
        If lFuncTeste = False Then lTeste = False
    End With
    'imprimi
    If Destino = 0 Then  'imprimir
        lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_Imprimir(Tela_Cotacao.DLL_FUNCS.NomeImpressora("IT_CotacaoEstoque"), NumVias)
        If lFuncTeste = False Then lTeste = False
    ElseIf Destino = 1 Then  'fax

    ElseIf Destino = 2 Then  'e-mail
    
    End If
End Sub
Private Function MontaCotacao_DivideDescricao(Texto As String, Indice As Integer) As Boolean
    MontaCotacao_DivideDescricao = False
    Dim sTmp1 As String, sTmp2 As String, nTmp As Integer, nNumLin As Integer, lT As Boolean, lTeste As Boolean, lFuncTeste As Boolean
    lTeste = True
    nNumLin = 1
    nTmp = 1
    sTmp1 = Texto
    Do While True
        With Tela_Cotacao.FG
            If Len(sTmp1) > 70 Then
                sTmp2 = Trim(Mid(sTmp1, (((nTmp - 1) * 70) + 1), 70))
                sTmp1 = Trim(Right(sTmp1, (Len(sTmp1) - (nTmp * 70))))
                nTmp = nTmp + 1
                If nNumLin = 1 Then
                    lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_Itens(nLin, Tela_Cotacao.DLL_FUNCS.PegaNumeroItem(Indice), Tela_Cotacao.DLL_FUNCS.PegaUnidade(CDbl(.TextMatrix(Indice, 1)), 0), .TextMatrix(Indice, 2), sTmp2, "", "", "", "")
                    If lFuncTeste = False Then lTeste = False
                Else
                    lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_Itens(nLin, "", "", "", sTmp2, "", "", "", "")
                    If lFuncTeste = False Then lTeste = False
                End If
                nNumLin = nNumLin + 1
                nLin = nLin + 1
            Else
                sTmp2 = sTmp1
                lFuncTeste = Tela_Cotacao.DLL_IMP.CotacaoEstoque_Itens(nLin, "", "", "", sTmp2, Format(Trim(.TextMatrix(Indice, 13)), "###,###,###,##0.00"), .TextMatrix(Indice, 16), .TextMatrix(Indice, 15), .TextMatrix(Indice, 14))
                Exit Do
            End If
        End With
    Loop
    MontaCotacao_DivideDescricao = lT
End Function
Private Function PegaQuantidade(Valor As Currency) As String
    If Valor <= 1 Then
        PegaQuantidade = Str(Format(Valor, "###,##0.00")) & " pç."
    Else
        PegaQuantidade = Str(Format(Valor, "###,##0.00")) & " pçs."
    End If
End Function
Private Sub DivideDescricao(Texto As String)
    Dim sTmp1 As String, sTmp2 As String, nTmp As Integer
    nTmp = 0
    sTmp1 = Texto
    Do While True
        If Len(sTmp1) > 70 Then
            sTmp2 = Mid(sTmp1, ((nTmp * 70) + 1), 70)
            sTmp1 = Right(sTmp1, (Len(sTmp1) - (nTmp * 70)))
            nTmp = nTmp + 1
            'Tela_Cotacao_IT.LB_Descricao(nLin).Caption = sTmp2
            nLin = nLin + 1
        Else
            sTmp2 = sTmp1
            'Tela_Cotacao_IT.LB_Descricao(nLin).Caption = sTmp2
            Exit Do
        End If
    Loop
End Sub
Private Static Function PegaCondPagto() As String
    With Tela_Cotacao
        If .TXT_D1.Text <> "" And .TXT_D2.Text = "" And .TXT_D3.Text = "" And .TXT_D4.Text = "" Then
            PegaCondPagto = Trim(.TXT_D1.Text) & " dd"
        ElseIf .TXT_D1.Text <> "" And .TXT_D2.Text <> "" And .TXT_D3.Text = "" And .TXT_D4.Text = "" Then
            PegaCondPagto = Trim(.TXT_D1.Text) & " dd / " & Trim(.TXT_D2.Text) & " dd"
        ElseIf .TXT_D1.Text <> "" And .TXT_D2.Text <> "" And .TXT_D3.Text <> "" And .TXT_D4.Text = "" Then
            PegaCondPagto = Trim(.TXT_D1.Text) & " dd / " & Trim(.TXT_D2.Text) & " dd / " & Trim(.TXT_D3.Text) & " dd"
        ElseIf .TXT_D1.Text <> "" And .TXT_D2.Text <> "" And .TXT_D3.Text <> "" And .TXT_D4.Text <> "" Then
            PegaCondPagto = Trim(.TXT_D1.Text) & " dd / " & Trim(.TXT_D2.Text) & " dd / " & Trim(.TXT_D3.Text) & " dd / " & Trim(.TXT_D4.Text) & " dd"
        End If
    End With
End Function
