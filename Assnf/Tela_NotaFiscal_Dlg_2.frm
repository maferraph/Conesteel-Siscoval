VERSION 5.00
Begin VB.Form Tela_NotaFiscal_Dlg_2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inserir ítens manualmente"
   ClientHeight    =   3375
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5655
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_1 
      Caption         =   "Dados do Produto:"
      Height          =   3372
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   5652
      Begin VB.TextBox TXT_RBC 
         Height          =   288
         Left            =   960
         MaxLength       =   10
         TabIndex        =   32
         ToolTipText     =   "Digite a porcentagem da Redução da Base de Cálculo do I.C.M.S."
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox TXT_PesoTotal 
         Height          =   288
         Left            =   1560
         TabIndex        =   12
         ToolTipText     =   "Valor total do I.C.M.S. deste produto"
         Top             =   2760
         Width           =   1332
      End
      Begin VB.TextBox TXT_PesoUnitario 
         Height          =   288
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Valor total do I.C.M.S. deste produto"
         Top             =   2760
         Width           =   1332
      End
      Begin VB.CommandButton BT_Inserir 
         Caption         =   "&Inserir"
         Height          =   732
         Left            =   3960
         Picture         =   "Tela_NotaFiscal_Dlg_2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Insere o produto na nota fiscal"
         Top             =   2520
         Width           =   732
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "Apa&gar"
         Height          =   732
         Left            =   3120
         Picture         =   "Tela_NotaFiscal_Dlg_2.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Apaga todos os campos"
         Top             =   2520
         Width           =   732
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   4800
         Picture         =   "Tela_NotaFiscal_Dlg_2.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Volta a tela da nota fiscal."
         Top             =   2520
         Width           =   732
      End
      Begin VB.TextBox TXT_ValorIPI 
         Height          =   288
         Left            =   4680
         TabIndex        =   13
         ToolTipText     =   "Valor total do I.P.I. deste produto"
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox TXT_ValorICMS 
         Height          =   288
         Left            =   3600
         TabIndex        =   10
         ToolTipText     =   "Valor total do I.C.M.S. deste produto"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox TXT_BaseCalcICMS 
         Height          =   288
         Left            =   2160
         TabIndex        =   9
         ToolTipText     =   "Digite a base de cálculo do I.C.M.S. do produto"
         Top             =   2040
         Width           =   1332
      End
      Begin VB.TextBox TXT_PorcIPI 
         Height          =   288
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   8
         ToolTipText     =   "Digite a porcentagem do I.P.I. do produto"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox TXT_PorcICMS 
         Height          =   288
         Left            =   120
         MaxLength       =   2
         TabIndex        =   7
         ToolTipText     =   "Digite a porcentagem do I.C.M.S. do produto"
         Top             =   2040
         Width           =   732
      End
      Begin VB.TextBox TXT_PrecoTotal 
         Height          =   288
         Left            =   4200
         TabIndex        =   6
         ToolTipText     =   "Preço total do produto"
         Top             =   1320
         Width           =   1332
      End
      Begin VB.TextBox TXT_PrecoUnitario 
         Height          =   288
         Left            =   2760
         TabIndex        =   5
         ToolTipText     =   "Digite o valor unitário"
         Top             =   1320
         Width           =   1332
      End
      Begin VB.TextBox TXT_Quantidade 
         Height          =   288
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "Digite a quantidade"
         Top             =   1320
         Width           =   972
      End
      Begin VB.TextBox TXT_Unidade 
         Height          =   288
         Left            =   1080
         TabIndex        =   3
         ToolTipText     =   "Digite a unidade do produto"
         Top             =   1320
         Width           =   492
      End
      Begin VB.TextBox TXT_ST 
         Height          =   288
         Left            =   600
         MaxLength       =   3
         TabIndex        =   2
         ToolTipText     =   "Digite a situação tributária do produto"
         Top             =   1320
         Width           =   372
      End
      Begin VB.TextBox TXT_CF 
         Height          =   288
         Left            =   120
         MaxLength       =   1
         TabIndex        =   1
         ToolTipText     =   "Digite a classificação fiscal do produto"
         Top             =   1320
         Width           =   372
      End
      Begin VB.TextBox TXT_Descricao 
         Height          =   288
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Digite os dados do produto"
         Top             =   600
         Width           =   5412
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "% RBC"
         Height          =   195
         Left            =   960
         TabIndex        =   33
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label LB_PesoTotal 
         AutoSize        =   -1  'True
         Caption         =   "Peso Total:"
         Height          =   192
         Left            =   1560
         TabIndex        =   31
         Top             =   2520
         Width           =   828
      End
      Begin VB.Label LB_PesoUnitario 
         AutoSize        =   -1  'True
         Caption         =   "Peso Unitário:"
         Height          =   192
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   1008
      End
      Begin VB.Label LB_ValorIPI 
         AutoSize        =   -1  'True
         Caption         =   "Valor I.P.I.:"
         Height          =   195
         Left            =   4680
         TabIndex        =   29
         Top             =   1800
         Width           =   750
      End
      Begin VB.Label LB_ValorICMS 
         AutoSize        =   -1  'True
         Caption         =   "Valor I.C.M.S.:"
         Height          =   195
         Left            =   3600
         TabIndex        =   28
         Top             =   1800
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base Calc.ICMS:"
         Height          =   195
         Left            =   2160
         TabIndex        =   27
         Top             =   1800
         Width           =   1200
      End
      Begin VB.Label LB_PorcIPI 
         AutoSize        =   -1  'True
         Caption         =   "% I.P.I.:"
         Height          =   195
         Left            =   1560
         TabIndex        =   26
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label LB_PorcICMS 
         AutoSize        =   -1  'True
         Caption         =   "% I.C.M.S.:"
         Height          =   192
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   744
      End
      Begin VB.Label LB_PrecoTotal 
         AutoSize        =   -1  'True
         Caption         =   "Preço Total:"
         Height          =   192
         Left            =   4200
         TabIndex        =   24
         Top             =   1080
         Width           =   876
      End
      Begin VB.Label LB_PrecoUnitario 
         AutoSize        =   -1  'True
         Caption         =   "Preço Unitário:"
         Height          =   192
         Left            =   2760
         TabIndex        =   23
         Top             =   1080
         Width           =   1056
      End
      Begin VB.Label LB_Quantidade 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   192
         Left            =   1680
         TabIndex        =   22
         Top             =   1080
         Width           =   876
      End
      Begin VB.Label LB_Unidade 
         AutoSize        =   -1  'True
         Caption         =   "Unid.:"
         Height          =   192
         Left            =   1080
         TabIndex        =   21
         Top             =   1080
         Width           =   408
      End
      Begin VB.Label LB_ST 
         AutoSize        =   -1  'True
         Caption         =   "S.T.:"
         Height          =   192
         Left            =   600
         TabIndex        =   20
         Top             =   1080
         Width           =   324
      End
      Begin VB.Label LB_CF 
         AutoSize        =   -1  'True
         Caption         =   "C.F.:"
         Height          =   192
         Left            =   120
         TabIndex        =   19
         Top             =   1080
         Width           =   312
      End
      Begin VB.Label LB_Descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   192
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   780
      End
   End
End
Attribute VB_Name = "Tela_NotaFiscal_Dlg_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NOMEAPLIC As String = "Assistente Manual de Ítens"
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.Text = ""
    TXT_CF.Text = ""
    TXT_ST.Text = ""
    TXT_Unidade.Text = ""
    TXT_PrecoTotal.Text = ""
    TXT_PrecoUnitario.Text = ""
    TXT_Quantidade.Text = ""
    TXT_PorcICMS.Text = ""
    TXT_PorcIPI.Text = ""
    TXT_BaseCalcICMS.Text = ""
    TXT_RBC.Text = ""
    TXT_ValorICMS.Text = ""
    TXT_ValorIPI.Text = ""
    TXT_PesoUnitario.Text = ""
    TXT_PesoTotal.Text = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    BT_Apagar.Value = True
    Tela_NotaFiscal_Dlg_2.Hide
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Inserir_Click()
    On Error GoTo ERRO_SISCOVAL
    Dim NumLinha, NumErro As Integer
    NumLinha = NumErro = 0
    
    If TXT_Descricao.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a descrição da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_Descricao.SetFocus
        Exit Sub
    ElseIf TXT_CF.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a classificação fiscal da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_CF.SetFocus
        Exit Sub
    ElseIf TXT_ST.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a situação tributária da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_ST.SetFocus
        Exit Sub
    ElseIf TXT_ST.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a situação tributária da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_ST.SetFocus
        Exit Sub
    ElseIf TXT_Unidade.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a unidade da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_Unidade.SetFocus
        Exit Sub
    ElseIf TXT_Quantidade.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a quantidade da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_Quantidade.SetFocus
        Exit Sub
    ElseIf TXT_PrecoUnitario.Text = "" Then
        RespMsg = MsgBox("Não foi digitado o preço unitário da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PrecoUnitario.SetFocus
        Exit Sub
    ElseIf TXT_PrecoTotal.Text = "" Then
        RespMsg = MsgBox("Não foi digitado o preço total da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PrecoTotal.SetFocus
        Exit Sub
    ElseIf TXT_PorcICMS.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a porcetagem do I.C.M.S. da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PorcICMS.SetFocus
        Exit Sub
    ElseIf TXT_PorcIPI.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a porcetagem do I.P.I. da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PorcIPI.SetFocus
        Exit Sub
    ElseIf TXT_BaseCalcICMS.Text = "" Then
        RespMsg = MsgBox("Não foi digitado a base de cálculo do I.C.M.S. da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_BaseCalcICMS.SetFocus
        Exit Sub
    ElseIf TXT_ValorICMS.Text = "" Then
        RespMsg = MsgBox("Não foi digitado o valor do I.C.M.S. da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_ValorICMS.SetFocus
        Exit Sub
    ElseIf TXT_ValorIPI.Text = "" Then
        RespMsg = MsgBox("Não foi digitado o valor do I.P.I. da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_ValorIPI.SetFocus
        Exit Sub
    ElseIf TXT_PesoUnitario.Text = "" Then
        RespMsg = MsgBox("Não foi digitado o peso unitário da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PesoUnitario.SetFocus
        Exit Sub
    ElseIf TXT_PesoTotal.Text = "" Then
        RespMsg = MsgBox("Não foi digitado o peso total da peça.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PesoTotal.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_Quantidade.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um número válido.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_Quantidade.SelLength = Len(Trim(TXT_Quantidade.Text))
        TXT_Quantidade.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_PrecoUnitario.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um preço unitário válido.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PrecoUnitario.SelLength = Len(Trim(TXT_PrecoUnitario.Text))
        TXT_PrecoUnitario.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_PrecoTotal.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um preço total válido.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PrecoTotal.SelLength = Len(Trim(TXT_PrecoTotal.Text))
        TXT_PrecoTotal.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_PorcICMS.Text) = False Then
        RespMsg = MsgBox("Não foi digitado uma porcentagem do I.C.M.S. válida.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PorcICMS.SelLength = Len(Trim(TXT_PorcICMS.Text))
        TXT_PorcICMS.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_PorcIPI.Text) = False Then
        RespMsg = MsgBox("Não foi digitado uma porcentagem do I.P.I. válida.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PorcIPI.SelLength = Len(Trim(TXT_PorcIPI.Text))
        TXT_PorcIPI.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_BaseCalcICMS.Text) = False Then
        RespMsg = MsgBox("Não foi digitado uma base de cálculo do I.C.M.S. válida.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_BaseCalcICMS.SelLength = Len(Trim(TXT_BaseCalcICMS.Text))
        TXT_BaseCalcICMS.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_ValorICMS.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um valor do I.C.M.S. válido.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_ValorICMS.SelLength = Len(Trim(TXT_ValorICMS.Text))
        TXT_ValorICMS.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_ValorIPI.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um valor do I.P.I. válido.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_ValorIPI.SelLength = Len(Trim(TXT_ValorIPI.Text))
        TXT_ValorIPI.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_PesoUnitario.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um peso unitário válido.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PesoUnitario.SelLength = Len(Trim(TXT_PesoUnitario.Text))
        TXT_PesoUnitario.SetFocus
        Exit Sub
    ElseIf IsNumeric(TXT_PesoTotal.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um peso total válido.", vbOKOnly, "Assistente Manual de Produtos")
        TXT_PesoTotal.SelLength = Len(Trim(TXT_PesoTotal.Text))
        TXT_PesoTotal.SetFocus
        Exit Sub
    End If
    
    For I = 1 To 20
        If Tela_NotaFiscal.FG_1.TextMatrix(I, 1) = "" And _
           Tela_NotaFiscal.FG_1.TextMatrix(I, 2) = "" Then
            NumLinha = I
            Exit For
        Else
            If I = 20 Then
                NumErro = 21
            End If
        End If
    Next I
    If NumErro = 21 Then
        RespMsg = MsgBox("Não existem mais linha em branco para serem preenchidas no quadro dados do produto na nota fiscal.", vbOKOnly, "Assistente Manual de Produtos")
        Tela_NotaFiscal_Dlg_2.Hide
        Exit Sub
    End If

    'Funçao que envia dados ao assistente da NF
    EnviaItemNF2 (NumLinha)
    
    If NumLinha = 20 Then
        Tela_NotaFiscal_Dlg_2.Hide
        Tela_NotaFiscal.BT_InserirItem.Enabled = False
    End If
    BT_Apagar.Value = True
    TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Activate()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    DLL_FUNCS.RegistraEvento "Abrir Inserir Ítem Manual para Notas Fiscais", Tela_NotaFiscal.TXT_NF.Text
    BT_Apagar.Value = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseCalcICMS_Change()
    On Error GoTo ERRO_SISCOVAL
    If IsNumeric(TXT_PorcICMS.Text) And TXT_PorcICMS.Text <> "" And TXT_PorcICMS.Text <> "0" Then
        If TXT_PorcICMS.Text >= 10 Then
            TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0." & TXT_PorcICMS.Text), "##,##0.00")
        ElseIf TXT_PorcICMS.Text < 10 And TXT_PorcICMS.Text > 0 Then
            TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0.0" & TXT_PorcICMS.Text), "##,##0.00")
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseCalcICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("A base de cálculo é inserida automaticamente...", vbOKOnly, "Assistente Manual de Produtos")
    TXT_BaseCalcICMS.SelLength = Len(TXT_BaseCalcICMS.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseCalcICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        PassaFoco (TXT_BaseCalcICMS.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseCalcICMS_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_BaseCalcICMS.Text = Format(TXT_BaseCalcICMS.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CF_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_CF.SelLength = Len(TXT_CF.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CF_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        PassaFoco (TXT_CF.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CF_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_CF.Text = UCase(TXT_CF.Text)
    'Procura pela RBC de ICMS da CF
    If TXT_CF.Text <> "" Then
        DLL_BD.BDSIS_TBEAL.Seek "=", "ICMS", Tela_NotaFiscal.CB_Estado.Text, TXT_CF.Text
        If DLL_BD.BDSIS_TBEAL.NoMatch Then
            TXT_RBC.Text = "0,00"
        Else
            TXT_RBC.Text = Format(DLL_BD.BDSIS_TBEAL_CPRBC.Value, "##,##0.00")
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.SelLength = Len(TXT_Descricao.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        PassaFoco (TXT_Descricao.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoTotal_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O peso total é calculado automaticamente...", vbOKOnly, "Assistente Manual de Produtos")
    TXT_PesoTotal.SelLength = Len(TXT_PesoTotal.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoTotal_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        PassaFoco (TXT_PesoTotal.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoTotal_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PesoTotal.Text = Format(TXT_PesoTotal.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoUnitario_Change()
    On Error GoTo ERRO_SISCOVAL
    If IsNumeric(TXT_PesoUnitario.Text) And _
       TXT_PesoUnitario.Text <> "" And _
       TXT_Quantidade.Text <> "" Then
        TXT_PesoTotal.Text = CDbl(TXT_PesoUnitario.Text) * CDbl(TXT_Quantidade.Text)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoUnitario_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PesoUnitario.SelLength = Len(TXT_PesoUnitario.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoUnitario_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then BT_Inserir.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoUnitario_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PesoUnitario.Text = Format(TXT_PesoUnitario.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcICMS_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_BaseCalcICMS.Text = "" Or TXT_BaseCalcICMS.Text = "0" Then Exit Sub
    If IsNumeric(TXT_PorcICMS.Text) And _
       TXT_PorcICMS.Text <> "" And _
       TXT_PorcICMS.Text <> "0" Then
        If IsNumeric(TXT_RBC.Text) And TXT_RBC.Text <> "" And Val(TXT_RBC.Text) > 0 Then
            'se tiver RBC
            Dim ValRed As Double
            ValRed = (100 - CDbl(TXT_RBC.Text)) / 100
            TXT_BaseCalcICMS.Text = Format((CDbl(TXT_PrecoTotal.Text) * ValRed), "##,##0.00")
        Else
            TXT_BaseCalcICMS.Text = TXT_PrecoTotal.Text
        End If
        If TXT_PorcICMS.Text >= 10 Then
            TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * Val("0." & Str((TXT_PorcICMS.Text))), "##,##0.00")
        ElseIf TXT_PorcICMS.Text < 10 Then
            TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * Val("0.0" & Str((TXT_PorcICMS.Text))), "##,##0.00")
        End If
    Else
        TXT_RBC.Text = "0,00"
        TXT_BaseCalcICMS.Text = "0,00"
        TXT_ValorICMS.Text = "0,00"
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PorcICMS.SelLength = Len(TXT_PorcICMS.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_PorcIPI.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcICMS_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PorcICMS.Text = Format(TXT_PorcICMS.Text, "###,###,##0")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcIPI_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_PrecoTotal.Text = "" Or TXT_PrecoTotal.Text = "0" Then Exit Sub
    If IsNumeric(TXT_PorcIPI.Text) And _
       TXT_PorcIPI.Text <> "" And _
       TXT_PorcIPI.Text <> "0" Then
        If TXT_PorcIPI.Text >= 10 Then
            TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * Val("0." & Str((TXT_PorcIPI.Text))), "##,##0.00")
        ElseIf TXT_PorcIPI.Text < 10 Then
            TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * Val("0.0" & Str((TXT_PorcIPI.Text))), "##,##0.00")
        End If
    Else
        TXT_ValorIPI.Text = "0,00"
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcIPI_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PorcIPI.SelLength = Len(TXT_PorcIPI.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcIPI_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_PesoUnitario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcIPI_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PorcIPI.Text = Format(TXT_PorcIPI.Text, "###,###,##0")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoTotal_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_PrecoTotal.Text <> "" Then
        TXT_BaseCalcICMS.Text = Format(TXT_PrecoTotal.Text, "###,##0.00")
        If IsNumeric(TXT_PorcIPI.Text) And TXT_PorcIPI.Text <> "" And TXT_PorcIPI.Text <> "0" Then
            If TXT_PorcIPI.Text >= 10 Then
                TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * CDbl("0." & TXT_PorcIPI.Text), "##,##0.00")
            ElseIf TXT_PorcIPI.Text < 10 Then
                TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * CDbl("0.0" & TXT_PorcIPI.Text), "##,##0.00")
            End If
        End If
        If IsNumeric(TXT_PorcICMS.Text) And TXT_PorcICMS.Text <> "" And TXT_PorcICMS.Text <> "0" Then
            If TXT_PorcICMS.Text >= 10 Then
                TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0." & TXT_PorcICMS.Text), "##,##0.00")
            ElseIf TXT_PorcICMS.Text < 10 Then
                TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0.0" & TXT_PorcICMS.Text), "##,##0.00")
            End If
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoTotal_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O Preço Total é calculado automaticamente...", vbOKOnly, "Assistente Manual de Produtos")
    TXT_PrecoTotal.SelLength = Len(TXT_PrecoTotal.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoTotal_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_PorcICMS.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoTotal_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PrecoTotal.Text = Format(TXT_PrecoTotal.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoUnitario_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Quantidade.Text = "" Or TXT_Quantidade.Text = "0" Then Exit Sub
    If IsNumeric(TXT_PrecoUnitario.Text) And TXT_PrecoUnitario.Text <> "" Then
        TXT_PrecoTotal.Text = Format(CDbl(TXT_Quantidade.Text) * CDbl(TXT_PrecoUnitario.Text), "###,##0.00")
        TXT_BaseCalcICMS.Text = Format(TXT_PrecoTotal.Text, "###,##0.00")
        If IsNumeric(TXT_PorcIPI.Text) And TXT_PorcIPI.Text <> "" And TXT_PorcIPI.Text <> "0" Then
            If TXT_PorcIPI.Text >= 10 Then
                TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * CDbl("0." & Str((TXT_PorcIPI.Text))), "##,##0.00")
            ElseIf TXT_PorcIPI.Text < 10 Then
                TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * CDbl("0.0" & Str((TXT_PorcIPI.Text))), "##,##0.00")
            End If
        End If
        If IsNumeric(TXT_PorcICMS.Text) And TXT_PorcICMS.Text <> "" And TXT_PorcICMS.Text <> "0" Then
            If TXT_PorcICMS.Text >= 10 Then
                TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0." & Str((TXT_PorcICMS.Text))), "##,##0.00")
            ElseIf TXT_PorcICMS.Text < 10 Then
                TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0.0" & Str((TXT_PorcICMS.Text))), "##,##0.00")
            End If
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoUnitario_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PrecoUnitario.SelLength = Len(TXT_PrecoUnitario.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoUnitario_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_PorcICMS.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoUnitario_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PrecoUnitario.Text = Format(TXT_PrecoUnitario.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Quantidade_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Quantidade.Text = "" Or TXT_Quantidade.Text = "0" Then
        TXT_PrecoTotal.Text = "0,00"
        Exit Sub
    End If
    If IsNumeric(TXT_PrecoUnitario.Text) And TXT_PrecoUnitario.Text <> "" And TXT_PrecoUnitario.Text <> "0" Then
        TXT_PrecoTotal.Text = Format(CDbl(TXT_Quantidade.Text) * CDbl(TXT_PrecoUnitario.Text), "###,##0.00")
        TXT_BaseCalcICMS.Text = Format(TXT_PrecoTotal.Text, "###,##0.00")
        If IsNumeric(TXT_PorcIPI.Text) And TXT_PorcIPI.Text <> "" And TXT_PorcIPI.Text <> "0" Then
            If TXT_PorcIPI.Text >= 10 Then
                TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * CDbl("0." & TXT_PorcIPI.Text), "##,##0.00")
            ElseIf TXT_PorcIPI.Text < 10 Then
                TXT_ValorIPI.Text = Format(CDbl(TXT_PrecoTotal.Text) * CDbl("0.0" & TXT_PorcIPI.Text), "##,##0.00")
            End If
        End If
        If IsNumeric(TXT_PorcICMS.Text) And TXT_PorcICMS.Text <> "" And TXT_PorcICMS.Text <> "0" Then
            If TXT_PorcICMS.Text >= 10 Then
                TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0." & TXT_PorcICMS.Text), "##,##0.00")
            ElseIf TXT_PorcICMS.Text < 10 Then
                TXT_ValorICMS.Text = Format(CDbl(TXT_BaseCalcICMS.Text) * CDbl("0.0" & TXT_PorcICMS.Text), "##,##0.00")
            End If
        End If
        If TXT_PesoUnitario.Text <> "" And TXT_PesoUnitario.Text <> "0" And _
           TXT_Quantidade.Text <> "" And TXT_Quantidade.Text <> "0" Then
            TXT_PesoTotal.Text = CDbl(TXT_PesoUnitario.Text) * CDbl(TXT_Quantidade.Text)
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Quantidade_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Quantidade.SelLength = Len(TXT_Quantidade.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Quantidade_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        PassaFoco (TXT_Quantidade.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_RBC_Change()
    TXT_PorcICMS_Change
End Sub
Private Sub TXT_RBC_GotFocus()
    TXT_RBC.SelLength = Len(TXT_RBC.Text)
End Sub
Private Sub TXT_RBC_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_PorcIPI.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_RBC_LostFocus()
    If Val(TXT_RBC.Text) >= 100 Then
        MsgBox "Digite um valor menor que 100% para a redução da base de cálculo do ICMS.", vbCritical + vbOKOnly, NOMEAPLIC
    End If
    TXT_PorcICMS_Change
End Sub
Private Sub TXT_ST_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ST.SelLength = Len(TXT_ST.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ST_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        PassaFoco (TXT_ST.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Unidade_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Unidade.SelLength = Len(TXT_Unidade.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Unidade_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        PassaFoco (TXT_Unidade.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O valor do I.C.M.S. é calculado automaticamente...", vbOKOnly, "Assistente Manual de Produtos")
    TXT_ValorICMS.SelLength = Len(TXT_ValorICMS.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        PassaFoco (TXT_ValorICMS.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorICMS.Text = Format(TXT_ValorICMS.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorIPI_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O valor do I.P.I. é calculado automaticamente...", vbOKOnly, "Assistente Manual de Produtos")
    TXT_ValorIPI.SelLength = Len(TXT_ValorIPI.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorIPI_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        PassaFoco (TXT_ValorIPI.TabIndex)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorIPI_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorIPI.Text = Format(TXT_ValorIPI.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub PassaFoco(IndiceTab As Integer)
    On Error GoTo ERRO_SISCOVAL
    Dim NumTab As Integer
    NumTab = IndiceTab + 1
    For I = 0 To Tela_NotaFiscal_Dlg_2.Controls.Count - 1
        If Tela_NotaFiscal_Dlg_2.Controls(I).TabIndex = NumTab Then
            Tela_NotaFiscal_Dlg_2.Controls(I).SetFocus
            Exit Sub
        End If
    Next I
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub EnviaItemNF2(NumLin As Integer)
    On Error GoTo ERRO_SISCOVAL
    Dim Linha As Integer
    Linha = NumLin
    'Verifica se a descricao é maior que 38 caracteres
    If Len(TXT_Descricao.Text) <= 38 Then
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = TXT_Descricao.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 3) = TXT_CF.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 4) = TXT_ST.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 5) = TXT_Unidade.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 6) = Format(TXT_Quantidade.Text, "###,##0")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 7) = Format(TXT_PrecoUnitario.Text, "###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 8) = Format(TXT_PrecoTotal.Text, "###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 9) = Format(TXT_PorcICMS.Text, "##0")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 10) = Format(TXT_PorcIPI.Text, "##0")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 11) = Format(TXT_ValorIPI.Text, "##,##0.00")
        
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "MN"
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 3) = Format(TXT_BaseCalcICMS.Text, "##,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 4) = Format(TXT_ValorICMS.Text, "##,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 5) = Format(TXT_PesoTotal.Text, "###,##0.00")
    Else
        Dim NumLinLivre As Integer
        NumLinLivre = 0
        'verifica quantas linhas ainda estão livres
        For I = NumLinLivre To 20
            If Tela_NotaFiscal.FG_1.TextMatrix(I, 1) = "" And _
               Tela_NotaFiscal.FG_1.TextMatrix(I, 2) = "" Then
                NumLinLivre = NumLinLivre + 1
            End If
        Next I
        'funcao de divisao de linhas... permite até 5 linhas
        Dim TamDes As Integer
        TamDes = Len(TXT_Descricao.Text)
        Dim DLR, DL1, DL2, DL3, DL4, DL5 As String
        DLR = TXT_Descricao.Text
        DL1 = ""
        DL2 = ""
        DL3 = ""
        DL4 = ""
        DL5 = ""
        Do
            If Len(DLR) > 38 Then
                For I = 38 To 1 Step -1
                    If Mid(DLR, I, 1) = " " Then
                        If DL1 = "" Then
                            DL1 = Left(DLR, I)
                            DLR = Mid(DLR, I + 1, TamDes - I)
                            Exit For
                        ElseIf DL2 = "" Then
                            DL2 = Left(DLR, I)
                            DLR = Mid(DLR, I + 1, TamDes - I)
                            Exit For
                        ElseIf DL3 = "" Then
                            DL3 = Left(DLR, I)
                            DLR = Mid(DLR, I + 1, TamDes - I)
                            Exit For
                        ElseIf DL4 = "" Then
                            DL4 = Left(DLR, I)
                            DLR = Mid(DLR, I + 1, TamDes - I)
                            Exit For
                        ElseIf DL5 = "" Then
                            DL5 = Left(DLR, I)
                            DLR = Mid(DLR, I + 1, TamDes - I)
                            Exit For
                        End If

End If
                Next I
            Else
                If DL2 = "" Then
                    DL2 = DLR
                ElseIf DL3 = "" Then
                
                    DL3 = DLR
                ElseIf DL4 = "" Then
                    DL4 = DLR
                ElseIf DL5 = "" Then
                    DL5 = DLR
                End If
                Exit Do
            End If
        Loop
        'confirma se há possibilidade de inserir essas linhas
        Dim LinErro As Boolean
        LinErro = False
        If DL2 <> "" And _
           DL3 = "" And _
           NumLinLivre < 2 Then
            LinErro = True
        ElseIf DL3 <> "" And _
           DL4 = "" And _
           NumLinLivre < 3 Then
            LinErro = True
        ElseIf DL4 <> "" And _
           DL5 = "" And _
           NumLinLivre < 4 Then
            LinErro = True
        ElseIf DL5 <> "" And _
           NumLinLivre < 5 Then
            LinErro = True
        End If
        If LinErro = True Then
            RespMsg = MsgBox("Como a descrição do produto ultrapassou o número de caracteres de uma linha, o sistema foi obrigado à dividi-lá em várias linhas, porém não existem mais linhas suficientes para esta operação. Não será possível inserir este ítem.", vbOKOnly, "Assistente Manual de Produtos")
            Tela_NotaFiscal_Dlg_2.BT_Cancelar.SetFocus
            Exit Sub
        End If
        'Insere itens
        If DL2 <> "" And DL3 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "MN"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        ElseIf DL3 <> "" And DL4 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "MN"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL3
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        ElseIf DL4 <> "" And DL5 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "MN"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL3
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL4
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        ElseIf DL5 <> "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "MN"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL3
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL4
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL5
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        End If

        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 3) = TXT_CF.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 4) = TXT_ST.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 5) = TXT_Unidade.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 6) = Format(TXT_Quantidade.Text, "###,##0")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 7) = Format(TXT_PrecoUnitario.Text, "###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 8) = Format(TXT_PrecoTotal.Text, "###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 9) = Format(TXT_PorcICMS.Text, "##0")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 10) = Format(TXT_PorcIPI.Text, "##0")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 11) = Format(TXT_ValorIPI.Text, "##,##0.00")
        
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 3) = Format(TXT_BaseCalcICMS.Text, "##,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 4) = Format(TXT_ValorICMS.Text, "##,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 5) = Format(TXT_PesoTotal.Text, "###,##0.00")
End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
