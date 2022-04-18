VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Tela_NotaFiscal_Dlg_4 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edição de Ítens"
   ClientHeight    =   4215
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5670
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_1 
      Caption         =   "Dados do Produto:"
      Height          =   4212
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5652
      Begin VB.TextBox TXT_PesoTotal 
         Height          =   288
         Left            =   120
         MaxLength       =   20
         TabIndex        =   14
         ToolTipText     =   "Peso total parcial."
         Top             =   3840
         Width           =   1332
      End
      Begin VB.TextBox TXT_PesoUnitario 
         Height          =   288
         Left            =   120
         MaxLength       =   20
         TabIndex        =   13
         ToolTipText     =   "Digite o peso unitário deste ítem."
         Top             =   3240
         Width           =   1332
      End
      Begin VB.OptionButton RD_Cortar 
         Caption         =   "Cortar"
         Height          =   192
         Left            =   2880
         TabIndex        =   21
         ToolTipText     =   "Ao inserir os dados no assistente da N.F., a descrição será cortada (se necessário) por caracteres."
         Top             =   3960
         Width           =   852
      End
      Begin VB.OptionButton RD_Dividir 
         Caption         =   "Dividir"
         Height          =   192
         Left            =   3960
         TabIndex        =   22
         ToolTipText     =   "Ao inserir os dados no assistente da N.F., a descrição será dividida (se necessário) por palavras."
         Top             =   3960
         Width           =   852
      End
      Begin MSFlexGridLib.MSFlexGrid FG_2 
         Height          =   852
         Left            =   4320
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   1508
         _Version        =   393216
         Rows            =   0
         Cols            =   12
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         TextStyleFixed  =   4
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid FG_1 
         Height          =   852
         Left            =   4080
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   1508
         _Version        =   393216
         Rows            =   0
         Cols            =   12
         FixedRows       =   0
         FixedCols       =   0
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         TextStyleFixed  =   4
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.HScrollBar BH_1 
         Height          =   252
         Left            =   2160
         TabIndex        =   0
         Top             =   480
         Width           =   3372
      End
      Begin VB.TextBox TXT_Figura 
         Height          =   288
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Figura do ítem"
         Top             =   480
         Width           =   1692
      End
      Begin VB.TextBox TXT_Descricao 
         Height          =   612
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         ToolTipText     =   "Descrição deste ítem"
         Top             =   1080
         Width           =   5412
      End
      Begin VB.TextBox TXT_CF 
         Height          =   288
         Left            =   120
         MaxLength       =   1
         TabIndex        =   3
         ToolTipText     =   "Digite a classificação fiscal do produto"
         Top             =   2040
         Width           =   372
      End
      Begin VB.TextBox TXT_ST 
         Height          =   288
         Left            =   600
         MaxLength       =   3
         TabIndex        =   4
         ToolTipText     =   "Digite a situação tributária do produto"
         Top             =   2040
         Width           =   372
      End
      Begin VB.TextBox TXT_Unidade 
         Height          =   288
         Left            =   1080
         TabIndex        =   5
         ToolTipText     =   "Digite a unidade do produto"
         Top             =   2040
         Width           =   492
      End
      Begin VB.TextBox TXT_Quantidade 
         Height          =   288
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Digite a quantidade"
         Top             =   2040
         Width           =   972
      End
      Begin VB.TextBox TXT_PrecoUnitario 
         Height          =   288
         Left            =   2760
         TabIndex        =   7
         ToolTipText     =   "Digite o valor unitário"
         Top             =   2040
         Width           =   1332
      End
      Begin VB.TextBox TXT_PrecoTotal 
         Height          =   288
         Left            =   4200
         TabIndex        =   8
         ToolTipText     =   "Preço total do produto"
         Top             =   2040
         Width           =   1332
      End
      Begin VB.TextBox TXT_PorcICMS 
         Height          =   288
         Left            =   120
         MaxLength       =   2
         TabIndex        =   9
         ToolTipText     =   "Digite a porcentagem do I.C.M.S. do produto"
         Top             =   2640
         Width           =   732
      End
      Begin VB.TextBox TXT_PorcIPI 
         Height          =   288
         Left            =   3600
         MaxLength       =   2
         TabIndex        =   12
         ToolTipText     =   "Digite a porcentagem do I.P.I. do produto"
         Top             =   2640
         Width           =   732
      End
      Begin VB.TextBox TXT_BaseCalcICMS 
         Height          =   288
         Left            =   960
         TabIndex        =   10
         ToolTipText     =   "Digite a base de cálculo do I.C.M.S. do produto"
         Top             =   2640
         Width           =   1212
      End
      Begin VB.TextBox TXT_ValorICMS 
         Height          =   288
         Left            =   2280
         TabIndex        =   11
         ToolTipText     =   "Valor total do I.C.M.S. deste produto"
         Top             =   2640
         Width           =   1092
      End
      Begin VB.TextBox TXT_ValorIPI 
         Height          =   288
         Left            =   4440
         TabIndex        =   15
         ToolTipText     =   "Valor total do I.P.I. deste produto"
         Top             =   2640
         Width           =   1092
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   4800
         Picture         =   "Tela_NotaFiscal_Dlg_4.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Volta ao assistente da nota fiscal."
         Top             =   3120
         Width           =   732
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   732
         Left            =   4080
         Picture         =   "Tela_NotaFiscal_Dlg_4.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancela edição."
         Top             =   3120
         Width           =   732
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "Apa&gar"
         Height          =   732
         Left            =   3360
         Picture         =   "Tela_NotaFiscal_Dlg_4.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Apaga todos os campos."
         Top             =   3120
         Width           =   732
      End
      Begin VB.CommandButton BT_Inserir 
         Caption         =   "&Inserir"
         Height          =   732
         Left            =   2640
         Picture         =   "Tela_NotaFiscal_Dlg_4.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Insere estes dados no assistente da nota fiscal."
         Top             =   3120
         Width           =   732
      End
      Begin VB.CommandButton BT_Editar 
         Caption         =   "&Editar"
         Height          =   732
         Left            =   1920
         Picture         =   "Tela_NotaFiscal_Dlg_4.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Edita este ítem."
         Top             =   3120
         Width           =   732
      End
      Begin VB.Label LB_Item 
         AutoSize        =   -1  'True
         Caption         =   "Item"
         Height          =   192
         Left            =   2160
         TabIndex        =   39
         ToolTipText     =   "Altere os valores da barra de rolagem para navegar entre um ítem e outro do assistente da nota fiscal."
         Top             =   240
         Width           =   300
      End
      Begin VB.Label LB_Figura 
         AutoSize        =   -1  'True
         Caption         =   "Figura:"
         Height          =   192
         Left            =   120
         TabIndex        =   38
         Top             =   240
         Width           =   492
      End
      Begin VB.Label LB_Descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   192
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   780
      End
      Begin VB.Label LB_CF 
         AutoSize        =   -1  'True
         Caption         =   "C.F.:"
         Height          =   192
         Left            =   120
         TabIndex        =   36
         Top             =   1800
         Width           =   312
      End
      Begin VB.Label LB_ST 
         AutoSize        =   -1  'True
         Caption         =   "S.T.:"
         Height          =   192
         Left            =   600
         TabIndex        =   35
         Top             =   1800
         Width           =   324
      End
      Begin VB.Label LB_Unidade 
         AutoSize        =   -1  'True
         Caption         =   "Unid.:"
         Height          =   192
         Left            =   1080
         TabIndex        =   34
         Top             =   1800
         Width           =   408
      End
      Begin VB.Label LB_Quantidade 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   192
         Left            =   1680
         TabIndex        =   33
         Top             =   1800
         Width           =   876
      End
      Begin VB.Label LB_PrecoUnitario 
         AutoSize        =   -1  'True
         Caption         =   "Preço Unitário:"
         Height          =   192
         Left            =   2760
         TabIndex        =   32
         Top             =   1800
         Width           =   1056
      End
      Begin VB.Label LB_PrecoTotal 
         AutoSize        =   -1  'True
         Caption         =   "Preço Total:"
         Height          =   192
         Left            =   4200
         TabIndex        =   31
         Top             =   1800
         Width           =   876
      End
      Begin VB.Label LB_PorcICMS 
         AutoSize        =   -1  'True
         Caption         =   "% I.C.M.S.:"
         Height          =   192
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   744
      End
      Begin VB.Label LB_PorcIPI 
         AutoSize        =   -1  'True
         Caption         =   "% I.P.I.:"
         Height          =   192
         Left            =   3600
         TabIndex        =   29
         Top             =   2400
         Width           =   504
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Base Calc.ICMS:"
         Height          =   192
         Left            =   960
         TabIndex        =   28
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label LB_ValorICMS 
         AutoSize        =   -1  'True
         Caption         =   "Valor I.C.M.S.:"
         Height          =   192
         Left            =   2280
         TabIndex        =   27
         Top             =   2400
         Width           =   984
      End
      Begin VB.Label LB_ValorIPI 
         AutoSize        =   -1  'True
         Caption         =   "Valor I.P.I.:"
         Height          =   192
         Left            =   4440
         TabIndex        =   26
         Top             =   2400
         Width           =   744
      End
      Begin VB.Label LB_PesoUnitario 
         AutoSize        =   -1  'True
         Caption         =   "Peso Unitário:"
         Height          =   192
         Left            =   120
         TabIndex        =   25
         Top             =   3000
         Width           =   1008
      End
      Begin VB.Label LB_PesoTotal 
         AutoSize        =   -1  'True
         Caption         =   "Peso Total:"
         Height          =   192
         Left            =   120
         TabIndex        =   24
         Top             =   3600
         Width           =   828
      End
   End
End
Attribute VB_Name = "Tela_NotaFiscal_Dlg_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NOMEAPLIC As String = "Edição de Ítens de Nota Fiscal"
Private Sub BH_1_Change()
    On Error GoTo ERRO_SISCOVAL
    CarregaTexto
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Figura.Text = ""
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
    TXT_ValorICMS.Text = ""
    TXT_ValorIPI.Text = ""
    TXT_PesoUnitario.Text = ""
    TXT_PesoTotal.Text = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaTelaEmEdicao (True)
    ValRed = TXT_BaseCalcICMS.Text / TXT_PrecoTotal.Text
    TXT_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Inserir_Click()
    On Error GoTo ERRO_SISCOVAL
    EditaItemNF (BH_1.Value + 1)
    AtivaTelaEmEdicao (False)
    RD_Cortar.Value = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_NotaFiscal_Dlg_4
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    BT_Apagar.Value = True
    AtivaTelaEmEdicao (False)
    
    'Verifica quantos itens tem na NF
    Dim NumItem As Integer
    NumItem = 0
    For I = 1 To 20
        If Tela_NotaFiscal.FG_1.TextMatrix(I, 3) <> "" Then
            NumItem = NumItem + 1
        End If
    Next I
    BH_1.Max = NumItem - 1
    BH_1.Value = 0
    CarregaTexto
    FG_1.Visible = False
    FG_2.Visible = False
    DLL_FUNCS.RegistraEvento "Abrir Edição de Ítens de Notas Fiscais", Tela_NotaFiscal.TXT_NF.Text
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseCalcICMS_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_BaseCalcICMS.Text = "" And TXT_BaseCalcICMS.Text = "0" Then
        TXT_ValorICMS.Text = "0,00"
        Exit Sub
    Else
        If TXT_PorcICMS.Text = "" Or TXT_PorcICMS.Text = "0" Then Exit Sub
        If TXT_PorcICMS.Text >= 10 Then
            TXT_ValorICMS.Text = Format(TXT_BaseCalcICMS.Text * Val("0." & Trim(TXT_PorcICMS.Text)), "##,##0.00")
        ElseIf TXT_PorcICMS.Text < 10 Then
            TXT_ValorICMS.Text = Format(TXT_BaseCalcICMS.Text * Val("0.0" & Trim(TXT_PorcICMS.Text)), "##,##0.00")
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseCalcICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("A base de cálculo é inserida automaticamente...", vbOKOnly, "Assistente de Edição")
    TXT_BaseCalcICMS.SelLength = Len(TXT_BaseCalcICMS.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseCalcICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_ValorICMS.SetFocus
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
    If KeyAscii = 13 Then TXT_ST.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CF_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_CF.Text = UCase(TXT_CF.Text)
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
        KeyAscii = 27
        TXT_CF.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Figura_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Figura.SelLength = Len(TXT_Figura.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Figura_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoTotal_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O peso total é calculado automaticamente...", vbOKOnly, "Assistente de Edição")
    TXT_PesoTotal.SelLength = Len(TXT_PesoTotal.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoTotal_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then BT_Inserir.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoTotal_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PesoTotal.Text = Format(TXT_PesoTotal.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoUnitario_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_PesoUnitario.Text = "" And TXT_PesoUnitario.Text = "0" Then
        TXT_PesoTotal.Text = "0,00"
        Exit Sub
    Else
        If TXT_Quantidade.Text = "" Or TXT_Quantidade.Text = "0" Then Exit Sub
        TXT_PesoTotal.Text = TXT_PesoUnitario.Text * TXT_Quantidade.Text
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
    If TXT_PorcICMS.Text = "" And TXT_PorcICMS.Text = "0" Then
        TXT_ValorICMS.Text = "0,00"
        TXT_BaseCalcICMS.Text = "0,00"
        Exit Sub
    Else
        TXT_BaseCalcICMS.Text = Format(TXT_PrecoTotal.Text * ValRed, "###,##0.00")
        If TXT_PorcICMS.Text >= 10 Then
            TXT_ValorICMS.Text = Format(TXT_BaseCalcICMS.Text * Val("0." & Trim(TXT_PorcICMS.Text)), "##,##0.00")
        ElseIf TXT_PorcICMS.Text < 10 Then
            TXT_ValorICMS.Text = Format(TXT_BaseCalcICMS.Text * Val("0.0" & Trim(TXT_PorcICMS.Text)), "##,##0.00")
        End If
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
    If TXT_PorcIPI.Text = "" And TXT_PorcIPI.Text = "0" Then
        TXT_ValorIPI.Text = "0,00"
        Exit Sub
    Else
        If TXT_PrecoTotal.Text = "" Or TXT_PrecoTotal.Text = "0" Then Exit Sub
        If TXT_PorcIPI.Text >= 10 Then
            TXT_ValorIPI.Text = Format(TXT_PrecoTotal.Text * Val("0." & Trim(TXT_PorcIPI.Text)), "##,##0.00")
        ElseIf TXT_PorcIPI.Text < 10 Then
            TXT_ValorIPI.Text = Format(TXT_PrecoTotal.Text * Val("0.0" & Trim(TXT_PorcIPI.Text)), "##,##0.00")
        End If
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
    TXT_ValorICMS.Text = Format(TXT_ValorICMS.Text, "###,###,#00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoTotal_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_PrecoTotal.Text = "" Or TXT_PrecoTotal.Text = "0" Then
        TXT_BaseCalcICMS.Text = "0,00"
        TXT_ValorICMS.Text = "0,00"
        TXT_ValorIPI.Text = "0,00"
        Exit Sub
    End If
    TXT_BaseCalcICMS.Text = Format(TXT_PrecoTotal.Text * ValRed, "###,##0.00")
    If TXT_PorcIPI.Text <> "" And TXT_PorcIPI.Text <> "0" Then
        If TXT_PorcIPI.Text >= 10 Then
            TXT_ValorIPI.Text = Format(TXT_PrecoTotal.Text * Val("0." & Trim(TXT_PorcIPI.Text)), "##,##0.00")
        ElseIf TXT_PorcIPI.Text < 10 Then
            TXT_ValorIPI.Text = Format(TXT_PrecoTotal.Text * Val("0.0" & Trim(TXT_PorcIPI.Text)), "##,##0.00")
        End If
    End If
    If TXT_PorcICMS.Text <> "" And TXT_PorcICMS.Text <> "0" Then
        If TXT_PorcICMS.Text >= 10 Then
            TXT_ValorICMS.Text = Format(TXT_BaseCalcICMS.Text * Val("0." & Trim(TXT_PorcICMS.Text)), "##,##0.00")
        ElseIf TXT_PorcICMS.Text < 10 Then
            TXT_ValorICMS.Text = Format(TXT_BaseCalcICMS.Text * Val("0.0" & Trim(TXT_PorcICMS.Text)), "##,##0.00")
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PrecoTotal_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O Preço Total é calculado automaticamente...", vbOKOnly, "Assistente de Edição")
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
    If TXT_PrecoUnitario.Text = "" Or TXT_PrecoUnitario.Text = "0" Then
        TXT_PrecoTotal.Text = "0,00"
        TXT_BaseCalcICMS.Text = "0,00"
        TXT_ValorICMS.Text = "0,00"
        TXT_ValorIPI.Text = "0,00"
        Exit Sub
    End If
    TXT_PrecoTotal.Text = Format(TXT_Quantidade.Text * TXT_PrecoUnitario.Text, "###,##0.00")
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
        TXT_BaseCalcICMS.Text = "0,00"
        TXT_ValorICMS.Text = "0,00"
        TXT_ValorIPI.Text = "0,00"
        TXT_PesoTotal.Text = "0,00"
        Exit Sub
        If TXT_Quantidade.Text <> "" Or TXT_PrecoUnitario.Text <> "0" Then TXT_PrecoTotal.Text = Format(TXT_Quantidade.Text * TXT_PrecoUnitario.Text, "###,##0.00")
        If TXT_Quantidade.Text <> "" Or TXT_PesoUnitario.Text <> "0" Then TXT_PesoTotal.Text = TXT_PesoUnitario.Text * TXT_Quantidade.Text
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
    If KeyAscii = 13 Then TXT_PrecoUnitario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ST_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ST.SelLength = Len(TXT_ST.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ST_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Unidade.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Unidade_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Unidade.SelLength = Len(TXT_Unidade.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Unidade_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Quantidade.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O valor do I.C.M.S. é calculado automaticamente...", vbOKOnly, "Assistente de Edição")
    TXT_ValorICMS.SelLength = Len(TXT_ValorICMS.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_PorcIPI.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorICMS.Text = Format(TXT_ValorICMS.Text, "###,###,##0.00")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorIPI_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("O valor do I.P.I. é calculado automaticamente...", vbOKOnly, "Assistente de Edição")
    TXT_ValorIPI.SelLength = Len(TXT_ValorIPI.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorIPI_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then TXT_PesoUnitario.SetFocus
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

Private Sub AtivaTelaEmEdicao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    TXT_Figura.Enabled = Valor
    TXT_Descricao.Enabled = Valor
    TXT_CF.Enabled = Valor
    TXT_ST.Enabled = Valor
    TXT_Unidade.Enabled = Valor
    TXT_PrecoTotal.Enabled = Valor
    TXT_PrecoUnitario.Enabled = Valor
    TXT_Quantidade.Enabled = Valor
    TXT_PorcICMS.Enabled = Valor
    TXT_PorcIPI.Enabled = Valor
    TXT_BaseCalcICMS.Enabled = Valor
    TXT_ValorICMS.Enabled = Valor
    TXT_ValorIPI.Enabled = Valor
    TXT_PesoUnitario.Enabled = Valor
    TXT_PesoTotal.Enabled = Valor
    If Valor = False Then
        BT_Editar.Enabled = True
        BT_Apagar.Enabled = False
        BT_Inserir.Enabled = False
        BT_Cancelar.Enabled = False
        RD_Cortar.Enabled = False
        RD_Dividir.Enabled = False
        BT_Voltar.Enabled = True
        BH_1.Enabled = True
        LB_Item.Enabled = True
    Else
        BT_Editar.Enabled = False
        BT_Apagar.Enabled = True
        BT_Inserir.Enabled = True
        BT_Cancelar.Enabled = True
        RD_Cortar.Enabled = True
        RD_Dividir.Enabled = True
        BT_Voltar.Enabled = False
        BH_1.Enabled = False
        LB_Item.Enabled = False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub EditaItemNF(NumItem As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Procura item
    Dim NumLinha As Integer
    NumLinha = 0
    For I = 1 To 20
        If Tela_NotaFiscal.FG_1.TextMatrix(I, 3) <> "" Then
            NumLinha = NumLinha + 1
            If NumLinha = NumItem Then
                NumLinha = I
                Exit For
            End If
        End If
    Next I
    Dim MaterialItem As String
    MaterialItem = Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 2)
    Dim UltLin, PriLin As Integer
    UltLin = NumLinha
    'Procura primeira linha do item
    For I = NumLinha To 1 Step -1
        If Tela_NotaFiscal.FG_2.TextMatrix(I, 1) <> "Idem" Then
            NumLinha = I
            Exit For
        End If
    Next I
    'Verifica tipo
    Dim TipoItem As String
    TipoItem = Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 1)
    PriLin = NumLinha
    'Recupera descricao do item original
    DescricaoItem = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 2)
    For I = (NumLinha + 1) To 20
        If Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = "Idem" Then
            DescricaoItem = DescricaoItem & Tela_NotaFiscal.FG_1.TextMatrix(I, 2)
        Else
            Exit For
        End If
    Next I
    'Verifica se a descricao do produto foi alterado
    Dim TamDesVel, TamDesNov As Integer
    Dim AlteraItem As Boolean
    AlteraItem = False
    If Len(DescricaoItem) <> Len(TXT_Descricao.Text) Then
        TamDesVel = Int(Len(DescricaoItem) / 38)
        TamDesNov = Int(Len(TXT_Descricao.Text) / 38)
        If TamDesVel <> TamDesNov Then
            'Verifica linhas em branco
            Dim NumLinBra As Integer
            NumLinBra = 0
            For I = 1 To 20
                If Tela_NotaFiscal.FG_2.TextMatrix(I, 1) <> "" Then NumLinBra = NumLinBra + 1
            Next I
            'Verifica se existe possibilidade de inserir novo item
            If NumLinBra < TamDesNov - TamDesVel Then
                RespMsg = MsgBox("Ao alterar a descrição do ítem, o número de linhas da descrição aumentou, porém não existe mais linhas em branco. Portanto, nesta condições é impossível realizar esta operação.", vbOKOnly, "Assistente de Edição")
                Exit Sub
            End If
            FG_1.Cols = 12
            FG_1.Rows = 21
            FG_2.Cols = 6
            FG_2.Rows = 21
            'limpa tabela
            For I = 0 To 20
                For J = 0 To 11
                    FG_1.TextMatrix(I, J) = ""
                Next J
                For J = 0 To 5
                    FG_2.TextMatrix(I, J) = ""
                Next J
            Next I
            'remove dados NF temporariamente
            Dim NumItensRem As Integer
            NumItensRem = 0
            For I = (UltLin + 1) To 20
                If Tela_NotaFiscal.FG_1.TextMatrix(I, 2) = "" Then Exit For
                For J = 1 To 11
                    FG_1.TextMatrix(I, J) = Tela_NotaFiscal.FG_1.TextMatrix(I, J)
                    Tela_NotaFiscal.FG_1.TextMatrix(I, J) = ""
                    NumItensRem = NumItensRem + 1
                Next J
                For J = 1 To 5
                    FG_2.TextMatrix(I, J) = Tela_NotaFiscal.FG_2.TextMatrix(I, J)
                    Tela_NotaFiscal.FG_2.TextMatrix(I, J) = ""
                Next J
            Next I
            AlteraItem = True
            'Apagar item velho
            For I = PriLin To UltLin
                For J = 1 To 11
                    Tela_NotaFiscal.FG_1.TextMatrix(I, J) = ""
                Next J
                For J = 1 To 5
                    Tela_NotaFiscal.FG_2.TextMatrix(I, J) = ""
                Next J
            Next I
        End If
    End If
    
    Dim Linha As Integer
    Linha = NumLinha
    'Verifica se a descricao é maior que 38 caracteres
    If Len(TXT_Descricao.Text) <= 38 Then
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 1) = TXT_Figura.Text
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
        
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = TipoItem
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 2) = MaterialItem
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
        Dim DLR, DL1, DL2, DL3, DL4, DL5, DL6, DL7, DL8, DL9, DL10, DL11, DL12, DL13, DL14, DL15, DL16, DL17, DL18, DL19, DL20 As String
        DLR = TXT_Descricao.Text
        DL1 = ""
        DL2 = ""
        DL3 = ""
        DL4 = ""
        DL5 = ""
        DL6 = ""
        DL7 = ""
        DL8 = ""
        DL9 = ""
        DL10 = ""
        DL11 = ""
        DL12 = ""
        DL13 = ""
        DL14 = ""
        DL15 = ""
        DL16 = ""
        DL17 = ""
        DL18 = ""
        DL19 = ""
        DL20 = ""
        'Verifica se o usuario optou por cortar ou dividir a descricao
        If RD_Cortar.Value = True Then
            Do
                If Len(DLR) > 38 Then
                    If DL1 = "" Then
                        DL1 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 38)
                    ElseIf DL2 = "" Then
                        DL2 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 76)
                    ElseIf DL3 = "" Then
                        DL3 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 114)
                    ElseIf DL4 = "" Then
                        DL4 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 152)
                    ElseIf DL5 = "" Then
                        DL5 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 190)
                    ElseIf DL6 = "" Then
                        DL6 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 228)
                    ElseIf DL7 = "" Then
                        DL7 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 266)
                    ElseIf DL8 = "" Then
                        DL8 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 304)
                    ElseIf DL9 = "" Then
                        DL9 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 342)
                    ElseIf DL10 = "" Then
                        DL10 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 380)
                    ElseIf DL11 = "" Then
                        DL11 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 418)
                    ElseIf DL12 = "" Then
                        DL12 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 456)
                    ElseIf DL13 = "" Then
                        DL13 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 494)
                    ElseIf DL14 = "" Then
                        DL14 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 532)
                    ElseIf DL15 = "" Then
                        DL15 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 570)
                    ElseIf DL16 = "" Then
                        DL16 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 608)
                    ElseIf DL17 = "" Then
                        DL17 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 646)
                    ElseIf DL18 = "" Then
                        DL18 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 684)
                    ElseIf DL19 = "" Then
                        DL19 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 722)
                    ElseIf DL20 = "" Then
                        DL20 = Left(DLR, 38)
                        DLR = Right(DLR, TamDes - 760)
                    End If
                Else
                    If DL2 = "" Then
                        DL2 = DLR
                    ElseIf DL3 = "" Then
                        DL3 = DLR
                    ElseIf DL4 = "" Then
                        DL4 = DLR
                    ElseIf DL5 = "" Then
                        DL5 = DLR
                    ElseIf DL6 = "" Then
                        DL6 = DLR
                    ElseIf DL7 = "" Then
                        DL7 = DLR
                    ElseIf DL8 = "" Then
                        DL8 = DLR
                    ElseIf DL9 = "" Then
                        DL9 = DLR
                    ElseIf DL10 = "" Then
                        DL10 = DLR
                    ElseIf DL11 = "" Then
                        DL11 = DLR
                    ElseIf DL12 = "" Then
                        DL12 = DLR
                    ElseIf DL13 = "" Then
                        DL13 = DLR
                    ElseIf DL14 = "" Then
                        DL14 = DLR
                    ElseIf DL15 = "" Then
                        DL15 = DLR
                    ElseIf DL16 = "" Then
                        DL16 = DLR
                    ElseIf DL17 = "" Then
                        DL17 = DLR
                    ElseIf DL18 = "" Then
                        DL18 = DLR
                    ElseIf DL19 = "" Then
                        DL19 = DLR
                    ElseIf DL20 = "" Then
                        DL20 = DLR
                    End If
                    Exit Do
                End If
            Loop
        ElseIf RD_Dividir.Value = True Then
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
        End If
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
           DL6 = "" And _
           NumLinLivre < 5 Then
            LinErro = True
        ElseIf DL6 <> "" And _
           DL7 = "" And _
           NumLinLivre < 6 Then
            LinErro = True
        ElseIf DL7 <> "" And _
           DL8 = "" And _
           NumLinLivre < 7 Then
            LinErro = True
        ElseIf DL8 <> "" And _
           DL9 = "" And _
           NumLinLivre < 8 Then
            LinErro = True
        ElseIf DL9 <> "" And _
           DL10 = "" And _
           NumLinLivre < 9 Then
            LinErro = True
        ElseIf DL10 <> "" And _
           DL11 = "" And _
           NumLinLivre < 10 Then
            LinErro = True
        ElseIf DL11 <> "" And _
           DL12 = "" And _
           NumLinLivre < 11 Then
            LinErro = True
        ElseIf DL12 <> "" And _
           DL13 = "" And _
           NumLinLivre < 12 Then
            LinErro = True
        ElseIf DL13 <> "" And _
           DL14 = "" And _
           NumLinLivre < 13 Then
            LinErro = True
        ElseIf DL14 <> "" And _
           DL15 = "" And _
           NumLinLivre < 14 Then
            LinErro = True
        ElseIf DL15 <> "" And _
           DL16 = "" And _
           NumLinLivre < 15 Then
            LinErro = True
        ElseIf DL16 <> "" And _
           DL17 = "" And _
           NumLinLivre < 16 Then
            LinErro = True
        ElseIf DL17 <> "" And _
           DL18 = "" And _
           NumLinLivre < 17 Then
            LinErro = True
        ElseIf DL18 <> "" And _
           DL19 = "" And _
           NumLinLivre < 18 Then
            LinErro = True
        ElseIf DL19 <> "" And _
           DL20 = "" And _
           NumLinLivre < 19 Then
            LinErro = True
        ElseIf DL20 <> "" And _
           NumLinLivre < 20 Then
            LinErro = True
        End If
        If LinErro = True Then
            RespMsg = MsgBox("Como a descrição do produto ultrapassou o número de caracteres de uma linha, o sistema foi obrigado à dividi-lá em várias linhas, porém não existem mais linhas suficientes para esta operação. Não será possível inserir este ítem.", vbOKOnly, "Assistente de Edição")
            Exit Sub
        End If
        'Insere itens
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 1) = TXT_Figura.Text
        If DL2 <> "" And DL3 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = TipoItem
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        ElseIf DL3 <> "" And DL4 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = TipoItem
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL3
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        ElseIf DL4 <> "" And DL5 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = TipoItem
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
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = TipoItem
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
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 6) = TXT_Quantidade.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 7) = Format(TXT_PrecoUnitario.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 8) = Format(TXT_PrecoTotal.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 9) = TXT_PorcICMS.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 10) = TXT_PorcIPI.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 11) = Format(TXT_ValorIPI.Text, "###,###,###,##0.00")
        
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 2) = MaterialItem
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 3) = Format(TXT_BaseCalcICMS.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 4) = Format(TXT_ValorICMS.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 5) = Format(TXT_PesoUnitario.Text, "###,##0.00")
    End If
    
    If AlteraItem = True Then
        Linha = Linha + 1
        Dim NumUltLinNF As Integer
        NumUltLinNF = 0
        'Procura primeira linha em branco
        For I = 1 To 20
            If Tela_NotaFiscal.FG_2.TextMatrix(I, 1) <> "" Then
                NumUltLinNF = I
                Exit For
            End If
        Next I
        'verifica quantas linhas ainda estão livres
        NumLinLivre = 0
        For I = 1 To 20
            If Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = "" Then
                NumLinLivre = NumLinLivre + 1
            End If
        Next I
        'verifica quantos itens estao no FG2 temporario
        Dim NumLinFG2 As Integer
        NumLinFG2 = 0
        For I = 1 To 20
            If FG_2.TextMatrix(I, 1) <> "" Then
                NumLinFG2 = NumLinFG2 + 1
            End If
        Next I
        If NumLinFG2 > NumLinLivre Then
            RespMsg = MsgBox("Ocorreu um erro quando o programa transportava os dados da NF temporários.", vbOKOnly, "Assistente de Edição")
            Exit Sub
        End If
        'Repassa dados temporarios
        For I = 1 To 20
            If FG_2.TextMatrix(I, 1) <> "" Then
                For J = 1 To 11
                    Tela_NotaFiscal.FG_1.TextMatrix(Linha, J) = FG_1.TextMatrix(I, J)
                Next J
                For J = 1 To 5
                    Tela_NotaFiscal.FG_2.TextMatrix(Linha, J) = FG_2.TextMatrix(I, J)
                Next J
                Linha = Linha + 1
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub CarregaTexto()
    On Error GoTo ERRO_SISCOVAL
    LB_Item.Caption = "Item " & Str(BH_1.Value + 1) & "/" & Str(BH_1.Max + 1)
    'Procura item
    Dim NumLinha As Integer
    NumLinha = 0
    For I = 1 To 20
        If Tela_NotaFiscal.FG_1.TextMatrix(I, 3) <> "" Then
            NumLinha = NumLinha + 1
            If NumLinha = (BH_1.Value + 1) Then
                NumLinha = I
                Exit For
            End If
        End If
    Next I
    'Procura primeira linha do item
    For I = NumLinha To 1 Step -1
        If Tela_NotaFiscal.FG_2.TextMatrix(I, 1) <> "Idem" Then
            NumLinha = I
            Exit For
        End If
    Next I
    'Insere dados do primeiro ítem
    TXT_Figura.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 1)
    TXT_Descricao.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 2)
    If Tela_NotaFiscal.FG_2.TextMatrix(NumLinha + 1, 1) <> "Idem" Then
        'Se nao for multiplas linhas
        TXT_CF.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 3)
        TXT_ST.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 4)
        TXT_Unidade.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 5)
        TXT_Quantidade.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 6)
        TXT_PrecoUnitario.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 7)
        TXT_PrecoTotal.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 8)
        TXT_PorcICMS.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 9)
        TXT_PorcIPI.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 10)
        TXT_ValorIPI.Text = Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 11)
                
        TXT_BaseCalcICMS.Text = Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 3)
        TXT_ValorICMS.Text = Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 4)
        TXT_PesoTotal.Text = Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 5)
        TXT_PesoUnitario.Text = Tela_NotaFiscal.FG_2.TextMatrix(NumLinha, 5) / Tela_NotaFiscal.FG_1.TextMatrix(NumLinha, 6)
    ElseIf Tela_NotaFiscal.FG_2.TextMatrix(NumLinha + 1, 1) = "Idem" And _
       Tela_NotaFiscal.FG_1.TextMatrix(NumLinha + 1, 2) <> "" Then
        For I = NumLinha To 20
            If Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = "Idem" Then
                TXT_Descricao.Text = Trim(TXT_Descricao.Text) & Tela_NotaFiscal.FG_1.TextMatrix(I, 2)
                TXT_CF.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 3)
                TXT_ST.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 4)
                TXT_Unidade.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 5)
                TXT_Quantidade.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 6)
                TXT_PrecoUnitario.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 7)
                TXT_PrecoTotal.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 8)
                TXT_PorcICMS.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 9)
                TXT_PorcIPI.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 10)
                TXT_ValorIPI.Text = Tela_NotaFiscal.FG_1.TextMatrix(I, 11)
                
                TXT_BaseCalcICMS.Text = Tela_NotaFiscal.FG_2.TextMatrix(I, 3)
                TXT_ValorICMS.Text = Tela_NotaFiscal.FG_2.TextMatrix(I, 4)
                TXT_PesoTotal.Text = Tela_NotaFiscal.FG_2.TextMatrix(I, 5)
                If Tela_NotaFiscal.FG_2.TextMatrix(I, 5) <> "" And Tela_NotaFiscal.FG_1.TextMatrix(I, 6) <> "" Then _
                    TXT_PesoUnitario.Text = CDbl(Tela_NotaFiscal.FG_2.TextMatrix(I, 5)) / CDbl(Tela_NotaFiscal.FG_1.TextMatrix(I, 6))
            ElseIf Tela_NotaFiscal.FG_2.TextMatrix(I, 1) <> "Idem" And I <> NumLinha Then
                Exit For
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
