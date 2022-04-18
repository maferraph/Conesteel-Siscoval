VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Tela_NotaFiscal_Dlg_1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Digite os dados do item da nota fiscal:"
   ClientHeight    =   5430
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton RD_Dividir 
      Caption         =   "Dividir"
      Height          =   192
      Left            =   4560
      TabIndex        =   24
      ToolTipText     =   "Ao inserir os dados no assistente da N.F., a descrição será dividida (se necessário) por palavras."
      Top             =   1680
      Width           =   852
   End
   Begin VB.OptionButton RD_Cortar 
      Caption         =   "Cortar"
      Height          =   192
      Left            =   4560
      TabIndex        =   23
      ToolTipText     =   "Ao inserir os dados no assistente da N.F., a descrição será cortada (se necessário) por caracteres."
      Top             =   1440
      Width           =   852
   End
   Begin VB.CommandButton BT_Cancelar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   4560
      Picture         =   "Tela_NotaFiscal_Dlg_1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Volta à edição da nota fiscal."
      Top             =   3960
      Width           =   732
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "Apa&gar"
      Height          =   732
      Left            =   4560
      Picture         =   "Tela_NotaFiscal_Dlg_1.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Apaga este ítem"
      Top             =   3000
      Width           =   732
   End
   Begin VB.CommandButton BT_Inserir 
      Caption         =   "&Inserir"
      Height          =   732
      Left            =   4560
      Picture         =   "Tela_NotaFiscal_Dlg_1.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Insere dados deste ítem no Assistente da Nota Fiscal"
      Top             =   2040
      Width           =   732
   End
   Begin VB.CheckBox CK_EditaCampos 
      Caption         =   "Editar os campos"
      Height          =   435
      Left            =   3240
      TabIndex        =   8
      ToolTipText     =   "Se você quiser editar os campos abaixo"
      Top             =   1560
      Width           =   972
   End
   Begin VB.Frame FR_1 
      Caption         =   "Peça:"
      Height          =   1332
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   5292
      Begin VB.CommandButton BT_AssitenteFigura 
         Caption         =   "Assistente &Figura"
         Height          =   972
         Left            =   3960
         Picture         =   "Tela_NotaFiscal_Dlg_1.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Assistente de Figuras de Estoque"
         Top             =   240
         Width           =   972
      End
      Begin VB.CommandButton BT_Procurar 
         Caption         =   "Pr&ocurar ítem"
         Height          =   972
         Left            =   2640
         Picture         =   "Tela_NotaFiscal_Dlg_1.frx":0E98
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Procura este ítem no estoque"
         Top             =   240
         Width           =   972
      End
      Begin VB.ComboBox CB_Material 
         Height          =   315
         Left            =   840
         TabIndex        =   2
         ToolTipText     =   "Selecione um material"
         Top             =   960
         Width           =   1452
      End
      Begin VB.ComboBox CB_Bitola 
         Height          =   315
         Left            =   840
         TabIndex        =   1
         ToolTipText     =   "Sselecione uma bitola"
         Top             =   600
         Width           =   1452
      End
      Begin VB.ComboBox CB_Figura 
         Height          =   315
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   0
         ToolTipText     =   "Selecione uma figura"
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label LB_Material 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   192
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   612
      End
      Begin VB.Label LB_Bitola 
         AutoSize        =   -1  'True
         Caption         =   "Bitola:"
         Height          =   192
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   444
      End
      Begin VB.Label LB_Figura 
         AutoSize        =   -1  'True
         Caption         =   "Figura:"
         Height          =   192
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   492
      End
   End
   Begin VB.Frame FR_2 
      Height          =   3372
      Left            =   0
      TabIndex        =   35
      Top             =   1320
      Width           =   4452
      Begin VB.OptionButton RD_Completa 
         Caption         =   "Completa"
         Height          =   192
         Left            =   3360
         TabIndex        =   11
         ToolTipText     =   "Se deseja inserir uma descrição completa do material."
         Top             =   1320
         Width           =   972
      End
      Begin VB.OptionButton RD_Normal 
         Caption         =   "Normal"
         Height          =   192
         Left            =   2400
         TabIndex        =   10
         ToolTipText     =   "Se deseja inserir uma descrição normal do material."
         Top             =   1320
         Width           =   972
      End
      Begin VB.OptionButton RD_Compacta 
         Caption         =   "Compacta"
         Height          =   192
         Left            =   1200
         TabIndex        =   9
         ToolTipText     =   "Se deseja inserir uma descrição compacta do material."
         Top             =   1320
         Width           =   1092
      End
      Begin VB.ComboBox CB_Tratamento 
         Height          =   288
         ItemData        =   "Tela_NotaFiscal_Dlg_1.frx":11A2
         Left            =   120
         List            =   "Tela_NotaFiscal_Dlg_1.frx":11AC
         TabIndex        =   7
         ToolTipText     =   "Escolha um tratamento ou dados adicionais deste ítem."
         Top             =   960
         Width           =   4212
      End
      Begin VB.TextBox TXT_Descricao 
         Height          =   492
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   12
         ToolTipText     =   "Descrição deste ítem."
         Top             =   1560
         Width           =   4212
      End
      Begin VB.TextBox TXT_CF 
         Height          =   288
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Digite a classificação fiscal da peça"
         Top             =   2400
         Width           =   492
      End
      Begin VB.TextBox TXT_ST 
         Height          =   288
         Left            =   720
         TabIndex        =   14
         ToolTipText     =   "Digite a situação tributária da peça"
         Top             =   2400
         Width           =   612
      End
      Begin VB.TextBox TXT_Unidade 
         Height          =   288
         Left            =   1440
         TabIndex        =   15
         ToolTipText     =   "Digite a unidade da peça"
         Top             =   2400
         Width           =   492
      End
      Begin VB.TextBox TXT_PorcICMS 
         Height          =   288
         Left            =   120
         MaxLength       =   2
         TabIndex        =   18
         ToolTipText     =   "Digite a porcentagem do I.C.M.S."
         Top             =   3000
         Width           =   372
      End
      Begin VB.TextBox TXT_PorcIPI 
         Height          =   288
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   21
         ToolTipText     =   "Porcentagem de IPI"
         Top             =   3000
         Width           =   372
      End
      Begin VB.TextBox TXT_Quantidade 
         Height          =   288
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Digite a quantidade de peças"
         Top             =   360
         Width           =   1332
      End
      Begin MSMask.MaskEdBox TXT_ValorUnitario 
         Height          =   288
         Left            =   1680
         TabIndex        =   6
         ToolTipText     =   "Valor unitário desta peça"
         Top             =   360
         Width           =   1332
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   20
         Format          =   "$###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_BaseICMS 
         Height          =   288
         Left            =   600
         TabIndex        =   19
         ToolTipText     =   "Base de cálculo do ICMS"
         Top             =   3000
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   20
         Format          =   "$###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_ValorTotal 
         Height          =   288
         Left            =   2040
         TabIndex        =   16
         ToolTipText     =   "Valor total deste ítem"
         Top             =   2400
         Width           =   1092
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   20
         Format          =   "$###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_ValICMS 
         Height          =   288
         Left            =   1680
         TabIndex        =   20
         ToolTipText     =   "Valor do ICMS"
         Top             =   3000
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   20
         Format          =   "$###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_PesoUnitario 
         Height          =   288
         Left            =   3240
         TabIndex        =   17
         ToolTipText     =   "Valor total parcial do peso deste ítem"
         Top             =   2400
         Width           =   1092
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_ValorIPI 
         Height          =   288
         Left            =   3360
         TabIndex        =   22
         ToolTipText     =   "Valor do IPI"
         Top             =   3000
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   503
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   20
         Format          =   "$###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LB_Tratamento 
         AutoSize        =   -1  'True
         Caption         =   "Tratamento / Dados adicionais:"
         Height          =   192
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   2256
      End
      Begin VB.Label LB_Descricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   192
         Left            =   120
         TabIndex        =   51
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label LB_CF 
         AutoSize        =   -1  'True
         Caption         =   "C.F.:"
         Height          =   192
         Left            =   120
         TabIndex        =   50
         Top             =   2160
         Width           =   312
      End
      Begin VB.Label LB_ST 
         AutoSize        =   -1  'True
         Caption         =   "S.Trib.:"
         Height          =   192
         Left            =   720
         TabIndex        =   49
         Top             =   2160
         Width           =   504
      End
      Begin VB.Label LB_Unidade 
         AutoSize        =   -1  'True
         Caption         =   "Unid."
         Height          =   192
         Left            =   1440
         TabIndex        =   48
         Top             =   2160
         Width           =   372
      End
      Begin VB.Label LB_ValorTotal 
         AutoSize        =   -1  'True
         Caption         =   "Valor Total:"
         Height          =   192
         Left            =   2040
         TabIndex        =   47
         Top             =   2160
         Width           =   828
      End
      Begin VB.Label LB_PorcICMS 
         AutoSize        =   -1  'True
         Caption         =   "ICMS:"
         Height          =   192
         Left            =   120
         TabIndex        =   46
         Top             =   2760
         Width           =   420
      End
      Begin VB.Label LB_PorcIPI 
         AutoSize        =   -1  'True
         Caption         =   "IPI:"
         Height          =   192
         Left            =   2880
         TabIndex        =   45
         Top             =   2760
         Width           =   216
      End
      Begin VB.Label LB_ValorIPI 
         AutoSize        =   -1  'True
         Caption         =   "Valor IPI:"
         Height          =   192
         Left            =   3360
         TabIndex        =   44
         Top             =   2760
         Width           =   636
      End
      Begin VB.Label LB_BaseICMS 
         AutoSize        =   -1  'True
         Caption         =   "Base ICMS:"
         Height          =   192
         Left            =   600
         TabIndex        =   43
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label LB_ValorICMS 
         AutoSize        =   -1  'True
         Caption         =   "Valor ICMS:"
         Height          =   192
         Left            =   1680
         TabIndex        =   42
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label LB_PesoUnitario 
         AutoSize        =   -1  'True
         Caption         =   "Peso Parcial:"
         Height          =   192
         Left            =   3240
         TabIndex        =   41
         Top             =   2160
         Width           =   960
      End
      Begin VB.Label LB_ValorUnitario 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unitário:"
         Height          =   192
         Left            =   1680
         TabIndex        =   37
         Top             =   120
         Width           =   1008
      End
      Begin VB.Label LB_Quantidade 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   192
         Left            =   120
         TabIndex        =   36
         Top             =   120
         Width           =   876
      End
   End
   Begin VB.Frame FR_PrecoPeso 
      Caption         =   "Confirme o preço (da lista) e peso unitário deste ítem:"
      Height          =   732
      Left            =   0
      TabIndex        =   38
      Top             =   4680
      Width           =   5292
      Begin VB.CommandButton BT_OK 
         Caption         =   "O&K"
         Height          =   252
         Left            =   4440
         TabIndex        =   30
         ToolTipText     =   "Confirma preço e peso."
         Top             =   360
         Width           =   732
      End
      Begin MSMask.MaskEdBox TXT_PreUnit 
         Height          =   252
         Left            =   1200
         TabIndex        =   28
         ToolTipText     =   "Valor unitário deste ítem (preço líquido da tabela)"
         Top             =   360
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "$###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TXT_PesoUnit 
         Height          =   252
         Left            =   3360
         TabIndex        =   29
         ToolTipText     =   "Peso unitário deste ítem."
         Top             =   360
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         Format          =   "###,##0.00"
         PromptChar      =   "_"
      End
      Begin VB.Label LB_PesoUnit 
         AutoSize        =   -1  'True
         Caption         =   "Peso Unitário:"
         Height          =   192
         Left            =   2280
         TabIndex        =   40
         Top             =   360
         Width           =   1008
      End
      Begin VB.Label LB_PreUnit 
         AutoSize        =   -1  'True
         Caption         =   "Valor Unitário:"
         Height          =   192
         Left            =   120
         TabIndex        =   39
         Top             =   360
         Width           =   1008
      End
   End
End
Attribute VB_Name = "Tela_NotaFiscal_Dlg_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NOMEAPLIC As String = "Assistente de Ítens de Estoque"

Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaFrameValores (False)
    AtivaFrameDescricao (False)
    CK_EditaCampos.Enabled = False
    BT_Procurar.Enabled = True
    BT_AssitenteFigura.Enabled = True
    BT_Inserir.Enabled = False
    BT_Apagar.Enabled = False
    LB_Figura.Enabled = True
    LB_Bitola.Enabled = True
    LB_Material.Enabled = True
    CB_Figura.Enabled = True
    CB_Bitola.Enabled = True
    CB_Material.Enabled = True
    CB_Figura.Text = ""
    CB_Material.Text = ""
    CB_Bitola.Text = ""
    TXT_Descricao.Text = ""
    TXT_CF.Text = ""
    TXT_ST.Text = ""
    TXT_ValorTotal.Text = ""
    TXT_ValorIPI.Text = ""
    TXT_PorcIPI.Text = ""
    TXT_ValICMS.Text = ""
    TXT_PorcICMS.Text = ""
    TXT_BaseICMS.Text = ""
    TXT_Unidade.Text = ""
    TXT_PesoUnitario.Text = ""
    CB_Tratamento.Text = ""
    CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_AssitenteFigura_Click()
    On Error GoTo ERRO_SISCOVAL
    CB_Figura.Text = DLL_ASFIG.AssistenteFigura(App.ProductName, "Assfig", App.LegalCopyright)
    CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    AtivaFrameValores (False)
    AtivaFrameDescricao (False)
    CK_EditaCampos.Enabled = False
    BT_Procurar.Enabled = False
    BT_Apagar.Enabled = False
    BT_Inserir.Enabled = False
    BT_Cancelar.Enabled = True
    CB_Figura.Text = ""
    CB_Material.Text = ""
    CB_Bitola.Text = ""
    CB_Figura.Enabled = True
    CB_Material.Enabled = True
    CB_Bitola.Enabled = True
    FR_1.Enabled = True
    LB_Figura.Enabled = True
    LB_Bitola.Enabled = True
    LB_Material.Enabled = True
    Tela_NotaFiscal_Dlg_1.Hide
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Inserir_Click()
    On Error GoTo ERRO_SISCOVAL
    Dim NumLinha, NumErro As Integer
    NumLinha = NumErro = 0
    
    AtivaFrameDescricao (True)
    If IsNumeric(TXT_Quantidade.Text) = False Then
        RespMsg = MsgBox("Não foi digitado um número válido.", vbOKOnly, "Assistente de Produtos")
        TXT_Quantidade.SelLength = Len(Trim(TXT_Quantidade.Text))
        TXT_Quantidade.SetFocus
        Exit Sub
    ElseIf TXT_Quantidade.Text = 0 Then
        RespMsg = MsgBox("Não é possível inserir Zero peças na nota fiscal.", vbOKOnly, "Assistente de Produtos")
        TXT_Quantidade.SetFocus
        Exit Sub
    ElseIf TXT_ValorUnitario.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_ValorUnitario.SetFocus
        Exit Sub
    ElseIf TXT_ValorUnitario.Text = 0 Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_ValorUnitario.SetFocus
        Exit Sub
    ElseIf TXT_Descricao.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_Descricao.SetFocus
        Exit Sub
    ElseIf TXT_CF.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_CF.SetFocus
        Exit Sub
    ElseIf TXT_ST.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_ST.SetFocus
        Exit Sub
    ElseIf TXT_Unidade.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_Unidade.SetFocus
        Exit Sub
    ElseIf TXT_ValorTotal.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_ValorTotal.SetFocus
        Exit Sub
    ElseIf TXT_ValorTotal.Text = 0 Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_ValorTotal.SetFocus
        Exit Sub
    ElseIf TXT_PesoUnitario.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_PesoUnitario.SetFocus
        Exit Sub
    ElseIf TXT_PesoUnitario.Text = 0 Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_PesoUnitario.SetFocus
        Exit Sub
    ElseIf TXT_PorcICMS.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_PorcICMS.SetFocus
        Exit Sub
    ElseIf TXT_ValICMS.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_ValICMS.SetFocus
        Exit Sub
    ElseIf TXT_BaseICMS.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_BaseICMS.SetFocus
        Exit Sub
    ElseIf TXT_ValorIPI.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_ValorIPI.SetFocus
        Exit Sub
    ElseIf TXT_PorcIPI.Text = "" Then
        RespMsg = MsgBox("Não é possível inserir dados na nota fiscal sem estar todas campos preenchidos.", vbOKOnly, "Assistente de Produtos")
        TXT_PorcIPI.SetFocus
        Exit Sub
    End If
    AtivaFrameDescricao (False)
    
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
        RespMsg = MsgBox("Não existem mais linha em branco para serem preenchidas no quadro dados do produto na nota fiscal.", vbOKOnly, "Assistente de Produtos")
        DesativaTela
        Tela_NotaFiscal_Dlg_1.Hide
        Exit Sub
    End If

    'Funçao que envia dados ao assistente da NF
    EnviaItemNF (NumLinha)
    
    'Funcao que desativa a tela
    DesativaTela
    
    If Me.Visible = True Then CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_OK_Click()
    On Error GoTo ERRO_SISCOVAL
    If TXT_PreUnit.Text <> "" And TXT_PesoUnit.Text <> "" Then
        DLL_BD.BDSIS_TBEST.Edit
        DLL_BD.BDSIS_TBEST_CPVUN.Value = CDbl(TXT_PreUnit.Text)
        DLL_BD.BDSIS_TBEST_CPPUN.Value = CDbl(TXT_PesoUnit.Text)
        DLL_BD.BDSIS_TBEST.Update
    End If
    AtivaFrameValores (True)
    CK_EditaCampos.Enabled = True
    BT_Apagar.Enabled = True
    BT_Inserir.Enabled = True
    BT_Cancelar.Enabled = True
    FR_1.Enabled = True
    If TXT_PreUnit.Text = "" And TXT_PesoUnit.Text = "" Then
        AtivaFrameValores (False)
        AtivaFrameDescricao (False)
        BT_Inserir.Enabled = False
        BT_Procurar.Enabled = False
        CK_EditaCampos.Enabled = False
        AtivaFramePrecoPeso (False)
        BT_Apagar.Value = True
        CB_Figura.SetFocus
        Exit Sub
    End If
    TXT_ValorUnitario.Text = TXT_PreUnit.Text
    PesoItem = CDbl(TXT_PesoUnit.Text)
    AtivaFramePrecoPeso (False)
    If Me.Visible = True Then TXT_Quantidade.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Procurar_Click()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        RespMsg = MsgBox("Não foi selecionado uma figura para procura.", vbOKOnly, "Assistente de Produtos")
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        RespMsg = MsgBox("Não foi selecionado uma bitola para procura.", vbOKOnly, "Assistente de Produtos")
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Material.Text = "" Then
        RespMsg = MsgBox("Não foi selecionado um material para procura.", vbOKOnly, "Assistente de Produtos")
        CB_Material.SetFocus
        Exit Sub
    End If
    'Ativa Campos
    CK_EditaCampos.Enabled = True
    BT_Procurar.Enabled = False
    BT_AssitenteFigura.Enabled = False
    BT_Inserir.Enabled = True
    BT_Apagar.Enabled = True
    LB_Figura.Enabled = False
    LB_Bitola.Enabled = False
    LB_Material.Enabled = False
    CB_Figura.Enabled = False
    CB_Bitola.Enabled = False
    CB_Material.Enabled = False
    AtivaFrameValores (True)
    'Procura Figura
    Dim gInd, gCla, gExt, gTer As String
    DLL_BD.BDSIS_TBEFG.Seek "=", Trim(CB_Figura.Text)
    If DLL_BD.BDSIS_TBEFG.NoMatch Then
        RespMsg = MsgBox("Ocorreu algum problema na procura da figura.", vbOKOnly, "Assistente de Produtos")
        Exit Sub
    Else
        gInd = DLL_BD.BDSIS_TBEFG_CPIFG.Value
        gCla = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGCL.Value))
        gExt = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGEX.Value))
        gTer = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGIN.Value))
    End If
    'Procura Indice de Figura
    DLL_BD.BDSIS_TBEID.Seek "=", gInd
    If DLL_BD.BDSIS_TBEID.NoMatch Then
        RespMsg = MsgBox("Ocorreu algum problema na procura dos dados sobre a figura.", vbOKOnly, "Assistente de Produtos")
        Exit Sub
    Else
        If DLL_BD.BDSIS_TBEID_CPGIN.Value = "" Then 'Nao tem Internos
            If DLL_BD.BDSIS_TBEID_CPTRE = "" Then 'Nao Tem Tipo
                'Descricoes Reduzidas
                DesRed = Trim(DLL_BD.BDSIS_TBEID_CPDRE.Value) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Normais
                DesNor = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Completas
                DesCom = Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) & ", Classe de Pressão " & _
                            gCla & ", Extremidade " & gExt & ", em " & _
                            Trim(CB_Material.Text) & " e Dn.: " & _
                            Trim(CB_Bitola.Text)
            Else 'Tem Tipos
                'Descricoes Reduzidas
                DesRed = Trim(DLL_BD.BDSIS_TBEID_CPDRE.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTRE.Value) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Normais
                DesNor = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTNO.Value) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Completas
                DesCom = Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTCO.Value) & ", Classe de Pressão " & _
                            gCla & ", Extremidade " & gExt & ", em " & _
                            Trim(CB_Material.Text) & " e Dn.: " & _
                            Trim(CB_Bitola.Text)
            End If
        Else 'Tem Internos
            If DLL_BD.BDSIS_TBEID_CPTRE = "" Then 'Nao Tem Tipo
                'Descricoes Reduzidas
                DesRed = Trim(DLL_BD.BDSIS_TBEID_CPDRE.Value) & " " & _
                            Trim(gTer) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Normais
                DesNor = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " Int." & _
                            Trim(gTer) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Completas
                DesCom = Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) & ", Internos em " & _
                            Trim(gTer) & ", Classe de Pressão " & _
                            gCla & ", Extremidade " & gExt & ", em " & _
                            Trim(CB_Material.Text) & " e Dn.: " & _
                            Trim(CB_Bitola.Text)
            Else 'Tem Tipos
                'Descricoes Reduzidas
                DesRed = Trim(DLL_BD.BDSIS_TBEID_CPDRE.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTRE.Value) & " " & _
                            Trim(gTer) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Normais
                DesNor = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTNO.Value) & " Int." & _
                            Trim(gTer) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
                'Descricoes Completas
                DesCom = Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTCO.Value) & ", Internos em " & _
                            Trim(gTer) & ", Classe de Pressão " & _
                            gCla & ", Extremidade " & gExt & ", em " & _
                            Trim(CB_Material.Text) & " e Dn.: " & _
                            Trim(CB_Bitola.Text)
            End If
        End If
    End If
    RD_Normal.Value = True
    TXT_Descricao.Text = DesNor
    'Procura pela CF e ST da Figura
    DLL_BD.BDSIS_TBCFS.Seek "=", CB_Figura.Text, DLL_FUNCS.ProcuraValorGrupo(CB_Material.Text, "MAT")
    If DLL_BD.BDSIS_TBCFS.NoMatch Then
        RespMsg = MsgBox("Ocorreu algum problema na procura da CF e ST da figura.", vbOKOnly, "Assistente de Produtos")
        Exit Sub
    Else
        TXT_CF.Text = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBCFS_CPGCF.Value))
        TXT_ST.Text = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBCFS_CPGST.Value))
    End If
    'Procura pela ficha da Figura
    DLL_BD.BDSIS_TBEST.Seek "=", CB_Figura.Text, CB_Bitola.Text, CB_Material.Text
    If DLL_BD.BDSIS_TBEST.NoMatch Then
        RespMsg = MsgBox("Ocorreu algum problema na procura da ficha da figura.", vbOKOnly, "Assistente de Produtos")
        Exit Sub
    Else
        TXT_ValorUnitario.Text = DLL_BD.BDSIS_TBEST_CPVUN.Value
        TXT_PreUnit.Text = DLL_BD.BDSIS_TBEST_CPVUN.Value
        PesoItem = DLL_BD.BDSIS_TBEST_CPPUN.Value
        TXT_Quantidade.Text = 0
        PesoItem = CDbl(DLL_BD.BDSIS_TBEST_CPPUN.Value)
        TXT_PesoUnit.Text = PesoItem
        TXT_PesoUnitario.Text = 0
    End If
    'Procura pela alíquota de IPI da Figura
    DLL_BD.BDSIS_TBEAL.Seek "=", "IPI", Tela_NotaFiscal.CB_Estado.Text, TXT_CF.Text
    If DLL_BD.BDSIS_TBEAL.NoMatch Then
        RespMsg = MsgBox("Ocorreu algum problema na procura da alíquota de IPI da figura.", vbOKOnly, "Assistente de Produtos")
        Exit Sub
    Else
        TXT_PorcIPI.Text = DLL_BD.BDSIS_TBEAL_CPPOR.Value
    End If
    'Procura pela alíquota de ICMS da Figura
    DLL_BD.BDSIS_TBEAL.Seek "=", "ICMS", Tela_NotaFiscal.CB_Estado.Text, TXT_CF.Text
    If DLL_BD.BDSIS_TBEAL.NoMatch Then
        RespMsg = MsgBox("Ocorreu algum problema na procura da alíquota de ICMS da figura.", vbOKOnly, "Assistente de Produtos")
        Exit Sub
    Else
        If Tela_NotaFiscal.CB_Estado.Text = "SP" Then
            TXT_PorcICMS.Text = DLL_BD.BDSIS_TBEAL_CPPOR.Value
            RBC_ICMS = DLL_BD.BDSIS_TBEAL_CPRBC.Value
        Else
            TXT_PorcICMS.Text = DLL_BD.BDSIS_TBEAL_CPPOR.Value
            RBC_ICMS = DLL_BD.BDSIS_TBEAL_CPRBC.Value
            'Procura valor em SP para calcular preço da peça
            Dim PorSP As Long
            DLL_BD.BDSIS_TBEAL.Seek "=", "ICMS", "SP", TXT_CF.Text
            If Not DLL_BD.BDSIS_TBEAL.NoMatch Then
                PorSP = DLL_BD.BDSIS_TBEAL_CPPOR.Value
                If TXT_ValorUnitario.Text > 0 Then
                    If (PorSP - TXT_PorcICMS.Text) >= 10 Then
                        TXT_ValorUnitario.Text = TXT_ValorUnitario.Text * (1 - Val("0." & Str((PorSP - TXT_PorcICMS.Text))))
                    Else
                        TXT_ValorUnitario.Text = TXT_ValorUnitario.Text * (1 - Val("0.0" & Str((PorSP - TXT_PorcICMS.Text))))
                    End If
                    TXT_ValorIPI.Text = TXT_ValorUnitario.Text * Val("0.0" & Str((TXT_PorcIPI.Text)))
                End If
            End If
        End If
    End If
    'Finalizando
    RD_Cortar.Value = True
    If TXT_ValorUnitario.Text <= 0 Or PesoItem <= 0 Then
        AtivaFramePrecoPeso (True)
    Else
        If Me.Visible = True Then TXT_Quantidade.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        MsgBox "Selecione primeiro uma figura.", vbOKOnly, "Assistente de Produtos"
        CB_Figura.SetFocus
    End If
    CB_Bitola.SelLength = Len(CB_Bitola.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And CB_Bitola.Text <> "" Then
        CB_Material.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    CB_Bitola.Text = UCase(CB_Bitola.Text)
    If CB_Bitola.Text <> "" Then
        For I = 0 To CB_Bitola.ListCount - 1
            If CB_Bitola.Text = CB_Bitola.List(I) Then
                Exit For
            ElseIf CB_Bitola.Text <> CB_Bitola.List(I) And I = CB_Bitola.ListCount - 1 Then
                RespMsg = MsgBox("Essa bitola digitada não existe - consulte esta lista.", vbOKOnly, "Assistente de Produtos")
                CB_Bitola.SetFocus
                Exit Sub
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_Click()
    CarregaFIGBITMAT
End Sub
Private Sub CB_Figura_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    CB_Figura.SelLength = Len(CB_Figura.Text)
    AtivaFrameValores (False)
    AtivaFrameDescricao (False)
    CK_EditaCampos.Enabled = False
    Dim X As String
    X = CB_Figura.Text
    BT_Apagar.Value = True
    BT_Inserir.Enabled = False
    BT_Procurar.Enabled = False
    CB_Figura.Text = X
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And CB_Figura.Text <> "" Then
        CB_Bitola.SetFocus
    ElseIf KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        CarregaFIGBITMAT
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    CarregaFIGBITMAT
    CB_Material.ListIndex = 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_Change()
    On Error GoTo ERRO_SISCOVAL
    If CB_Material.Text = "" Then
        BT_Procurar.Enabled = False
    Else
        BT_Procurar.Enabled = True
    End If
    CB_Material.SelLength = Len(CB_Material.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_Click()
    On Error GoTo ERRO_SISCOVAL
    BT_Procurar.Enabled = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma figura.", vbOKOnly, "Assistente de Produtos")
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma bitola.", vbOKOnly, "Assistente de Produtos")
        CB_Bitola.SetFocus
        Exit Sub
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And CB_Material.Text <> "" Then
        BT_Procurar.Enabled = True
        BT_Procurar.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    CB_Material.Text = UCase(CB_Material.Text)
    If CB_Material.Text <> "" Then
        For I = 0 To CB_Material.ListCount - 1
            If CB_Material.Text = CB_Material.List(I) Then
                Exit For
            ElseIf CB_Material.Text <> CB_Material.List(I) And I = CB_Material.ListCount - 1 Then
                RespMsg = MsgBox("Esse material digitado não existe - consulte esta lista.", vbOKOnly, "Assistente de Produtos")
                CB_Material.SetFocus
                Exit Sub
            End If
        Next I
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Tratamento_Change()
    On Error GoTo ERRO_SISCOVAL
    If RD_Compacta.Value = True Then
        TXT_Descricao.Text = DesRed & " " & CB_Tratamento.Text
    ElseIf RD_Normal.Value = True Then
        TXT_Descricao.Text = DesNor & " " & CB_Tratamento.Text
    ElseIf RD_Completa.Value = True Then
        TXT_Descricao.Text = DesCom & " " & CB_Tratamento.Text
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Tratamento_Click()
    On Error GoTo ERRO_SISCOVAL
    If RD_Compacta.Value = True Then
        TXT_Descricao.Text = DesRed & " " & CB_Tratamento.Text
    ElseIf RD_Normal.Value = True Then
        TXT_Descricao.Text = DesNor & " " & CB_Tratamento.Text
    ElseIf RD_Completa.Value = True Then
        TXT_Descricao.Text = DesCom & " " & CB_Tratamento.Text
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Tratamento_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then BT_Inserir.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EditaCampos_Click()
    On Error GoTo ERRO_SISCOVAL
    If CK_EditaCampos.Value = 1 Then
        AtivaFrameDescricao (True)
        TXT_Descricao.SetFocus
    Else
        AtivaFrameDescricao (False)
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EditaCampos_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And CK_EditaCampos.Value = 1 Then
        TXT_Descricao.SetFocus
    ElseIf KeyAscii = 13 And CK_EditaCampos.Value = 0 Then
        BT_Inserir.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Activate()
    On Error GoTo ERRO_SISCOVAL
    CK_EditaCampos.Value = 0
    CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    AtivaFrameValores (False)
    AtivaFrameDescricao (False)
    BT_Inserir.Enabled = False
    BT_Procurar.Enabled = False
    BT_AssitenteFigura.Enabled = True
    CK_EditaCampos.Enabled = False
    AtivaFramePrecoPeso (False)
    CB_Figura.Text = ""
    CB_Material.Text = ""
    CB_Bitola.Text = ""
    TXT_Descricao.Text = ""
    TXT_CF.Text = ""
    TXT_ST.Text = ""
    TXT_ValorTotal.Text = ""
    TXT_ValorIPI.Text = ""
    TXT_PorcIPI.Text = ""
    TXT_PorcICMS.Text = ""
    TXT_Unidade.Text = ""
    CK_EditaCampos.Value = 0
    DLL_FUNCS.RegistraEvento "Abrir Assistente de Ítens para Notas Fiscais", Tela_NotaFiscal.TXT_NF.Text
    If Tela_NotaFiscal_Dlg_1.Visible = True Then CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub

Private Sub RD_Compacta_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.Text = DesRed
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Compacta_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Completa_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.Text = DesCom
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Completa_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Cortar_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then BT_Inserir.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Dividir_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then BT_Inserir.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Normal_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.Text = DesNor
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Normal_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Descricao.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_BaseICMS.SelLength = Len(TXT_BaseICMS.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CF_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_CF.SelLength = Len(TXT_CF.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CF_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_CF.Text <> "" Then TXT_ST.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.SelLength = Len(TXT_Descricao.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Descricao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_Descricao.Text <> "" Then TXT_CF.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoUnit_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PesoUnit.SelLength = Len(TXT_PesoUnit.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoUnit_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 And TXT_PesoUnit.Text <> "" Then BT_OK.SetFocus
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
    If KeyAscii = 13 And TXT_PesoUnitario.Text <> "" Then TXT_PorcICMS.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcICMS_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_PorcICMS.Text = "" Then
        Exit Sub
    ElseIf TXT_PorcICMS.Text = 0 Then
        Exit Sub
    ElseIf TXT_PorcICMS.Text >= 10 Then
        TXT_ValICMS.Text = CDbl(TXT_BaseICMS.Text) * Val("0." & Str((TXT_PorcICMS.Text)))
    ElseIf TXT_PorcICMS.Text < 10 Then
        TXT_ValICMS.Text = CDbl(TXT_BaseICMS.Text) * Val("0.0" & Str((TXT_PorcICMS.Text)))
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
    If KeyAscii = 13 And TXT_PorcICMS.Text <> "" Then TXT_PorcIPI.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PorcIPI_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_PorcIPI.Text = "" Then
        TXT_ValorIPI.Text = 0
    ElseIf Val(TXT_PorcIPI.Text) >= 10 Then
        TXT_ValorIPI.Text = CDbl(TXT_ValorTotal.Text) * Val("0." & Str((TXT_PorcIPI.Text)))
    ElseIf Val(TXT_PorcIPI.Text) < 10 Then
        TXT_ValorIPI.Text = CDbl(TXT_ValorTotal.Text) * Val("0.0" & Str((TXT_PorcIPI.Text)))
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
    If KeyAscii = 13 And TXT_PorcIPI.Text <> "" Then CK_EditaCampos.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PreUnit_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PreUnit.SelLength = Len(TXT_PreUnit.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PreUnit_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 And TXT_PreUnit.Text <> "" Then TXT_PesoUnit.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Quantidade_Change()
    On Error GoTo ERRO_SISCOVAL
    If Not IsNumeric(TXT_Quantidade.Text) Then Exit Sub

    TXT_ValorTotal.Text = Val(TXT_Quantidade.Text) * CDbl(TXT_ValorUnitario.Text)
    
    If TXT_PorcIPI.Text = "" Then
        TXT_ValorIPI.Text = 0
    ElseIf Val(TXT_PorcIPI.Text) >= 10 Then
        TXT_ValorIPI.Text = CDbl(TXT_ValorTotal.Text) * Val("0." & Str((TXT_PorcIPI.Text)))
    ElseIf Val(TXT_PorcIPI.Text) < 10 Then
        TXT_ValorIPI.Text = CDbl(TXT_ValorTotal.Text) * Val("0.0" & Str((TXT_PorcIPI.Text)))
    End If
    If Val(TXT_Quantidade.Text) = 1 Then
        TXT_Unidade.Text = "pç."
    ElseIf Val(TXT_Quantidade.Text) = 0 Then
        TXT_Unidade.Text = ""
    ElseIf Val(TXT_Quantidade.Text) > 1 Then
        TXT_Unidade.Text = "pçs."
    End If
    If RBC_ICMS = 0 Then
        TXT_BaseICMS.Text = TXT_ValorTotal.Text
    Else
        ValRed = (100 - RBC_ICMS) / 100
        TXT_BaseICMS.Text = CDbl(TXT_ValorTotal.Text) * ValRed
    End If
    If TXT_PorcICMS.Text = "" Then
        TXT_PorcICMS.Text = 0
    ElseIf TXT_PorcICMS.Text >= 10 Then
        TXT_ValICMS.Text = CDbl(TXT_BaseICMS.Text) * Val("0." & Str((TXT_PorcICMS.Text)))
    ElseIf TXT_PorcICMS.Text < 10 Then
        TXT_ValICMS.Text = CDbl(TXT_BaseICMS.Text) * Val("0.0" & Str((TXT_PorcICMS.Text)))
    End If
    If TXT_Quantidade.Text = "" Then
        TXT_PesoUnitario.Text = 0
    ElseIf TXT_Quantidade.Text = 0 Then
        TXT_PesoUnitario.Text = 0
    Else
        If PesoItem = 0 Then PesoItem = 1
        TXT_PesoUnitario.Text = CDbl(TXT_Quantidade.Text) * CDbl(PesoItem)
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
    If KeyAscii = 13 And TXT_Quantidade.Text <> "" Then TXT_ValorUnitario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ST_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ST.SelLength = Len(TXT_ST.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ST_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_ST.Text <> "" Then TXT_Unidade.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Unidade_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Unidade.SelLength = Len(TXT_Unidade.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Unidade_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_PesoUnitario.Text <> "" Then TXT_PesoUnitario.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValICMS.SelLength = Len(TXT_ValICMS.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorIPI_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorIPI.SelLength = Len(TXT_ValorIPI.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorIPI_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotal_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorTotal.SelLength = Len(TXT_ValorTotal.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotal_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorUnitario_Change()
    On Error GoTo ERRO_SISCOVAL
    If Not IsNumeric(TXT_ValorUnitario.Text) Then
        If TXT_ValorUnitario.Text = "" Then Exit Sub
        RespMsg = MsgBox("Campos não tem um número válido.", vbOKOnly, "Assistente de Produtos")
        TXT_ValorUnitario.SetFocus
        Exit Sub
    End If
    
    Dim ValRed
    TXT_ValorTotal.Text = Val(TXT_Quantidade.Text) * CDbl(TXT_ValorUnitario.Text)
    
    If TXT_PorcIPI.Text = "" Then
        TXT_ValorIPI.Text = 0
    ElseIf Val(TXT_PorcIPI.Text) >= 10 Then
        TXT_ValorIPI.Text = CDbl(TXT_ValorTotal.Text) * Val("0." & Str((TXT_PorcIPI.Text)))
    ElseIf Val(TXT_PorcIPI.Text) < 10 Then
        TXT_ValorIPI.Text = CDbl(TXT_ValorTotal.Text) * Val("0.0" & Str((TXT_PorcIPI.Text)))
    End If
    If RBC_ICMS = 0 Then
        TXT_BaseICMS.Text = TXT_ValorTotal.Text
    Else
        ValRed = (100 - RBC_ICMS) / 100
        TXT_BaseICMS.Text = CDbl(TXT_ValorTotal.Text) * ValRed
    End If
    If TXT_PorcICMS.Text = "" Then
        TXT_PorcICMS.Text = 0
    ElseIf TXT_PorcICMS.Text >= 10 Then
        TXT_ValICMS.Text = CDbl(TXT_BaseICMS.Text) * Val("0." & Str((TXT_PorcICMS.Text)))
    ElseIf TXT_PorcICMS.Text < 10 Then
        TXT_ValICMS.Text = CDbl(TXT_BaseICMS.Text) * Val("0.0" & Str((TXT_PorcICMS.Text)))
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorUnitario_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorUnitario.SelLength = Len(TXT_ValorUnitario.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorUnitario_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 And TXT_ValorUnitario.Text <> "" Then
        CB_Tratamento.SetFocus
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub AtivaFrameValores(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    FR_2.Enabled = Valor
    TXT_Quantidade.Enabled = Valor
    TXT_ValorUnitario.Enabled = Valor
    LB_Quantidade.Enabled = Valor
    LB_ValorUnitario.Enabled = Valor
    CB_Tratamento.Enabled = Valor
    LB_Tratamento.Enabled = Valor
    If Valor = False Then
        TXT_Quantidade.Text = ""
        TXT_ValorUnitario.Text = ""
        CB_Tratamento.Text = ""
    End If
    RD_Cortar.Enabled = Valor
    RD_Dividir.Enabled = Valor
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub AtivaFrameDescricao(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    TXT_Descricao.Enabled = Valor
    TXT_CF.Enabled = Valor
    TXT_ST.Enabled = Valor
    TXT_ValorTotal.Enabled = False
    TXT_ValorIPI.Enabled = False
    TXT_PorcIPI.Enabled = Valor
    TXT_PorcICMS.Enabled = Valor
    TXT_Unidade.Enabled = Valor
    TXT_BaseICMS.Enabled = False
    TXT_ValICMS.Enabled = False
    TXT_PesoUnitario.Enabled = Valor
    CB_Tratamento.Enabled = Valor
    LB_Tratamento.Enabled = Valor
    LB_Descricao.Enabled = Valor
    LB_CF.Enabled = Valor
    LB_ST.Enabled = Valor
    LB_ValorTotal.Enabled = False
    LB_ValorIPI.Enabled = False
    LB_PorcIPI.Enabled = Valor
    LB_PorcICMS.Enabled = Valor
    LB_Unidade.Enabled = Valor
    LB_ValorICMS.Enabled = False
    LB_BaseICMS.Enabled = False
    LB_PesoUnitario.Enabled = Valor
    RD_Compacta.Enabled = Valor
    RD_Normal.Enabled = Valor
    RD_Completa.Enabled = Valor
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub AtivaFramePrecoPeso(Valor As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Valor = False Then
        Tela_NotaFiscal_Dlg_1.Height = Tela_NotaFiscal_Dlg_1.Height - FR_PrecoPeso.Height
    Else
        AtivaFrameValores (False)
        AtivaFrameDescricao (False)
        CK_EditaCampos.Enabled = False
        BT_Apagar.Enabled = False
        BT_Inserir.Enabled = False
        BT_Cancelar.Enabled = False
        FR_1.Enabled = False
        Tela_NotaFiscal_Dlg_1.Height = Tela_NotaFiscal_Dlg_1.Height + FR_PrecoPeso.Height
    End If
    FR_PrecoPeso.Enabled = Valor
    LB_PreUnit.Enabled = Valor
    TXT_PreUnit.Enabled = Valor
    LB_PesoUnit.Enabled = Valor
    TXT_PesoUnit.Enabled = Valor
    BT_OK.Enabled = Valor
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub EnviaItemNF(NumLin As Integer)
    On Error GoTo ERRO_SISCOVAL
    Dim Linha As Integer
    Linha = NumLin
    'Verifica se a descricao é maior que 38 caracteres
    If Len(TXT_Descricao.Text) <= 38 Then
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 1) = CB_Figura.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = TXT_Descricao.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 3) = TXT_CF.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 4) = TXT_ST.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 5) = TXT_Unidade.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 6) = TXT_Quantidade.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 7) = Format(TXT_ValorUnitario.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 8) = Format(TXT_ValorTotal.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 9) = TXT_PorcICMS.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 10) = TXT_PorcIPI.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 11) = Format(TXT_ValorIPI.Text, "###,###,###,##0.00")
        
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = CB_Bitola.Text
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 2) = CB_Material.Text
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 3) = Format(TXT_BaseICMS.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 4) = Format(TXT_ValICMS.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 5) = Format(TXT_PesoUnitario.Text, "####,##0.00")
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
        'Verifica se o DLL_FUNCS.PegaUsuario optou por cortar ou dividir a descricao
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
            RespMsg = MsgBox("Como a descrição do produto ultrapassou o número de caracteres de uma linha, o sistema foi obrigado à dividi-lá em várias linhas, porém não existem mais linhas suficientes para esta operação. Não será possível inserir este ítem.", vbOKOnly, "Assistente de Produtos")
            Tela_NotaFiscal_Dlg_1.BT_Cancelar.SetFocus
            Exit Sub
        End If
        'Insere itens
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 1) = CB_Figura.Text
        If DL2 <> "" And DL3 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = CB_Bitola.Text
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        ElseIf DL3 <> "" And DL4 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = CB_Bitola.Text
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL2
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
            Linha = Linha + 1
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL3
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = "Idem"
        ElseIf DL4 <> "" And DL5 = "" Then
            Tela_NotaFiscal.FG_1.TextMatrix(Linha, 2) = DL1
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = CB_Bitola.Text
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
            Tela_NotaFiscal.FG_2.TextMatrix(Linha, 1) = CB_Bitola.Text
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
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 7) = Format(TXT_ValorUnitario.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 8) = Format(TXT_ValorTotal.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 9) = TXT_PorcICMS.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 10) = TXT_PorcIPI.Text
        Tela_NotaFiscal.FG_1.TextMatrix(Linha, 11) = Format(TXT_ValorIPI.Text, "###,###,###,##0.00")
        
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 2) = CB_Material.Text
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 3) = Format(TXT_BaseICMS.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 4) = Format(TXT_ValICMS.Text, "###,###,###,##0.00")
        Tela_NotaFiscal.FG_2.TextMatrix(Linha, 5) = Format(TXT_PesoUnitario.Text, "###,##0.00")
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub DesativaTela()
    On Error GoTo ERRO_SISCOVAL
    AtivaFrameValores (False)
    AtivaFrameDescricao (False)
    CK_EditaCampos.Enabled = False
    BT_Procurar.Enabled = True
    BT_AssitenteFigura.Enabled = True
    BT_Inserir.Enabled = False
    BT_Apagar.Enabled = False
    LB_Figura.Enabled = True
    LB_Bitola.Enabled = True
    LB_Material.Enabled = True
    CB_Figura.Enabled = True
    CB_Bitola.Enabled = True
    CB_Material.Enabled = True
    CB_Figura.Text = ""
    CB_Material.Text = ""
    CB_Bitola.Text = ""
    TXT_Descricao.Text = ""
    TXT_CF.Text = ""
    TXT_ST.Text = ""
    TXT_ValorTotal.Text = ""
    TXT_ValorIPI.Text = ""
    TXT_PorcIPI.Text = ""
    TXT_PorcICMS.Text = ""
    TXT_Unidade.Text = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub CarregaFIGBITMAT(Optional ColocaBIT As String, Optional ColocaMAT As String)
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then Exit Sub
    CB_Figura.Text = UCase(CB_Figura.Text)
    For I = 0 To CB_Figura.ListCount - 1
        If CB_Figura.Text = CB_Figura.List(I) Then
            Exit For
        ElseIf CB_Figura.Text <> CB_Figura.List(I) And I = CB_Figura.ListCount - 1 Then
            MsgBox "Essa figura digitada não existe - consulte esta lista.", vbOKOnly + vbInformation, NOMEAPLIC
            CB_Figura.SetFocus
            Exit Sub
        End If
    Next I
    'procura figura
    DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figura.Text
    've se o indice de figura da nova consulta é igual a figura anteriormente consultada
    If DLL_BD.BDSIS_TBEFG_CPIFG.Value = ESTIND Then Exit Sub
    ESTIND = DLL_BD.BDSIS_TBEFG_CPIFG.Value
    DLL_BD.BDSIS_TBEID.Seek "=", DLL_BD.BDSIS_TBEFG_CPIFG.Value
    If DLL_BD.BDSIS_TBEFG.NoMatch And DLL_BD.BDSIS_TBEID.NoMatch Then
        MsgBox "Ocorreu algum erro durante a procura do índice da figura.", vbOKOnly + vbInformation, NOMEAPLIC
        Exit Sub
    End If
    'Como são tabelas relacionadas, a procura acima ja acha o indice de figura
    Dim cA As String
    'Montando lista de bitolas
    cA = ""
    CB_Bitola.Clear
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGBI.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) = ";" Then
            CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    If IsEmpty(ColocaBIT) = False Then CB_Bitola.Text = ColocaBIT
    'Montando lista de materiais
    cA = ""
    CB_Material.Clear
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
            CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    'seleciona material A-105
    CB_Material.ListIndex = 0
    If IsEmpty(ColocaMAT) = False Then CB_Material.Text = ColocaMAT
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
