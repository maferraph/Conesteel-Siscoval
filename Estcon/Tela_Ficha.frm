VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Tela_Ficha 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ficha de Estoque"
   ClientHeight    =   3300
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   7695
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   6840
      Picture         =   "Tela_Ficha.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   732
   End
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   732
      Left            =   6840
      Picture         =   "Tela_Ficha.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir ficha de cadastro de uma empresa."
      Top             =   600
      Width           =   732
   End
   Begin MSFlexGridLib.MSFlexGrid FG 
      Height          =   3252
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6732
      _ExtentX        =   11880
      _ExtentY        =   5741
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      SelectionMode   =   1
   End
End
Attribute VB_Name = "Tela_Ficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Ficha de Estoque"
Dim I, J As Integer
Dim RespMsg

Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload BT_Voltar.Parent
ERRO_SISCOVAL: If Err Then If Tela_FichaEstoque.DLL_FUNCS.MensagemErro(Tela_FichaEstoque.DLL_FUNCS.PegaUsuario, Tela_FichaEstoque.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    FG.ColAlignment(0) = flexAlignCenterCenter
    FG.ColAlignment(1) = flexAlignLeftCenter
    FG.ColAlignment(2) = flexAlignCenterCenter
    FG.ColAlignment(3) = flexAlignCenterCenter
    FG.ColAlignment(4) = flexAlignCenterCenter
    FG.ColAlignment(5) = flexAlignCenterCenter
    FG.ColAlignment(6) = flexAlignLeftCenter
    FG.ColWidth(0) = 900
    FG.ColWidth(1) = 2000
    FG.ColWidth(2) = 900
    FG.ColWidth(3) = 900
    FG.ColWidth(4) = 900
    FG.ColWidth(5) = 1000
    FG.ColWidth(6) = 2000
    FG.TextArray(0) = "Data"
    FG.TextArray(1) = "Movimento"
    FG.TextArray(2) = "Nota Fiscal"
    FG.TextArray(3) = "Entrada"
    FG.TextArray(4) = "Saída"
    FG.TextArray(5) = "Saldo"
    FG.TextArray(6) = "Observações"
    FG.TextMatrix(1, 0) = "24/11/99"
    FG.TextMatrix(1, 1) = "ATIVAL"
    FG.TextMatrix(1, 2) = "4900"
    FG.TextMatrix(1, 3) = "100"
    FG.TextMatrix(1, 4) = "200"
    FG.TextMatrix(1, 5) = "1000"
    FG.TextMatrix(1, 6) = "Não tem"
ERRO_SISCOVAL: If Err Then If Tela_FichaEstoque.DLL_FUNCS.MensagemErro(Tela_FichaEstoque.DLL_FUNCS.PegaUsuario, Tela_FichaEstoque.DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
