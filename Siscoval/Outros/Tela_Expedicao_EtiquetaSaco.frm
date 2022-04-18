VERSION 5.00
Begin VB.Form Tela_Expedicao_EtiquetaSaco 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Etiquetas para Sacos Plásticos (Peças à Granel)"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   855
      Left            =   4200
      Picture         =   "Tela_Expedicao_EtiquetaSaco.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Imprimir Pedido"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton BT_Procurar 
      Caption         =   "&Deletar"
      Height          =   855
      Left            =   2280
      Picture         =   "Tela_Expedicao_EtiquetaSaco.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Deletar Pedido"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox TXT_PE 
      Height          =   285
      Left            =   600
      TabIndex        =   31
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton BT_Apagar 
      Caption         =   "&Apagar"
      Height          =   855
      Left            =   5280
      Picture         =   "Tela_Expedicao_EtiquetaSaco.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Apaga campos"
      Top             =   3000
      Width           =   855
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   6240
      Picture         =   "Tela_Expedicao_EtiquetaSaco.frx":0A56
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   3000
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dados para Etiqueta:"
      Height          =   2895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.Label LB_SP 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   29
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label LB_PE 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   28
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label LB_Transportadora 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label LB_Bairro 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   26
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label LB_Cidade 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   25
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label LB_Estado 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   24
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label LB_CEP 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6120
         TabIndex        =   23
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label LB_Fax 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   22
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label LB_Fone 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3960
         TabIndex        =   21
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label LB_Endereco 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   585
      End
      Begin VB.Label LB_IE 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         TabIndex        =   19
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label LB_CNPJ 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label LB_Contato 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5640
         TabIndex        =   17
         Top             =   480
         Width           =   585
      End
      Begin VB.Label LB_Empresa 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Pedido Cliente nº:"
         Height          =   195
         Index           =   14
         Left            =   5400
         TabIndex        =   15
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nosso Pedido nº:"
         Height          =   195
         Index           =   13
         Left            =   3960
         TabIndex        =   14
         Top             =   2040
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contato:"
         Height          =   195
         Index           =   12
         Left            =   5640
         TabIndex        =   13
         Top             =   240
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Transportadora:"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fax:"
         Height          =   195
         Index           =   10
         Left            =   5640
         TabIndex        =   11
         Top             =   840
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fone:"
         Height          =   195
         Index           =   9
         Left            =   3960
         TabIndex        =   10
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CEP:"
         Height          =   195
         Index           =   8
         Left            =   6120
         TabIndex        =   9
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
         Height          =   195
         Index           =   7
         Left            =   5400
         TabIndex        =   8
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cidade:"
         Height          =   195
         Index           =   6
         Left            =   3960
         TabIndex        =   7
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Bairro:"
         Height          =   195
         Index           =   5
         Left            =   2760
         TabIndex        =   6
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Endereço:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual:"
         Height          =   195
         Index           =   3
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C.N.P.J.:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   660
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número do Pedido:"
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   32
      Top             =   3120
      Width           =   1365
   End
End
Attribute VB_Name = "Tela_Expedicao_EtiquetaSaco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** VARIÁVEIS DLL's ****************
Dim DLL_BD As Scvbd.Classe_Scvbd

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Sistema Siscoval"

Private Sub BT_Apagar_Click()
    LB_Empresa.Caption = ""
    LB_Contato.Caption = ""
    LB_CNPJ.Caption = ""
    LB_IE.Caption = ""
    LB_Fone.Caption = ""
    LB_Fax.Caption = ""
    LB_Endereco.Caption = ""
    LB_Bairro.Caption = ""
    LB_Cidade.Caption = ""
    LB_Estado.Caption = ""
    LB_CEP.Caption = ""
    LB_Transportadora.Caption = ""
    LB_PE.Caption = ""
    LB_SP.Caption = ""
End Sub
Private Sub BT_Imprimir_Click()
    Tela_Expedicao_EtiquetaSaco_Imprimir.Show vbModal
End Sub
Private Sub BT_Procurar_Click()
    On Error Resume Next
    BT_Apagar_Click
    Dim Emp, Tra As String
    With DLL_BD
       .BDSIS_TBPED.Seek "=", TXT_PE.Text
       If .BDSIS_TBPED.NoMatch = False Then
          Emp = .BDSIS_TBPED_CPEMP.Value
          LB_Contato.Caption = .BDSIS_TBPED_CPCON.Value
          LB_PE.Caption = .BDSIS_TBPED_CPIND.Value
          LB_SP.Caption = .BDSIS_TBPED_CPNSP.Value
          Tra = .BDSIS_TBPED_CPTRA.Value
          .BDSIS_TBEMP.Seek "=", Emp
          If .BDSIS_TBEMP.NoMatch = False Then
             LB_Empresa.Caption = .BDSIS_TBEMP_CPEMP.Value
             LB_CNPJ.Caption = .BDSIS_TBEMP_CPCGC.Value
             LB_IE.Caption = .BDSIS_TBEMP_CPINE.Value
             LB_Fone.Caption = .BDSIS_TBEMP_CPFON.Value
             LB_Fax.Caption = .BDSIS_TBEMP_CPFAX.Value
             LB_Endereco.Caption = .BDSIS_TBEMP_CPEND.Value
             LB_Bairro.Caption = .BDSIS_TBEMP_CPBAI.Value
             LB_Cidade.Caption = .BDSIS_TBEMP_CPCID.Value
             LB_Estado.Caption = .BDSIS_TBEMP_CPEST.Value
             LB_CEP.Caption = .BDSIS_TBEMP_CPCEP.Value
          End If
          .BDSIS_TBEMP.Seek "=", Tra
          If .BDSIS_TBEMP.NoMatch = False Then
             LB_Transportadora.Caption = .BDSIS_TBEMP_CPEMP.Value
          End If
       Else
          MsgBox "Não existe dados sobre o Pedido nº " & TXT_PE.Text & ".", vbExclamation + vbOKOnly, "Pedido não existe"
       End If
    End With
End Sub
Private Sub BT_Voltar_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    'Abre bancos de dados
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_SISCOVAL
    'Abrindo Tabelas
    If DLL_BD.AbreTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_SISCOVAL
    If DLL_BD.AbreTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_SISCOVAL
    'Abre Campos
    If DLL_BD.AbreCampos_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_SISCOVAL
    If DLL_BD.AbreCampos_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_SISCOVAL
    BT_Apagar_Click
    Exit Sub
ERRO_SISCOVAL:
    If Err Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Fecha tabelas
    If DLL_BD.FechaTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
End Sub
