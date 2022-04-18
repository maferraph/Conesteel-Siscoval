VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Tela_AssistenteFigura 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Assistente de procura de figuras"
   ClientHeight    =   4290
   ClientLeft      =   30
   ClientTop       =   255
   ClientWidth     =   5085
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ST 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Selecione o tipo de peça nas guias acima"
      Top             =   0
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5530
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Conexão"
      TabPicture(0)   =   "Tela_AssistenteFigura.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FR"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Válvula"
      TabPicture(1)   =   "Tela_AssistenteFigura.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Flange"
      TabPicture(2)   =   "Tela_AssistenteFigura.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Componente"
      TabPicture(3)   =   "Tela_AssistenteFigura.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Produção Andam."
      TabPicture(4)   =   "Tela_AssistenteFigura.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame4"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Matéria-Prima"
      TabPicture(5)   =   "Tela_AssistenteFigura.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame5"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      Begin VB.Frame Frame5 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   60
         Top             =   840
         Width           =   4455
         Begin VB.ComboBox CB_DesMat 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   29
            ToolTipText     =   "Selecione uma das peças nesta lista"
            Top             =   720
            Width           =   4215
         End
         Begin VB.ComboBox CB_TipMat 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   30
            ToolTipText     =   "Selecione o tipo conforme a descrição"
            Top             =   1320
            Width           =   4215
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Descrição da Matéria-Prima:"
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   480
            Width           =   1995
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   1080
            Width           =   375
         End
      End
      Begin VB.Frame Frame4 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   55
         Top             =   840
         Width           =   4455
         Begin VB.ComboBox CB_ComPro 
            Height          =   315
            Left            =   2280
            Sorted          =   -1  'True
            TabIndex        =   26
            ToolTipText     =   "Selecione um complemento"
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox CB_DesPro 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   24
            ToolTipText     =   "Selecione uma das peças nesta lista"
            Top             =   360
            Width           =   4215
         End
         Begin VB.ComboBox CB_TipPro 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   25
            ToolTipText     =   "Selecione o tipo conforme a descrição"
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox CB_ExtPro 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   27
            ToolTipText     =   "Selecione uma extremidade"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.ComboBox CB_ClaPro 
            Height          =   315
            Left            =   2280
            Sorted          =   -1  'True
            TabIndex        =   28
            ToolTipText     =   "Selecione uma classe"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            Height          =   195
            Index           =   2
            Left            =   2280
            TabIndex        =   65
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Descrição da Produção em Andamento:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   2835
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   58
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Extremidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
            Height          =   195
            Left            =   2280
            TabIndex        =   56
            Top             =   1320
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   50
         Top             =   840
         Width           =   4455
         Begin VB.ComboBox CB_ComCom 
            Height          =   315
            Left            =   2280
            Sorted          =   -1  'True
            TabIndex        =   21
            ToolTipText     =   "Selecione um complemento"
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox CB_DesCom 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   19
            ToolTipText     =   "Selecione uma das peças nesta lista"
            Top             =   360
            Width           =   4215
         End
         Begin VB.ComboBox CB_TipCom 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   20
            ToolTipText     =   "Selecione o tipo conforme a descrição"
            Top             =   960
            Width           =   2055
         End
         Begin VB.ComboBox CB_ExtCom 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   22
            ToolTipText     =   "Selecione uma extremidade"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.ComboBox CB_ClaCom 
            Height          =   315
            Left            =   2280
            Sorted          =   -1  'True
            TabIndex        =   23
            ToolTipText     =   "Selecione uma classe"
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            Height          =   195
            Index           =   1
            Left            =   2280
            TabIndex        =   64
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Descrição do Componente:"
            Height          =   195
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Extremidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
            Height          =   195
            Left            =   2280
            TabIndex        =   51
            Top             =   1320
            Width           =   540
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   44
         Top             =   840
         Width           =   4455
         Begin VB.ComboBox CB_ComFla 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   15
            ToolTipText     =   "Selecione um complemento"
            Top             =   960
            Width           =   4215
         End
         Begin VB.ComboBox CB_ClaFla 
            Height          =   315
            Left            =   3000
            Sorted          =   -1  'True
            TabIndex        =   18
            ToolTipText     =   "Selecione uma classe"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox CB_ExtFla 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   16
            ToolTipText     =   "Selecione uma extremidade"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox CB_TipFla 
            Height          =   315
            Left            =   2280
            Sorted          =   -1  'True
            TabIndex        =   14
            ToolTipText     =   "Selecione o tipo conforme a descrição"
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox CB_DesFla 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   13
            ToolTipText     =   "Selecione uma das peças nesta lista"
            Top             =   360
            Width           =   2055
         End
         Begin VB.ComboBox CB_RanFla 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   17
            ToolTipText     =   "Selecione uma ranhura"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
            Height          =   195
            Left            =   3000
            TabIndex        =   49
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Extremidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   2280
            TabIndex        =   47
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Descrição da Flange:"
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Ranhuras:"
            Height          =   195
            Left            =   1560
            TabIndex        =   45
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   38
         Top             =   840
         Width           =   4455
         Begin VB.ComboBox CB_IntVal 
            Height          =   315
            Left            =   1560
            Sorted          =   -1  'True
            TabIndex        =   11
            ToolTipText     =   "Selecione um interno"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox CB_DesVal 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Selecione uma das peças nesta lista"
            Top             =   360
            Width           =   4215
         End
         Begin VB.ComboBox CB_TipVal 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   9
            ToolTipText     =   "Selecione o tipo conforme a descrição"
            Top             =   960
            Width           =   4215
         End
         Begin VB.ComboBox CB_ExtVal 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   10
            ToolTipText     =   "Selecione uma extremidade"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox CB_ClaVal 
            Height          =   315
            Left            =   3000
            Sorted          =   -1  'True
            TabIndex        =   12
            ToolTipText     =   "Selecione uma classe"
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label LB_Internos 
            AutoSize        =   -1  'True
            Caption         =   "Internos:"
            Height          =   195
            Left            =   1560
            TabIndex        =   43
            Top             =   1320
            Width           =   600
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Descrição da Válvula:"
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   1560
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   41
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Extremidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
            Height          =   195
            Left            =   3000
            TabIndex        =   39
            Top             =   1320
            Width           =   540
         End
      End
      Begin VB.Frame FR 
         Height          =   2055
         Left            =   240
         TabIndex        =   33
         Top             =   840
         Width           =   4455
         Begin VB.ComboBox CB_ClaCon 
            Height          =   315
            Left            =   2400
            Sorted          =   -1  'True
            TabIndex        =   7
            ToolTipText     =   "Selecione uma classe"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.ComboBox CB_ExtCon 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   6
            ToolTipText     =   "Selecione uma extremidade"
            Top             =   1560
            Width           =   1935
         End
         Begin VB.ComboBox CB_TipCon 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   5
            ToolTipText     =   "Selecione o tipo conforme a descrição"
            Top             =   960
            Width           =   4215
         End
         Begin VB.ComboBox CB_DesCon 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   4
            ToolTipText     =   "Selecione uma das peças nesta lista"
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Classe:"
            Height          =   195
            Left            =   2400
            TabIndex        =   37
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Extremidade:"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Descrição da Conexão:"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   1665
         End
      End
   End
   Begin VB.Frame FR_Tela 
      Height          =   1095
      Left            =   0
      TabIndex        =   31
      Top             =   3120
      Width           =   5055
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "Apa&gar"
         Height          =   732
         Left            =   3360
         Picture         =   "Tela_AssistenteFigura.frx":00A8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Apaga campos"
         Top             =   240
         Width           =   732
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   732
         Left            =   4200
         Picture         =   "Tela_AssistenteFigura.frx":04EA
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Volta à edição"
         Top             =   240
         Width           =   732
      End
      Begin VB.TextBox TXT_Figura 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   120
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Figura da peça"
         Top             =   600
         Width           =   3135
      End
      Begin VB.Label LB_Figura 
         AutoSize        =   -1  'True
         Caption         =   "Figura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   690
      End
   End
End
Attribute VB_Name = "Tela_AssistenteFigura"
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
Dim ModoEdicao As Boolean, RespMsg, cResp, I As Integer, ESTIND As String
Const NOMEAPLIC As String = "Assistente de Figuras de Estoque"
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    With Tela_AssistenteFigura
        .TXT_Figura.Text = ""
        .CB_DesCon.Text = ""
        .CB_TipCon.Clear
        .CB_ExtCon.Clear
        .CB_ClaCon.Clear
        .CB_DesVal.Text = ""
        .CB_TipVal.Clear
        .CB_IntVal.Clear
        .CB_ExtVal.Clear
        .CB_ClaVal.Clear
        .CB_DesFla.Text = ""
        .CB_TipFla.Clear
        .CB_RanFla.Clear
        .CB_ExtFla.Clear
        .CB_ClaFla.Clear
        .CB_ComFla.Clear
        .CB_DesCom.Text = ""
        .CB_TipCom.Clear
        .CB_ExtCom.Clear
        .CB_ClaCom.Clear
        .CB_ComCom.Clear
        .CB_DesPro.Text = ""
        .CB_TipPro.Clear
        .CB_ExtPro.Clear
        .CB_ClaPro.Clear
        .CB_ComPro.Clear
        .CB_DesMat.Text = ""
        .CB_TipMat.Clear
    End With
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    FIG = TXT_Figura.Text
    Unload Tela_AssistenteFigura
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_ClaCom_Click()
    ProcuraItem
End Sub
Private Sub CB_ClaCom_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesCom.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesCom.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ClaCom_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ClaCon_Click()
    ProcuraItem
End Sub
Private Sub CB_ClaCon_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesCon.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesCon.SetFocus
        Exit Sub
    ElseIf CB_DesCon.Text <> "" And (CB_ExtCon.ListCount = 0 Or CB_ClaCon.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipCon.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ClaCon_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ClaFla_Click()
    ProcuraItem
End Sub
Private Sub CB_ClaFla_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesFla.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesFla.SetFocus
        Exit Sub
    ElseIf CB_DesFla.Text <> "" And (CB_ExtFla.ListCount = 0 And CB_ClaFla.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipFla.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ClaFla_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ClaPro_Click()
    ProcuraItem
End Sub
Private Sub CB_ClaPro_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesPro.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesPro.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ClaPro_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ClaVal_Click()
    ProcuraItem
End Sub
Private Sub CB_ClaVal_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesVal.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesVal.SetFocus
        Exit Sub
    ElseIf CB_DesVal.Text <> "" And (CB_ExtVal.ListCount = 0 Or CB_ClaVal.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipVal.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ClaVal_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ComCom_Click()
    ProcuraItem
End Sub
Private Sub CB_ComCom_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ComFla_Click()
    ProcuraItem
End Sub
Private Sub CB_ComFla_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_DesCom_Click()
    TelaEmEspera True
    CarregaComboTipo
    CarregaCombos
    TelaEmEspera False
    ProcuraItem
End Sub
Private Sub CB_DesCom_GotFocus()
    With Tela_AssistenteFigura
        .TXT_Figura.Text = ""
        .CB_TipCom.Clear
        .CB_ExtCom.Clear
        .CB_ClaCom.Clear
        .CB_TipCom.Enabled = False
        .CB_ExtCom.Enabled = False
        .CB_ClaCom.Enabled = False
    End With
End Sub
Private Sub CB_DesCom_LostFocus()
    If CB_DesCom.Text = "" Then BT_Apagar.Value = True
    ProcuraItem
End Sub
Private Sub CB_DesCon_Click()
    TelaEmEspera True
    CarregaComboTipo
    CarregaCombos
    TelaEmEspera False
End Sub
Private Sub CB_DesCon_GotFocus()
    With Tela_AssistenteFigura
        .TXT_Figura.Text = ""
        .CB_TipCon.Clear
        .CB_ExtCon.Clear
        .CB_ClaCon.Clear
        .CB_TipCon.Enabled = False
        .CB_ExtCon.Enabled = False
        .CB_ClaCon.Enabled = False
    End With
End Sub
Private Sub CB_DesCon_LostFocus()
    If CB_DesCon.Text = "" Then BT_Apagar.Value = True
    ProcuraItem
End Sub
Private Sub CB_DesFla_Click()
    TelaEmEspera True
    CarregaComboTipo
    CarregaCombos
    TelaEmEspera False
End Sub
Private Sub CB_DesFla_GotFocus()
    With Tela_AssistenteFigura
        .TXT_Figura.Text = ""
        .CB_TipFla.Clear
        .CB_ComFla.Clear
        .CB_RanFla.Clear
        .CB_ExtFla.Clear
        .CB_ClaFla.Clear
        .CB_TipFla.Enabled = False
        .CB_ComFla.Enabled = False
        .CB_RanFla.Enabled = False
        .CB_ExtFla.Enabled = False
        .CB_ClaFla.Enabled = False
    End With
End Sub
Private Sub CB_DesFla_LostFocus()
    If CB_DesFla.Text = "" Then BT_Apagar.Value = True
    ProcuraItem
End Sub
Private Sub CB_DesMat_Click()
    TelaEmEspera True
    CarregaComboTipo
    CarregaCombos
    TelaEmEspera False
    ProcuraItem
End Sub
Private Sub CB_DesMat_GotFocus()
    With Tela_AssistenteFigura
        .TXT_Figura.Text = ""
        .CB_TipMat.Clear
        .CB_TipMat.Enabled = False
    End With
End Sub
Private Sub CB_DesMat_LostFocus()
    If CB_DesMat.Text = "" Then BT_Apagar.Value = True
    ProcuraItem
End Sub
Private Sub CB_DesPro_Click()
    TelaEmEspera True
    CarregaComboTipo
    CarregaCombos
    TelaEmEspera False
    ProcuraItem
End Sub
Private Sub CB_DesPro_GotFocus()
    With Tela_AssistenteFigura
        .TXT_Figura.Text = ""
        .CB_TipPro.Clear
        .CB_ExtPro.Clear
        .CB_ClaPro.Clear
        .CB_TipPro.Enabled = False
        .CB_ExtPro.Enabled = False
        .CB_ClaPro.Enabled = False
    End With
End Sub
Private Sub CB_DesPro_LostFocus()
    If CB_DesPro.Text = "" Then BT_Apagar.Value = True
    ProcuraItem
End Sub
Private Sub CB_DesVal_Click()
    TelaEmEspera True
    CarregaComboTipo
    CarregaCombos
    TelaEmEspera False
End Sub
Private Sub CB_DesVal_GotFocus()
    With Tela_AssistenteFigura
        .TXT_Figura.Text = ""
        .CB_TipVal.Clear
        .CB_IntVal.Clear
        .CB_ExtVal.Clear
        .CB_ClaVal.Clear
        .CB_TipVal.Enabled = False
        .CB_IntVal.Enabled = False
        .CB_ExtVal.Enabled = False
        .CB_ClaVal.Enabled = False
    End With
End Sub
Private Sub CB_DesVal_LostFocus()
    If CB_DesVal.Text = "" Then BT_Apagar.Value = True
    ProcuraItem
End Sub
Private Sub CB_ExtCom_Click()
    ProcuraItem
End Sub
Private Sub CB_ExtCom_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesCom.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesCom.SetFocus
        Exit Sub
    ElseIf CB_DesCom.Text <> "" And (CB_ExtCom.ListCount = 0 Or CB_ClaCom.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipCom.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ExtCom_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ExtCon_Click()
    ProcuraItem
End Sub
Private Sub CB_ExtCon_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesCon.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesCon.SetFocus
        Exit Sub
    ElseIf CB_DesCon.Text <> "" And (CB_ExtCon.ListCount = 0 Or CB_ClaCon.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipCon.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ExtCon_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ExtFla_Click()
    ProcuraItem
End Sub
Private Sub CB_ExtFla_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesFla.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesFla.SetFocus
        Exit Sub
    ElseIf CB_DesFla.Text <> "" And (CB_ExtFla.ListCount = 0 Or CB_ClaFla.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipFla.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ExtFla_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ExtPro_Click()
    ProcuraItem
End Sub
Private Sub CB_ExtPro_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesPro.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesPro.SetFocus
        Exit Sub
    ElseIf CB_DesPro.Text <> "" And (CB_ExtPro.ListCount = 0 Or CB_ClaPro.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipPro.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ExtPro_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_ExtVal_Click()
    ProcuraItem
End Sub
Private Sub CB_ExtVal_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesVal.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesVal.SetFocus
        Exit Sub
    ElseIf CB_DesVal.Text <> "" And (CB_ExtVal.ListCount = 0 Or CB_ClaVal.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipVal.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_ExtVal_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_IntVal_Click()
    ProcuraItem
End Sub
Private Sub CB_IntVal_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesVal.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesVal.SetFocus
        Exit Sub
    ElseIf CB_DesVal.Text <> "" And (CB_ExtVal.ListCount = 0 Or CB_ClaVal.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipVal.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_IntVal_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_RanFla_Click()
    ProcuraItem
End Sub
Private Sub CB_RanFla_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesFla.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesFla.SetFocus
        Exit Sub
    ElseIf CB_DesFla.Text <> "" And (CB_ExtFla.ListCount = 0 And CB_ClaFla.ListCount = 0) Then
        MsgBox "Selecione primeiro um tipo...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_TipFla.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_RanFla_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_TipCom_Click()
    TelaEmEspera True
    CarregaCombos
    TelaEmEspera False
    ProcuraItem
End Sub
Private Sub CB_TipCom_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesCom.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesCom.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_TipCom_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_TipCon_Click()
    TelaEmEspera True
    CarregaCombos
    TelaEmEspera False
End Sub
Private Sub CB_TipCon_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesCon.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesCon.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_TipCon_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_TipFla_Click()
    TelaEmEspera True
    CarregaCombos
    TelaEmEspera False
End Sub
Private Sub CB_TipFla_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesFla.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesFla.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_TipFla_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_TipMat_Click()
    TelaEmEspera True
    CarregaCombos
    TelaEmEspera False
    ProcuraItem
End Sub
Private Sub CB_TipMat_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesMat.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesMat.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_TipMat_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_TipPro_Click()
    TelaEmEspera True
    CarregaCombos
    TelaEmEspera False
    ProcuraItem
End Sub
Private Sub CB_TipPro_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesPro.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesPro.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_TipPro_LostFocus()
    ProcuraItem
End Sub
Private Sub CB_TipVal_Click()
    TelaEmEspera True
    CarregaCombos
    TelaEmEspera False
End Sub
Private Sub CB_TipVal_GotFocus()
    TXT_Figura.Text = ""
    If CB_DesVal.Text = "" Then
        MsgBox "Selecione primeiro uma descrição...", vbInformation + vbOKOnly, NOMEAPLIC
        CB_DesVal.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_TipVal_LostFocus()
    ProcuraItem
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc

    'Abre bancos de dados
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreCampos_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreCampos_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Carrega combo de descricao
    CB_DesCon.Clear
    CB_DesVal.Clear
    CB_DesFla.Clear
    CB_DesCom.Clear
    CB_DesPro.Clear
    CB_DesMat.Clear
    If DLL_BD.BDSIS_TBEID.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEID.MoveFirst
        Do While Not DLL_BD.BDSIS_TBEID.EOF
            If Trim(DLL_BD.BDSIS_TBEID_CPGPE.Value) = "PEC_CON" Then
                CarregaComboDescricao CB_DesCon, DLL_BD.BDSIS_TBEID_CPDCO.Value
            ElseIf Trim(DLL_BD.BDSIS_TBEID_CPGPE.Value) = "PEC_VAL" Then
                CarregaComboDescricao CB_DesVal, DLL_BD.BDSIS_TBEID_CPDCO.Value
            ElseIf Trim(DLL_BD.BDSIS_TBEID_CPGPE.Value) = "PEC_FLA" Then
                CarregaComboDescricao CB_DesFla, DLL_BD.BDSIS_TBEID_CPDCO.Value
            ElseIf Trim(DLL_BD.BDSIS_TBEID_CPGPE.Value) = "PEC_COM" Then
                CarregaComboDescricao CB_DesCom, DLL_BD.BDSIS_TBEID_CPDCO.Value
            ElseIf Trim(DLL_BD.BDSIS_TBEID_CPGPE.Value) = "PEC_PA" Then
                CarregaComboDescricao CB_DesPro, DLL_BD.BDSIS_TBEID_CPDCO.Value
            ElseIf Trim(DLL_BD.BDSIS_TBEID_CPGPE.Value) = "PEC_MP" Then
                CarregaComboDescricao CB_DesMat, DLL_BD.BDSIS_TBEID_CPDCO.Value
            End If
            DLL_BD.BDSIS_TBEID.MoveNext
        Loop
    End If
    'Limpa outras combos
    BT_Apagar.Value = True
    ST.Tab = 0
    DLL_FUNCS.RegistraEvento "Abrir Assistente de Figuras", ""
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Tela_AssistenteFigura.Hide
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
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



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub TelaEmEspera(Estado As Boolean)
    If Estado = True Then
        Me.MousePointer = vbHourglass
        Me.Enabled = False
    Else
        Me.MousePointer = vbDefault
        Me.Enabled = True
    End If
End Sub
Private Static Sub CarregaComboDescricao(ByRef NomeCombo As ComboBox, Valor As String)
    If NomeCombo.ListCount > 0 Then
        For I = 0 To (NomeCombo.ListCount - 1)
            If NomeCombo.List(I) = Valor Then Exit Sub
        Next I
    End If
    NomeCombo.AddItem Valor
End Sub
Private Static Sub CarregaComboTipo()
    CB_TipCon.Clear
    CB_TipVal.Clear
    CB_TipFla.Clear
    CB_TipCom.Clear
    CB_TipPro.Clear
    CB_TipMat.Clear
    CB_TipCon.Enabled = False
    CB_TipVal.Enabled = False
    CB_TipFla.Enabled = False
    CB_TipCom.Enabled = False
    CB_TipPro.Enabled = False
    CB_TipMat.Enabled = False
    DLL_BD.BDSIS_TBEID.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEID.EOF
        'Procura descricao da ficha
        If ST.Tab = 0 And Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) = Trim(CB_DesCon.Text) Then
            If DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_CON" Then CarregaComboTipo_Aux CB_TipCon
        ElseIf ST.Tab = 1 And Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) = Trim(CB_DesVal.Text) Then
            If DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_VAL" Then CarregaComboTipo_Aux CB_TipVal
        ElseIf ST.Tab = 2 And Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) = Trim(CB_DesFla.Text) Then
            If DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_FLA" Then CarregaComboTipo_Aux CB_TipFla
        ElseIf ST.Tab = 3 And Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) = Trim(CB_DesCom.Text) Then
            If DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_COM" Then CarregaComboTipo_Aux CB_TipCom
        ElseIf ST.Tab = 4 And Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) = Trim(CB_DesPro.Text) Then
            If DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_PA" Then CarregaComboTipo_Aux CB_TipPro
        ElseIf ST.Tab = 5 And Trim(DLL_BD.BDSIS_TBEID_CPDCO.Value) = Trim(CB_DesMat.Text) Then
            If DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_MP" Then CarregaComboTipo_Aux CB_TipMat
        End If
        DLL_BD.BDSIS_TBEID.MoveNext
    Loop
End Sub
Private Static Sub CarregaComboTipo_Aux(NomeCombo As ComboBox)
    If DLL_BD.BDSIS_TBEID_CPTCO.Value <> "" Then
        If NomeCombo.ListCount = 0 Then
            NomeCombo.Enabled = True
            NomeCombo.AddItem (DLL_BD.BDSIS_TBEID_CPTCO.Value)
        Else
            For I = 0 To NomeCombo.ListCount - 1
                If NomeCombo.List(I) = DLL_BD.BDSIS_TBEID_CPTCO.Value Then
                    Exit For
                ElseIf I = NomeCombo.ListCount - 1 And _
                    NomeCombo.List(I) <> DLL_BD.BDSIS_TBEID_CPTCO.Value Then
                    NomeCombo.AddItem (DLL_BD.BDSIS_TBEID_CPTCO.Value)
                End If
            Next I
        End If
    End If
End Sub
Private Static Sub CarregaCombos()
    CB_ExtCon.Clear
    CB_ClaCon.Clear
    CB_IntVal.Clear
    CB_ExtVal.Clear
    CB_ClaVal.Clear
    CB_ComFla.Clear
    CB_RanFla.Clear
    CB_ExtFla.Clear
    CB_ClaFla.Clear
    CB_ExtCom.Clear
    CB_ClaCom.Clear
    CB_ComCom.Clear
    CB_ExtPro.Clear
    CB_ClaPro.Clear
    CB_ComPro.Clear
    CB_ExtCon.Enabled = False
    CB_ClaCon.Enabled = False
    CB_IntVal.Enabled = False
    CB_ExtVal.Enabled = False
    CB_ClaVal.Enabled = False
    CB_ComFla.Enabled = False
    CB_RanFla.Enabled = False
    CB_ExtFla.Enabled = False
    CB_ClaFla.Enabled = False
    CB_ExtCom.Enabled = False
    CB_ClaCom.Enabled = False
    CB_ComCom.Enabled = False
    CB_ExtPro.Enabled = False
    CB_ClaPro.Enabled = False
    CB_ComPro.Enabled = False
    Dim cPro As String, sDCO As String
    cPro = ""
    If ST.Tab = 0 Then
        If CB_TipCon.ListCount > 0 Then cPro = CB_TipCon.List(CB_TipCon.ListIndex)
    ElseIf ST.Tab = 1 Then
        If CB_TipVal.ListCount > 0 Then cPro = CB_TipVal.List(CB_TipVal.ListIndex)
    ElseIf ST.Tab = 2 Then
        If CB_TipFla.ListCount > 0 Then cPro = CB_TipFla.List(CB_TipFla.ListIndex)
        CarregaComplementoFlange
    ElseIf ST.Tab = 3 Then
        If CB_TipCom.ListCount > 0 Then cPro = CB_TipCom.List(CB_TipCom.ListIndex)
        CarregaComplementoComponente
    ElseIf ST.Tab = 4 Then
        If CB_TipPro.ListCount > 0 Then cPro = CB_TipPro.List(CB_TipPro.ListIndex)
        CarregaComplementoProducao
    ElseIf ST.Tab = 5 Then
        If CB_TipMat.ListCount > 0 Then cPro = CB_TipMat.List(CB_TipMat.ListIndex)
    End If
    DLL_BD.BDSIS_TBEID.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEID.EOF
        sDCO = DLL_BD.BDSIS_TBEID_CPDCO.Value
        If ((ST.Tab = 0 And sDCO = CB_DesCon.Text) Or (ST.Tab = 1 And sDCO = CB_DesVal.Text) Or (ST.Tab = 2 And sDCO = CB_DesFla.Text) Or (ST.Tab = 3 And sDCO = CB_DesCom.Text) Or (ST.Tab = 4 And sDCO = CB_DesPro.Text) Or (ST.Tab = 5 And sDCO = CB_DesMat.Text)) And _
           DLL_BD.BDSIS_TBEID_CPTCO.Value = cPro Then
            cResp = DivideTexto("I", DLL_BD.BDSIS_TBEID_CPGIN.Value)
            cResp = DivideTexto("E", DLL_BD.BDSIS_TBEID_CPGEX.Value)
            cResp = DivideTexto("C", DLL_BD.BDSIS_TBEID_CPGCL.Value)
            If CB_ExtCon.ListCount > 0 Then CB_ExtCon.Enabled = True
            If CB_ClaCon.ListCount > 0 Then CB_ClaCon.Enabled = True
            If CB_IntVal.ListCount > 0 Then CB_IntVal.Enabled = True
            If CB_ExtVal.ListCount > 0 Then CB_ExtVal.Enabled = True
            If CB_ClaVal.ListCount > 0 Then CB_ClaVal.Enabled = True
            If CB_RanFla.ListCount > 0 Then CB_RanFla.Enabled = True
            If CB_ExtFla.ListCount > 0 Then CB_ExtFla.Enabled = True
            If CB_ClaFla.ListCount > 0 Then CB_ClaFla.Enabled = True
            If CB_ExtCom.ListCount > 0 Then CB_ExtCom.Enabled = True
            If CB_ClaCom.ListCount > 0 Then CB_ClaCom.Enabled = True
            If CB_ExtPro.ListCount > 0 Then CB_ExtPro.Enabled = True
            If CB_ClaPro.ListCount > 0 Then CB_ClaPro.Enabled = True
            If CB_ComFla.ListCount > 0 Then CB_ComFla.Enabled = True
            If CB_ComCom.ListCount > 0 Then CB_ComCom.Enabled = True
            If CB_ComPro.ListCount > 0 Then CB_ComPro.Enabled = True
            Exit Sub
        End If
        DLL_BD.BDSIS_TBEID.MoveNext
    Loop
End Sub
Private Static Function DivideTexto(Tipo As String, Texto As String) As String
    If Tipo = "" Or Texto = "" Then Exit Function
    Dim cA As String
    cA = ""
    For I = 1 To Len(Texto)
        If Mid(Texto, I, 1) <> ";" Then
            cA = cA & Mid(Texto, I, 1)
        ElseIf Mid(Texto, I, 1) = ";" Then
            If Tipo = "I" Then
                If ST.Tab = 1 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_IntVal.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 2 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_RanFla.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                End If
            ElseIf Tipo = "E" Then
                If ST.Tab = 0 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtCon.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 1 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtVal.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 2 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtFla.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 3 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtCom.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 4 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtPro.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                End If
            ElseIf Tipo = "C" Then
                If ST.Tab = 0 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaCon.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 1 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaVal.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 2 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaFla.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 3 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaCom.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                ElseIf ST.Tab = 4 Then
                    If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaPro.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
                End If
            End If
            cA = ""
        End If
    Next I
    If Tipo = "I" Then
        If ST.Tab = 1 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_IntVal.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 2 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_RanFla.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        End If
    ElseIf Tipo = "E" Then
        If ST.Tab = 0 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtCon.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 1 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtVal.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 2 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtFla.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 3 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtCom.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 4 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ExtPro.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        End If
    ElseIf Tipo = "C" Then
        If ST.Tab = 0 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaCon.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 1 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaVal.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 2 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaFla.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 3 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaCom.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        ElseIf ST.Tab = 4 Then
            If Trim(DLL_FUNCS.ProcuraGrupo(cA)) <> "" Then CB_ClaPro.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
        End If
    End If
End Function
Private Static Sub ProcuraItem()
    If ST.Tab = 0 Then
        If CB_DesCon.Text = "" Or CB_ExtCon.Text = "" Or CB_ClaCon.Text = "" Then Exit Sub
    ElseIf ST.Tab = 1 Then
        If CB_DesVal.Text = "" Or CB_ExtVal.Text = "" Or CB_IntVal.Text = "" Or CB_ClaVal.Text = "" Then Exit Sub
    ElseIf ST.Tab = 2 Then
        If CB_DesFla.Text = "" Or CB_ClaFla.Text = "" Or CB_ComFla.Text = "" Then Exit Sub
    ElseIf ST.Tab = 3 Then
        If CB_DesCom.Text = "" Then Exit Sub
    ElseIf ST.Tab = 4 Then
        If CB_DesPro.Text = "" Then Exit Sub
    ElseIf ST.Tab = 5 Then
        If CB_DesMat.Text = "" Then Exit Sub
    End If
    Dim cInd As String, cTer As String, cExt As String, cCla As String, sCom As String
    'Procura Indice de Figura
    cInd = ""
    cTer = ""
    cExt = ""
    cCla = ""
    sCom = ""
    Dim sDCO As String, sTCO As String
    DLL_BD.BDSIS_TBEID.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEID.EOF
        sDCO = DLL_BD.BDSIS_TBEID_CPDCO.Value
        sTCO = ""
        If IsNull(DLL_BD.BDSIS_TBEID_CPTCO.Value) = False Then sTCO = DLL_BD.BDSIS_TBEID_CPTCO.Value
        If (ST.Tab = 0 And DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_CON" And sDCO = CB_DesCon.Text And sTCO = CB_TipCon.Text) Or _
           (ST.Tab = 1 And DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_VAL" And sDCO = CB_DesVal.Text And sTCO = CB_TipVal.Text) Or _
           (ST.Tab = 2 And DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_FLA" And sDCO = CB_DesFla.Text And sTCO = CB_TipFla.Text) Or _
           (ST.Tab = 3 And DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_COM" And sDCO = CB_DesCom.Text And sTCO = CB_TipCom.Text) Or _
           (ST.Tab = 4 And DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_PA" And sDCO = CB_DesPro.Text And sTCO = CB_TipPro.Text) Or _
           (ST.Tab = 5 And DLL_BD.BDSIS_TBEID_CPGPE.Value = "PEC_MP" And sDCO = CB_DesMat.Text And sTCO = CB_TipMat.Text) Then
            cInd = DLL_BD.BDSIS_TBEID_CPIFI.Value
            Exit Do
        End If
        DLL_BD.BDSIS_TBEID.MoveNext
    Loop
    If IsNull(cInd) = True Then Exit Sub
    If cInd = "" Then Exit Sub
    'Procura grupos
    If ST.Tab = 0 Then
        cExt = DLL_FUNCS.ProcuraValorGrupo(CB_ExtCon.Text, "EXT")
        cCla = DLL_FUNCS.ProcuraValorGrupo(CB_ClaCon.Text, "CLA")
    ElseIf ST.Tab = 1 Then
        cExt = DLL_FUNCS.ProcuraValorGrupo(CB_ExtVal.Text, "EXT")
        cCla = DLL_FUNCS.ProcuraValorGrupo(CB_ClaVal.Text, "CLA")
        cTer = DLL_FUNCS.ProcuraValorGrupo(CB_IntVal.Text, "INT")
    ElseIf ST.Tab = 2 Then
        cExt = DLL_FUNCS.ProcuraValorGrupo(CB_ExtFla.Text, "EXT")
        cCla = DLL_FUNCS.ProcuraValorGrupo(CB_ClaFla.Text, "CLA")
        cTer = DLL_FUNCS.ProcuraValorGrupo(CB_RanFla.Text, "RAN")
        sCom = Trim(CB_ComFla.Text)
    ElseIf ST.Tab = 3 Then
        cExt = DLL_FUNCS.ProcuraValorGrupo(CB_ExtCom.Text, "EXT")
        cCla = DLL_FUNCS.ProcuraValorGrupo(CB_ClaCom.Text, "CLA")
        sCom = Trim(CB_ComCom.Text)
    ElseIf ST.Tab = 4 Then
        cExt = DLL_FUNCS.ProcuraValorGrupo(CB_ExtPro.Text, "EXT")
        cCla = DLL_FUNCS.ProcuraValorGrupo(CB_ClaPro.Text, "CLA")
    End If
    If cExt = "" Then cExt = "EXT_NADA"
    If cCla = "" Then cCla = "CLA_NADA"
    'Procura pela figura
    Dim c2Ter As String, c2Ext As String, c2Cla As String, s2Com As String
    DLL_BD.BDSIS_TBEFG.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEFG.EOF
        c2Ter = ""
        c2Ext = ""
        c2Cla = ""
        s2Com = ""
        If IsNull(DLL_BD.BDSIS_TBEFG_CPGIN.Value) = False And DLL_BD.BDSIS_TBEFG_CPGIN.Value <> "" Then c2Ter = DLL_BD.BDSIS_TBEFG_CPGIN.Value
        If IsNull(DLL_BD.BDSIS_TBEFG_CPGEX.Value) = False And DLL_BD.BDSIS_TBEFG_CPGEX.Value <> "" Then c2Ext = DLL_BD.BDSIS_TBEFG_CPGEX.Value
        If IsNull(DLL_BD.BDSIS_TBEFG_CPGCL.Value) = False And DLL_BD.BDSIS_TBEFG_CPGCL.Value <> "" Then c2Cla = DLL_BD.BDSIS_TBEFG_CPGCL.Value
        If IsNull(DLL_BD.BDSIS_TBEFG_CPCOM.Value) = False And DLL_BD.BDSIS_TBEFG_CPCOM.Value <> "" Then s2Com = DLL_BD.BDSIS_TBEFG_CPCOM.Value
        If DLL_BD.BDSIS_TBEFG_CPIFG.Value = cInd And _
           c2Ter = cTer And _
           c2Ext = cExt And _
           c2Cla = cCla And _
           s2Com = sCom Then
            TXT_Figura.Text = DLL_BD.BDSIS_TBEFG_CPFIG.Value
            BT_Voltar.SetFocus
            Exit Sub
        End If
        DLL_BD.BDSIS_TBEFG.MoveNext
    Loop
    'Se não achou...
    'MsgBox ("Não foi possível procurar a figura para esta peça. É possível que esta descrição não exista. Confira e tente novamente.")
    TXT_Figura.Text = ""
End Sub
Private Static Sub CarregaComplementoFlange()
    CB_ComFla.Clear
    DLL_BD.BDSIS_TBEFG.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEFG.EOF
        If CInt(DLL_BD.BDSIS_TBEFG_CPIFG.Value) > 49 And CInt(DLL_BD.BDSIS_TBEFG_CPIFG.Value) < 55 Then
            If CB_ComFla.ListCount = 0 Then
                CB_ComFla.AddItem (DLL_BD.BDSIS_TBEFG_CPCOM.Value)
            Else
                For I = 0 To (CB_ComFla.ListCount - 1)
                    If CB_ComFla.List(I) = DLL_BD.BDSIS_TBEFG_CPCOM.Value Then
                        Exit For
                    ElseIf I = CB_ComFla.ListCount - 1 And _
                        CB_ComFla.List(I) <> DLL_BD.BDSIS_TBEFG_CPCOM.Value Then
                        CB_ComFla.AddItem (DLL_BD.BDSIS_TBEFG_CPCOM.Value)
                    End If
                Next I
            End If
        End If
        DLL_BD.BDSIS_TBEFG.MoveNext
    Loop
    CB_ComFla.Enabled = True
End Sub
Private Static Sub CarregaComplementoComponente()
    CB_ComCom.Clear
    DLL_BD.BDSIS_TBEFG.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEFG.EOF
        If Left(CStr(DLL_BD.BDSIS_TBEFG_CPIFG.Value), 2) = "CP" Then
            If CB_ComCom.ListCount = 0 Then
                CB_ComCom.AddItem (DLL_BD.BDSIS_TBEFG_CPCOM.Value)
            Else
                For I = 0 To (CB_ComCom.ListCount - 1)
                    If CB_ComCom.List(I) = DLL_BD.BDSIS_TBEFG_CPCOM.Value Then
                        Exit For
                    ElseIf I = CB_ComCom.ListCount - 1 And _
                        CB_ComCom.List(I) <> DLL_BD.BDSIS_TBEFG_CPCOM.Value Then
                        CB_ComCom.AddItem (DLL_BD.BDSIS_TBEFG_CPCOM.Value)
                    End If
                Next I
            End If
        End If
        DLL_BD.BDSIS_TBEFG.MoveNext
    Loop
    CB_ComCom.Enabled = True
End Sub
Private Static Sub CarregaComplementoProducao()
    CB_ComPro.Clear
    DLL_BD.BDSIS_TBEFG.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEFG.EOF
        If Left(CStr(DLL_BD.BDSIS_TBEFG_CPIFG.Value), 2) = "PA" Then
            If CB_ComPro.ListCount = 0 Then
                CB_ComPro.AddItem (DLL_BD.BDSIS_TBEFG_CPCOM.Value)
            Else
                For I = 0 To (CB_ComPro.ListCount - 1)
                    If CB_ComPro.List(I) = DLL_BD.BDSIS_TBEFG_CPCOM.Value Then
                        Exit For
                    ElseIf I = CB_ComCom.ListCount - 1 And _
                        CB_ComPro.List(I) <> DLL_BD.BDSIS_TBEFG_CPCOM.Value Then
                        CB_ComPro.AddItem (DLL_BD.BDSIS_TBEFG_CPCOM.Value)
                    End If
                Next I
            End If
        End If
        DLL_BD.BDSIS_TBEFG.MoveNext
    Loop
    CB_ComPro.Enabled = True
End Sub
