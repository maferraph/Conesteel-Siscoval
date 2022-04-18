VERSION 5.00
Begin VB.Form Tela_Convunit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conversor de Unidades"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FR_1 
      Caption         =   "Valor para ser convertido:"
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox CB_UA 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   120
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox CB_Unidade 
         Height          =   315
         ItemData        =   "Tela_Convunit.frx":0000
         Left            =   4800
         List            =   "Tela_Convunit.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   430
         Width           =   1815
      End
      Begin VB.ComboBox CB_Grandeza 
         Height          =   315
         ItemData        =   "Tela_Convunit.frx":0004
         Left            =   2880
         List            =   "Tela_Convunit.frx":001A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   430
         Width           =   1815
      End
      Begin VB.TextBox TXT_V 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   2
         Text            =   "SSSS"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Unidade:"
         Height          =   195
         Left            =   4800
         TabIndex        =   6
         Top             =   240
         Width           =   645
      End
      Begin VB.Label LB_1 
         AutoSize        =   -1  'True
         Caption         =   "Grandeza:"
         Height          =   195
         Left            =   2880
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   732
      Left            =   6960
      Picture         =   "Tela_Convunit.frx":0056
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volta � Tela Principal."
      Top             =   120
      Width           =   732
   End
   Begin VB.Frame FR_2 
      Height          =   4335
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   7575
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   59
         Top             =   150
         Width           =   600
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   58
         Top             =   150
         Width           =   450
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   57
         Top             =   150
         Width           =   1290
      End
      Begin VB.Line L2 
         BorderColor     =   &H80000005&
         Index           =   2
         X1              =   10
         X2              =   7550
         Y1              =   375
         Y2              =   375
      End
      Begin VB.Line L1 
         BorderColor     =   &H80000003&
         Index           =   2
         X1              =   10
         X2              =   7550
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   56
         Top             =   480
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   0
         Left            =   3840
         TabIndex        =   55
         Top             =   480
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   54
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   53
         Top             =   720
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   1
         Left            =   3840
         TabIndex        =   52
         Top             =   720
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   51
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   50
         Top             =   960
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   2
         Left            =   3840
         TabIndex        =   49
         Top             =   960
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   2
         Left            =   1680
         TabIndex        =   48
         Top             =   960
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   47
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   3
         Left            =   3840
         TabIndex        =   46
         Top             =   1200
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   3
         Left            =   1680
         TabIndex        =   45
         Top             =   1200
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   44
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   4
         Left            =   3840
         TabIndex        =   43
         Top             =   1440
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   4
         Left            =   1680
         TabIndex        =   42
         Top             =   1440
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   41
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   5
         Left            =   3840
         TabIndex        =   40
         Top             =   1680
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   5
         Left            =   1680
         TabIndex        =   39
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   38
         Top             =   1920
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   6
         Left            =   3840
         TabIndex        =   37
         Top             =   1920
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   36
         Top             =   1920
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   35
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   7
         Left            =   3840
         TabIndex        =   34
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   7
         Left            =   1680
         TabIndex        =   33
         Top             =   2160
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   32
         Top             =   2400
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   8
         Left            =   3840
         TabIndex        =   31
         Top             =   2400
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   8
         Left            =   1680
         TabIndex        =   30
         Top             =   2400
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   9
         Left            =   240
         TabIndex        =   29
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   9
         Left            =   3840
         TabIndex        =   28
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   9
         Left            =   1680
         TabIndex        =   27
         Top             =   2640
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   10
         Left            =   240
         TabIndex        =   26
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   10
         Left            =   3840
         TabIndex        =   25
         Top             =   2880
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   10
         Left            =   1680
         TabIndex        =   24
         Top             =   2880
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   23
         Top             =   3120
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   11
         Left            =   3840
         TabIndex        =   22
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   11
         Left            =   1680
         TabIndex        =   21
         Top             =   3120
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   12
         Left            =   3840
         TabIndex        =   19
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   12
         Left            =   1680
         TabIndex        =   18
         Top             =   3360
         Width           =   1290
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   13
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Width           =   600
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   13
         Left            =   3840
         TabIndex        =   16
         Top             =   3600
         Width           =   450
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   13
         Left            =   1680
         TabIndex        =   15
         Top             =   3600
         Width           =   1290
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   14
         Left            =   1680
         TabIndex        =   14
         Top             =   3840
         Width           =   1290
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   14
         Left            =   3840
         TabIndex        =   13
         Top             =   3840
         Width           =   450
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   12
         Top             =   3840
         Width           =   600
      End
      Begin VB.Label LB_N 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Unidade"
         Height          =   195
         Index           =   15
         Left            =   1680
         TabIndex        =   11
         Top             =   4080
         Width           =   1290
      End
      Begin VB.Label LB_V 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
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
         Index           =   15
         Left            =   3840
         TabIndex        =   10
         Top             =   4080
         Width           =   450
      End
      Begin VB.Label LB_U 
         AutoSize        =   -1  'True
         Caption         =   "Unidade"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   9
         Top             =   4080
         Width           =   600
      End
   End
End
Attribute VB_Name = "Tela_Convunit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** VARI�VEIS DLL's ****************
Dim DLL_CARGA As Scvcarr.Classe_Scvcarr
Dim DLL_FUNCS As Scvfunc.Classe_Scvfunc

' ****************** DECLARA��ES ****************
Const NOMEAPLIC As String = "Conversor de Unidades"
Dim I As Long, J As Long, RespMsg
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_Convunit
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Grandeza_Click()
    LimpaUnidade
    CB_Unidade.Clear
    CB_UA.Clear
    If CB_Grandeza.Text = "Temperatura" Then
        CB_Unidade.AddItem ("Celsius")
        CB_Unidade.AddItem ("Farenheit")
        CB_Unidade.AddItem ("Kelvin")
        CB_Unidade.AddItem ("Rankine")
        CB_UA.AddItem ("�C")
        CB_UA.AddItem ("�F")
        CB_UA.AddItem ("K")
        CB_UA.AddItem ("R")
    ElseIf CB_Grandeza.Text = "Comprimento" Then
        CB_Unidade.AddItem ("quil�metro")
        CB_Unidade.AddItem ("hect�metro")
        CB_Unidade.AddItem ("dec�metro")
        CB_Unidade.AddItem ("metro")
        CB_Unidade.AddItem ("dec�metro")
        CB_Unidade.AddItem ("cent�metro")
        CB_Unidade.AddItem ("mil�metro")
        CB_Unidade.AddItem ("micr�metro")
        CB_Unidade.AddItem ("jarda")
        CB_Unidade.AddItem ("p�")
        CB_Unidade.AddItem ("polegada")
        CB_Unidade.AddItem ("milha mar�tima")
        CB_Unidade.AddItem ("milha terrestre")
        CB_Unidade.AddItem ("ano luz")
        CB_UA.AddItem ("km")
        CB_UA.AddItem ("hm")
        CB_UA.AddItem ("dam")
        CB_UA.AddItem ("m")
        CB_UA.AddItem ("dm")
        CB_UA.AddItem ("cm")
        CB_UA.AddItem ("mm")
        CB_UA.AddItem ("�m")
        CB_UA.AddItem ("yd")
        CB_UA.AddItem ("ft")
        CB_UA.AddItem ("in")
        CB_UA.AddItem ("-")
        CB_UA.AddItem ("-")
        CB_UA.AddItem ("-")
    ElseIf CB_Grandeza.Text = "Press�o" Then
        CB_Unidade.AddItem ("Pascal")
        CB_Unidade.AddItem ("megaPascal")
        CB_Unidade.AddItem ("Atmosf�rica")
        CB_Unidade.AddItem ("bar")
        CB_Unidade.AddItem ("b�ria")
        CB_Unidade.AddItem ("kgf por m�")
        CB_Unidade.AddItem ("Atmosfera t�cnica")
        CB_Unidade.AddItem ("kgf por mm�")
        CB_Unidade.AddItem ("lbf por ft�")
        CB_Unidade.AddItem ("PSI")
        CB_Unidade.AddItem ("Torricelli")
        CB_Unidade.AddItem ("in de merc�rio")
        CB_Unidade.AddItem ("ft de �gua")
        CB_Unidade.AddItem ("m de �gua")
        CB_UA.AddItem ("Pa (N/m�)")
        CB_UA.AddItem ("MPa")
        CB_UA.AddItem ("atm")
        CB_UA.AddItem ("bar")
        CB_UA.AddItem ("ba")
        CB_UA.AddItem ("kgf/m�")
        CB_UA.AddItem ("at (kgf/cm�)")
        CB_UA.AddItem ("kgf/mm�")
        CB_UA.AddItem ("lbf/ft�")
        CB_UA.AddItem ("PSI (lbf/in�)")
        CB_UA.AddItem ("Torr (mm Hg)")
        CB_UA.AddItem ("in Hg")
        CB_UA.AddItem ("ft H2O")
        CB_UA.AddItem ("m H2O")
    ElseIf CB_Grandeza.Text = "�rea" Then
        CB_Unidade.AddItem ("quil�metro quadrado")
        CB_Unidade.AddItem ("hect�metro quadrado")
        CB_Unidade.AddItem ("dec�metro quadrado")
        CB_Unidade.AddItem ("metro quadrado")
        CB_Unidade.AddItem ("dec�metro quadrado")
        CB_Unidade.AddItem ("cent�metro quadrado")
        CB_Unidade.AddItem ("mil�metro quadrado")
        CB_Unidade.AddItem ("micr�metro quadrado")
        CB_Unidade.AddItem ("jarda quadrada")
        CB_Unidade.AddItem ("p� quadrado")
        CB_Unidade.AddItem ("polegada quadrada")
        CB_UA.AddItem ("km�")
        CB_UA.AddItem ("hm�")
        CB_UA.AddItem ("dam�")
        CB_UA.AddItem ("m�")
        CB_UA.AddItem ("dm�")
        CB_UA.AddItem ("cm�")
        CB_UA.AddItem ("mm�")
        CB_UA.AddItem ("�m�")
        CB_UA.AddItem ("yd�")
        CB_UA.AddItem ("ft�")
        CB_UA.AddItem ("in�")
    ElseIf CB_Grandeza.Text = "Volume" Then
        CB_Unidade.AddItem ("quil�metro c�bico")
        CB_Unidade.AddItem ("hect�metro c�bico")
        CB_Unidade.AddItem ("dec�metro c�bico")
        CB_Unidade.AddItem ("metro c�bico")
        CB_Unidade.AddItem ("dec�metro c�bico")
        CB_Unidade.AddItem ("cent�metro c�bico")
        CB_Unidade.AddItem ("mil�metro c�bico")
        CB_Unidade.AddItem ("micr�metro c�bico")
        CB_Unidade.AddItem ("jarda c�bica")
        CB_Unidade.AddItem ("p� c�bico")
        CB_Unidade.AddItem ("polegada c�bica")
        CB_Unidade.AddItem ("Litro")
        CB_Unidade.AddItem ("miliLitro")
        CB_Unidade.AddItem ("Gal�o ingl�s")
        CB_Unidade.AddItem ("Gal�o americano")
        CB_UA.AddItem ("km�")
        CB_UA.AddItem ("hm�")
        CB_UA.AddItem ("dam�")
        CB_UA.AddItem ("m�")
        CB_UA.AddItem ("dm�")
        CB_UA.AddItem ("cm�")
        CB_UA.AddItem ("mm�")
        CB_UA.AddItem ("�m�")
        CB_UA.AddItem ("yd�")
        CB_UA.AddItem ("ft�")
        CB_UA.AddItem ("in�")
        CB_UA.AddItem ("l")
        CB_UA.AddItem ("ml")
        CB_UA.AddItem ("U.K. gal")
        CB_UA.AddItem ("U.S. gal")
    ElseIf CB_Grandeza.Text = "Massa" Then
        CB_Unidade.AddItem ("quilograma")
        CB_Unidade.AddItem ("grama")
        CB_Unidade.AddItem ("Unidade T�cnica Massa")
        CB_Unidade.AddItem ("libra")
        CB_Unidade.AddItem ("on�a")
        CB_Unidade.AddItem ("slug")
        CB_Unidade.AddItem ("stone")
        CB_Unidade.AddItem ("tonelada")
        CB_Unidade.AddItem ("tonelada brit�nica")
        CB_Unidade.AddItem ("tonelada americana")
        CB_Unidade.AddItem ("Hundred Weight brit�nica")
        CB_Unidade.AddItem ("Hundred Weight americana")
        CB_UA.AddItem ("kg")
        CB_UA.AddItem ("g")
        CB_UA.AddItem ("utm")
        CB_UA.AddItem ("lb")
        CB_UA.AddItem ("oz")
        CB_UA.AddItem ("slug")
        CB_UA.AddItem ("stone")
        CB_UA.AddItem ("ton")
        CB_UA.AddItem ("U.K. ton")
        CB_UA.AddItem ("U.S. ton")
        CB_UA.AddItem ("U.K. cwt")
        CB_UA.AddItem ("U.S. cwt")
    
    End If
    For I = 0 To CB_Unidade.ListCount - 1
        LB_U(I).Caption = CB_UA.List(I)
        LB_V(I).Caption = ""
        LB_N(I).Caption = CB_Unidade.List(I)
    Next I
End Sub
Private Sub CB_Grandeza_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_Unidade.SetFocus
End Sub
Private Sub CB_Unidade_Click()
    ConverteUnidades
End Sub
Private Sub CB_Unidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_V.SetFocus
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    'Abre tela carregamento
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (2)
    DLL_CARGA.ResetaBP
    'Montando tela
    DLL_CARGA.CarregaTexto ("Organizando tela...")
    LimpaUnidade
    CB_UA.Visible = False
    TXT_V.Text = ""
    CB_Grandeza.ListIndex = -1
    CB_Unidade.ListIndex = -1
    DLL_FUNCS.RegistraEvento "Abrir Conversor de Unidades", ""
    DLL_CARGA.CarregaTexto ("Finalizando...")
    DLL_CARGA.Exibe (False)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_V_Change()
    ConverteUnidades
End Sub
Private Sub TXT_V_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CB_Grandeza.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUN��ES E ROTINAS
'***************************************************

Private Static Sub LimpaUnidade()
    On Error GoTo ERRO_SISCOVAL
    For I = 0 To 15
        LB_U(I).Caption = ""
        LB_V(I).Caption = ""
        LB_N(I).Caption = ""
    Next I
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ConverteUnidades()
    If TXT_V.Text = "" Then Exit Sub
    If IsNumeric(TXT_V.Text) = False Then Exit Sub
    If CB_Grandeza.Text = "" Then Exit Sub
    If CB_Unidade.Text = "" Then Exit Sub
    On Error GoTo ERRO_SISCOVAL
    'come�a conversoes
    Dim NumCon As Double
    NumCon = TXT_V.Text
    If CB_Grandeza.Text = "Temperatura" Then
        If CB_Unidade.Text = "Celsius" Then
            LB_V(0).Caption = CDbl(NumCon)
            LB_V(1).Caption = CDbl((9 / 5) * NumCon + 32)
            LB_V(2).Caption = CDbl(NumCon + 273.15)
            LB_V(3).Caption = CDbl(Val(LB_V(1).Caption) + 459.67)
        ElseIf CB_Unidade.Text = "Farenheit" Then
            LB_V(1).Caption = CDbl(NumCon)
            LB_V(0).Caption = CDbl(5 / 9 * (NumCon - 32))
            LB_V(2).Caption = CDbl(Val(LB_V(0).Caption) + 273.15)
            LB_V(3).Caption = CDbl(Val(NumCon) + 459.67)
        ElseIf CB_Unidade.Text = "Kelvin" Then
            LB_V(2).Caption = CDbl(NumCon)
            LB_V(0).Caption = CDbl(NumCon - 273.15)
            LB_V(1).Caption = CDbl((9 / 5) * Val(LB_V(0).Caption) + 32)
            LB_V(3).Caption = CDbl(Val(LB_V(1).Caption) + 459.67)
        ElseIf CB_Unidade.Text = "Rankine" Then
            LB_V(3).Caption = CDbl(NumCon)
            LB_V(1).Caption = CDbl(NumCon - 459.67)
            LB_V(0).Caption = CDbl(5 / 9 * (Val(LB_V(1).Caption) - 32))
            LB_V(2).Caption = CDbl(Val(LB_V(0).Caption) + 273.15)
        End If
    ElseIf CB_Grandeza.Text = "Comprimento" Then
        'usei unidade princiapal - metro
        If CB_Unidade.Text = "quil�metro" Then
            NumCon = NumCon * 1000#
        ElseIf CB_Unidade.Text = "hect�metro" Then
            NumCon = NumCon * 100#
        ElseIf CB_Unidade.Text = "dec�metro" Then
            NumCon = NumCon * 10#
        ElseIf CB_Unidade.Text = "metro" Then
            NumCon = NumCon * 1#
        ElseIf CB_Unidade.Text = "dec�metro" Then
            NumCon = NumCon * 0.1
        ElseIf CB_Unidade.Text = "cent�metro" Then
            NumCon = NumCon * 0.01
        ElseIf CB_Unidade.Text = "mil�metro" Then
            NumCon = NumCon * 0.001
        ElseIf CB_Unidade.Text = "micr�metro" Then
            NumCon = NumCon * 0.000001
        ElseIf CB_Unidade.Text = "jarda" Then
            NumCon = NumCon * 0.9144
        ElseIf CB_Unidade.Text = "p�" Then
            NumCon = NumCon * 0.3048
        ElseIf CB_Unidade.Text = "polegada" Then
            NumCon = NumCon * 0.00254
        ElseIf CB_Unidade.Text = "milha mar�tima" Then
            NumCon = NumCon * 1852#
        ElseIf CB_Unidade.Text = "milha terrestre" Then
            NumCon = NumCon * 1609.3
        ElseIf CB_Unidade.Text = "ano luz" Then
            NumCon = NumCon * 9.46E+15
        End If
        LB_V(0).Caption = CDbl(NumCon * 0.001) 'quil�metro
        LB_V(1).Caption = CDbl(NumCon * 0.01) 'hect�metro
        LB_V(2).Caption = CDbl(NumCon * 0.1) 'dec�metro
        LB_V(3).Caption = CDbl(NumCon) 'metro
        LB_V(4).Caption = CDbl(NumCon * 10#) 'dec�metro
        LB_V(5).Caption = CDbl(NumCon * 100#) 'cent�metro
        LB_V(6).Caption = CDbl(NumCon * 1000#) 'mil�metro
        LB_V(7).Caption = CDbl(NumCon * 1000000#) 'micr�metro
        LB_V(8).Caption = CDbl(NumCon * 1.094) 'jarda
        LB_V(9).Caption = CDbl(NumCon * 3.2808) 'p�
        LB_V(10).Caption = CDbl(NumCon * 39.37) 'polegada
        LB_V(11).Caption = CDbl(NumCon * (1 / 1852#)) 'milha mar�tima
        LB_V(12).Caption = CDbl(NumCon * (1 / 1609.3)) 'milha terrestre
        LB_V(13).Caption = CDbl(NumCon * (1 / 9.46E+15)) 'ano luz
    ElseIf CB_Grandeza.Text = "Press�o" Then
        'usei unidade de Press�o principal - kgf/cm�
        If CB_Unidade.Text = "Pascal" Then
            NumCon = NumCon * 0.0000102
        ElseIf CB_Unidade.Text = "megaPascal" Then
            NumCon = NumCon * 0.0102
        ElseIf CB_Unidade.Text = "Atmosf�rica" Then
            NumCon = NumCon * 1.033
        ElseIf CB_Unidade.Text = "bar" Then
            NumCon = NumCon * 1.02
        ElseIf CB_Unidade.Text = "b�ria" Then
            NumCon = NumCon * 0.00000102
        ElseIf CB_Unidade.Text = "kgf por m�" Then
            NumCon = NumCon * 0.0001
        ElseIf CB_Unidade.Text = "Atmosfera t�cnica" Then
            NumCon = NumCon * 1
        ElseIf CB_Unidade.Text = "kgf por mm�" Then
            NumCon = NumCon * 10000#
        ElseIf CB_Unidade.Text = "lbf por ft�" Then
            NumCon = NumCon * 0.00049
        ElseIf CB_Unidade.Text = "PSI" Then
            NumCon = NumCon * 0.0704
        ElseIf CB_Unidade.Text = "Torricelli" Then
            NumCon = NumCon * 0.00136
        ElseIf CB_Unidade.Text = "in de merc�rio" Then
            NumCon = NumCon * 0.0345
        ElseIf CB_Unidade.Text = "ft de �gua" Then
            NumCon = NumCon * 0.0305
        ElseIf CB_Unidade.Text = "m de �gua" Then
            NumCon = NumCon * 0.1
        End If
        LB_V(0).Caption = CDbl(NumCon * 98066.5) 'Pascal
        LB_V(1).Caption = CDbl(NumCon * 98.0665) 'mega Pascal
        LB_V(2).Caption = CDbl(NumCon * 0.96784)   'Atmosf�rica
        LB_V(3).Caption = CDbl(NumCon * 0.98)   'bar
        LB_V(4).Caption = CDbl(NumCon * 9800000#) 'b�ria
        LB_V(5).Caption = CDbl(NumCon * 10000#) 'kgf por m�
        LB_V(6).Caption = CDbl(NumCon) 'Atmosfera t�cnica
        LB_V(7).Caption = CDbl(NumCon * 0.0001) 'kgf por mm�
        LB_V(8).Caption = CDbl(NumCon * 2048) 'lbf por ft�
        LB_V(9).Caption = CDbl(NumCon * 14.2) 'PSI
        LB_V(10).Caption = CDbl(NumCon * 735.56) 'Torricelli
        LB_V(11).Caption = CDbl(NumCon * 28.958) 'in de merc�rio
        LB_V(12).Caption = CDbl(NumCon * 32.808) 'ft de �gua
        LB_V(13).Caption = CDbl(NumCon * 10) 'm de �gua
    ElseIf CB_Grandeza.Text = "�rea" Then
        'usei unidade princiapal - metro�
        If CB_Unidade.Text = "quil�metro quadrado" Then
            NumCon = NumCon * (1000# ^ 2)
        ElseIf CB_Unidade.Text = "hect�metro quadrado" Then
            NumCon = NumCon * (100# ^ 2)
        ElseIf CB_Unidade.Text = "dec�metro quadrado" Then
            NumCon = NumCon * (10# ^ 2)
        ElseIf CB_Unidade.Text = "metro quadrado" Then
            NumCon = NumCon * (1# ^ 2)
        ElseIf CB_Unidade.Text = "dec�metro quadrado" Then
            NumCon = NumCon * (0.1 ^ 2)
        ElseIf CB_Unidade.Text = "cent�metro quadrado" Then
            NumCon = NumCon * (0.01 ^ 2)
        ElseIf CB_Unidade.Text = "mil�metro quadrado" Then
            NumCon = NumCon * (0.001 ^ 2)
        ElseIf CB_Unidade.Text = "micr�metro quadrado" Then
            NumCon = NumCon * (0.000001 ^ 2)
        ElseIf CB_Unidade.Text = "jarda quadrada" Then
            NumCon = NumCon * (0.9144 ^ 2)
        ElseIf CB_Unidade.Text = "p� quadrado" Then
            NumCon = NumCon * (0.3048 ^ 2)
        ElseIf CB_Unidade.Text = "polegada quadrada" Then
            NumCon = NumCon * (0.00254 ^ 2)
        End If
        LB_V(0).Caption = CDbl(NumCon * (0.001 ^ 2)) 'quil�metro quadrado
        LB_V(1).Caption = CDbl(NumCon * (0.01 ^ 2)) 'hect�metro quadrado
        LB_V(2).Caption = CDbl(NumCon * (0.1 ^ 2)) 'dec�metro quadrado
        LB_V(3).Caption = CDbl(NumCon) 'metro quadrado
        LB_V(4).Caption = CDbl(NumCon * (10# ^ 2)) 'dec�metro quadrado
        LB_V(5).Caption = CDbl(NumCon * (100# ^ 2)) 'cent�metro quadrado
        LB_V(6).Caption = CDbl(NumCon * (1000# ^ 2)) 'mil�metro quadrado
        LB_V(7).Caption = CDbl(NumCon * (1000000# ^ 2)) 'micr�metro quadrado
        LB_V(8).Caption = CDbl(NumCon * (1.094 ^ 2)) 'jarda quadrada
        LB_V(9).Caption = CDbl(NumCon * (3.2808 ^ 2)) 'p� quadrado
        LB_V(10).Caption = CDbl(NumCon * (39.37 ^ 2)) 'polegada quadrada
    ElseIf CB_Grandeza.Text = "Volume" Then
        'usei unidade princiapal - metro�
        If CB_Unidade.Text = "quil�metro c�bico" Then
            NumCon = NumCon * (1000# ^ 3)
        ElseIf CB_Unidade.Text = "hect�metro c�bico" Then
            NumCon = NumCon * (100# ^ 3)
        ElseIf CB_Unidade.Text = "dec�metro c�bico" Then
            NumCon = NumCon * (10# ^ 3)
        ElseIf CB_Unidade.Text = "metro c�bico" Then
            NumCon = NumCon * (1# ^ 3)
        ElseIf CB_Unidade.Text = "dec�metro c�bico" Then
            NumCon = NumCon * (0.1 ^ 3)
        ElseIf CB_Unidade.Text = "cent�metro c�bico" Then
            NumCon = NumCon * (0.01 ^ 3)
        ElseIf CB_Unidade.Text = "mil�metro c�bico" Then
            NumCon = NumCon * (0.001 ^ 3)
        ElseIf CB_Unidade.Text = "micr�metro c�bico" Then
            NumCon = NumCon * (0.000001 ^ 3)
        ElseIf CB_Unidade.Text = "jarda c�bica" Then
            NumCon = NumCon * (0.9144 ^ 3)
        ElseIf CB_Unidade.Text = "p� c�bico" Then
            NumCon = NumCon * (0.3048 ^ 3)
        ElseIf CB_Unidade.Text = "polegada c�bica" Then
            NumCon = NumCon * (0.00254 ^ 3)
        ElseIf CB_Unidade.Text = "Litro" Then
            NumCon = NumCon * 0.001
        ElseIf CB_Unidade.Text = "miliLitro" Then
            NumCon = NumCon * 0.000001
        ElseIf CB_Unidade.Text = "Gal�o ingl�s" Then
            NumCon = NumCon * 0.00455
        ElseIf CB_Unidade.Text = "Gal�o americano" Then
            NumCon = NumCon * 0.00378
        End If
        LB_V(0).Caption = CDbl(NumCon * (0.001 ^ 3)) 'quil�metro c�bico
        LB_V(1).Caption = CDbl(NumCon * (0.01 ^ 3)) 'hect�metro c�bico
        LB_V(2).Caption = CDbl(NumCon * (0.1 ^ 3)) 'dec�metro c�bico
        LB_V(3).Caption = CDbl(NumCon) 'metro c�bico
        LB_V(4).Caption = CDbl(NumCon * (10# ^ 3)) 'dec�metro c�bico
        LB_V(5).Caption = CDbl(NumCon * (100# ^ 3)) 'cent�metro c�bico
        LB_V(6).Caption = CDbl(NumCon * (1000# ^ 3)) 'mil�metro c�bico
        LB_V(7).Caption = CDbl(NumCon * (1000000# ^ 3)) 'micr�metro c�bico
        LB_V(8).Caption = CDbl(NumCon * (1.094 ^ 3)) 'jarda c�bica
        LB_V(9).Caption = CDbl(NumCon * (3.2808 ^ 3)) 'p� c�bico
        LB_V(10).Caption = CDbl(NumCon * (39.37 ^ 3)) 'polegada c�bica
        LB_V(11).Caption = CDbl(NumCon * 1000) 'Litro
        LB_V(12).Caption = CDbl(NumCon * 1000000#) 'miliLitro
        LB_V(13).Caption = CDbl(NumCon * 220) 'Gal�o ingl�s
        LB_V(14).Caption = CDbl(NumCon * 264.2) 'Gal�o americano
    ElseIf CB_Grandeza.Text = "Massa" Then
        'usei unidade princiapal - grama
        If CB_Unidade.Text = "quilograma" Then
            NumCon = NumCon * 1000#
        ElseIf CB_Unidade.Text = "grama" Then
            NumCon = NumCon * 1
        ElseIf CB_Unidade.Text = "Unidade T�cnica Massa" Then
            NumCon = NumCon * 9806.65
        ElseIf CB_Unidade.Text = "libra" Then
            NumCon = NumCon * 1
        ElseIf CB_Unidade.Text = "on�a" Then
            NumCon = NumCon * 28.35
        ElseIf CB_Unidade.Text = "slug" Then
            NumCon = NumCon * 14591
        ElseIf CB_Unidade.Text = "stone" Then
            NumCon = NumCon * 6350
        ElseIf CB_Unidade.Text = "tonelada" Then
            NumCon = NumCon * 1000000#
        ElseIf CB_Unidade.Text = "tonelada brit�nica" Then
            NumCon = NumCon * 1016050
        ElseIf CB_Unidade.Text = "tonelada americana" Then
            NumCon = NumCon * 907185
        ElseIf CB_Unidade.Text = "Hundred Weight brit�nica" Then
            NumCon = NumCon * 50802
        ElseIf CB_Unidade.Text = "Hundred Weight americana" Then
            NumCon = NumCon * 45359
        End If
        LB_V(0).Caption = CDbl(NumCon * 0.001) 'quilograma
        LB_V(1).Caption = CDbl(NumCon * 1) 'grama
        LB_V(2).Caption = CDbl(NumCon * 0.000102) 'Unidade T�cnica Massa
        LB_V(3).Caption = CDbl(NumCon * 0.0022) 'libra
        LB_V(4).Caption = CDbl(NumCon * 0.0353)  'on�a
        LB_V(5).Caption = CDbl(NumCon * 0.0000685) 'slug
        LB_V(6).Caption = CDbl(NumCon * 0.00015748) 'stone
        LB_V(7).Caption = CDbl(NumCon * 0.000001) 'tonelada
        LB_V(8).Caption = CDbl(NumCon * (1 / 1016050)) 'tonelada brit�nica
        LB_V(9).Caption = CDbl(NumCon * (1 / 907185)) 'tonelada americana
        LB_V(10).Caption = CDbl(NumCon * (1 / 50802)) 'Hundred Weight brit�nica
        LB_V(11).Caption = CDbl(NumCon * (1 / 45359)) 'Hundred Weight americana
 
 
    
    End If
    For I = 0 To 15
        LB_V(I).ToolTipText = LB_V(I).Caption
    Next I
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
