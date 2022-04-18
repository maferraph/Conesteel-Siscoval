VERSION 5.00
Begin VB.UserControl Valvula 
   ClientHeight    =   3645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10710
   ScaleHeight     =   3645
   ScaleWidth      =   10710
   Begin VB.Frame FR_Dados 
      Caption         =   "Dados sobre a válvula selecionada:"
      Height          =   855
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   10695
      Begin VB.TextBox TXT_Quantidade 
         Height          =   285
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LB_Revestimento 
         AutoSize        =   -1  'True
         Caption         =   "XU - Stellite"
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
         Left            =   7440
         TabIndex        =   71
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label LB_Internos 
         AutoSize        =   -1  'True
         Caption         =   "Inconel"
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
         Left            =   6600
         TabIndex        =   70
         Top             =   480
         Width           =   645
      End
      Begin VB.Label LB_Material 
         AutoSize        =   -1  'True
         Caption         =   "A105"
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
         Left            =   5520
         TabIndex        =   69
         Top             =   480
         Width           =   450
      End
      Begin VB.Label LB_Bitola 
         AutoSize        =   -1  'True
         Caption         =   "1.1/2"""
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
         Left            =   4680
         TabIndex        =   68
         Top             =   480
         Width           =   570
      End
      Begin VB.Label LB_Extremidade 
         AutoSize        =   -1  'True
         Caption         =   "Flange"
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
         Left            =   2880
         TabIndex        =   67
         Top             =   480
         Width           =   585
      End
      Begin VB.Label LB_Classe 
         AutoSize        =   -1  'True
         Caption         =   "1500"
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
         TabIndex        =   66
         Top             =   480
         Width           =   435
      End
      Begin VB.Label LB_Valvula 
         AutoSize        =   -1  'True
         Caption         =   "Retenção Portinhola"
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
         Left            =   960
         TabIndex        =   65
         Top             =   480
         Width           =   1755
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Revestimento:"
         Height          =   195
         Index           =   17
         Left            =   7440
         TabIndex        =   64
         Top             =   240
         Width           =   1020
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Internos:"
         Height          =   195
         Index           =   16
         Left            =   6600
         TabIndex        =   63
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   195
         Index           =   15
         Left            =   5520
         TabIndex        =   62
         Top             =   240
         Width           =   600
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Bitola:"
         Height          =   195
         Index           =   14
         Left            =   4680
         TabIndex        =   61
         Top             =   240
         Width           =   435
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Extremidade:"
         Height          =   195
         Index           =   13
         Left            =   2880
         TabIndex        =   60
         Top             =   240
         Width           =   915
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Classe:"
         Height          =   195
         Index           =   12
         Left            =   3960
         TabIndex        =   59
         Top             =   240
         Width           =   510
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Válvula:"
         Height          =   195
         Index           =   11
         Left            =   960
         TabIndex        =   58
         Top             =   240
         Width           =   570
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Quant."
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   57
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame FR_Valvula 
      Caption         =   "VÁLVULA:"
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   1935
      Begin VB.CommandButton BT_PontaAgulha 
         Caption         =   "Ponta de Agulha"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton BT_Portinhola 
         Caption         =   "Retenção Portinhola"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton BT_Pistao 
         Caption         =   "Retenção Pistão"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton BT_Globo 
         Caption         =   "Globo"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton BT_Gaveta 
         Caption         =   "Gaveta"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FR_Extremidade 
      Caption         =   "EXTREM.:"
      Height          =   1575
      Left            =   1920
      TabIndex        =   14
      Top             =   960
      Width           =   975
      Begin VB.CommandButton BT_OutraExtremidade 
         Caption         =   "Outra"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton BT_Flange 
         Caption         =   "Flange"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton BT_BSP 
         Caption         =   "BSP"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton BT_SW 
         Caption         =   "SW"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton BT_NPT 
         Caption         =   "NPT"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FR_Classe 
      Caption         =   "CLASSE:"
      Height          =   1575
      Left            =   2880
      TabIndex        =   6
      Top             =   960
      Width           =   1455
      Begin VB.CommandButton BT_OutraClasse 
         Caption         =   "Outra"
         Height          =   255
         Left            =   720
         TabIndex        =   47
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton BT_3000 
         Caption         =   "3000"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton BT_1500 
         Caption         =   "1500"
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton BT_900 
         Caption         =   "900"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton BT_800 
         Caption         =   "800"
         Height          =   255
         Left            =   720
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_600 
         Caption         =   "600"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_300 
         Caption         =   "300"
         Height          =   255
         Left            =   720
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton BT_150 
         Caption         =   "150"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame FR_Bitola 
      Caption         =   "BITOLA:"
      Height          =   1575
      Left            =   4320
      TabIndex        =   20
      Top             =   960
      Width           =   1455
      Begin VB.CommandButton BT_OutraBitola 
         Caption         =   "Outra"
         Height          =   255
         Left            =   720
         TabIndex        =   28
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton BT_2 
         Caption         =   "2"""
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton BT_1_1_2 
         Caption         =   "1.1/2"""
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_1_1_4 
         Caption         =   "1.1/4"""
         Height          =   255
         Left            =   720
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton BT_1 
         Caption         =   "1"""
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   615
      End
      Begin VB.CommandButton BT_3_4 
         Caption         =   "3/4"""
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton BT_1_2 
         Caption         =   "1/2"""
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton BT_3_8 
         Caption         =   "3/8"""
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame FR_Material 
      Caption         =   "MATERIAL:"
      Height          =   1575
      Left            =   5760
      TabIndex        =   29
      Top             =   960
      Width           =   1455
      Begin VB.CommandButton BT_OutroMaterial 
         Caption         =   "Outro"
         Height          =   255
         Left            =   720
         TabIndex        =   48
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton BT_LF2 
         Caption         =   "LF2"
         Height          =   255
         Left            =   720
         TabIndex        =   30
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton BT_F11 
         Caption         =   "F11"
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton BT_F5 
         Caption         =   "F5"
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton BT_F316L 
         Caption         =   "F316L"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   615
      End
      Begin VB.CommandButton BT_F304L 
         Caption         =   "F304L"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   615
      End
      Begin VB.CommandButton BT_F316 
         Caption         =   "F316"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
      Begin VB.CommandButton BT_F304 
         Caption         =   "F304"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   615
      End
      Begin VB.CommandButton BT_A105 
         Caption         =   "A105"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FR_Internos 
      Caption         =   "INTERNOS:"
      Height          =   1575
      Left            =   7200
      TabIndex        =   38
      Top             =   960
      Width           =   1695
      Begin VB.CommandButton BT_OutroInterno 
         Caption         =   "Outro"
         Height          =   255
         Left            =   840
         TabIndex        =   46
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton BT_Alloy20 
         Caption         =   "Alloy20"
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton BT_Inconel 
         Caption         =   "Inconel"
         Height          =   255
         Left            =   840
         TabIndex        =   49
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton BT_Monel 
         Caption         =   "Monel"
         Height          =   255
         Left            =   840
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton BT_316L 
         Caption         =   "316L"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   1200
         Width           =   735
      End
      Begin VB.CommandButton BT_304L 
         Caption         =   "304L"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton BT_316 
         Caption         =   "316"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton BT_304 
         Caption         =   "304"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton BT_410 
         Caption         =   "410"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame FR_Revestimento 
      Caption         =   "REVESTIMENTO:"
      Height          =   1575
      Left            =   8880
      TabIndex        =   50
      Top             =   960
      Width           =   1815
      Begin VB.CommandButton BT_OutroRevestimento 
         Caption         =   "Outro"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton BT_UU 
         Caption         =   "Stellite ( UU )"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton BT_XU 
         Caption         =   "Stellite ( XU )"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton BT_XX 
         Caption         =   "Não aplicado ( XX )"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FR_Especial 
      Caption         =   "Casos Especiais:"
      Height          =   1095
      Left            =   0
      TabIndex        =   72
      Top             =   2520
      Width           =   10695
      Begin VB.CommandButton Command10 
         Caption         =   "Agulha"
         Height          =   255
         Left            =   4320
         TabIndex        =   85
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Normal"
         Height          =   255
         Left            =   4320
         TabIndex        =   84
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "PTFE"
         Height          =   255
         Left            =   3240
         TabIndex        =   83
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Grafite"
         Height          =   255
         Left            =   3240
         TabIndex        =   82
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Outra"
         Height          =   255
         Left            =   2280
         TabIndex        =   80
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "F.Carbono"
         Height          =   255
         Left            =   2280
         TabIndex        =   79
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "PTFE"
         Height          =   255
         Left            =   1440
         TabIndex        =   78
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Grafite"
         Height          =   255
         Left            =   1440
         TabIndex        =   77
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Soldado"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aparafusado"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Globo:"
         Height          =   195
         Index           =   3
         Left            =   4320
         TabIndex        =   86
         Top             =   240
         Width           =   465
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Junta:"
         Height          =   195
         Index           =   2
         Left            =   3240
         TabIndex        =   81
         Top             =   240
         Width           =   435
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Gaxeta:"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   76
         Top             =   240
         Width           =   555
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Castelo:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   73
         Top             =   240
         Width           =   570
      End
   End
End
Attribute VB_Name = "Valvula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'***********************************
'            VARIAVEIS
'***********************************
Private VALVULA_TIPO As String
Private VALVULA_CLASSE As String
Private VALVULA_EXTREMIDADE As String
Private VALVULA_BITOLA As String
Private VALVULA_MATERIAL As String
Private VALVULA_INTERNOS As String
Private VALVULA_REVESTIMENTO As String

Public C01_Q As Integer
Private C01_U As Double
Public C01_C As String
Public C01_N As String
Public C01_B As String
Public C01_M As String
Public C02_Q As Integer
Private C02_U As Double
Public C02_C As String
Public C02_N As String
Public C02_B As String
Public C02_M As String
Public C03_Q As Integer
Private C03_U As Double
Public C03_C As String
Public C03_N As String
Public C03_B As String
Public C03_M As String
Public C04_Q As Integer
Private C04_U As Double
Public C04_C As String
Public C04_N As String
Public C04_B As String
Public C04_M As String
Public C05_Q As Integer
Private C05_U As Double
Public C05_C As String
Public C05_N As String
Public C05_B As String
Public C05_M As String
Public C06_Q As Integer
Private C06_U As Double
Public C06_C As String
Public C06_N As String
Public C06_B As String
Public C06_M As String
Public C07_Q As Integer
Private C07_U As Double
Public C07_C As String
Public C07_N As String
Public C07_B As String
Public C07_M As String
Public C08_Q As Integer
Private C08_U As Double
Public C08_C As String
Public C08_N As String
Public C08_B As String
Public C08_M As String
Public C09_Q As Integer
Private C09_U As Double
Public C09_C As String
Public C09_N As String
Public C09_B As String
Public C09_M As String
Public C10_Q As Integer
Private C10_U As Double
Public C10_C As String
Public C10_N As String
Public C10_B As String
Public C10_M As String
Public C11_Q As Integer
Private C11_U As Double
Public C11_C As String
Public C11_N As String
Public C11_B As String
Public C11_M As String
Public C12_Q As Integer
Private C12_U As Double
Public C12_C As String
Public C12_N As String
Public C12_B As String
Public C12_M As String
Public C13_Q As Integer
Private C13_U As Double
Public C13_C As String
Public C13_N As String
Public C13_B As String
Public C13_M As String
Public C14_Q As Integer
Private C14_U As Double
Public C14_C As String
Public C14_N As String
Public C14_B As String
Public C14_M As String
Public C15_Q As Integer
Private C15_U As Double
Public C15_C As String
Public C15_N As String
Public C15_B As String
Public C15_M As String
Public C16_Q As Integer
Private C16_U As Double
Public C16_C As String
Public C16_N As String
Public C16_B As String
Public C16_M As String
Public C17_Q As Integer
Private C17_U As Double
Public C17_C As String
Public C17_N As String
Public C17_B As String
Public C17_M As String
Public C18_Q As Integer
Private C18_U As Double
Public C18_C As String
Public C18_N As String
Public C18_B As String
Public C18_M As String
Public C19_Q As Integer
Private C19_U As Double
Public C19_C As String
Public C19_N As String
Public C19_B As String
Public C19_M As String
Public C20_Q As Integer
Private C20_U As Double
Public C20_C As String
Public C20_N As String
Public C20_B As String
Public C20_M As String

Dim tmp As String


Private Sub BT_1_1_2_Click()
    VALVULA_BITOLA = "1.1/2"
    LB_Bitola.Caption = "1.1/2" & Chr(34)
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_1_1_4_Click()
    VALVULA_BITOLA = "1.1/4"
    LB_Bitola.Caption = "1.1/4" & Chr(34)
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_1_2_Click()
    VALVULA_BITOLA = "1/2"
    LB_Bitola.Caption = "1/2" & Chr(34)
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_1_Click()
    VALVULA_BITOLA = "1"
    LB_Bitola.Caption = "1" & Chr(34)
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_150_Click()
    VALVULA_CLASSE = "150"
    LB_Classe.Caption = "150"
    LigaFRBitola True
    If VALVULA_TIPO = "Gaveta" Then
        C18_C = "FLANGE 150"
        C19_C = "SOLDA"
        C18_N = "Flange"
        C19_N = "Solda de Adaptação"
        C18_U = 2
        C19_U = 2
    ElseIf VALVULA_TIPO = "Globo" Then
        C19_C = "FLANGE 150"
        C20_C = "SOLDA"
        C19_N = "Flange"
        C20_N = "Solda de Adaptação"
        C19_U = 2
        C20_U = 2
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        C11_C = "FLANGE 150"
        C12_C = "SOLDA"
        C11_N = "Flange"
        C12_N = "Solda de Adaptação"
        C11_U = 2
        C12_U = 2
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        C14_C = "FLANGE 150"
        C15_C = "SOLDA"
        C14_N = "Flange"
        C15_N = "Solda de Adaptação"
        C14_U = 2
        C15_U = 2
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub BT_1500_Click()
    VALVULA_CLASSE = "1500"
    LB_Classe.Caption = "1500"
    LigaFRBitola True
    If VALVULA_EXTREMIDADE = "Flangeada" Then
        If VALVULA_TIPO = "Gaveta" Then
            C18_C = "FLANGE 1500"
            C19_C = "SOLDA"
            C18_N = "Flange"
            C19_N = "Solda de Adaptação"
            C18_U = 2
            C19_U = 2
        ElseIf VALVULA_TIPO = "Globo" Then
            C19_C = "FLANGE 1500"
            C20_C = "SOLDA"
            C19_N = "Flange"
            C20_N = "Solda de Adaptação"
            C19_U = 2
            C20_U = 2
        ElseIf VALVULA_TIPO = "Retenção Pistão" Then
            C11_C = "FLANGE 1500"
            C12_C = "SOLDA"
            C11_N = "Flange"
            C12_N = "Solda de Adaptação"
            C11_U = 2
            C12_U = 2
        ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
            C14_C = "FLANGE 1500"
            C15_C = "SOLDA"
            C14_N = "Flange"
            C15_N = "Solda de Adaptação"
            C14_U = 2
            C15_U = 2
        ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
        
        End If
    Else
        If VALVULA_TIPO = "Gaveta" Then
            C18_C = ""
            C19_C = ""
            C18_N = ""
            C19_N = ""
            C18_U = 0
            C19_U = 0
        ElseIf VALVULA_TIPO = "Globo" Then
            C19_C = ""
            C20_C = ""
            C19_N = ""
            C20_N = ""
            C19_U = 0
            C20_U = 0
        ElseIf VALVULA_TIPO = "Retenção Pistão" Then
            C11_C = ""
            C12_C = ""
            C11_N = ""
            C12_N = ""
            C11_U = 0
            C12_U = 0
        ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
            C14_C = ""
            C15_C = ""
            C14_N = ""
            C15_N = ""
            C14_U = 0
            C15_U = 0
        ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
        
        End If
    End If
End Sub
Private Sub BT_2_Click()
    VALVULA_BITOLA = "2"
    LB_Bitola.Caption = "2" & Chr(34)
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_3_4_Click()
    VALVULA_BITOLA = "3/4"
    LB_Bitola.Caption = "3/4" & Chr(34)
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_3_8_Click()
    VALVULA_BITOLA = "3/8"
    LB_Bitola.Caption = "3/8" & Chr(34)
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_300_Click()
    VALVULA_CLASSE = "300"
    LB_Classe.Caption = "300"
    LigaFRBitola True
    If VALVULA_TIPO = "Gaveta" Then
        C18_C = "FLANGE 300"
        C19_C = "SOLDA"
        C18_N = "Flange"
        C19_N = "Solda de Adaptação"
        C18_U = 2
        C19_U = 2
    ElseIf VALVULA_TIPO = "Globo" Then
        C19_C = "FLANGE 300"
        C20_C = "SOLDA"
        C19_N = "Flange"
        C20_N = "Solda de Adaptação"
        C19_U = 2
        C20_U = 2
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        C11_C = "FLANGE 300"
        C12_C = "SOLDA"
        C11_N = "Flange"
        C12_N = "Solda de Adaptação"
        C11_U = 2
        C12_U = 2
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        C14_C = "FLANGE 300"
        C15_C = "SOLDA"
        C14_N = "Flange"
        C15_N = "Solda de Adaptação"
        C14_U = 2
        C15_U = 2
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub BT_3000_Click()
    VALVULA_CLASSE = "3000"
    LB_Classe.Caption = "3000"
    LigaFRBitola True
    If VALVULA_TIPO = "Gaveta" Then
        C18_C = ""
        C19_C = ""
        C18_N = ""
        C19_N = ""
        C18_U = 0
        C19_U = 0
    ElseIf VALVULA_TIPO = "Globo" Then
        C19_C = ""
        C20_C = ""
        C19_N = ""
        C20_N = ""
        C19_U = 0
        C20_U = 0
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        C11_C = ""
        C12_C = ""
        C11_N = ""
        C12_N = ""
        C11_U = 0
        C12_U = 0
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        C14_C = ""
        C15_C = ""
        C14_N = ""
        C15_N = ""
        C14_U = 0
        C15_U = 0
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub BT_304_Click()
    VALVULA_INTERNOS = "304"
    LB_Internos.Caption = "304"
    InternosComponentes
End Sub
Private Sub BT_304L_Click()
    VALVULA_INTERNOS = "304L"
    LB_Internos.Caption = "304L"
    InternosComponentes
End Sub
Private Sub BT_316_Click()
    VALVULA_INTERNOS = "316"
    LB_Internos.Caption = "316"
    InternosComponentes
End Sub
Private Sub BT_316L_Click()
    VALVULA_INTERNOS = "316L"
    LB_Internos.Caption = "316L"
    InternosComponentes
End Sub
Private Sub BT_410_Click()
    VALVULA_INTERNOS = "410"
    LB_Internos.Caption = "410"
    InternosComponentes
End Sub
Private Sub BT_600_Click()
    VALVULA_CLASSE = "600"
    LB_Classe.Caption = "600"
    LigaFRBitola True
    If VALVULA_TIPO = "Gaveta" Then
        C18_C = "FLANGE 600"
        C19_C = "SOLDA"
        C18_N = "Flange"
        C19_N = "Solda de Adaptação"
        C18_U = 2
        C19_U = 2
    ElseIf VALVULA_TIPO = "Globo" Then
        C19_C = "FLANGE 600"
        C20_C = "SOLDA"
        C19_N = "Flange"
        C20_N = "Solda de Adaptação"
        C19_U = 2
        C20_U = 2
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        C11_C = "FLANGE 600"
        C12_C = "SOLDA"
        C11_N = "Flange"
        C12_N = "Solda de Adaptação"
        C11_U = 2
        C12_U = 2
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        C14_C = "FLANGE 600"
        C15_C = "SOLDA"
        C14_N = "Flange"
        C15_N = "Solda de Adaptação"
        C14_U = 2
        C15_U = 2
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub BT_800_Click()
    VALVULA_CLASSE = "800"
    LB_Classe.Caption = "800"
    LigaFRBitola True
    If VALVULA_TIPO = "Gaveta" Then
        C18_C = ""
        C19_C = ""
        C18_N = ""
        C19_N = ""
        C18_U = 0
        C19_U = 0
    ElseIf VALVULA_TIPO = "Globo" Then
        C19_C = ""
        C20_C = ""
        C19_N = ""
        C20_N = ""
        C19_U = 0
        C20_U = 0
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        C11_C = ""
        C12_C = ""
        C11_N = ""
        C12_N = ""
        C11_U = 0
        C12_U = 0
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        C14_C = ""
        C15_C = ""
        C14_N = ""
        C15_N = ""
        C14_U = 0
        C15_U = 0
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub BT_900_Click()
    VALVULA_CLASSE = "900"
    LB_Classe.Caption = "900"
    LigaFRBitola True
    If VALVULA_TIPO = "Gaveta" Then
        C18_C = "FLANGE 900"
        C19_C = "SOLDA"
        C18_N = "Flange"
        C19_N = "Solda de Adaptação"
        C18_U = 2
        C19_U = 2
    ElseIf VALVULA_TIPO = "Globo" Then
        C19_C = "FLANGE 900"
        C20_C = "SOLDA"
        C19_N = "Flange"
        C20_N = "Solda de Adaptação"
        C19_U = 2
        C20_U = 2
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        C11_C = "FLANGE 900"
        C12_C = "SOLDA"
        C11_N = "Flange"
        C12_N = "Solda de Adaptação"
        C11_U = 2
        C12_U = 2
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        C14_C = "FLANGE 900"
        C15_C = "SOLDA"
        C14_N = "Flange"
        C15_N = "Solda de Adaptação"
        C14_U = 2
        C15_U = 2
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub BT_A105_Click()
    VALVULA_MATERIAL = "A105"
    LB_Material.Caption = "A105"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_Alloy20_Click()
    VALVULA_INTERNOS = "Alloy20"
    LB_Internos.Caption = "Alloy20"
    InternosComponentes
End Sub
Private Sub BT_BSP_Click()
    BT_NPT_Click
    VALVULA_EXTREMIDADE = "BSP"
    LB_Extremidade.Caption = "BSP"
End Sub
Private Sub BT_F11_Click()
    VALVULA_MATERIAL = "A182 F11"
    LB_Material.Caption = "A182 F11"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_F304_Click()
    VALVULA_MATERIAL = "A182 F304"
    LB_Material.Caption = "A182 F304"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_F304L_Click()
    VALVULA_MATERIAL = "A182 F304L"
    LB_Material.Caption = "A182 F304L"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_F316_Click()
    VALVULA_MATERIAL = "A182 F316"
    LB_Material.Caption = "A182 F316"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_F316L_Click()
    VALVULA_MATERIAL = "A182 F316L"
    LB_Material.Caption = "A182 F316L"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_F5_Click()
    VALVULA_MATERIAL = "A182 F5"
    LB_Material.Caption = "A182 F5"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_Flange_Click()
    VALVULA_EXTREMIDADE = "Flangeada"
    LB_Extremidade.Caption = "Flangeada"
    FR_Classe.Enabled = True
    BT_150.Enabled = True
    BT_300.Enabled = True
    BT_600.Enabled = True
    BT_800.Enabled = False
    BT_900.Enabled = True
    BT_1500.Enabled = True
    BT_3000.Enabled = False
    BT_OutraClasse.Enabled = True
End Sub
Private Sub BT_Gaveta_Click()
    On Error Resume Next
    'Liga botoes
    LigaFRExtremidade True
    LigaFRClasse False
    LigaFRBitola False
    LigaFRMaterial False
    LigaFRInternos False
    'Carrega campos
    VALVULA_TIPO = "Gaveta"
    VALVULA_CLASSE = ""
    VALVULA_EXTREMIDADE = ""
    VALVULA_BITOLA = ""
    VALVULA_MATERIAL = ""
    VALVULA_INTERNOS = ""
    VALVULA_REVESTIMENTO = ""
    LB_Valvula.Caption = "Gaveta"
    LB_Classe.Caption = ""
    LB_Extremidade.Caption = ""
    LB_Bitola.Caption = ""
    LB_Material.Caption = ""
    LB_Internos.Caption = ""
    LB_Revestimento.Caption = ""
    'Componentes
    C01_C = "CORPO GAVETA"
    C02_C = "CASTELO"
    C03_C = "PREME"
    C04_C = "CUNHA"
    C05_C = "REDONDO"
    C06_C = "REDONDO"
    C07_C = "JUNTA ESPIROTÁLICA"
    C08_C = "GAXETA"
    C09_C = "PRISIONEIRO CORPO"
    C10_C = "PORCA PRISIONEIRO CORPO"
    C11_C = "PRISIONEIRO PREME"
    C12_C = "PORCA PRISIONEIRO PREME"
    C13_C = "BUCHA DE MOVIMENTO GAVETA"
    C14_C = "SEXTAVADO"
    C15_C = "VOLANTE GAVETA"
    C16_C = "PLACA DE IDENTIFICAÇÃO"
    C17_C = "REVESTIMENTO"
    C18_C = "FLANGE"
    C19_C = "SOLDA"
    C20_C = ""
    'Nome do Componente
    C01_N = "Corpo"
    C02_N = "Castelo"
    C03_N = "Preme"
    C04_N = "Cunha"
    C05_N = "Anél"
    C06_N = "Haste"
    C07_N = "Junta Espirotálica"
    C08_N = "Gaxeta"
    C09_N = "Prisioneiro Corpo"
    C10_N = "Porca Corpo"
    C11_N = "Prisioneiro Preme"
    C12_N = "Porca Preme"
    C13_N = "Bucha de Movimento"
    C14_N = "Porca Volante"
    C15_N = "Volante"
    C16_N = "Placa de Identificação"
    C17_N = "Revestimento"
    C18_N = "Flange"
    C19_N = "Solda de Adaptação"
    C20_N = ""
    'Número de Componente
    C01_U = 1
    C02_U = 1
    C03_U = 1
    C04_U = 1
    C05_U = 2
    C06_U = 1
    C07_U = 1
    C08_U = 1
    C09_U = 4
    C10_U = 4
    C11_U = 2
    C12_U = 4
    C13_U = 1
    C14_U = 1
    C15_U = 1
    C16_U = 1
    C17_U = 0
    C18_U = 0
    C19_U = 0
    C20_U = 0
    'Carrega quantidades
    QuantidadeComponentes
End Sub
Private Sub BT_Globo_Click()
    On Error Resume Next
    'Liga botoes
    LigaFRExtremidade True
    LigaFRClasse False
    LigaFRBitola False
    LigaFRMaterial False
    LigaFRInternos False
    'Carrega campos
    VALVULA_TIPO = "Globo"
    VALVULA_CLASSE = ""
    VALVULA_EXTREMIDADE = ""
    VALVULA_BITOLA = ""
    VALVULA_MATERIAL = ""
    VALVULA_INTERNOS = ""
    VALVULA_REVESTIMENTO = ""
    LB_Valvula.Caption = "Globo"
    LB_Classe.Caption = ""
    LB_Extremidade.Caption = ""
    LB_Bitola.Caption = ""
    LB_Material.Caption = ""
    LB_Internos.Caption = ""
    LB_Revestimento.Caption = ""
    'Componentes
    C01_C = "CORPO GAVETA"
    C02_C = "CASTELO"
    C03_C = "PREME"
    C04_C = "CONTRA-SEDE"
    C05_C = "REDONDO"
    C06_C = "REDONDO"
    C07_C = "JUNTA ESPIROTÁLICA"
    C08_C = "GAXETA"
    C09_C = "PRISIONEIRO CORPO"
    C10_C = "PORCA PRISIONEIRO CORPO"
    C11_C = "PRISIONEIRO PREME"
    C12_C = "PORCA PRISIONEIRO PREME"
    C13_C = "BUCHA DE MOVIMENTO GLOBO"
    C14_C = "PORCA VOLANTE GLOBO"
    C15_C = "ARRUELA VOLANTE GLOBO"
    C16_C = "VOLANTE GLOBO"
    C17_C = "PLACA DE IDENTIFICAÇÃO"
    C18_C = "REVESTIMENTO"
    C19_C = "FLANGE"
    C20_C = "SOLDA ADAPTAÇÃO"
    'Nome do Componente
    C01_N = "Corpo"
    C02_N = "Castelo"
    C03_N = "Preme"
    C04_N = "Contra-Sede"
    C05_N = "Sede"
    C06_N = "Haste"
    C07_N = "Junta Espirotálica"
    C08_N = "Gaxeta"
    C09_N = "Prisioneiro Corpo"
    C10_N = "Porca Corpo"
    C11_N = "Prisioneiro Preme"
    C12_N = "Porca Preme"
    C13_N = "Bucha de Movimento"
    C14_N = "Porca Volante"
    C15_N = "Arruela Volante"
    C16_N = "Volante"
    C17_N = "Placa de Identificação"
    C18_N = "Revestimento"
    C19_N = "Flange"
    C20_N = "Solda de Adaptação"
    'Número de Componente
    C01_U = 1
    C02_U = 1
    C03_U = 1
    C04_U = 1
    C05_U = 1
    C06_U = 1
    C07_U = 1
    C08_U = 1
    C09_U = 4
    C10_U = 4
    C11_U = 2
    C12_U = 4
    C13_U = 1
    C14_U = 1
    C15_U = 1
    C16_U = 1
    C17_U = 1
    C18_U = 0
    C19_U = 0
    C20_U = 0
    'Carrega quantidades
    QuantidadeComponentes
End Sub
Private Sub BT_Inconel_Click()
    VALVULA_INTERNOS = "Inconel"
    LB_Internos.Caption = "Inconel"
    InternosComponentes
End Sub
Private Sub BT_LF2_Click()
    VALVULA_MATERIAL = "A350 LF2"
    LB_Material.Caption = "A350 LF2"
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_Monel_Click()
    VALVULA_INTERNOS = "Monel"
    LB_Internos.Caption = "Monel"
    InternosComponentes
End Sub
Private Sub BT_NPT_Click()
    VALVULA_EXTREMIDADE = "NPT"
    LB_Extremidade.Caption = "NPT"
    FR_Classe.Enabled = True
    BT_150.Enabled = False
    BT_300.Enabled = False
    BT_600.Enabled = False
    BT_800.Enabled = True
    BT_900.Enabled = False
    BT_1500.Enabled = True
    BT_3000.Enabled = False
    BT_OutraClasse.Enabled = True
    If VALVULA_TIPO = "Ponta de Agulha" Then
        BT_800.Enabled = False
        BT_1500.Enabled = False
        BT_3000.Enabled = True
    End If
End Sub
Private Sub BT_OutraBitola_Click()
    tmp = InputBox("Digite a bitola da válvula", "BITOLA")
    VALVULA_BITOLA = tmp
    LB_Bitola.Caption = tmp
    LigaFRMaterial True
    BitolaComponentes
End Sub
Private Sub BT_OutraClasse_Click()
    tmp = InputBox("Digite a classe da válvula", "CLASSE")
    VALVULA_CLASSE = tmp
    LB_Classe.Caption = tmp
    LigaFRBitola True
End Sub
Private Sub BT_OutraExtremidade_Click()
    tmp = InputBox("Digite a extremidade da válvula", "EXTREMIDADE")
    VALVULA_EXTREMIDADE = tmp
    LB_Extremidade.Caption = tmp
    LigaFRClasse True
End Sub
Private Sub BT_OutroMaterial_Click()
    tmp = InputBox("Digite o material da válvula", "MATERIAL")
    VALVULA_MATERIAL = tmp
    LB_Material.Caption = tmp
    LigaFRInternos True
    MaterialComponentes
End Sub
Private Sub BT_Pistao_Click()
    On Error Resume Next
    'Liga botoes
    LigaFRExtremidade True
    LigaFRClasse False
    LigaFRBitola False
    LigaFRMaterial False
    LigaFRInternos False
    'Carrega campos
    VALVULA_TIPO = "Retenção Pistão"
    VALVULA_CLASSE = ""
    VALVULA_EXTREMIDADE = ""
    VALVULA_BITOLA = ""
    VALVULA_MATERIAL = ""
    VALVULA_INTERNOS = ""
    VALVULA_REVESTIMENTO = ""
    LB_Valvula.Caption = "Retenção Pistão"
    LB_Classe.Caption = ""
    LB_Extremidade.Caption = ""
    LB_Bitola.Caption = ""
    LB_Material.Caption = ""
    LB_Internos.Caption = ""
    LB_Revestimento.Caption = ""
    'Componentes
    C01_C = "CORPO GAVETA"
    C02_C = "TAMPA"
    C03_C = "REDONDO"
    C04_C = "REDONDO"
    C05_C = "JUNTA ESPIROTÁLICA"
    C06_C = "PRISIONEIRO CORPO"
    C07_C = "PORCA PRISIONEIRO CORPO"
    C08_C = "MOLA"
    C09_C = "PLACA DE IDENTIFICAÇÃO"
    C10_C = "REVESTIMENTO"
    C11_C = "FLANGE"
    C12_C = "SOLDA"
    C13_C = ""
    C14_C = ""
    C15_C = ""
    C16_C = ""
    C17_C = ""
    C18_C = ""
    C19_C = ""
    C20_C = ""
    'Nome do Componente
    C01_N = "Corpo"
    C02_N = "Tampa"
    C03_N = "Pistão"
    C04_N = "Sede"
    C05_N = "Junta Espirotálica"
    C06_N = "Prisioneiro Corpo"
    C07_N = "Porca Corpo"
    C08_N = "Mola"
    C09_N = "Placa de Identificação"
    C10_N = "Revestimento"
    C11_N = "Flange"
    C12_N = "Solda de Adaptação"
    C13_N = ""
    C14_N = ""
    C15_N = ""
    C16_N = ""
    C17_N = ""
    C18_N = ""
    C19_N = ""
    C20_N = ""
    'Número de Componente
    C01_U = 1
    C02_U = 1
    C03_U = 1
    C04_U = 1
    C05_U = 1
    C06_U = 4
    C07_U = 4
    C08_U = 1
    C09_U = 1
    C10_U = 1
    C11_U = 2
    C12_U = 1
    C13_U = 0
    C14_U = 0
    C15_U = 0
    C16_U = 0
    C17_U = 0
    C18_U = 0
    C19_U = 0
    C20_U = 0
    'Carrega quantidades
    QuantidadeComponentes
End Sub
Private Sub BT_PontaAgulha_Click()
    On Error Resume Next
    'Liga botoes
    LigaFRExtremidade True
    LigaFRClasse False
    LigaFRBitola False
    LigaFRMaterial False
    LigaFRInternos False
    'Carrega campos
    VALVULA_TIPO = "Ponta de Agulha"
    VALVULA_CLASSE = ""
    VALVULA_EXTREMIDADE = ""
    VALVULA_BITOLA = ""
    VALVULA_MATERIAL = ""
    VALVULA_INTERNOS = ""
    VALVULA_REVESTIMENTO = ""
    LB_Valvula.Caption = "Ponta de Agulha"
    LB_Classe.Caption = ""
    LB_Extremidade.Caption = ""
    LB_Bitola.Caption = ""
    LB_Material.Caption = ""
    LB_Internos.Caption = ""
    LB_Revestimento.Caption = ""
    
    QuantidadeComponentes
End Sub
Private Sub BT_Portinhola_Click()
    On Error Resume Next
    'Liga botoes
    LigaFRExtremidade True
    LigaFRClasse False
    LigaFRBitola False
    LigaFRMaterial False
    LigaFRInternos False
    'Carrega campos
    VALVULA_TIPO = "Retenção Portinhola"
    VALVULA_CLASSE = ""
    VALVULA_EXTREMIDADE = ""
    VALVULA_BITOLA = ""
    VALVULA_MATERIAL = ""
    VALVULA_INTERNOS = ""
    VALVULA_REVESTIMENTO = ""
    LB_Valvula.Caption = "Retenção Portinhola"
    LB_Classe.Caption = ""
    LB_Extremidade.Caption = ""
    LB_Bitola.Caption = ""
    LB_Material.Caption = ""
    LB_Internos.Caption = ""
    LB_Revestimento.Caption = ""
    'Componentes
    C01_C = "CORPO GAVETA"
    C02_C = "TAMPA"
    C03_C = "REDONDO"
    C04_C = "REDONDO"
    C05_C = "PÊNDULO"
    C06_C = "REDONDO"
    C07_C = "PORCA EIXO"
    C08_C = "QUADRADO"
    C09_C = "JUNTA ESPIROTÁLICA"
    C10_C = "PRISIONEIRO CORPO"
    C11_C = "PORCA PRISIONEIRO CORPO"
    C12_C = "PLACA DE IDENTIFICAÇÃO"
    C13_C = "REVESTIMENTO"
    C14_C = "FLANGE"
    C15_C = "SOLDA"
    C16_C = ""
    C17_C = ""
    C18_C = ""
    C19_C = ""
    C20_C = ""
    'Nome do Componente
    C01_N = "Corpo"
    C02_N = "Tampa"
    C03_N = "Disco"
    C04_N = "Anél"
    C05_N = "Pêndulo"
    C06_N = "Eixo"
    C07_N = "Porca Eixo"
    C08_N = "Braço"
    C09_N = "Junta Espirotálica"
    C10_N = "Prisioneiro Corpo"
    C11_N = "Porca Corpo"
    C12_N = "Placa de Identificação"
    C13_N = "Revestimento"
    C14_N = "Flange"
    C15_N = "Solda de Adaptação"
    C16_N = ""
    C17_N = ""
    C18_N = ""
    C19_N = ""
    C20_N = ""
    'Número de Componente
    C01_U = 1
    C02_U = 1
    C03_U = 1
    C04_U = 1
    C05_U = 1
    C06_U = 1
    C07_U = 1
    C08_U = 1
    C09_U = 1
    C10_U = 4
    C11_U = 4
    C12_U = 1
    C13_U = 0
    C14_U = 0
    C15_U = 0
    C16_U = 0
    C17_U = 0
    C18_U = 0
    C19_U = 0
    C20_U = 0
    'Carrega quantidades
    QuantidadeComponentes
End Sub
Private Sub BT_SW_Click()
    BT_NPT_Click
    VALVULA_EXTREMIDADE = "SW"
    LB_Extremidade.Caption = "SW"
End Sub
Private Sub BT_XU_Click()
    VALVULA_REVESTIMENTO = "NA"
    LB_Revestimento.Caption = "Não aplicado"
    InternosComponentes
End Sub
Private Sub BT_XX_Click()
    VALVULA_REVESTIMENTO = "NA"
    LB_Revestimento.Caption = "Não aplicado"
    InternosComponentes
End Sub
Private Sub TXT_Quantidade_Change()
    QuantidadeComponentes
End Sub
Private Sub UserControl_Initialize()
    VALVULA_TIPO = ""
    VALVULA_CLASSE = ""
    VALVULA_EXTREMIDADE = ""
    VALVULA_BITOLA = ""
    VALVULA_MATERIAL = ""
    VALVULA_INTERNOS = ""
    VALVULA_REVESTIMENTO = ""
    TXT_Quantidade.Text = ""
    LB_Valvula.Caption = ""
    LB_Classe.Caption = ""
    LB_Extremidade.Caption = ""
    LB_Bitola.Caption = ""
    LB_Material.Caption = ""
    LB_Internos.Caption = ""
    LB_Revestimento.Caption = ""
    LigaFRExtremidade False
    LigaFRClasse False
    LigaFRBitola False
    LigaFRMaterial False
    LigaFRInternos False
End Sub


'***********************************
'          PROPRIEDADES
'***********************************
Public Property Get Tipo() As String
    Tipo = VALVULA_TIPO
End Property
Public Property Let Tipo(ByVal Valor As String)
    VALVULA_TIPO = Valor
End Property
Public Property Get Classe() As String
    Classe = VALVULA_CLASSE
End Property
Public Property Let Classe(ByVal Valor As String)
    VALVULA_CLASSE = Valor
End Property
Public Property Get Extremidade() As String
    Extremidade = VALVULA_EXTREMIDADE
End Property
Public Property Let Extremidade(ByVal Valor As String)
    VALVULA_EXTREMIDADE = Valor
End Property
Public Property Get Bitola() As String
    Bitola = VALVULA_BITOLA
End Property
Public Property Let Bitola(ByVal Valor As String)
    VALVULA_BITOLA = Valor
End Property
Public Property Get Material() As String
    Material = VALVULA_MATERIAL
End Property
Public Property Let Material(ByVal Valor As String)
    VALVULA_MATERIAL = Valor
End Property
Public Property Get Internos() As String
    Internos = VALVULA_INTERNOS
End Property
Public Property Let Internos(ByVal Valor As String)
    VALVULA_INTERNOS = Valor
End Property
Public Property Get Revestimento() As String
    Revestimento = VALVULA_REVESTIMENTO
End Property
Public Property Let Revestimento(ByVal Valor As String)
    VALVULA_REVESTIMENTO = Valor
End Property




'***********************************
'            FUNCOES
'***********************************
Private Sub QuantidadeComponentes()
    If TXT_Quantidade.Text = "" Then Exit Sub
    On Error Resume Next
    C01_Q = TXT_Quantidade.Text * C01_U
    C02_Q = TXT_Quantidade.Text * C02_U
    C03_Q = TXT_Quantidade.Text * C03_U
    C04_Q = TXT_Quantidade.Text * C04_U
    C05_Q = TXT_Quantidade.Text * C05_U
    C06_Q = TXT_Quantidade.Text * C06_U
    C07_Q = TXT_Quantidade.Text * C07_U
    C08_Q = TXT_Quantidade.Text * C08_U
    C09_Q = TXT_Quantidade.Text * C09_U
    C10_Q = TXT_Quantidade.Text * C10_U
    C11_Q = TXT_Quantidade.Text * C11_U
    C12_Q = TXT_Quantidade.Text * C12_U
    C13_Q = TXT_Quantidade.Text * C13_U
    C14_Q = TXT_Quantidade.Text * C14_U
    C15_Q = TXT_Quantidade.Text * C15_U
    C16_Q = TXT_Quantidade.Text * C16_U
    C17_Q = TXT_Quantidade.Text * C17_U
    C18_Q = TXT_Quantidade.Text * C18_U
    C19_Q = TXT_Quantidade.Text * C19_U
    C20_Q = TXT_Quantidade.Text * C20_U
End Sub
Private Sub BitolaComponentes()
    On Error Resume Next
    If VALVULA_TIPO = "Gaveta" Then
        'Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C01_B = "1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C01_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C01_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C01_B = "1.1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C01_B = "2" & Chr(34)
        End If
        'Castelo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C02_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C02_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C02_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Preme
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C03_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C03_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C03_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Cunha
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C04_B = "1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C04_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C04_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C04_B = "1.1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C04_B = "2" & Chr(34)
        End If
        'Anél
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C05_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C05_B = "7/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C05_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C05_B = "1.5/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C05_B = "1.7/8" & Chr(34)
        End If
        'Haste
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C06_B = "9/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C06_B = "3/4" & Chr(34)
        End If
        'Junta
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C07_B = "41,5 x 33 x 3,2"
        ElseIf VALVULA_BITOLA = "1" Then
            C07_B = "48,5 x 38 x 3,2"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C07_B = "69 x 59 x 3,2"
        End If
        'Gaxeta
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C08_B = "17,5 x 11 x 5"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C08_B = "25,5 x 16 x 5"
        End If
        'Prisioneiro Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C09_B = "3/8" & Chr(34) & " x 45"
        ElseIf VALVULA_BITOLA = "1" Then
            C09_B = "7/16" & Chr(34) & " x 52"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C09_B = "1/2" & Chr(34) & " x 60"
        End If
        'Porca Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C10_B = "3/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C10_B = "7/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C10_B = "1/2" & Chr(34)
        End If
        'Prisioneiro Preme
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C11_B = "5/16" & Chr(34) & " x 50"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C11_B = "3/8" & Chr(34) & " x 60"
        End If
        'Porca Preme
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C12_B = "5/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C12_B = "3/8" & Chr(34)
        End If
        'Bucha
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C13_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C13_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C13_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Porca Volante
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C14_B = "7/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C14_B = "1.1/8" & Chr(34)
        End If
        'Volante
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C15_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C15_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C15_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Placa de Identificação
        C16_B = ""
        'VALVULA_REVESTIMENTO
        C17_B = ""
        'Flange
        C18_B = VALVULA_BITOLA & Chr(34)
        'Solda
        C19_B = ""
        'Nenhum
        C20_B = ""
    ElseIf VALVULA_TIPO = "Globo" Then
        'Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C01_B = "1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C01_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C01_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C01_B = "1.1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C01_B = "2" & Chr(34)
        End If
        'Castelo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C02_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C02_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C02_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Preme
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C03_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C03_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C03_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Contra-Sede
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C04_B = "1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C04_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C04_B = "1" & Chr(34)
        End If
        'Sede
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C05_B = "7/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C05_B = "1.1/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C05_B = "1.3/8" & Chr(34)
        End If
        'Haste
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C06_B = "9/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C06_B = "3/4" & Chr(34)
        End If
        'Junta
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C07_B = "41,5 x 33 x 3,2"
        ElseIf VALVULA_BITOLA = "1" Then
            C07_B = "48,5 x 38 x 3,2"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C07_B = "69 x 59 x 3,2"
        End If
        'Gaxeta
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C08_B = "17,5 x 11 x 5"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C08_B = "25,5 x 16 x 5"
        End If
        'Prisioneiro Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C09_B = "3/8" & Chr(34) & " x 45"
        ElseIf VALVULA_BITOLA = "1" Then
            C09_B = "7/16" & Chr(34) & " x 52"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C09_B = "1/2" & Chr(34) & " x 60"
        End If
        'Porca Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C10_B = "3/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C10_B = "7/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C10_B = "1/2" & Chr(34)
        End If
        'Prisioneiro Preme
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C11_B = "5/16" & Chr(34) & " x 50"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C11_B = "3/8" & Chr(34) & " x 60"
        End If
        'Porca Preme
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C12_B = "5/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C12_B = "3/8" & Chr(34)
        End If
        'Bucha
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C13_B = "1/2" & Chr(34) & " e 3/4" & Chr(34) & " e 1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C13_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Porca Volante
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C14_B = "1/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C14_B = "5/16" & Chr(34)
        End If
        'Arruela Volante
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Or VALVULA_BITOLA = "1" Then
            C15_B = "1/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C15_B = "5/16" & Chr(34)
        End If
        'Volante
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C16_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C16_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C16_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Placa de Identificação
        C17_B = ""
        'VALVULA_REVESTIMENTO
        C18_B = ""
        'Flange
        C19_B = VALVULA_BITOLA & Chr(34)
        'Solda
        C20_B = ""
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        'Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C01_B = "1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C01_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C01_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C01_B = "1.1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C01_B = "2" & Chr(34)
        End If
        'Tampa
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C02_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C02_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C02_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Pistão
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C03_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C03_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C03_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Sede
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C04_B = "7/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C04_B = "1.1/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C04_B = "1.3/8" & Chr(34)
        End If
        'Junta Espirotálica
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C05_B = "41,5 x 33 x 3,2"
        ElseIf VALVULA_BITOLA = "1" Then
            C05_B = "48,5 x 38 x 3,2"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C05_B = "69 x 59 x 3,2"
        End If
        'Prisioneiro Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C06_B = "3/8" & Chr(34) & " x 45"
        ElseIf VALVULA_BITOLA = "1" Then
            C06_B = "7/16" & Chr(34) & " x 52"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C06_B = "1/2" & Chr(34) & " x 60"
        End If
        'Porca Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C07_B = "3/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C07_B = "7/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C07_B = "1/2" & Chr(34)
        End If
        'Mola
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C08_B = "22 x 32 x 1,2"
        ElseIf VALVULA_BITOLA = "1" Then
            C08_B = "30 x 33 x 1,5"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C08_B = "36 x 46 x 1,7"
        End If
        'Placa de Identificação
        C09_B = ""
        'VALVULA_REVESTIMENTO
        C10_B = ""
        'Flange
        C11_B = VALVULA_BITOLA & Chr(34)
        'Solda
        C12_B = ""
        'Nenhum
        C13_N = ""
        C14_N = ""
        C15_N = ""
        C16_N = ""
        C17_N = ""
        C18_N = ""
        C19_N = ""
        C20_N = ""
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        'Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C01_B = "1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C01_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C01_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C01_B = "1.1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C01_B = "2" & Chr(34)
        End If
        'Tampa
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C02_B = "1/2" & Chr(34) & " e 3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C02_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C02_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Disco
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C03_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C03_B = "7/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C03_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C03_B = "1.5/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C03_B = "1.7/8" & Chr(34)
        End If
        'Anél
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C04_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C04_B = "7/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C04_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Then
            C04_B = "1.5/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "2" Then
            C04_B = "1.7/8" & Chr(34)
        End If
        'Pêndulo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Then
            C05_B = "1/2" & Chr(34)
        ElseIf VALVULA_BITOLA = "3/4" Then
            C05_B = "3/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C05_B = "1" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C05_B = "1.1/2" & Chr(34) & " e 2" & Chr(34)
        End If
        'Eixo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C06_B = "1/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C06_B = "3/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C06_B = "5/16" & Chr(34)
        End If
        'Porca Eixo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C07_B = "3/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C07_B = "1/4" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C07_B = "5/16" & Chr(34)
        End If
        'Braço
        If VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C08_B = "1" & Chr(34)
        Else
            C08_B = ""
        End If
        'Junta Espirotálica
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C09_B = "41,5 x 33 x 3,2"
        ElseIf VALVULA_BITOLA = "1" Then
            C09_B = "48,5 x 38 x 3,2"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C09_B = "69 x 59 x 3,2"
        End If
        'Prisioneiro Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C10_B = "3/8" & Chr(34) & " x 45"
        ElseIf VALVULA_BITOLA = "1" Then
            C10_B = "7/16" & Chr(34) & " x 52"
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C10_B = "1/2" & Chr(34) & " x 60"
        End If
        'Porca Corpo
        If VALVULA_BITOLA = "3/8" Or VALVULA_BITOLA = "1/2" Or VALVULA_BITOLA = "3/4" Then
            C11_B = "3/8" & Chr(34)
        ElseIf VALVULA_BITOLA = "1" Then
            C11_B = "7/16" & Chr(34)
        ElseIf VALVULA_BITOLA = "1.1/4" Or VALVULA_BITOLA = "1.1/2" Or VALVULA_BITOLA = "2" Then
            C11_B = "1/2" & Chr(34)
        End If
        'Placa de Identificação
        C12_B = ""
        'VALVULA_REVESTIMENTO
        C13_B = ""
        'Flange
        C14_B = VALVULA_BITOLA & Chr(34)
        'Solda
        C15_B = ""
        'Nenhum
        C16_N = ""
        C17_N = ""
        C18_N = ""
        C19_N = ""
        C20_N = ""
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub MaterialComponentes()
    On Error Resume Next
    If VALVULA_TIPO = "Gaveta" Then
        If VALVULA_MATERIAL = "A105" Then
            C01_M = "A105"
            C02_M = "A105"
            C03_M = "A105"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B7"
            C10_M = "A194 2H"
            C11_M = "A193 B7"
            C12_M = "A194 2H"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = ""
            C18_M = "A105"
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F304" Then
            C01_M = "A182 F304"
            C02_M = "A182 F304"
            C03_M = "A182 F304"
            C04_M = "A351 CF8"
            C05_M = "A276 T304"
            C06_M = "A276 T304"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = ""
            C18_M = "A182 F304"
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F316" Then
            C01_M = "A182 F316"
            C02_M = "A182 F316"
            C03_M = "A182 F316"
            C04_M = "A351 CF8M"
            C05_M = "A276 T316"
            C06_M = "A276 T316"
            C07_M = "316 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = ""
            C18_M = "A182 F316"
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F304L" Then
            C01_M = "A182 F304L"
            C02_M = "A182 F304L"
            C03_M = "A182 F304L"
            C04_M = "A351 CF3"
            C05_M = "A276 T304L"
            C06_M = "A276 T304L"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = ""
            C18_M = "A182 F304L"
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F316L" Then
            C01_M = "A182 F316L"
            C02_M = "A182 F316L"
            C03_M = "A182 F316L"
            C04_M = "A351 CF3M"
            C05_M = "A276 T316L"
            C06_M = "A276 T316L"
            C07_M = "316 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = ""
            C18_M = "A182 F316L"
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F5" Then
            C01_M = "A182 F5"
            C02_M = "A182 F5"
            C03_M = "A182 F5"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B16"
            C10_M = "A194 8M"
            C11_M = "A193 B16"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = "AWS A5.13 E CoCrA"
            C18_M = "A182 F5"
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F11" Then
            C01_M = "A182 F11"
            C02_M = "A182 F11"
            C03_M = "A182 F11"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B16"
            C10_M = "A194 8M"
            C11_M = "A193 B16"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = "AWS A5.13 E CoCrA"
            C18_M = "A182 F11"
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A350 LF2" Then
            C01_M = "A350 LF2"
            C02_M = "A350 LF2"
            C03_M = "A350 LF2"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "A216 WCB"
            C16_M = "Alumínio"
            C17_M = "AWS A5.13 E CoCrA"
            C18_M = "A350 LF2"
            C19_M = ""
            C20_M = ""
        End If
    ElseIf VALVULA_TIPO = "Globo" Then
        If VALVULA_MATERIAL = "A105" Then
            C01_M = "A105"
            C02_M = "A105"
            C03_M = "A105"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B7"
            C10_M = "A194 2H"
            C11_M = "A193 B7"
            C12_M = "A194 2H"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A105"
            C20_M = "Solda de Adaptação"
        ElseIf VALVULA_MATERIAL = "A182 F304" Then
            C01_M = "A182 F304"
            C02_M = "A182 F304"
            C03_M = "A182 F304"
            C04_M = "A351 CF8"
            C05_M = "A276 T304"
            C06_M = "A276 T304"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A182 F304"
            C20_M = "Solda de Adaptação"
        ElseIf VALVULA_MATERIAL = "A182 F316" Then
            C01_M = "A182 F316"
            C02_M = "A182 F316"
            C03_M = "A182 F316"
            C04_M = "A351 CF8M"
            C05_M = "A276 T316"
            C06_M = "A276 T316"
            C07_M = "316 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A182 F316"
            C20_M = "Solda de Adaptação"
        ElseIf VALVULA_MATERIAL = "A182 F304L" Then
            C01_M = "A182 F304L"
            C02_M = "A182 F304L"
            C03_M = "A182 F304L"
            C04_M = "A351 CF3"
            C05_M = "A276 T304L"
            C06_M = "A276 T304L"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A182 F304L"
            C20_M = "Solda de Adaptação"
        ElseIf VALVULA_MATERIAL = "A182 F316L" Then
            C01_M = "A182 F316L"
            C02_M = "A182 F316L"
            C03_M = "A182 F316L"
            C04_M = "A351 CF3M"
            C05_M = "A276 T316L"
            C06_M = "A276 T316L"
            C07_M = "316 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A182 F316L"
            C20_M = "Solda de Adaptação"
        ElseIf VALVULA_MATERIAL = "A182 F5" Then
            C01_M = "A182 F5"
            C02_M = "A182 F5"
            C03_M = "A182 F5"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B16"
            C10_M = "A194 8M"
            C11_M = "A193 B16"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A182 F316"
            C20_M = "Solda de Adaptação"
        ElseIf VALVULA_MATERIAL = "A182 F11" Then
            C01_M = "A182 F11"
            C02_M = "A182 F11"
            C03_M = "A182 F11"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B16"
            C10_M = "A194 8M"
            C11_M = "A193 B16"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A182 F11"
            C20_M = "Solda de Adaptação"
        ElseIf VALVULA_MATERIAL = "A350 LF2" Then
            C01_M = "A350 LF2"
            C02_M = "A350 LF2"
            C03_M = "A350 LF2"
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
            C07_M = "304 / Grafite"
            C08_M = "Grafite"
            C09_M = "A193 B8M-CL2"
            C10_M = "A194 8M"
            C11_M = "A193 B8M-CL2"
            C12_M = "A194 8M"
            C13_M = "ASTM B16"
            C14_M = "SAE 1010/1020"
            C15_M = "SAE 1010/1020"
            C16_M = "A216 WCB"
            C17_M = "Alumínio"
            C18_M = "AWS A5.13 E CoCrA"
            C19_M = "A182 F11"
            C20_M = "Solda de Adaptação"
        End If
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        If VALVULA_MATERIAL = "A105" Then
            C01_M = "A105"
            C02_M = "A105"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "304 / Grafite"
            C06_M = "A193 B7"
            C07_M = "A194 2H"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A105"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F304" Then
            C01_M = "A182 F304"
            C02_M = "A182 F304"
            C03_M = "A276 T304"
            C04_M = "A276 T304"
            C05_M = "304 / Grafite"
            C06_M = "A193 B8M-CL2"
            C07_M = "A194 8M"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A182 F304"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F316" Then
            C01_M = "A182 F316"
            C02_M = "A182 F316"
            C03_M = "A276 T316"
            C04_M = "A276 T316"
            C05_M = "316 / Grafite"
            C06_M = "A193 B8M-CL2"
            C07_M = "A194 8M"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A182 F316"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F304L" Then
            C01_M = "A182 F304L"
            C02_M = "A182 F304L"
            C03_M = "A276 T304L"
            C04_M = "A276 T304L"
            C05_M = "304 / Grafite"
            C06_M = "A193 B8M-CL2"
            C07_M = "A194 8M"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A182 F304L"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F316L" Then
            C01_M = "A182 F316L"
            C02_M = "A182 F316L"
            C03_M = "A276 T316L"
            C04_M = "A276 T316L"
            C05_M = "316 / Grafite"
            C06_M = "A193 B8M-CL2"
            C07_M = "A194 8M"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A182 F316L"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F5" Then
            C01_M = "A182 F5"
            C02_M = "A182 F5"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "304 / Grafite"
            C06_M = "A193 B16"
            C07_M = "A194 8M"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A182 F5"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F11" Then
            C01_M = "A182 F11"
            C02_M = "A182 F11"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "304 / Grafite"
            C06_M = "A193 B16"
            C07_M = "A194 8M"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A182 F11"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A350 LF2" Then
            C01_M = "A350 LF2"
            C02_M = "A350 LF2"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "304 / Grafite"
            C06_M = "A193 B8M-CL2"
            C07_M = "A194 8M"
            C08_M = "A276 T302"
            C09_M = "Alumínio"
            C10_M = "AWS A5.13 E CoCrA"
            C11_M = "A350 LF2"
            C12_M = "Solda de Adaptação"
            C13_M = ""
            C14_M = ""
            C15_M = ""
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        End If
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        If VALVULA_MATERIAL = "A105" Then
            C01_M = "A105"
            C02_M = "A105"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "A216 WCB"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
            C10_M = "A193 B7"
            C11_M = "A194 2H"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A105"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F304" Then
            C01_M = "A182 F304"
            C02_M = "A182 F304"
            C03_M = "A276 T304"
            C04_M = "A276 T304"
            C05_M = "A351 CF8"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T304"
            C09_M = "304 / Grafite"
            C10_M = "A193 B8M-CL2"
            C11_M = "A194 8M"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A182 F304"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F316" Then
            C01_M = "A182 F316"
            C02_M = "A182 F316"
            C03_M = "A276 T316"
            C04_M = "A276 T316"
            C05_M = "A351 CF8M"
            C06_M = "A276 T316"
            C07_M = "A276 T316"
            C08_M = "A276 T316"
            C09_M = "316 / Grafite"
            C10_M = "A193 B8M-CL2"
            C11_M = "A194 8M"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A182 F316"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F304L" Then
            C01_M = "A182 F304L"
            C02_M = "A182 F304L"
            C03_M = "A276 T304L"
            C04_M = "A276 T304L"
            C05_M = "A351 CF3"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T304"
            C09_M = "304 / Grafite"
            C10_M = "A193 B8M-CL2"
            C11_M = "A194 8M"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A182 F304L"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F316L" Then
            C01_M = "A182 F316L"
            C02_M = "A182 F316L"
            C03_M = "A276 T316L"
            C04_M = "A276 T316L"
            C05_M = "A351 CF3M"
            C06_M = "A276 T316"
            C07_M = "A276 T316"
            C08_M = "A276 T316"
            C09_M = "316 / Grafite"
            C10_M = "A193 B8M-CL2"
            C11_M = "A194 8M"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A182 F316L"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F5" Then
            C01_M = "A182 F5"
            C02_M = "A182 F5"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "A216 WCB"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
            C10_M = "A193 B16"
            C11_M = "A194 8M"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A182 F5"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A182 F11" Then
            C01_M = "A182 F11"
            C02_M = "A182 F11"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "A216 WCB"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
            C10_M = "A193 B16"
            C11_M = "A194 8M"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A182 F11"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        ElseIf VALVULA_MATERIAL = "A350 LF2" Then
            C01_M = "A350 LF2"
            C02_M = "A350 LF2"
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "A216 WCB"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
            C10_M = "A193 B8M-CL2"
            C11_M = "A194 8M"
            C12_M = "Alumínio"
            C13_M = "AWS A5.13 E CoCrA"
            C14_M = "A350 LF2"
            C15_M = "Solda de Adaptação"
            C16_M = ""
            C17_M = ""
            C18_M = ""
            C19_M = ""
            C20_M = ""
        End If
    ElseIf VALVULA_TIPO = "Ponta de Agulha" Then
    
    End If
End Sub
Private Sub InternosComponentes()
    On Error Resume Next
    If VALVULA_TIPO = "Gaveta" Or VALVULA_TIPO = "Globo" Then
        If VALVULA_INTERNOS = "410" Then
            C04_M = "A217 CA15"
            C05_M = "A276 T410"
            C06_M = "A276 T410"
        ElseIf VALVULA_INTERNOS = "304" Then
            C04_M = "A351 CF8"
            C05_M = "A276 T304"
            C06_M = "A276 T304"
        ElseIf VALVULA_INTERNOS = "316" Then
            C04_M = "A351 CF8M"
            C05_M = "A276 T316"
            C06_M = "A276 T316"
        ElseIf VALVULA_INTERNOS = "304L" Then
            C04_M = "A351 CF3"
            C05_M = "A276 T304L"
            C06_M = "A276 T304L"
        ElseIf VALVULA_INTERNOS = "316L" Then
            C04_M = "A351 CF3M"
            C05_M = "A276 T316L"
            C06_M = "A276 T316L"
        ElseIf VALVULA_INTERNOS = "Monel" Then
            C04_M = "Monel"
            C05_M = "Monel"
            C06_M = "Monel"
        ElseIf VALVULA_INTERNOS = "Inconel" Then
            C04_M = "Inconel"
            C05_M = "Inconel"
            C06_M = "Inconel"
        ElseIf VALVULA_INTERNOS = "Alloy20" Then
            C04_M = "Alloy 20"
            C05_M = "Alloy 20"
            C06_M = "Alloy 20"
        End If
    ElseIf VALVULA_TIPO = "Retenção Pistão" Then
        If VALVULA_INTERNOS = "410" Then
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "304" Then
            C03_M = "A276 T304"
            C04_M = "A276 T304"
            C05_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "316" Then
            C03_M = "A276 T316"
            C04_M = "A276 T316"
            C05_M = "316 / Grafite"
        ElseIf VALVULA_INTERNOS = "304L" Then
            C03_M = "A276 T304L"
            C04_M = "A276 T304L"
            C05_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "316L" Then
            C03_M = "A276 T316L"
            C04_M = "A276 T316L"
            C05_M = "316 / Grafite"
        ElseIf VALVULA_INTERNOS = "Monel" Then
            C03_M = "Monel"
            C04_M = "Monel"
            C05_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "Inconel" Then
            C03_M = "Inconel"
            C04_M = "Inconel"
            C05_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "Alloy20" Then
            C03_M = "Alloy 20"
            C04_M = "Alloy 20"
            C05_M = "304 / Grafite"
        End If
    ElseIf VALVULA_TIPO = "Retenção Portinhola" Then
        If VALVULA_INTERNOS = "410" Then
            C03_M = "A276 T410"
            C04_M = "A276 T410"
            C05_M = "A217 CA15"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "304" Then
            C03_M = "A276 T304"
            C04_M = "A276 T304"
            C05_M = "A351 CF8"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T304"
            C09_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "316" Then
            C03_M = "A276 T316"
            C04_M = "A276 T316"
            C05_M = "A351 CF8M"
            C06_M = "A276 T316"
            C07_M = "A276 T316"
            C08_M = "A276 T316"
            C09_M = "316 / Grafite"
        ElseIf VALVULA_INTERNOS = "304L" Then
            C03_M = "A276 T304L"
            C04_M = "A276 T304L"
            C05_M = "A351 CF3"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T304"
            C09_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "316L" Then
            C03_M = "A276 T316L"
            C04_M = "A276 T316L"
            C05_M = "A351 CF3M"
            C06_M = "A276 T316"
            C07_M = "A276 T316"
            C08_M = "A276 T316"
            C09_M = "316 / Grafite"
        ElseIf VALVULA_INTERNOS = "Monel" Then
            C03_M = "Monel"
            C04_M = "Monel"
            C05_M = "A217 CA15"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "Inconel" Then
            C03_M = "Inconel"
            C04_M = "Inconel"
            C05_M = "A217 CA15"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
        ElseIf VALVULA_INTERNOS = "Alloy20" Then
            C03_M = "Alloy 20"
            C04_M = "Alloy 20"
            C05_M = "A217 CA15"
            C06_M = "A276 T304"
            C07_M = "A276 T304"
            C08_M = "A276 T410"
            C09_M = "304 / Grafite"
        End If
    End If
End Sub
Private Sub LigaFRExtremidade(Valor As Boolean)
    FR_Extremidade.Enabled = Valor
    BT_NPT.Enabled = Valor
    BT_SW.Enabled = Valor
    BT_Flange.Enabled = Valor
    BT_BSP.Enabled = Valor
    BT_OutraExtremidade.Enabled = Valor
End Sub
Private Sub LigaFRClasse(Valor As Boolean)
    FR_Classe.Enabled = Valor
    BT_150.Enabled = Valor
    BT_300.Enabled = Valor
    BT_600.Enabled = Valor
    BT_800.Enabled = Valor
    BT_900.Enabled = Valor
    BT_1500.Enabled = Valor
    BT_3000.Enabled = Valor
    BT_OutraClasse.Enabled = Valor
End Sub
Private Sub LigaFRBitola(Valor As Boolean)
    FR_Bitola.Enabled = Valor
    BT_3_8.Enabled = Valor
    BT_1_2.Enabled = Valor
    BT_3_4.Enabled = Valor
    BT_1.Enabled = Valor
    BT_1_1_4.Enabled = Valor
    BT_1_1_2.Enabled = Valor
    BT_2.Enabled = Valor
    BT_OutraBitola.Enabled = Valor
End Sub
Private Sub LigaFRMaterial(Valor As Boolean)
    FR_Material.Enabled = Valor
    BT_A105.Enabled = Valor
    BT_F304.Enabled = Valor
    BT_F316.Enabled = Valor
    BT_F304L.Enabled = Valor
    BT_F316L.Enabled = Valor
    BT_F5.Enabled = Valor
    BT_F11.Enabled = Valor
    BT_LF2.Enabled = Valor
    BT_OutroMaterial.Enabled = Valor
End Sub
Private Sub LigaFRInternos(Valor As Boolean)
    FR_Internos.Enabled = Valor
    BT_410.Enabled = Valor
    BT_304.Enabled = Valor
    BT_316.Enabled = Valor
    BT_304L.Enabled = Valor
    BT_316L.Enabled = Valor
    BT_Monel.Enabled = Valor
    BT_Inconel.Enabled = Valor
    BT_Alloy20.Enabled = Valor
    BT_OutroInterno.Enabled = Valor
    FR_Revestimento.Enabled = Valor
    BT_XX.Enabled = Valor
    BT_XU.Enabled = Valor
    BT_UU.Enabled = Valor
    BT_OutroRevestimento.Enabled = Valor
End Sub
