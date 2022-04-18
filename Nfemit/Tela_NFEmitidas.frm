VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Tela_NFEmitidas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Notas Fiscais Emitidas"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7650
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog DIMP 
      Left            =   0
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab ST 
      Height          =   3375
      Left            =   2040
      TabIndex        =   7
      ToolTipText     =   "Dados sobre a nota fiscal"
      Top             =   120
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Remetente"
      TabPicture(0)   =   "Tela_NFEmitidas.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FR_Nota"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ítens"
      TabPicture(1)   =   "Tela_NFEmitidas.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FG_1"
      Tab(1).Control(1)=   "FG_2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Transportador"
      TabPicture(2)   =   "Tela_NFEmitidas.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Valores"
      TabPicture(3)   =   "Tela_NFEmitidas.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin MSFlexGridLib.MSFlexGrid FG_2 
         Height          =   1695
         Left            =   -72600
         TabIndex        =   104
         TabStop         =   0   'False
         ToolTipText     =   "Dados que serão impressos na nota fiscal."
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   2990
         _Version        =   393216
         Rows            =   21
         Cols            =   6
         AllowBigSelection=   0   'False
         Enabled         =   -1  'True
         TextStyleFixed  =   4
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin MSFlexGridLib.MSFlexGrid FG_1 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   103
         Top             =   480
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   21
         Cols            =   20
         AllowBigSelection=   0   'False
         TextStyleFixed  =   4
         FocusRect       =   0
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Frame Frame2 
         Caption         =   "Base de Cálculo ICMS:"
         Height          =   2895
         Left            =   -74880
         TabIndex        =   69
         Top             =   360
         Width           =   5295
         Begin VB.Label LB_IP 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   106
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total I.P.I.:"
            Height          =   195
            Index           =   42
            Left            =   3960
            TabIndex        =   105
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "/D:"
            Height          =   195
            Index           =   41
            Left            =   2640
            TabIndex        =   102
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label LB_VD 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            Left            =   4200
            TabIndex        =   101
            Top             =   2520
            Width           =   705
         End
         Begin VB.Label LB_CD 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
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
            Left            =   3000
            TabIndex        =   100
            Top             =   2520
            Width           =   1035
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "/C:"
            Height          =   195
            Index           =   40
            Left            =   2640
            TabIndex        =   99
            Top             =   2280
            Width           =   225
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Index           =   39
            Left            =   4200
            TabIndex        =   98
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label LB_VAC 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            Left            =   4200
            TabIndex        =   97
            Top             =   2280
            Width           =   705
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            Height          =   195
            Index           =   36
            Left            =   3000
            TabIndex        =   96
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label LB_CC 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
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
            Left            =   3000
            TabIndex        =   95
            Top             =   2280
            Width           =   1035
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "/B:"
            Height          =   195
            Index           =   35
            Left            =   120
            TabIndex        =   94
            Top             =   2520
            Width           =   225
         End
         Begin VB.Label LB_VB 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            Left            =   1680
            TabIndex        =   93
            Top             =   2520
            Width           =   705
         End
         Begin VB.Label LB_CB 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
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
            Left            =   480
            TabIndex        =   92
            Top             =   2520
            Width           =   1035
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "/A:"
            Height          =   195
            Index           =   34
            Left            =   120
            TabIndex        =   91
            Top             =   2280
            Width           =   225
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            Height          =   195
            Index           =   31
            Left            =   480
            TabIndex        =   88
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Index           =   33
            Left            =   1680
            TabIndex        =   90
            Top             =   2040
            Width           =   405
         End
         Begin VB.Label LB_VA 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            Left            =   1680
            TabIndex        =   89
            Top             =   2280
            Width           =   705
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   10
            X1              =   15
            X2              =   5265
            Y1              =   2145
            Y2              =   2145
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   10
            X1              =   15
            X2              =   5265
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Label LB_CA 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
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
            Left            =   480
            TabIndex        =   87
            Top             =   2280
            Width           =   1035
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total dos Produtos:"
            Height          =   195
            Index           =   30
            Left            =   90
            TabIndex        =   84
            Top             =   1440
            Width           =   1785
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor Total da Nota Fiscal:"
            Height          =   195
            Index           =   32
            Left            =   2880
            TabIndex        =   86
            Top             =   1440
            Width           =   1875
         End
         Begin VB.Label LB_TO 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   85
            Top             =   1680
            Width           =   705
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   9
            X1              =   10
            X2              =   5260
            Y1              =   1545
            Y2              =   1545
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   9
            X1              =   10
            X2              =   5260
            Y1              =   1530
            Y2              =   1530
         End
         Begin VB.Label LB_PR 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   83
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor ICMS:"
            Height          =   195
            Index           =   26
            Left            =   2880
            TabIndex        =   82
            Top             =   0
            Width           =   840
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Outras Despesas:"
            Height          =   195
            Index           =   29
            Left            =   2640
            TabIndex        =   81
            Top             =   960
            Width           =   1260
         End
         Begin VB.Label LB_OD 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            Left            =   2640
            TabIndex        =   80
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Seguro:"
            Height          =   195
            Index           =   28
            Left            =   1320
            TabIndex        =   79
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor do Frete:"
            Height          =   195
            Index           =   27
            Left            =   105
            TabIndex        =   78
            Top             =   960
            Width           =   1035
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   7
            X1              =   10
            X2              =   5270
            Y1              =   1065
            Y2              =   1065
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   7
            X1              =   10
            X2              =   5270
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Label LB_SG 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            Left            =   1320
            TabIndex        =   77
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label LB_FT 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   76
            Top             =   1200
            Width           =   705
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor ICMS Subst.:"
            Height          =   195
            Index           =   38
            Left            =   2880
            TabIndex        =   74
            Top             =   510
            Width           =   1335
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Base de Cálculo ICMS Subst.:"
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   73
            Top             =   510
            Width           =   2130
         End
         Begin VB.Label LB_VC 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   75
            Top             =   240
            Width           =   705
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   11
            X1              =   10
            X2              =   5280
            Y1              =   615
            Y2              =   615
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   11
            X1              =   10
            X2              =   5280
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label LB_VS 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   72
            Top             =   720
            Width           =   705
         End
         Begin VB.Label LB_BS 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   71
            Top             =   720
            Width           =   705
         End
         Begin VB.Label LB_BI 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            TabIndex        =   70
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Transportadora:"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   5295
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Peso Líquido:"
            Height          =   195
            Index           =   25
            Left            =   1800
            TabIndex        =   68
            Top             =   2040
            Width           =   990
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Peso Bruto:"
            Height          =   195
            Index           =   19
            Left            =   105
            TabIndex        =   67
            Top             =   2040
            Width           =   825
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   5
            X1              =   15
            X2              =   5275
            Y1              =   2130
            Y2              =   2130
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   5
            X1              =   15
            X2              =   5275
            Y1              =   2145
            Y2              =   2145
         End
         Begin VB.Label LB_PQ 
            AutoSize        =   -1  'True
            Caption         =   "00,00"
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
            Left            =   1800
            TabIndex        =   66
            Top             =   2250
            Width           =   495
         End
         Begin VB.Label LB_PB 
            AutoSize        =   -1  'True
            Caption         =   "00,00"
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
            TabIndex        =   65
            Top             =   2250
            Width           =   495
         End
         Begin VB.Label LB_EV 
            AutoSize        =   -1  'True
            Caption         =   "SP"
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
            Left            =   3360
            TabIndex        =   64
            Top             =   840
            Width           =   255
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Estado Placa Veículo:"
            Height          =   195
            Index           =   10
            Left            =   3360
            TabIndex        =   63
            Top             =   630
            Width           =   1590
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Número:"
            Height          =   195
            Index           =   24
            Left            =   4200
            TabIndex        =   39
            Top             =   1350
            Width           =   600
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Marca:"
            Height          =   195
            Index           =   23
            Left            =   3000
            TabIndex        =   37
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Espécie:"
            Height          =   195
            Index           =   21
            Left            =   1320
            TabIndex        =   34
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   33
            Top             =   1350
            Width           =   870
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Placa Veículo:"
            Height          =   195
            Index           =   22
            Left            =   1920
            TabIndex        =   35
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Frete por conta de:"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   32
            Top             =   630
            Width           =   1350
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   6
            X1              =   15
            X2              =   5285
            Y1              =   1440
            Y2              =   1440
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   6
            X1              =   15
            X2              =   5285
            Y1              =   1455
            Y2              =   1455
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   4
            X1              =   15
            X2              =   5285
            Y1              =   735
            Y2              =   735
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   4
            X1              =   15
            X2              =   5285
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Label LB_NU 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            Left            =   4200
            TabIndex        =   40
            Top             =   1560
            Width           =   645
         End
         Begin VB.Label LB_MA 
            AutoSize        =   -1  'True
            Caption         =   "Conesteel"
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
            Left            =   3000
            TabIndex        =   38
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label LB_PL 
            AutoSize        =   -1  'True
            Caption         =   "AAA 0000"
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
            Left            =   1920
            TabIndex        =   36
            Top             =   840
            Width           =   855
         End
         Begin VB.Label LB_EP 
            AutoSize        =   -1  'True
            Caption         =   "Pacote"
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
            Left            =   1320
            TabIndex        =   31
            Top             =   1560
            Width           =   615
         End
         Begin VB.Label LB_QT 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            TabIndex        =   30
            Top             =   1560
            Width           =   645
         End
         Begin VB.Label LB_FR 
            AutoSize        =   -1  'True
            Caption         =   "Destinatário"
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
            TabIndex        =   29
            Top             =   840
            Width           =   1035
         End
         Begin VB.Label LB_TR 
            AutoSize        =   -1  'True
            Caption         =   "Conesteel Conexões de Aço Ltda."
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
            TabIndex        =   28
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.Frame FR_Nota 
         Caption         =   "Razão Social:"
         Height          =   2895
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   5295
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Setor:"
            Height          =   195
            Index           =   9
            Left            =   3705
            TabIndex        =   58
            Top             =   2310
            Width           =   420
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor Externo:"
            Height          =   195
            Index           =   17
            Left            =   1785
            TabIndex        =   59
            Top             =   2310
            Width           =   1320
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor Interno:"
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   60
            Top             =   2310
            Width           =   1275
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   8
            X1              =   10
            X2              =   5260
            Y1              =   2415
            Y2              =   2415
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   8
            X1              =   10
            X2              =   5260
            Y1              =   2405
            Y2              =   2405
         End
         Begin VB.Label LB_VI 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            Left            =   90
            TabIndex        =   62
            Top             =   2520
            Width           =   645
         End
         Begin VB.Label LB_VX 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            Left            =   1785
            TabIndex        =   61
            Top             =   2520
            Width           =   645
         End
         Begin VB.Label LB_SE 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            Left            =   3705
            TabIndex        =   57
            Top             =   2520
            Width           =   645
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Operação:"
            Height          =   195
            Index           =   12
            Left            =   3720
            TabIndex        =   52
            Top             =   1860
            Width           =   750
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Nº Seu Pedido:"
            Height          =   195
            Index           =   15
            Left            =   1800
            TabIndex        =   53
            Top             =   1860
            Width           =   1095
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Nº Pedido:"
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   54
            Top             =   1860
            Width           =   765
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   3
            X1              =   10
            X2              =   5270
            Y1              =   1965
            Y2              =   1965
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   3
            X1              =   10
            X2              =   5270
            Y1              =   1955
            Y2              =   1955
         End
         Begin VB.Label LB_NP 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            TabIndex        =   56
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LB_SP 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            Left            =   1800
            TabIndex        =   55
            Top             =   2070
            Width           =   645
         End
         Begin VB.Label LB_OP 
            AutoSize        =   -1  'True
            Caption         =   "01234"
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
            Left            =   3720
            TabIndex        =   51
            Top             =   2070
            Width           =   540
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Natureza da Operação:"
            Height          =   195
            Index           =   6
            Left            =   2400
            TabIndex        =   45
            Top             =   960
            Width           =   1665
         End
         Begin VB.Label LB_Tipo 
            AutoSize        =   -1  'True
            Caption         =   "Saída"
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
            TabIndex        =   50
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label LB_CFOP 
            AutoSize        =   -1  'True
            Caption         =   "5.11"
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
            Left            =   1200
            TabIndex        =   49
            Top             =   1200
            Width           =   390
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   48
            Top             =   960
            Width           =   360
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "CFOP:"
            Height          =   195
            Index           =   7
            Left            =   1200
            TabIndex        =   47
            Top             =   960
            Width           =   465
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   2
            X1              =   10
            X2              =   5270
            Y1              =   1065
            Y2              =   1065
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   2
            X1              =   10
            X2              =   5270
            Y1              =   1050
            Y2              =   1050
         End
         Begin VB.Label LB_NO 
            AutoSize        =   -1  'True
            Caption         =   "Venda para Comercialização"
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
            Left            =   2400
            TabIndex        =   46
            Top             =   1200
            Width           =   2430
         End
         Begin VB.Label LB_DS 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
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
            TabIndex        =   44
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label LB_HS 
            AutoSize        =   -1  'True
            Caption         =   "00:00:00"
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
            Left            =   4080
            TabIndex        =   43
            Top             =   720
            Width           =   765
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Data Saída:"
            Height          =   195
            Index           =   5
            Left            =   2760
            TabIndex        =   42
            Top             =   510
            Width           =   870
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Hora Saída:"
            Height          =   195
            Index           =   0
            Left            =   4080
            TabIndex        =   41
            Top             =   510
            Width           =   870
         End
         Begin VB.Label LB_RS 
            AutoSize        =   -1  'True
            Caption         =   "Conesteel Conexões de Aço Ltda."
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
            TabIndex        =   26
            Top             =   240
            Width           =   4935
         End
         Begin VB.Label LB_DE 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
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
            TabIndex        =   25
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label LB_HE 
            AutoSize        =   -1  'True
            Caption         =   "00:00:00"
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
            Left            =   1440
            TabIndex        =   24
            Top             =   720
            Width           =   765
         End
         Begin VB.Label LB_NF 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            TabIndex        =   23
            Top             =   1620
            Width           =   645
         End
         Begin VB.Label LB_V 
            AutoSize        =   -1  'True
            Caption         =   "1000,00"
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
            Left            =   3720
            TabIndex        =   22
            Top             =   1620
            Width           =   705
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Data Emissão:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   510
            Width           =   1020
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Hora Emissão:"
            Height          =   195
            Index           =   2
            Left            =   1440
            TabIndex        =   20
            Top             =   510
            Width           =   1020
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Nota Fiscal:"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   19
            Top             =   1410
            Width           =   840
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Valor:"
            Height          =   195
            Index           =   4
            Left            =   3720
            TabIndex        =   18
            Top             =   1410
            Width           =   405
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Nº Duplicata:"
            Height          =   195
            Index           =   13
            Left            =   1200
            TabIndex        =   17
            Top             =   1410
            Width           =   945
         End
         Begin VB.Label LB_DP 
            AutoSize        =   -1  'True
            Caption         =   "000000"
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
            Left            =   1200
            TabIndex        =   16
            Top             =   1620
            Width           =   645
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento:"
            Height          =   195
            Index           =   14
            Left            =   2400
            TabIndex        =   15
            Top             =   1410
            Width           =   885
         End
         Begin VB.Label LB_VE 
            AutoSize        =   -1  'True
            Caption         =   "00/00/0000"
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
            Left            =   2400
            TabIndex        =   14
            Top             =   1620
            Width           =   1035
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   0
            X1              =   10
            X2              =   5280
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   0
            X1              =   10
            X2              =   5280
            Y1              =   615
            Y2              =   615
         End
         Begin VB.Line L2 
            BorderColor     =   &H80000005&
            Index           =   1
            X1              =   15
            X2              =   5285
            Y1              =   1515
            Y2              =   1515
         End
         Begin VB.Line L1 
            BorderColor     =   &H80000003&
            Index           =   1
            X1              =   15
            X2              =   5285
            Y1              =   1500
            Y2              =   1500
         End
      End
   End
   Begin MSComctlLib.ProgressBar BP 
      Height          =   255
      Left            =   6120
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton BT_INF 
      Caption         =   "Re-imprimir"
      Height          =   855
      Left            =   3840
      Picture         =   "Tela_NFEmitidas.frx":0070
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Re-Imprime uma nota fiscal que tenha ocorrido problemas com a 1ª impressão"
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton BT_RNF 
      Caption         =   "N.F."
      Height          =   855
      Left            =   4800
      Picture         =   "Tela_NFEmitidas.frx":04B2
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Imprime relatório sobre a nota fiscal selecionada na lista"
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton BT_Relatorios 
      Caption         =   "Relatórios"
      Height          =   855
      Left            =   5760
      Picture         =   "Tela_NFEmitidas.frx":07BC
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Imprime relatório simplificado de todas notas fiscais na lista"
      Top             =   3600
      Width           =   855
   End
   Begin VB.ComboBox CB_Empresas 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Selecione uma empresa para filtrar as notas fiscais"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.ListBox LT_NF 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Lista de notas fiscais encontradas"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Frame FR 
      Caption         =   "Exibir N.F. por:"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton RB_Condicional 
         Caption         =   "Condicional"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton RB_Empresa 
         Caption         =   "Razão Social"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "Exibe notas fiscais pela empresa selecionada na lista abaixo"
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton RB_Todas 
         Caption         =   "Todas"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         ToolTipText     =   "Exibe todas notas fiscais"
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   6720
      Picture         =   "Tela_NFEmitidas.frx":0AC6
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volta à tela principal"
      Top             =   3600
      Width           =   855
   End
   Begin MSComctlLib.StatusBar BS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   4530
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Tela_NFEmitidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ****************** VARIÁVEIS DLL's ****************
Public DLL_BD As Scvbd.Classe_Scvbd
Dim DLL_CARGA As Scvcarr.Classe_Scvcarr
Public DLL_FUNCS As Scvfunc.Classe_Scvfunc
Dim DLL_ASFIG As Assfig.Classe_Assfig
Dim DLL_IMP As Impform.Classe_Impform

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Notas Fiscais Emitidas"
Dim TRANSPORTADORA As String, I As Integer, J As Integer, RespMsg
Public CRITERIO As String

Private Type EMPRESA
    RS As String
    APE As String
    CNPJ As String
    END As String
    BAI As String
    CEP As String
    MUN As String
    FON As String
    EST As String
    INE As String
    PRA As String
End Type
Dim REMETENTE As EMPRESA, TRANSPORTADOR As EMPRESA
Private Sub BT_INF_Click()
    On Error GoTo ERRO_SISCOVAL
    RespMsg = MsgBox("Se você prosseguir, você irá re-imprimir a nota fiscal selecionada. Somente prossiga se quando esta nota fiscal iria ser impressa, houve algum problema que interrompeu esta operação. Você tem certeza que deseja imprimir novamente esta nota fiscal ?", vbQuestion + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then
        DLL_FUNCS.RegistraEvento "Imprimir - Notas Fiscais Emitidas - ReImpressão", LT_NF.Text
        ReImprimirNF
        Unload Me
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Relatorios_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NF.ListCount < 1 Then
        MsgBox "Não existe ainda nenhuma nota fiscal na lista.", vbCritical + vbOKOnly, NOMEAPLIC
        LT_NF.SetFocus
        Exit Sub
    End If
    RespMsg = MsgBox("Esta operação irá imprimir um relatório simplificado com um resumo de todas notas fiscais listadas na lista ao lado. Você deseja continuar ?", vbInformation + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then
        TelaEmEspera (True)
        DLL_FUNCS.SelecionaImpressora (DLL_FUNCS.NomeImpressora("Tela_RelSimp"))
        BS.SimpleText = ""
        BP.Max = LT_NF.ListCount + 1
        BP.Value = 0
        Dim Num As Integer, Ind As Integer, Pags As Integer, Pg As Integer, I As Integer
        I = 0
        Num = 0
        Pags = LT_NF.ListCount / 10
        If Pags < 1 Then
            Pags = 1
        ElseIf Pags > Int(Pags) Then
            Pags = Int(Pags) + 1
        End If
        Pg = 1
        Do While True
            Ind = 0
            If (LT_NF.ListCount - 1) - Num > 10 Then
                For I = 0 To 9
                    RelSimp_Controles True, I
                Next I
                Tela_RelSimp.LB_FO2.Caption = Str(Pg) & "/" & Str(Pags)
                Tela_RelSimp.LB_DA2.Caption = Date
                Tela_RelSimp.LB_ME2.Caption = CRITERIO
                For I = Num To (Num + 9)
                    LT_NF.ListIndex = I
                    RelSimp_Dados (Ind)
                    BP.Value = BP.Value + 1
                    Ind = Ind + 1
                Next I
                Num = Num + 10
                Pg = Pg + 1
                Ind = 0
                Tela_RelSimp.Height = 15200
                Tela_RelSimp.PrintForm
            ElseIf (LT_NF.ListCount - 1) - Num <= 10 Then
                For I = 0 To 9
                    RelSimp_Controles False, I
                Next I
                Tela_RelSimp.LB_FO2.Caption = Str(Pg) & "/" & Str(Pags)
                Tela_RelSimp.LB_DA2.Caption = Date
                Tela_RelSimp.LB_ME2.Caption = CRITERIO
                For I = Num To LT_NF.ListCount - 1
                    LT_NF.ListIndex = I
                    RelSimp_Controles True, Ind
                    RelSimp_Dados (Ind)
                    BP.Value = BP.Value + 1
                    Ind = Ind + 1
                Next I
                Tela_RelSimp.Height = 15200
                Tela_RelSimp.PrintForm
                Exit Do
            End If
        Loop
        DLL_FUNCS.SelecionaImpressora (DLL_FUNCS.NomeImpressora("PADRÃO"))
        BS.SimpleText = ""
        BP.Value = 0
        DLL_FUNCS.RegistraEvento "Imprimir - Notas Fiscais Emitidas - Relatório Simplificado", ""
        Unload Me
    End If
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_RNF_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NF.ListIndex = -1 Then
        MsgBox "Não existe ainda nenhuma nota fiscal selecionada na lista.", vbCritical + vbOKOnly, NOMEAPLIC
        LT_NF.SetFocus
        Exit Sub
    End If
    RespMsg = MsgBox("Esta operação irá imprimir um relatório completo da nota fiscal " & Trim(LT_NF.Text) & ". Você deseja continuar ?", vbInformation + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then
        TelaEmEspera (True)
        DLL_FUNCS.SelecionaImpressora (DLL_FUNCS.NomeImpressora("Tela_Relatorio"))
        BS.SimpleText = "Preparando relatório completo de nota fiscal..."
        BP.Max = 10
        BP.Value = 0
        With Tela_Relatorio
            'Nota Fiscal Fatura
            BP.Value = BP.Value + 1
            .LB_NNF.Caption = LB_NF.Caption
            .LB_TP.Caption = LB_Tipo.Caption
            .LB_NO.Caption = LB_NO.Caption
            .LB_CF.Caption = LB_CFOP.Caption
            .LB_DE.Caption = LB_DE.Caption
            .LB_DS.Caption = LB_DS.Caption
            'Razao Social
            BP.Value = BP.Value + 1
            .LB_RS.Caption = REMETENTE.RS
            .LB_CN.Caption = REMETENTE.CNPJ
            .LB_IE.Caption = REMETENTE.INE
            .LB_EN.Caption = REMETENTE.END
            .LB_BA.Caption = REMETENTE.BAI
            .LB_MU.Caption = REMETENTE.MUN
            .LB_UF.Caption = REMETENTE.EST
            .LB_CE.Caption = REMETENTE.CEP
            'Fatura
            BP.Value = BP.Value + 1
            .LB_FE.Caption = LB_DE.Caption
            .LB_VA.Caption = LB_V.Caption
            .LB_DP.Caption = LB_DP.Caption
            .LB_VE.Caption = LB_NF.Caption
            .LB_PR.Caption = REMETENTE.PRA
            .LB_EX.Caption = "(" & DLL_FUNCS.ValorExtenso(LB_V.Caption) & ")"
            'Dados Gerais
            BP.Value = BP.Value + 1
            .LB_PI.Caption = LB_NP.Caption
            .LB_SP.Caption = LB_SP.Caption
            .LB_OP.Caption = LB_OP.Caption
            .LB_VI.Caption = LB_VI.Caption
            .LB_VX.Caption = LB_VX.Caption
            .LB_SE.Caption = LB_SE.Caption
            'Dados do Produto
            BP.Value = BP.Value + 1
            Dim NumItem As Integer
            NumItem = 0
            For I = 1 To 11
                For J = 1 To 20
                    .LB_CP(NumItem).Caption = FG_1.TextMatrix(J, I)
                    NumItem = NumItem + 1
                Next J
            Next I
            'Calculo do Imposto
            BP.Value = BP.Value + 1
            .LB_BI.Caption = LB_BI.Caption
            .LB_VM.Caption = LB_VC.Caption
            .LB_BS.Caption = LB_BS.Caption
            .LB_VS.Caption = LB_VS.Caption
            .LB_VT.Caption = LB_PR.Caption
            .LB_VF.Caption = LB_SE.Caption
            .LB_VS.Caption = LB_FT.Caption
            .LB_VO.Caption = LB_OD.Caption
            .LB_VP.Caption = LB_IP.Caption
            .LB_VN.Caption = LB_TO.Caption
            'Transportador
            BP.Value = BP.Value + 1
            .LB_TR.Caption = TRANSPORTADOR.RS
            .LB_FT.Caption = LB_FR.Caption
            .LB_PL.Caption = LB_PL.Caption
            .LB_VP.Caption = LB_EV.Caption
            .LB_CT.Caption = TRANSPORTADOR.CNPJ
            .LB_ET.Caption = TRANSPORTADOR.END
            .LB_MT.Caption = TRANSPORTADOR.MUN
            .LB_UT.Caption = TRANSPORTADOR.EST
            .LB_IT.Caption = TRANSPORTADOR.INE
            'Volumes Transportados
            BP.Value = BP.Value + 1
            .LB_QT.Caption = LB_QT.Caption
            .LB_ES.Caption = LB_EP.Caption
            .LB_MA.Caption = LB_MA.Caption
            .LB_NU.Caption = LB_NU.Caption
            .LB_PB.Caption = LB_PB.Caption
            .LB_PQ.Caption = LB_PQ.Caption
            'Desdobramento Duplicatas
            BP.Value = BP.Value + 1
            .LB_LA.Caption = LB_VA.Caption
            .LB_CA.Caption = LB_CA.Caption
            .LB_LB.Caption = LB_VB.Caption
            .LB_CB.Caption = LB_CB.Caption
            .LB_LC.Caption = LB_VAC.Caption
            .LB_CC.Caption = LB_CC.Caption
            .LB_LD.Caption = LB_VD.Caption
            .LB_CD.Caption = LB_CD.Caption
            'Prepara impressão
            BP.Value = BP.Value + 1
            .Height = 15500
            .PrintForm
        End With
        DLL_FUNCS.SelecionaImpressora (DLL_FUNCS.NomeImpressora("PADRÃO"))
        BS.SimpleText = ""
        BP.Value = 0
        TelaEmEspera (False)
        DLL_FUNCS.RegistraEvento "Imprimir - Notas Fiscais Emitidas - Relatório Completo", LB_NF.Caption
        Unload Me
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Me
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Empresas_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (False)
    'Carrega Notas Fiscais por Empresas
    ApagaDados
    LT_NF.Clear
    ResetaBP (DLL_BD.BDSIS_TBNTF.RecordCount + 5)
    DLL_BD.BDSIS_TBNTF.MoveFirst
    Do While Not DLL_BD.BDSIS_TBNTF.EOF
        CarregaBSEP ("Carregando notas fiscais de " & CB_Empresas.Text)
        If Trim(DLL_BD.BDSIS_TBNTF_CPEMP.Value) = Trim(CB_Empresas.Text) Then
            LT_NF.AddItem (DLL_BD.BDSIS_TBNTF_CPNNF.Value)
        End If
        DLL_BD.BDSIS_TBNTF.MoveNext
    Loop
ERRO_SISCOVAL:
    TelaEmEspera (False)
    ResetaBSEP
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub

Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    Set DLL_ASFIG = New Assfig.Classe_Assfig
    Set DLL_IMP = New Impform.Classe_Impform
    
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (24)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Empresas...")
    If DLL_BD.AbreTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Códigos Fiscais...")
    If DLL_BD.AbreTabela_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal...")
    If DLL_BD.AbreTabela_NotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Produtos...")
    If DLL_BD.AbreTabela_NotaFiscalProdutos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Declarações...")
    If DLL_BD.AbreTabela_NotaFiscalDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Bancos...")
    If DLL_BD.AbreTabela_NotaFiscalBancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Comentários...")
    If DLL_BD.AbreTabela_NotaFiscalComentarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Grupos...")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Bancos...")
    If DLL_BD.AbreTabela_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Configurações da Nota Fiscal...")
    If DLL_BD.AbreTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Abre Campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Empresas...")
    If DLL_BD.AbreCampos_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Códigos Fiscais...")
    If DLL_BD.AbreCampos_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal...")
    If DLL_BD.AbreCampos_NotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Produtos...")
    If DLL_BD.AbreCampos_NotaFiscalProdutos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Declarações...")
    If DLL_BD.AbreCampos_NotaFiscalDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Bancos...")
    If DLL_BD.AbreCampos_NotaFiscalBancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Comentários...")
    If DLL_BD.AbreCampos_NotaFiscalComentarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Bancos...")
    If DLL_BD.AbreCampos_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Configurações da Nota Fiscal")
    If DLL_BD.AbreCampos_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Apaga dados
    DLL_CARGA.CarregaTexto ("Apagando dados...")
    ApagaDados
    LT_NF.Clear
    CB_Empresas.Clear
    'Apaga dados
    DLL_CARGA.CarregaTexto ("Montando corpo da nota fiscal...")
    MontaFG
    'Carrega lista de empresas
    CB_Empresas.Clear
    DLL_BD.BDSIS_TBEMP.MoveFirst
    Do While Not DLL_BD.BDSIS_TBEMP.EOF
        If DLL_BD.BDSIS_TBEMP_CPAPE.Value <> "" Then
            CB_Empresas.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
        End If
        DLL_BD.BDSIS_TBEMP.MoveNext
    Loop
    
    'Finalizando
    ST.Tab = 0
    DLL_CARGA.CarregaTexto ("Finalizando")
    ResetaBSEP
    DLL_CARGA.Exibe (False)
    DLL_FUNCS.RegistraEvento "Abrir Notas Fiscais Emitidas", ""
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_NFEmitidas
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalProdutos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalBancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalComentarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
    Set DLL_FUNCS = Nothing
    Set DLL_ASFIG = Nothing
    Set DLL_IMP = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_NF_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    If LT_NF.Text = "5145" Then
        LT_NF.RemoveItem (LT_NF.ListIndex)
        Exit Sub
    End If
    ApagaDados
    MontaFG
    'Procura Nota Fiscal
    DLL_BD.BDSIS_TBNTF.Seek "=", LT_NF.Text
    If DLL_BD.BDSIS_TBNTF.NoMatch Then
        MsgBox "Não foi possível localizar a nota fiscal", vbCritical + vbOKOnly, NOMEAPLIC
        TelaEmEspera (False)
        Exit Sub
    End If
    'Aba Dados
    PegaEmpresa DLL_BD.BDSIS_TBNTF_CPEMP.Value, 1
    If DLL_BD.BDSIS_TBNTF_CPEMP.Value <> "" Then LB_RS.Caption = REMETENTE.RS
    If DLL_BD.BDSIS_TBNTF_CPDEM.Value <> "" Then LB_DE.Caption = Format(DLL_BD.BDSIS_TBNTF_CPDEM.Value, "dd/mm/yyyy")
    If DLL_BD.BDSIS_TBNTF_CPHEM.Value <> "" Then LB_HE.Caption = Format(DLL_BD.BDSIS_TBNTF_CPHEM.Value, "hh:mm:ss")
    If DLL_BD.BDSIS_TBNTF_CPDSA.Value <> "" Then LB_DS.Caption = Format(DLL_BD.BDSIS_TBNTF_CPDSA.Value, "dd/mm/yyyy")
    If DLL_BD.BDSIS_TBNTF_CPHSA.Value <> "" Then LB_HS.Caption = Format(DLL_BD.BDSIS_TBNTF_CPHSA.Value, "hh:mm:ss")
    If DLL_BD.BDSIS_TBNTF_CPTIP.Value <> "" Then LB_Tipo.Caption = DLL_BD.BDSIS_TBNTF_CPTIP.Value
    If DLL_BD.BDSIS_TBNTF_CPCFO.Value <> "" Then LB_CFOP.Caption = DLL_BD.BDSIS_TBNTF_CPCFO.Value
    If DLL_BD.BDSIS_TBNTF_CPNOP.Value <> "" Then LB_NO.Caption = DLL_BD.BDSIS_TBNTF_CPNOP.Value
    If DLL_BD.BDSIS_TBNTF_CPNNF.Value <> "" Then LB_NF.Caption = DLL_BD.BDSIS_TBNTF_CPNNF.Value
    If DLL_BD.BDSIS_TBNTF_CPNDP.Value <> "" Then LB_DP.Caption = DLL_BD.BDSIS_TBNTF_CPNDP.Value
    If DLL_BD.BDSIS_TBNTF_CPVEN.Value <> "" Then LB_VE.Caption = DLL_BD.BDSIS_TBNTF_CPVEN.Value
    If DLL_BD.BDSIS_TBNTF_CPVAL.Value <> "" Then LB_V.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVAL.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPNPI.Value <> "" Then LB_NP.Caption = DLL_BD.BDSIS_TBNTF_CPNPI.Value
    If DLL_BD.BDSIS_TBNTF_CPNSP.Value <> "" Then LB_SP.Caption = DLL_BD.BDSIS_TBNTF_CPNSP.Value
    If DLL_BD.BDSIS_TBNTF_CPOPE.Value <> "" Then LB_OP.Caption = DLL_BD.BDSIS_TBNTF_CPOPE.Value
    If DLL_BD.BDSIS_TBNTF_CPVIN.Value <> "" Then LB_VI.Caption = DLL_BD.BDSIS_TBNTF_CPVIN.Value
    If DLL_BD.BDSIS_TBNTF_CPVEX.Value <> "" Then LB_VX.Caption = DLL_BD.BDSIS_TBNTF_CPVEX.Value
    If DLL_BD.BDSIS_TBNTF_CPSET.Value <> "" Then LB_SE.Caption = DLL_BD.BDSIS_TBNTF_CPSET.Value
    'Aba Ítens
    ItemNF (DLL_BD.BDSIS_TBNTF_CPNNF.Value)
    'Se tiver declaracoes
    If DLL_BD.BDSIS_TBNTF_CPDEC.Value = True Then DecFiscalNF (DLL_BD.BDSIS_TBNTF_CPNNF.Value)
    'Se tiver bancos
    If DLL_BD.BDSIS_TBNTF_CPBAN.Value = True Then BancoNF (DLL_BD.BDSIS_TBNTF_CPNNF.Value)
    'Se tiver comentários
    If DLL_BD.BDSIS_TBNTF_CPCOM.Value = True Then ComentarioNF (DLL_BD.BDSIS_TBNTF_CPNNF.Value)
    'Aba Transportador
    PegaEmpresa DLL_BD.BDSIS_TBNTF_CPTRA.Value, 2
    If TRANSPORTADOR.APE = "SEU MOTORISTA" Then
        TRANSPORTADOR = REMETENTE
        TRANSPORTADOR.APE = "SEU MOTORISTA"
        TRANSPORTADOR.RS = "Seu motorista"
    End If
    If DLL_BD.BDSIS_TBNTF_CPTRA.Value <> "" Then LB_TR.Caption = TRANSPORTADOR.RS
    If DLL_BD.BDSIS_TBNTF_CPFCO.Value <> "" Then LB_FR.Caption = DLL_BD.BDSIS_TBNTF_CPFCO.Value
    If DLL_BD.BDSIS_TBNTF_CPPVE.Value <> "" Then LB_PL.Caption = DLL_BD.BDSIS_TBNTF_CPPVE.Value
    If DLL_BD.BDSIS_TBNTF_CPEPV.Value <> "" Then LB_EV.Caption = DLL_BD.BDSIS_TBNTF_CPEPV.Value
    If DLL_BD.BDSIS_TBNTF_CPQUA.Value <> "" Then LB_QT.Caption = Format(DLL_BD.BDSIS_TBNTF_CPQUA.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPESP.Value <> "" Then LB_EP.Caption = DLL_BD.BDSIS_TBNTF_CPESP.Value
    If DLL_BD.BDSIS_TBNTF_CPMAR.Value <> "" Then LB_MA.Caption = DLL_BD.BDSIS_TBNTF_CPMAR.Value
    If DLL_BD.BDSIS_TBNTF_CPNVO.Value <> "" Then LB_NU.Caption = DLL_BD.BDSIS_TBNTF_CPNVO.Value
    If DLL_BD.BDSIS_TBNTF_CPPBR.Value <> "" Then LB_PB.Caption = Format(DLL_BD.BDSIS_TBNTF_CPPBR.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPPLI.Value <> "" Then LB_PQ.Caption = Format(DLL_BD.BDSIS_TBNTF_CPPLI.Value, "###,###,###,##0.00")
    'Aba Valores
    If DLL_BD.BDSIS_TBNTF_CPBCI.Value <> "" Then LB_BI.Caption = Format(DLL_BD.BDSIS_TBNTF_CPBCI.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVIC.Value <> "" Then LB_VC.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVIC.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPBIS.Value <> "" Then LB_BS.Caption = Format(DLL_BD.BDSIS_TBNTF_CPBIS.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVIS.Value <> "" Then LB_VS.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVIS.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVFR.Value <> "" Then LB_FT.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVFR.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVSE.Value <> "" Then LB_SG.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVSE.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPODA.Value <> "" Then LB_OD.Caption = Format(DLL_BD.BDSIS_TBNTF_CPODA.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVTP.Value <> "" Then LB_PR.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVTP.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVTN.Value <> "" Then LB_TO.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVTN.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVIP.Value <> "" Then LB_IP.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVIP.Value, "###,###,###,##0.00")
    If DLL_BD.BDSIS_TBNTF_CPVCA.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVCA.Value <> "0" Then
        LB_CA.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVCA.Value, "dd/mm/yyyy")
    Else
        LB_CA.Caption = "-x-"
    End If
    If DLL_BD.BDSIS_TBNTF_CPVLA.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVLA.Value > 0 Then
        LB_VA.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVLA.Value, "###,###,###,##0.00")
    Else
        LB_VA.Caption = "-x-"
    End If
    If DLL_BD.BDSIS_TBNTF_CPVCB.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVCB.Value <> "0" Then
        LB_CB.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVCB.Value, "dd/mm/yyyy")
    Else
        LB_CB.Caption = "-x-"
    End If
    If DLL_BD.BDSIS_TBNTF_CPVLB.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVLB.Value > 0 Then
        LB_VB.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVLB.Value, "###,###,###,##0.00")
    Else
        LB_VB.Caption = "-x-"
    End If
    If DLL_BD.BDSIS_TBNTF_CPVCC.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVCC.Value <> "0" Then
        LB_CC.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVCC.Value, "dd/mm/yyyy")
    Else
        LB_CC.Caption = "-x-"
    End If
    If DLL_BD.BDSIS_TBNTF_CPVLC.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVLC.Value > 0 Then
        LB_VAC.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVLC.Value, "###,###,###,##0.00")
    Else
        LB_VAC.Caption = "-x-"
    End If
    If DLL_BD.BDSIS_TBNTF_CPVCD.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVCD.Value <> "0" Then
        LB_CD.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVCD.Value, "dd/mm/yyyy")
    Else
        LB_CD.Caption = "-x-"
    End If
    If DLL_BD.BDSIS_TBNTF_CPVLD.Value <> "" And DLL_BD.BDSIS_TBNTF_CPVLD.Value > 0 Then
        LB_VD.Caption = Format(DLL_BD.BDSIS_TBNTF_CPVLD.Value, "###,###,###,##0.00")
    Else
        LB_VD.Caption = "-x-"
    End If
ERRO_SISCOVAL:
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Condicional_Click()
    On Error GoTo ERRO_SISCOVAL
    Tela_Condicional.Show vbModal
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Empresa_Click()
    On Error GoTo ERRO_SISCOVAL
    CB_Empresas.Enabled = True
    ApagaDados
    LT_NF.Clear
    CRITERIO = "Por empresa: " & Trim(CB_Empresas.Text)
    CB_Empresas.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_Todas_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    CB_Empresas.Enabled = False
    ApagaDados
    LT_NF.Clear
    'Carrega Notas Fiscais
    ResetaBP (DLL_BD.BDSIS_TBNTF.RecordCount + 5)
    DLL_BD.BDSIS_TBNTF.MoveFirst
    Do While Not DLL_BD.BDSIS_TBNTF.EOF
        CarregaBSEP ("Carregando notas fiscais...")
        LT_NF.AddItem (DLL_BD.BDSIS_TBNTF_CPNNF.Value)
        DLL_BD.BDSIS_TBNTF.MoveNext
    Loop
    CRITERIO = "Todas notas fiscais"
    LT_NF.RemoveItem (0)
ERRO_SISCOVAL:
    ResetaBSEP
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub



'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub ApagaDados()
    On Error GoTo ERRO_SISCOVAL
    LB_RS.Caption = ""
    LB_DE.Caption = ""
    LB_HE.Caption = ""
    LB_DS.Caption = ""
    LB_HS.Caption = ""
    LB_Tipo.Caption = ""
    LB_CFOP.Caption = ""
    LB_NO.Caption = ""
    LB_NF.Caption = ""
    LB_DP.Caption = ""
    LB_VE.Caption = ""
    LB_V.Caption = ""
    LB_NP.Caption = ""
    LB_SP.Caption = ""
    LB_OP.Caption = ""
    LB_VI.Caption = ""
    LB_VX.Caption = ""
    LB_SE.Caption = ""
    LB_TR.Caption = ""
    LB_FR.Caption = ""
    LB_PL.Caption = ""
    LB_EV.Caption = ""
    LB_QT.Caption = ""
    LB_EP.Caption = ""
    LB_MA.Caption = ""
    LB_NU.Caption = ""
    LB_PB.Caption = ""
    LB_PQ.Caption = ""
    LB_BI.Caption = ""
    LB_VC.Caption = ""
    LB_BS.Caption = ""
    LB_VS.Caption = ""
    LB_FT.Caption = ""
    LB_SG.Caption = ""
    LB_OD.Caption = ""
    LB_PR.Caption = ""
    LB_TO.Caption = ""
    LB_CA.Caption = ""
    LB_VA.Caption = ""
    LB_CB.Caption = ""
    LB_VB.Caption = ""
    LB_CC.Caption = ""
    LB_VAC.Caption = ""
    LB_CD.Caption = ""
    LB_VD.Caption = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_NFEmitidas.MousePointer = vbHourglass
        Tela_NFEmitidas.Enabled = False
    Else
        Tela_NFEmitidas.MousePointer = vbDefault
        Tela_NFEmitidas.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub CarregaBSEP(Texto As String)
    On Error GoTo ERRO_SISCOVAL
    If Texto <> BS.SimpleText Then BS.SimpleText = Texto
    BP.Value = BP.Value + 1
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ResetaBP(Max As Integer)
    On Error GoTo ERRO_SISCOVAL
    BP.Max = Max
    BP.Value = 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ResetaBSEP()
    On Error GoTo ERRO_SISCOVAL
    BP.Value = 0
    BS.SimpleText = ""
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function PegaEmpresa(Apelido As String, Tipo As Integer)
    On Error GoTo ERRO_SISCOVAL
    If Apelido = "" Then Exit Function
    DLL_BD.BDSIS_TBEMP.Seek "=", Apelido
    If Not DLL_BD.BDSIS_TBEMP.NoMatch Then
        'Pega dados completos
        If Tipo = 1 Then
            REMETENTE.APE = Apelido
            If DLL_BD.BDSIS_TBEMP_CPEMP.Value <> "" Then REMETENTE.RS = DLL_BD.BDSIS_TBEMP_CPEMP.Value
            If DLL_BD.BDSIS_TBEMP_CPCGC.Value <> "" Then REMETENTE.CNPJ = DLL_BD.BDSIS_TBEMP_CPCGC.Value
            If DLL_BD.BDSIS_TBEMP_CPEND.Value <> "" Then REMETENTE.END = DLL_BD.BDSIS_TBEMP_CPEND.Value
            If DLL_BD.BDSIS_TBEMP_CPBAI.Value <> "" Then REMETENTE.BAI = DLL_BD.BDSIS_TBEMP_CPBAI.Value
            If DLL_BD.BDSIS_TBEMP_CPCEP.Value <> "" Then REMETENTE.CEP = DLL_BD.BDSIS_TBEMP_CPCEP.Value
            If DLL_BD.BDSIS_TBEMP_CPCID.Value <> "" Then REMETENTE.MUN = DLL_BD.BDSIS_TBEMP_CPCID.Value
            If DLL_BD.BDSIS_TBEMP_CPFON.Value <> "" Then REMETENTE.FON = DLL_BD.BDSIS_TBEMP_CPFON.Value
            If DLL_BD.BDSIS_TBEMP_CPEST.Value <> "" Then REMETENTE.EST = DLL_BD.BDSIS_TBEMP_CPEST.Value
            If DLL_BD.BDSIS_TBEMP_CPINE.Value <> "" Then REMETENTE.INE = DLL_BD.BDSIS_TBEMP_CPINE.Value
            If DLL_BD.BDSIS_TBEMP_CPPRA.Value <> "" Then REMETENTE.PRA = DLL_BD.BDSIS_TBEMP_CPPRA.Value
        ElseIf Tipo = 2 Then
            TRANSPORTADOR.APE = Apelido
            If DLL_BD.BDSIS_TBEMP_CPEMP.Value <> "" Then TRANSPORTADOR.RS = DLL_BD.BDSIS_TBEMP_CPEMP.Value
            If DLL_BD.BDSIS_TBEMP_CPCGC.Value <> "" Then TRANSPORTADOR.CNPJ = DLL_BD.BDSIS_TBEMP_CPCGC.Value
            If DLL_BD.BDSIS_TBEMP_CPEND.Value <> "" Then TRANSPORTADOR.END = DLL_BD.BDSIS_TBEMP_CPEND.Value
            If DLL_BD.BDSIS_TBEMP_CPBAI.Value <> "" Then TRANSPORTADOR.BAI = DLL_BD.BDSIS_TBEMP_CPBAI.Value
            If DLL_BD.BDSIS_TBEMP_CPCEP.Value <> "" Then TRANSPORTADOR.CEP = DLL_BD.BDSIS_TBEMP_CPCEP.Value
            If DLL_BD.BDSIS_TBEMP_CPCID.Value <> "" Then TRANSPORTADOR.MUN = DLL_BD.BDSIS_TBEMP_CPCID.Value
            If DLL_BD.BDSIS_TBEMP_CPFON.Value <> "" Then TRANSPORTADOR.FON = DLL_BD.BDSIS_TBEMP_CPFON.Value
            If DLL_BD.BDSIS_TBEMP_CPEST.Value <> "" Then TRANSPORTADOR.EST = DLL_BD.BDSIS_TBEMP_CPEST.Value
            If DLL_BD.BDSIS_TBEMP_CPINE.Value <> "" Then TRANSPORTADOR.INE = DLL_BD.BDSIS_TBEMP_CPINE.Value
            If DLL_BD.BDSIS_TBEMP_CPPRA.Value <> "" Then TRANSPORTADOR.PRA = DLL_BD.BDSIS_TBEMP_CPPRA.Value
        End If
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Static Sub MontaFG()
    On Error GoTo ERRO_SISCOVAL
    FG_1.Clear
    FG_2.Clear
    FG_1.Cols = 12
    FG_1.ColAlignment(0) = flexAlignCenterCenter
    FG_1.ColAlignment(1) = flexAlignLeftCenter
    FG_1.ColAlignment(2) = flexAlignLeftCenter
    FG_1.ColAlignment(3) = flexAlignCenterCenter
    FG_1.ColAlignment(4) = flexAlignCenterCenter
    FG_1.ColAlignment(5) = flexAlignCenterCenter
    FG_1.ColAlignment(9) = flexAlignCenterCenter
    FG_1.ColAlignment(10) = flexAlignCenterCenter
    FG_1.ColWidth(0) = 500
    FG_1.ColWidth(1) = 800
    FG_1.ColWidth(2) = 3300
    FG_1.ColWidth(3) = 400
    FG_1.ColWidth(4) = 400
    FG_1.ColWidth(5) = 700
    FG_1.ColWidth(6) = 900
    FG_1.ColWidth(7) = 1100
    FG_1.ColWidth(8) = 1100
    FG_1.ColWidth(9) = 800
    FG_1.ColWidth(10) = 600
    FG_1.TextArray(0) = "Linha"
    FG_1.TextArray(1) = "Figura"
    FG_1.TextArray(2) = "Descrição dos Produtos"
    FG_1.TextArray(3) = "C.F."
    FG_1.TextArray(4) = "S.T."
    FG_1.TextArray(5) = "Unidade"
    FG_1.TextArray(6) = "Quantidade"
    FG_1.TextArray(7) = "Preço Unitário"
    FG_1.TextArray(8) = "Preço Total"
    FG_1.TextArray(9) = "% I.C.M.S."
    FG_1.TextArray(10) = "% I.P.I."
    FG_1.TextArray(11) = "Valor I.P.I."
    For I = 1 To 20
        FG_1.TextMatrix(I, 0) = I
        FG_2.TextMatrix(I, 0) = I
    Next I
    FG_2.Cols = 6
    FG_2.TextArray(0) = "Linha"
    FG_2.TextArray(1) = "Bitola"
    FG_2.TextArray(2) = "Material"
    FG_2.TextArray(3) = "Base Cálculo I.C.M.S."
    FG_2.TextArray(4) = "Valor I.C.M.S."
    FG_2.TextArray(5) = "Peso Unitário"
    
    FG_1.Visible = True
    FG_2.Visible = False
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ItemNF(NumNF As Integer)
    On Error GoTo ERRO_SISCOVAL
    If NumNF < 0 Then Exit Sub
    Dim Linha As Integer
    
    DLL_BD.BDSIS_TBNFP.Index = "Número da Nota Fiscal"
    Dim NumReg As Integer
    NumReg = 0
    DLL_BD.BDSIS_TBNFP.MoveFirst
    'Procura primeiro item
    DLL_BD.BDSIS_TBNFP.Seek "=", NumNF
    Do While Not DLL_BD.BDSIS_TBNFP.EOF
        'Como os registros estão em ordem de numero da NF, qd. o nº for
        'maior que o nº NumNF, então acaba a pesquisa.
        If NumReg <> 0 Then
            If DLL_BD.BDSIS_TBNFP_CPNNF.Value > NumNF Then Exit Sub
        End If
        NumReg = 1
        
        For I = 1 To 20
            If FG_1.TextMatrix(I, 1) = "" And _
               FG_1.TextMatrix(I, 2) = "" Then
                Linha = I
                Exit For
            ElseIf I = 21 Then
                Exit Sub
            End If
        Next I
        
        'Verifica se a descricao é maior que 38 caracteres
        If Len(DLL_BD.BDSIS_TBNFP_CPDES.Value) <= 38 Then
            FG_1.TextMatrix(Linha, 1) = DLL_BD.BDSIS_TBNFP_CPFIG.Value
            FG_1.TextMatrix(Linha, 2) = DLL_BD.BDSIS_TBNFP_CPDES.Value
            FG_1.TextMatrix(Linha, 3) = DLL_BD.BDSIS_TBNFP_CPCCF.Value
            FG_1.TextMatrix(Linha, 4) = DLL_BD.BDSIS_TBNFP_CPCST.Value
            FG_1.TextMatrix(Linha, 5) = DLL_BD.BDSIS_TBNFP_CPUNI.Value
            FG_1.TextMatrix(Linha, 6) = DLL_BD.BDSIS_TBNFP_CPQUA.Value
            FG_1.TextMatrix(Linha, 7) = Format(DLL_BD.BDSIS_TBNFP_CPPUN.Value, "###,###,###,##0.00")
            FG_1.TextMatrix(Linha, 8) = Format(DLL_BD.BDSIS_TBNFP_CPPTO.Value, "###,###,###,##0.00")
            FG_1.TextMatrix(Linha, 9) = DLL_BD.BDSIS_TBNFP_CPAIC.Value
            FG_1.TextMatrix(Linha, 10) = DLL_BD.BDSIS_TBNFP_CPAIP.Value
            FG_1.TextMatrix(Linha, 11) = Format(DLL_BD.BDSIS_TBNFP_CPVIP.Value, "###,###,###,##0.00")
            
            FG_2.TextMatrix(Linha, 1) = DLL_BD.BDSIS_TBNFP_CPBIT.Value
            FG_2.TextMatrix(Linha, 2) = DLL_BD.BDSIS_TBNFP_CPMAT.Value
            FG_2.TextMatrix(Linha, 3) = Format(DLL_BD.BDSIS_TBNFP_CPBCI.Value, "###,###,###,##0.00")
            FG_2.TextMatrix(Linha, 4) = Format(DLL_BD.BDSIS_TBNFP_CPVIC.Value, "###,###,###,##0.00")
            FG_2.TextMatrix(Linha, 5) = Format(DLL_BD.BDSIS_TBNFP_CPPPA.Value, "####,##0.00")
        Else
            Dim NumLinLivre As Integer
            NumLinLivre = 0
            'verifica quantas linhas ainda estão livres
            For I = NumLinLivre To 20
                If FG_1.TextMatrix(I, 1) = "" And _
                   FG_1.TextMatrix(I, 2) = "" Then
                    NumLinLivre = NumLinLivre + 1
                End If
            Next I
            'funcao de divisao de linhas... permite até 5 linhas
            Dim TamDes As Integer
            TamDes = Len(DLL_BD.BDSIS_TBNFP_CPDES.Value)
            Dim DLR, DL1, DL2, DL3, DL4, DL5, DL6, DL7, DL8, DL9, DL10, DL11, DL12, DL13, DL14, DL15, DL16, DL17, DL18, DL19, DL20 As String
            DLR = DLL_BD.BDSIS_TBNFP_CPDES.Value
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
            'Corta itens
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
            'Insere itens
            FG_1.TextMatrix(Linha, 1) = DLL_BD.BDSIS_TBNFP_CPFIG.Value
            If DL2 <> "" And DL3 = "" Then
                FG_1.TextMatrix(Linha, 2) = DL1
                FG_2.TextMatrix(Linha, 1) = DLL_BD.BDSIS_TBNFP_CPBIT.Value
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL2
                FG_2.TextMatrix(Linha, 1) = "Idem"
            ElseIf DL3 <> "" And DL4 = "" Then
                FG_1.TextMatrix(Linha, 2) = DL1
                FG_2.TextMatrix(Linha, 1) = DLL_BD.BDSIS_TBNFP_CPBIT.Value
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL2
                FG_2.TextMatrix(Linha, 1) = "Idem"
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL3
                FG_2.TextMatrix(Linha, 1) = "Idem"
            ElseIf DL4 <> "" And DL5 = "" Then
                FG_1.TextMatrix(Linha, 2) = DL1
                FG_2.TextMatrix(Linha, 1) = DLL_BD.BDSIS_TBNFP_CPBIT.Value
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL2
                FG_2.TextMatrix(Linha, 1) = "Idem"
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL3
                FG_2.TextMatrix(Linha, 1) = "Idem"
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL4
                FG_2.TextMatrix(Linha, 1) = "Idem"
            ElseIf DL5 <> "" Then
                FG_1.TextMatrix(Linha, 2) = DL1
                FG_2.TextMatrix(Linha, 1) = DLL_BD.BDSIS_TBNFP_CPBIT.Value
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL2
                FG_2.TextMatrix(Linha, 1) = "Idem"
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL3
                FG_2.TextMatrix(Linha, 1) = "Idem"
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL4
                FG_2.TextMatrix(Linha, 1) = "Idem"
                Linha = Linha + 1
                FG_1.TextMatrix(Linha, 2) = DL5
                FG_2.TextMatrix(Linha, 1) = "Idem"
            End If
    
            FG_1.TextMatrix(Linha, 3) = DLL_BD.BDSIS_TBNFP_CPCCF.Value
            FG_1.TextMatrix(Linha, 4) = DLL_BD.BDSIS_TBNFP_CPCST.Value
            FG_1.TextMatrix(Linha, 5) = DLL_BD.BDSIS_TBNFP_CPUNI.Value
            FG_1.TextMatrix(Linha, 6) = DLL_BD.BDSIS_TBNFP_CPQUA.Value
            FG_1.TextMatrix(Linha, 7) = Format(DLL_BD.BDSIS_TBNFP_CPPUN.Value, "###,###,###,##0.00")
            FG_1.TextMatrix(Linha, 8) = Format(DLL_BD.BDSIS_TBNFP_CPPTO.Value, "###,###,###,##0.00")
            FG_1.TextMatrix(Linha, 9) = DLL_BD.BDSIS_TBNFP_CPAIC.Value
            FG_1.TextMatrix(Linha, 10) = DLL_BD.BDSIS_TBNFP_CPAIP.Value
            FG_1.TextMatrix(Linha, 11) = Format(DLL_BD.BDSIS_TBNFP_CPVIP.Value, "###,###,###,##0.00")
            
            FG_2.TextMatrix(Linha, 2) = DLL_BD.BDSIS_TBNFP_CPMAT.Value
            FG_2.TextMatrix(Linha, 3) = Format(DLL_BD.BDSIS_TBNFP_CPBCI.Value, "###,###,###,##0.00")
            FG_2.TextMatrix(Linha, 4) = Format(DLL_BD.BDSIS_TBNFP_CPVIC.Value, "###,###,###,##0.00")
            FG_2.TextMatrix(Linha, 5) = Format(DLL_BD.BDSIS_TBNFP_CPPPA.Value, "###,##0.00")
        End If
        DLL_BD.BDSIS_TBNFP.MoveNext
    Loop
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub DecFiscalNF(NumNF As Integer)
    On Error GoTo ERRO_SISCOVAL
    If NumNF < 0 Then Exit Sub
    DLL_BD.BDSIS_TBNFD.MoveFirst
    Do While Not DLL_BD.BDSIS_TBNFD.EOF
        If DLL_BD.BDSIS_TBNFD_CPNNF.Value < NumNF Then 'Iniciando procura
            DLL_BD.BDSIS_TBNFD.MoveNext
        ElseIf DLL_BD.BDSIS_TBNFD_CPNNF.Value > NumNF Then 'Acabou procura
            Exit Sub
        Else 'Se for a NumNF
            Dim nNumI As Integer
            nNumI = Val(DLL_BD.BDSIS_TBNFD_CPLIN.Value)
            If nNumI > 0 Then
                FG_1.TextMatrix(nNumI, 1) = DLL_BD.BDSIS_TBNFD_CPDEC.Value
                FG_2.TextMatrix(nNumI, 1) = "DF"
            Else
                For I = 20 To 1 Step -1
                    If FG_1.TextMatrix(I, 1) = "" And FG_2.TextMatrix(I, 1) = "" Then
                        FG_1.TextMatrix(I, 1) = DLL_BD.BDSIS_TBNFD_CPDEC.Value
                        FG_2.TextMatrix(I, 1) = "DF"
                        Exit For
                    End If
                Next I
            End If
            DLL_BD.BDSIS_TBNFD.MoveNext
        End If
    Loop
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub BancoNF(NumNF As Integer)
    On Error GoTo ERRO_SISCOVAL
    If NumNF < 0 Then Exit Sub
    DLL_BD.BDSIS_TBNFB.MoveFirst
    Do While Not DLL_BD.BDSIS_TBNFB.EOF
        If DLL_BD.BDSIS_TBNFB_CPNNF.Value < NumNF Then 'Iniciando procura
            DLL_BD.BDSIS_TBNFB.MoveNext
        ElseIf DLL_BD.BDSIS_TBNFB_CPNNF.Value > NumNF Then 'Acabou procura
            Exit Sub
        Else 'Se for a NumNF
            Dim nNumI As Integer
            nNumI = Val(DLL_BD.BDSIS_TBNFB_CPLIN.Value)
            If nNumI > 0 Then
                FG_1.TextMatrix(nNumI, 1) = DLL_BD.BDSIS_TBNFB_CPCON.Value
                FG_2.TextMatrix(nNumI, 1) = "DP"
            Else
                For I = 20 To 1 Step -1
                    If FG_1.TextMatrix(I, 1) = "" And FG_2.TextMatrix(I, 1) = "" Then
                        FG_1.TextMatrix(I, 1) = DLL_BD.BDSIS_TBNFB_CPCON.Value
                        FG_2.TextMatrix(I, 1) = "DP"
                        Exit For
                    End If
                Next I
            End If
            DLL_BD.BDSIS_TBNFB.MoveNext
        End If
    Loop
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ComentarioNF(NumNF As Integer)
    On Error GoTo ERRO_SISCOVAL
    If NumNF < 0 Then Exit Sub
    DLL_BD.BDSIS_TBNFC.MoveFirst
    Do While Not DLL_BD.BDSIS_TBNFC.EOF
        If DLL_BD.BDSIS_TBNFC_CPNNF.Value < NumNF Then 'Iniciando procura
            DLL_BD.BDSIS_TBNFC.MoveNext
        ElseIf DLL_BD.BDSIS_TBNFC_CPNNF.Value > NumNF Then 'Acabou procura
            Exit Sub
        Else 'Se for a NumNF
            Dim nNumI As Integer
            nNumI = Val(DLL_BD.BDSIS_TBNFC_CPLIN.Value)
            If nNumI > 0 Then
                FG_1.TextMatrix(nNumI, 1) = DLL_BD.BDSIS_TBNFC_CPCOM.Value
                FG_2.TextMatrix(nNumI, 1) = "CT"
            Else
                For I = 20 To 1 Step -1
                    If FG_1.TextMatrix(I, 1) = "" And FG_2.TextMatrix(I, 1) = "" Then
                        FG_1.TextMatrix(I, 1) = DLL_BD.BDSIS_TBNFC_CPCOM.Value
                        FG_2.TextMatrix(I, 1) = "CT"
                        Exit For
                    End If
                Next I
            End If
            DLL_BD.BDSIS_TBNFC.MoveNext
        End If
    Loop
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ReImprimirNF()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera (True)
    BS.SimpleText = "Aguarde... montando nota fiscal."
    ResetaBP DLL_IMP.NotaFiscal_LimpaItens
    'Envia dados para NF
        
    If LB_Tipo.Caption = "Saída" Then
        If EnviaDados("TIPO_SAIDA", "X") = False Then Exit Sub
    Else
        If EnviaDados("TIPO_ENTRADA", "X") = False Then Exit Sub
    End If
    If EnviaDados("NO", LB_NO.Caption) = False Then Exit Sub
    If EnviaDados("CFOP", LB_CFOP.Caption) = False Then Exit Sub
    'Destinatário
    If EnviaDados("RAZAO", LB_RS.Caption) = False Then Exit Sub
    If EnviaDados("CGC", REMETENTE.CNPJ) = False Then Exit Sub
    If EnviaDados("DATA_EMISSAO", Format(LB_DE.Caption, "dd/mm/yyyy")) = False Then Exit Sub
    If EnviaDados("ENDERECO", REMETENTE.END) = False Then Exit Sub
    If EnviaDados("BAIRRO", REMETENTE.BAI) = False Then Exit Sub
    If EnviaDados("CEP", REMETENTE.CEP) = False Then Exit Sub
    If LB_DS.Caption <> "" Then If EnviaDados("DATA_SAIDA", Format(LB_DS.Caption, "dd/mm/yyyy")) = False Then Exit Sub
    If LB_HS.Caption <> "" Then If EnviaDados("HORA_SAIDA", Format(LB_HS.Caption, "hh:mm:ss")) = False Then Exit Sub
    If EnviaDados("MUNICIPIO", REMETENTE.MUN) = False Then Exit Sub
    If EnviaDados("FONE", REMETENTE.FON) = False Then Exit Sub
    If EnviaDados("ESTADO", REMETENTE.EST) = False Then Exit Sub
    If EnviaDados("INS_EST", REMETENTE.INE) = False Then Exit Sub
    'Fatura
    If EnviaDados("DATA_EMISSAO_FATURA", Format(LB_DE.Caption, "dd/mm/yyyy")) = False Then Exit Sub
    If EnviaDados("NUM_NOTAFISCAL", LB_NF.Caption) = False Then Exit Sub
    If EnviaDados("VALOR_FATURA", Format(LB_V.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    If EnviaDados("NUM_DUPLICATA", LB_DP.Caption) = False Then Exit Sub
    If EnviaDados("DATA_VENCIMENTO", LB_VE.Caption) = False Then Exit Sub
    If EnviaDados("PRACA_PAGAMENTO", REMETENTE.PRA) = False Then Exit Sub
    Dim sExtenso As String
    sExtenso = "(" & DLL_FUNCS.ValorExtenso(LB_V.Caption) & ")" & DLL_FUNCS.MultiString("x ", 400)
    If EnviaDados("VALOR_EXTENSO_1", Mid(sExtenso, 1, 120)) = False Then Exit Sub
    If EnviaDados("VALOR_EXTENSO_2", Mid(sExtenso, 121, 240)) = False Then Exit Sub
    If LB_NP.Caption <> "" Then
        If EnviaDados("PI", LB_NP.Caption) = False Then Exit Sub Else
        If EnviaDados("PI", "-") = False Then Exit Sub
    End If
    If LB_SP.Caption <> "" Then
        If EnviaDados("SEU_PEDIDO", LB_SP.Caption) = False Then Exit Sub
    Else
        If EnviaDados("SEU_PEDIDO", "-") = False Then Exit Sub
    End If
    If LB_OP.Caption <> "" Then
        If EnviaDados("OP", LB_OP.Caption) = False Then Exit Sub
    Else
        If EnviaDados("OP", "-") = False Then Exit Sub
    End If
    'Dados dos Produtos
    Dim cDBTipoNF As String
    Dim cNumDBTipoNF As String
    Dim pCF_I As Boolean, pCF_J As Boolean
    pCF_I = False
    pCF_J = True
    cDBTipoNF = ""
    cNumDBTipoNF = ""
    For I = 1 To 20
        'Verifica o número da linha p/item
        If I < 10 Then
            cNumDBTipoNF = "0" & Trim(Str(I))
        Else
            cNumDBTipoNF = Trim(Str(I))
        End If
        'Aqui e para linhas de produtos
        If FG_2.TextMatrix(I, 1) <> "" And _
           FG_2.TextMatrix(I, 1) <> "CT" And _
           FG_2.TextMatrix(I, 1) <> "DF" And _
           FG_2.TextMatrix(I, 1) <> "DP" Then
            'Figura
            If FG_1.TextMatrix(I, 1) <> "" Then
                cDBTipoNF = "FIGURA_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 1)) = False Then Exit Sub
            End If
            'Descrição
            If FG_1.TextMatrix(I, 2) <> "" Then
                cDBTipoNF = "DESCRICAO_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 2)) = False Then Exit Sub
            End If
            'CF
            If FG_1.TextMatrix(I, 3) <> "" Then
                cDBTipoNF = "CF_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 3)) = False Then Exit Sub
            End If
            'ST
            If FG_1.TextMatrix(I, 4) <> "" Then
                cDBTipoNF = "ST_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 4)) = False Then Exit Sub
            End If
            'Unidade
            If FG_1.TextMatrix(I, 5) <> "" Then
                cDBTipoNF = "UNIDADE_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 5)) = False Then Exit Sub
            End If
            'Quantidade
            If FG_1.TextMatrix(I, 6) <> "" Then
                cDBTipoNF = "QUANTIDADE_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 6)) = False Then Exit Sub
            End If
            'Preço Unitário
            If FG_1.TextMatrix(I, 7) <> "" Then
                cDBTipoNF = "PRECO_UNITARIO_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 7)) = False Then Exit Sub
            End If
            'Preço Total
            If FG_1.TextMatrix(I, 8) <> "" Then
                cDBTipoNF = "PRECO_TOTAL_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 8)) = False Then Exit Sub
            End If
            'Aliquota ICMS
            If FG_1.TextMatrix(I, 9) <> "" Then
                cDBTipoNF = "ALIQUOTA_ICMS_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 9)) = False Then Exit Sub
            End If
            'Aliquota IPI
            If FG_1.TextMatrix(I, 10) <> "" Then
                cDBTipoNF = "ALIQUOTA_IPI_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 10)) = False Then Exit Sub
            End If
            'Valor IPI
            If FG_1.TextMatrix(I, 11) <> "" Then
                cDBTipoNF = "VALOR_IPI_" & cNumDBTipoNF
                If EnviaDados(cDBTipoNF, FG_1.TextMatrix(I, 11)) = False Then Exit Sub
            End If
        ElseIf FG_2.TextMatrix(I, 1) = "DF" Then 'Se for Declaracao Fiscal
            If FG_1.TextMatrix(I, 1) <> "" Then
                If EnviaDados("DECLARACAO", FG_1.TextMatrix(I, 1), I) = False Then Exit Sub
            End If
        ElseIf FG_2.TextMatrix(I, 1) = "CT" Then 'Se for Comentarios
            If FG_1.TextMatrix(I, 1) <> "" Then
                If EnviaDados("COMENTARIO", FG_1.TextMatrix(I, 1), I) = False Then Exit Sub
            End If
        ElseIf FG_2.TextMatrix(I, 1) = "DP" Then 'Se for Bancos
            If FG_1.TextMatrix(I, 1) <> "" Then
                If EnviaDados("BANCO", FG_1.TextMatrix(I, 1), I) = False Then Exit Sub
            End If
        End If
        'Verifica se existe CF I e J
        If FG_1.TextMatrix(I, 3) = "I" Then
            pCF_I = True
        ElseIf FG_1.TextMatrix(I, 3) = "J" Then
            pCF_J = True
        End If
    Next I
    'Cálculo do Imposto
    If LB_BI.Caption <> "0,00" Then
        If EnviaDados("BASE_CALCULO_ICMS", Format(LB_BI.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("BASE_CALCULO_ICMS", "-.-") = False Then Exit Sub
    End If
    If LB_VC.Caption <> "0,00" Then
        If EnviaDados("VALOR_ICMS", Format(LB_VC.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("VALOR_ICMS", "-.-") = False Then Exit Sub
    End If
    If LB_BS.Caption <> "0,00" Then
        If EnviaDados("BASE_CALCULO_ICMSSUB", Format(LB_BS.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("BASE_CALCULO_ICMSSUB", "-.-") = False Then Exit Sub
    End If
    If LB_VS.Caption <> "0,00" Then
        If EnviaDados("VALOR_ICMSSUB", Format(LB_VS.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("VALOR_ICMSSUB", "-.-") = False Then Exit Sub
    End If
    If LB_PR.Caption <> "0,00" Then
        If EnviaDados("VALOR_TOTAL_PRODUTOS", Format(LB_PR.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("VALOR_TOTAL_PRODUTOS", "-.-") = False Then Exit Sub
    End If
    If LB_FT.Caption <> "0,00" Then
        If EnviaDados("VALOR_FRETE", Format(LB_FT.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("VALOR_FRETE", "-.-") = False Then Exit Sub
    End If
    If LB_SG.Caption <> "0,00" Then
        If EnviaDados("VALOR_SEGURO", Format(LB_SG.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("VALOR_SEGURO", "-.-") = False Then Exit Sub
    End If
    If LB_OD.Caption <> "0,00" Then
        If EnviaDados("OUTRAS_DESPESAS", Format(LB_OD.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("OUTRAS_DESPESAS", "-.-") = False Then Exit Sub
    End If
    If LB_IP.Caption <> "0,00" Then
        If EnviaDados("VALOR_TOTAL_IPI", Format(LB_IP.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("VALOR_TOTAL_IPI", "-.-") = False Then Exit Sub
    End If
    If LB_TO.Caption <> "0,00" Then
        If EnviaDados("VALOR_TOTAL_NOTA", Format(LB_TO.Caption, "###,###,###,##0.00")) = False Then Exit Sub
    Else
        If EnviaDados("VALOR_TOTAL_NOTA", "-.-") = False Then Exit Sub
    End If
    'Transportador
    If EnviaDados("TRANSPORTADORA_NOME", TRANSPORTADOR.RS) = False Then Exit Sub
    If LB_FR.Caption = "Remetente" Then
        If EnviaDados("TRANSPORTADORA_FRETE", "1") = False Then Exit Sub
    ElseIf LB_FR.Caption = "Destinatário" Then
        If EnviaDados("TRANSPORTADORA_FRETE", "2") = False Then Exit Sub
    End If
    If LB_PL.Caption <> "" Then
        If EnviaDados("TRANSPORTADORA_PLACA", LB_PL.Caption) = False Then Exit Sub
    End If
    If LB_EV.Caption <> "" Then
        If EnviaDados("TRANSPORTADORA_UFVEI", LB_EV.Caption) = False Then Exit Sub
    End If
    If TRANSPORTADOR.CNPJ <> "" Then
        If EnviaDados("TRANSPORTADORA_CGC", TRANSPORTADOR.CNPJ) = False Then Exit Sub
    End If
    If TRANSPORTADOR.END <> "" Then
        If EnviaDados("TRANSPORTADORA_END", TRANSPORTADOR.END) = False Then Exit Sub
    End If
    If TRANSPORTADOR.MUN <> "" Then
        If EnviaDados("TRANSPORTADORA_MUN", TRANSPORTADOR.MUN) = False Then Exit Sub
    End If
    If TRANSPORTADOR.EST <> "" Then
        If EnviaDados("TRANSPORTADORA_UF", TRANSPORTADOR.EST) = False Then Exit Sub
    End If
    If TRANSPORTADOR.INE <> "" Then
        If EnviaDados("TRANSPORTADORA_IE", TRANSPORTADOR.INE) = False Then Exit Sub
    End If
    If LB_QT.Caption <> "" Then
        If EnviaDados("VOLUMES_QUANTIDADE", CDbl(LB_QT.Caption)) = False Then Exit Sub
    End If
    If LB_EP.Caption <> "" Then
        If EnviaDados("VOLUMES_ESPECIE", LB_EP.Caption) = False Then Exit Sub
    End If
    If LB_MA.Caption <> "" Then
        If EnviaDados("VOLUMES_MARCA", LB_MA.Caption) = False Then Exit Sub
    End If
    If LB_NU.Caption <> "" Then
        If EnviaDados("VOLUMES_NUMERO", LB_NU.Caption) = False Then Exit Sub
    End If
    If LB_PB.Caption <> "" Then
        If EnviaDados("VOLUMES_PESOBRUTO", Format(CDbl(LB_PB.Caption), "###,###,###,##0.00")) = False Then Exit Sub
    End If
    If LB_PQ.Caption <> "" Then
        If EnviaDados("VOLUMES_PESOLIQUIDO", Format(CDbl(LB_PQ.Caption), "###,###,###,##0.00")) = False Then Exit Sub
    End If
    'Dados Adicionais
    If LB_VI.Caption <> "" Then
        If EnviaDados("VEND_INT", LB_VI.Caption) = False Then Exit Sub
    End If
    If LB_VX.Caption <> "" Then
        If EnviaDados("VEND_EXT", LB_VX.Caption) = False Then Exit Sub
    End If
    If LB_SE.Caption <> "" Then
        If EnviaDados("VEND_SETOR", LB_SE.Caption) = False Then Exit Sub
    End If
    'Desdobramento DP
    If LB_CA.Caption <> "" Then
        If EnviaDados("DESD_DUP_A_VENC", LB_CA.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_A_VENC", "-x-") = False Then Exit Sub
    End If
    If LB_CB.Caption <> "" Then
        If EnviaDados("DESD_DUP_B_VENC", LB_CB.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_B_VENC", "-x-") = False Then Exit Sub
    End If
    If LB_CC.Caption <> "" Then
        If EnviaDados("DESD_DUP_C_VENC", LB_CC.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_C_VENC", "-x-") = False Then Exit Sub
    End If
    If LB_CD.Caption <> "" Then
        If EnviaDados("DESD_DUP_D_VENC", LB_CD.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_D_VENC", "-x-") = False Then Exit Sub
    End If
    If LB_VA.Caption <> "" Then
        If EnviaDados("DESD_DUP_A_VALOR", LB_VA.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_A_VALOR", "-x-") = False Then Exit Sub
    End If
    If LB_VB.Caption <> "" Then
        If EnviaDados("DESD_DUP_B_VALOR", LB_VB.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_B_VALOR", "-x-") = False Then Exit Sub
    End If
    If LB_VAC.Caption <> "" Then
        If EnviaDados("DESD_DUP_C_VALOR", LB_VAC.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_C_VALOR", "-x-") = False Then Exit Sub
    End If
    If LB_VD.Caption <> "" Then
        If EnviaDados("DESD_DUP_D_VALOR", LB_VD.Caption) = False Then Exit Sub
    Else
        If EnviaDados("DESD_DUP_D_VALOR", "-x-") = False Then Exit Sub
    End If
    'Declaracao do Simples
    'If EnviaDados("SIMPLES", "EMITENTE: Empresa Optante Pelo Simples.") = False Then Exit Sub
    'imprimir NF
    BS.SimpleText = "Imprimindo nota fiscal..."
    BP.Value = BP.Value + 1
    DLL_IMP.NotaFiscal_Imprimir (DLL_FUNCS.NomeImpressora("IT_NotaFiscal"))
ERRO_SISCOVAL:
    BS.SimpleText = ""
    BP.Value = 0
    TelaEmEspera (False)
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub RelSimp_Controles(Visivel As Boolean, Indice As Integer)
    On Error GoTo ERRO_SISCOVAL
    With Tela_RelSimp
        .L1(Indice).Visible = Visivel
        .L2(Indice).Visible = Visivel
        .LB_NF1(Indice).Visible = Visivel
        .LB_NF2(Indice).Visible = Visivel
        .LB_EM1(Indice).Visible = Visivel
        .LB_EM2(Indice).Visible = Visivel
        .LB_DE1(Indice).Visible = Visivel
        .LB_DE2(Indice).Visible = Visivel
        .LB_V1(Indice).Visible = Visivel
        .LB_V2(Indice).Visible = Visivel
        .LB_CA1(Indice).Visible = Visivel
        .LB_CA2(Indice).Visible = Visivel
        .LB_CB1(Indice).Visible = Visivel
        .LB_CB2(Indice).Visible = Visivel
        .LB_CC1(Indice).Visible = Visivel
        .LB_CC2(Indice).Visible = Visivel
        .LB_CD1(Indice).Visible = Visivel
        .LB_CD2(Indice).Visible = Visivel
        .LB_VA1(Indice).Visible = Visivel
        .LB_VA2(Indice).Visible = Visivel
        .LB_VB1(Indice).Visible = Visivel
        .LB_VB2(Indice).Visible = Visivel
        .LB_VC1(Indice).Visible = Visivel
        .LB_VC2(Indice).Visible = Visivel
        .LB_VD1(Indice).Visible = Visivel
        .LB_VD2(Indice).Visible = Visivel
        .LB_TP1(Indice).Visible = Visivel
        .LB_TP2(Indice).Visible = Visivel
        .LB_IC1(Indice).Visible = Visivel
        .LB_IC2(Indice).Visible = Visivel
        .LB_IP1(Indice).Visible = Visivel
        .LB_IP2(Indice).Visible = Visivel
        .LB_VN1(Indice).Visible = Visivel
        .LB_VN2(Indice).Visible = Visivel
        .LB_PI1(Indice).Visible = Visivel
        .LB_PI2(Indice).Visible = Visivel
        .LB_SP1(Indice).Visible = Visivel
        .LB_SP2(Indice).Visible = Visivel
        .LB_TR1(Indice).Visible = Visivel
        .LB_TR2(Indice).Visible = Visivel
        .LB_PB1(Indice).Visible = Visivel
        .LB_PB2(Indice).Visible = Visivel
    End With
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub RelSimp_Dados(I As Integer)
    On Error GoTo ERRO_SISCOVAL
    With Tela_RelSimp
        .LB_NF2(I).Caption = LB_NF.Caption
        .LB_EM2(I).Caption = LB_RS.Caption
        .LB_DE2(I).Caption = LB_DE.Caption
        .LB_V2(I).Caption = LB_V.Caption
        .LB_CA2(I).Caption = LB_CA.Caption
        .LB_CB2(I).Caption = LB_CB.Caption
        .LB_CC2(I).Caption = LB_CC.Caption
        .LB_CD2(I).Caption = LB_CD.Caption
        .LB_VA2(I).Caption = LB_VA.Caption
        .LB_VB2(I).Caption = LB_VB.Caption
        .LB_VC2(I).Caption = LB_VAC.Caption
        .LB_VD2(I).Caption = LB_VD.Caption
        .LB_TP2(I).Caption = LB_PR.Caption
        .LB_IC2(I).Caption = LB_VC.Caption
        .LB_IP2(I).Caption = LB_IP.Caption
        .LB_VN2(I).Caption = LB_TO.Caption
        .LB_PI2(I).Caption = LB_NP.Caption
        .LB_SP2(I).Caption = LB_SP.Caption
        .LB_TR2(I).Caption = LB_TR.Caption
        .LB_PB2(I).Caption = LB_PB.Caption
    End With
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function EnviaDados(Tipo As String, Texto As String, Optional NumI As Integer) As Boolean
    If NumI > 0 Then
        EnviaDados = DLL_IMP.NotaFiscal_MontaNF(Tipo, Texto, NumI)
    Else
        EnviaDados = DLL_IMP.NotaFiscal_MontaNF(Tipo, Texto)
    End If
    BP.Value = BP.Value + 1
End Function
