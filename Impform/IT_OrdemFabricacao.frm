VERSION 5.00
Begin VB.Form IT_OrdemFabricacao 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "4"
   ClientHeight    =   14700
   ClientLeft      =   0
   ClientTop       =   4500
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   14700
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PIC_LOGO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   2520
      Picture         =   "IT_OrdemFabricacao.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   112
      Top             =   240
      Width           =   600
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvulas e Conexões Industriais Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   240
      TabIndex        =   113
      Top             =   1250
      Width           =   5220
   End
   Begin VB.Line LH 
      Index           =   6
      X1              =   0
      X2              =   11160
      Y1              =   14310
      Y2              =   14310
   End
   Begin VB.Line LH 
      Index           =   7
      X1              =   0
      X2              =   11160
      Y1              =   14280
      Y2              =   14280
   End
   Begin VB.Line LH 
      Index           =   5
      X1              =   0
      X2              =   11160
      Y1              =   6870
      Y2              =   6870
   End
   Begin VB.Line LH 
      Index           =   4
      X1              =   0
      X2              =   11160
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line LH 
      Index           =   2
      X1              =   0
      X2              =   11160
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line LH 
      Index           =   26
      X1              =   0
      X2              =   11160
      Y1              =   6150
      Y2              =   6150
   End
   Begin VB.Line LH 
      Index           =   25
      X1              =   0
      X2              =   11160
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line LH 
      Index           =   24
      X1              =   0
      X2              =   11160
      Y1              =   5430
      Y2              =   5430
   End
   Begin VB.Line LH 
      Index           =   23
      X1              =   0
      X2              =   11160
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line LH 
      Index           =   22
      X1              =   0
      X2              =   11160
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line LH 
      Index           =   21
      X1              =   0
      X2              =   11160
      Y1              =   4710
      Y2              =   4710
   End
   Begin VB.Line LH 
      Index           =   20
      X1              =   0
      X2              =   11160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line LH 
      Index           =   8
      X1              =   0
      X2              =   11160
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line LH 
      Index           =   14
      X1              =   0
      X2              =   11160
      Y1              =   3990
      Y2              =   3990
   End
   Begin VB.Line LH 
      Index           =   13
      X1              =   0
      X2              =   11160
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line LH 
      Index           =   11
      X1              =   0
      X2              =   11160
      Y1              =   2250
      Y2              =   2250
   End
   Begin VB.Line LH 
      Index           =   3
      X1              =   0
      X2              =   11160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line LH 
      Index           =   10
      X1              =   0
      X2              =   11160
      Y1              =   90
      Y2              =   90
   End
   Begin VB.Line LH 
      Index           =   9
      X1              =   0
      X2              =   11160
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line LH 
      Index           =   0
      X1              =   0
      X2              =   11160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line LH 
      Index           =   1
      X1              =   0
      X2              =   11160
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line LV 
      Index           =   2
      X1              =   5520
      X2              =   5520
      Y1              =   120
      Y2              =   1560
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações sobre a peça para ser produzida:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   36
      Left            =   0
      TabIndex        =   5
      Top             =   1920
      Width           =   3960
   End
   Begin VB.Label LB_OFA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """E X I S T E     O F     A B E R T A"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6480
      TabIndex        =   111
      Top             =   1250
      Width           =   3705
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   1660
      TabIndex        =   110
      Top             =   13760
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   340
      TabIndex        =   109
      Top             =   13760
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   1660
      TabIndex        =   108
      Top             =   13040
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   340
      TabIndex        =   107
      Top             =   13040
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1660
      TabIndex        =   106
      Top             =   12320
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   340
      TabIndex        =   105
      Top             =   12320
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   1660
      TabIndex        =   104
      Top             =   11600
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   340
      TabIndex        =   103
      Top             =   11600
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1660
      TabIndex        =   102
      Top             =   10880
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   340
      TabIndex        =   101
      Top             =   10880
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   1660
      TabIndex        =   100
      Top             =   10160
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   340
      TabIndex        =   99
      Top             =   10160
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1660
      TabIndex        =   98
      Top             =   9440
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   340
      TabIndex        =   97
      Top             =   9440
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   1660
      TabIndex        =   96
      Top             =   8720
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   340
      TabIndex        =   95
      Top             =   8720
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1660
      TabIndex        =   94
      Top             =   8000
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   340
      TabIndex        =   93
      Top             =   8000
      Width           =   210
   End
   Begin VB.Label LB_M 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1660
      TabIndex        =   92
      Top             =   7280
      Width           =   210
   End
   Begin VB.Label LB_E 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   340
      TabIndex        =   91
      Top             =   7275
      Width           =   210
   End
   Begin VB.Label LB_DC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123456"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7080
      TabIndex        =   90
      Top             =   3480
      Width           =   630
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESENHO CONESTEEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   81
      Left            =   7080
      TabIndex        =   89
      Top             =   3150
      UseMnemonic     =   0   'False
      Width           =   1320
   End
   Begin VB.Line LV 
      Index           =   47
      X1              =   6960
      X2              =   6960
      Y1              =   3120
      Y2              =   3960
   End
   Begin VB.Label LB_Corrida 
      BackStyle       =   0  'Transparent
      Caption         =   "ABC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8760
      TabIndex        =   88
      Top             =   3360
      Width           =   2460
   End
   Begin VB.Label LB_QuantPro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 PÇ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9600
      TabIndex        =   87
      Top             =   2640
      Width           =   435
   End
   Begin VB.Line LV 
      Index           =   46
      X1              =   8640
      X2              =   8640
      Y1              =   3120
      Y2              =   3960
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CÓDIGOS DE CORRIDA UTILIZADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   80
      Left            =   8760
      TabIndex        =   86
      Top             =   3150
      UseMnemonic     =   0   'False
      Width           =   2055
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   9480
      TabIndex        =   85
      Top             =   13680
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   7920
      TabIndex        =   84
      Top             =   13680
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   4560
      TabIndex        =   83
      Top             =   13680
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   6120
      TabIndex        =   82
      Top             =   13680
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   2760
      TabIndex        =   81
      Top             =   13680
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   45
      X1              =   1920
      X2              =   1920
      Y1              =   13680
      Y2              =   14040
   End
   Begin VB.Line LH 
      Index           =   60
      X1              =   1560
      X2              =   1920
      Y1              =   14040
      Y2              =   14040
   End
   Begin VB.Line LH 
      Index           =   59
      X1              =   1560
      X2              =   1920
      Y1              =   13680
      Y2              =   13680
   End
   Begin VB.Line LV 
      Index           =   44
      X1              =   1560
      X2              =   1560
      Y1              =   13680
      Y2              =   14040
   End
   Begin VB.Line LV 
      Index           =   43
      X1              =   600
      X2              =   600
      Y1              =   13680
      Y2              =   14040
   End
   Begin VB.Line LH 
      Index           =   58
      X1              =   240
      X2              =   600
      Y1              =   14040
      Y2              =   14040
   End
   Begin VB.Line LH 
      Index           =   57
      X1              =   240
      X2              =   600
      Y1              =   13680
      Y2              =   13680
   End
   Begin VB.Line LV 
      Index           =   42
      X1              =   240
      X2              =   240
      Y1              =   13680
      Y2              =   14040
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   9480
      TabIndex        =   80
      Top             =   12960
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   7920
      TabIndex        =   79
      Top             =   12960
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   4560
      TabIndex        =   78
      Top             =   12960
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   6120
      TabIndex        =   77
      Top             =   12960
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   2760
      TabIndex        =   76
      Top             =   12960
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   41
      X1              =   1920
      X2              =   1920
      Y1              =   12960
      Y2              =   13320
   End
   Begin VB.Line LH 
      Index           =   56
      X1              =   1560
      X2              =   1920
      Y1              =   13320
      Y2              =   13320
   End
   Begin VB.Line LH 
      Index           =   55
      X1              =   1560
      X2              =   1920
      Y1              =   12960
      Y2              =   12960
   End
   Begin VB.Line LV 
      Index           =   40
      X1              =   1560
      X2              =   1560
      Y1              =   12960
      Y2              =   13320
   End
   Begin VB.Line LV 
      Index           =   39
      X1              =   600
      X2              =   600
      Y1              =   12960
      Y2              =   13320
   End
   Begin VB.Line LH 
      Index           =   54
      X1              =   240
      X2              =   600
      Y1              =   13320
      Y2              =   13320
   End
   Begin VB.Line LH 
      Index           =   53
      X1              =   240
      X2              =   600
      Y1              =   12960
      Y2              =   12960
   End
   Begin VB.Line LV 
      Index           =   38
      X1              =   240
      X2              =   240
      Y1              =   12960
      Y2              =   13320
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   9480
      TabIndex        =   75
      Top             =   12240
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   7920
      TabIndex        =   74
      Top             =   12240
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   4560
      TabIndex        =   73
      Top             =   12240
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   6120
      TabIndex        =   72
      Top             =   12240
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   2760
      TabIndex        =   71
      Top             =   12240
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   37
      X1              =   1920
      X2              =   1920
      Y1              =   12240
      Y2              =   12600
   End
   Begin VB.Line LH 
      Index           =   52
      X1              =   1560
      X2              =   1920
      Y1              =   12600
      Y2              =   12600
   End
   Begin VB.Line LH 
      Index           =   51
      X1              =   1560
      X2              =   1920
      Y1              =   12240
      Y2              =   12240
   End
   Begin VB.Line LV 
      Index           =   36
      X1              =   1560
      X2              =   1560
      Y1              =   12240
      Y2              =   12600
   End
   Begin VB.Line LV 
      Index           =   35
      X1              =   600
      X2              =   600
      Y1              =   12240
      Y2              =   12600
   End
   Begin VB.Line LH 
      Index           =   50
      X1              =   240
      X2              =   600
      Y1              =   12600
      Y2              =   12600
   End
   Begin VB.Line LH 
      Index           =   49
      X1              =   240
      X2              =   600
      Y1              =   12240
      Y2              =   12240
   End
   Begin VB.Line LV 
      Index           =   34
      X1              =   240
      X2              =   240
      Y1              =   12240
      Y2              =   12600
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   9480
      TabIndex        =   70
      Top             =   11520
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   7920
      TabIndex        =   69
      Top             =   11520
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   4560
      TabIndex        =   68
      Top             =   11520
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   6120
      TabIndex        =   67
      Top             =   11520
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   2760
      TabIndex        =   66
      Top             =   11520
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   33
      X1              =   1920
      X2              =   1920
      Y1              =   11520
      Y2              =   11880
   End
   Begin VB.Line LH 
      Index           =   48
      X1              =   1560
      X2              =   1920
      Y1              =   11880
      Y2              =   11880
   End
   Begin VB.Line LH 
      Index           =   47
      X1              =   1560
      X2              =   1920
      Y1              =   11520
      Y2              =   11520
   End
   Begin VB.Line LV 
      Index           =   32
      X1              =   1560
      X2              =   1560
      Y1              =   11520
      Y2              =   11880
   End
   Begin VB.Line LV 
      Index           =   31
      X1              =   600
      X2              =   600
      Y1              =   11520
      Y2              =   11880
   End
   Begin VB.Line LH 
      Index           =   46
      X1              =   240
      X2              =   600
      Y1              =   11880
      Y2              =   11880
   End
   Begin VB.Line LH 
      Index           =   45
      X1              =   240
      X2              =   600
      Y1              =   11520
      Y2              =   11520
   End
   Begin VB.Line LV 
      Index           =   30
      X1              =   240
      X2              =   240
      Y1              =   11520
      Y2              =   11880
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   9480
      TabIndex        =   65
      Top             =   10800
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   7920
      TabIndex        =   64
      Top             =   10800
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   4560
      TabIndex        =   63
      Top             =   10800
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   6120
      TabIndex        =   62
      Top             =   10800
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   2760
      TabIndex        =   61
      Top             =   10800
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   29
      X1              =   1920
      X2              =   1920
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Line LH 
      Index           =   44
      X1              =   1560
      X2              =   1920
      Y1              =   11160
      Y2              =   11160
   End
   Begin VB.Line LH 
      Index           =   43
      X1              =   1560
      X2              =   1920
      Y1              =   10800
      Y2              =   10800
   End
   Begin VB.Line LV 
      Index           =   28
      X1              =   1560
      X2              =   1560
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Line LV 
      Index           =   27
      X1              =   600
      X2              =   600
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Line LH 
      Index           =   42
      X1              =   240
      X2              =   600
      Y1              =   11160
      Y2              =   11160
   End
   Begin VB.Line LH 
      Index           =   41
      X1              =   240
      X2              =   600
      Y1              =   10800
      Y2              =   10800
   End
   Begin VB.Line LV 
      Index           =   26
      X1              =   240
      X2              =   240
      Y1              =   10800
      Y2              =   11160
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   9480
      TabIndex        =   60
      Top             =   10080
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   7920
      TabIndex        =   59
      Top             =   10080
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   4560
      TabIndex        =   58
      Top             =   10080
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   6120
      TabIndex        =   57
      Top             =   10080
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2760
      TabIndex        =   56
      Top             =   10080
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   25
      X1              =   1920
      X2              =   1920
      Y1              =   10080
      Y2              =   10440
   End
   Begin VB.Line LH 
      Index           =   40
      X1              =   1560
      X2              =   1920
      Y1              =   10440
      Y2              =   10440
   End
   Begin VB.Line LH 
      Index           =   39
      X1              =   1560
      X2              =   1920
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line LV 
      Index           =   24
      X1              =   1560
      X2              =   1560
      Y1              =   10080
      Y2              =   10440
   End
   Begin VB.Line LV 
      Index           =   23
      X1              =   600
      X2              =   600
      Y1              =   10080
      Y2              =   10440
   End
   Begin VB.Line LH 
      Index           =   38
      X1              =   240
      X2              =   600
      Y1              =   10440
      Y2              =   10440
   End
   Begin VB.Line LH 
      Index           =   37
      X1              =   240
      X2              =   600
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line LV 
      Index           =   19
      X1              =   240
      X2              =   240
      Y1              =   10080
      Y2              =   10440
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   9480
      TabIndex        =   55
      Top             =   9360
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   7920
      TabIndex        =   54
      Top             =   9360
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   4560
      TabIndex        =   53
      Top             =   9360
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   6120
      TabIndex        =   52
      Top             =   9360
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2760
      TabIndex        =   51
      Top             =   9360
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   18
      X1              =   1920
      X2              =   1920
      Y1              =   9360
      Y2              =   9720
   End
   Begin VB.Line LH 
      Index           =   36
      X1              =   1560
      X2              =   1920
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line LH 
      Index           =   35
      X1              =   1560
      X2              =   1920
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line LV 
      Index           =   17
      X1              =   1560
      X2              =   1560
      Y1              =   9360
      Y2              =   9720
   End
   Begin VB.Line LV 
      Index           =   16
      X1              =   600
      X2              =   600
      Y1              =   9360
      Y2              =   9720
   End
   Begin VB.Line LH 
      Index           =   34
      X1              =   240
      X2              =   600
      Y1              =   9720
      Y2              =   9720
   End
   Begin VB.Line LH 
      Index           =   33
      X1              =   240
      X2              =   600
      Y1              =   9360
      Y2              =   9360
   End
   Begin VB.Line LV 
      Index           =   15
      X1              =   240
      X2              =   240
      Y1              =   9360
      Y2              =   9720
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   9480
      TabIndex        =   50
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   7920
      TabIndex        =   49
      Top             =   8640
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   4560
      TabIndex        =   48
      Top             =   8640
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   6120
      TabIndex        =   47
      Top             =   8640
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2760
      TabIndex        =   46
      Top             =   8640
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   14
      X1              =   1920
      X2              =   1920
      Y1              =   8640
      Y2              =   9000
   End
   Begin VB.Line LH 
      Index           =   32
      X1              =   1560
      X2              =   1920
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line LH 
      Index           =   31
      X1              =   1560
      X2              =   1920
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line LV 
      Index           =   13
      X1              =   1560
      X2              =   1560
      Y1              =   8640
      Y2              =   9000
   End
   Begin VB.Line LV 
      Index           =   12
      X1              =   600
      X2              =   600
      Y1              =   8640
      Y2              =   9000
   End
   Begin VB.Line LH 
      Index           =   30
      X1              =   240
      X2              =   600
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line LH 
      Index           =   29
      X1              =   240
      X2              =   600
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line LV 
      Index           =   11
      X1              =   240
      X2              =   240
      Y1              =   8640
      Y2              =   9000
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   9480
      TabIndex        =   45
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   7920
      TabIndex        =   44
      Top             =   7920
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4560
      TabIndex        =   43
      Top             =   7920
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   6120
      TabIndex        =   42
      Top             =   7920
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2760
      TabIndex        =   41
      Top             =   7920
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   10
      X1              =   1920
      X2              =   1920
      Y1              =   7920
      Y2              =   8280
   End
   Begin VB.Line LH 
      Index           =   28
      X1              =   1560
      X2              =   1920
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line LH 
      Index           =   27
      X1              =   1560
      X2              =   1920
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line LV 
      Index           =   9
      X1              =   1560
      X2              =   1560
      Y1              =   7920
      Y2              =   8280
   End
   Begin VB.Line LV 
      Index           =   8
      X1              =   600
      X2              =   600
      Y1              =   7920
      Y2              =   8280
   End
   Begin VB.Line LH 
      Index           =   19
      X1              =   240
      X2              =   600
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line LH 
      Index           =   18
      X1              =   240
      X2              =   600
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line LV 
      Index           =   7
      X1              =   240
      X2              =   240
      Y1              =   7920
      Y2              =   8280
   End
   Begin VB.Label LB_HE 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_______________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   9480
      TabIndex        =   40
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label LB_HF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   7920
      TabIndex        =   39
      Top             =   7200
      Width           =   1110
   End
   Begin VB.Label LB_HI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "_____:_____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   4560
      TabIndex        =   38
      Top             =   7200
      Width           =   1110
   End
   Begin VB.Label LB_DF 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   6120
      TabIndex        =   37
      Top             =   7200
      Width           =   1380
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "____/____/____"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2760
      TabIndex        =   36
      Top             =   7200
      Width           =   1380
   End
   Begin VB.Line LV 
      Index           =   6
      X1              =   1920
      X2              =   1920
      Y1              =   7200
      Y2              =   7560
   End
   Begin VB.Line LH 
      Index           =   17
      X1              =   1560
      X2              =   1920
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line LH 
      Index           =   16
      X1              =   1560
      X2              =   1920
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line LV 
      Index           =   5
      X1              =   1560
      X2              =   1560
      Y1              =   7200
      Y2              =   7560
   End
   Begin VB.Line LV 
      Index           =   3
      X1              =   600
      X2              =   600
      Y1              =   7200
      Y2              =   7560
   End
   Begin VB.Line LH 
      Index           =   15
      X1              =   240
      X2              =   600
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line LH 
      Index           =   12
      X1              =   240
      X2              =   600
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line LV 
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   7200
      Y2              =   7560
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANT. H.EXTRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   18
      Left            =   9480
      TabIndex        =   35
      Top             =   6555
      Width           =   1590
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA FIM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   17
      Left            =   7980
      TabIndex        =   34
      Top             =   6555
      Width           =   945
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA FIM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   16
      Left            =   6360
      TabIndex        =   33
      Top             =   6555
      Width           =   900
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA INÍCIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   15
      Left            =   11520
      TabIndex        =   32
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA INÍCIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   14
      Left            =   4545
      TabIndex        =   31
      Top             =   6555
      Width           =   1170
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA INÍCIO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   13
      Left            =   2880
      TabIndex        =   30
      Top             =   6555
      Width           =   1125
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº MÁQUINA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   12
      Left            =   1185
      TabIndex        =   29
      Top             =   6555
      Width           =   1170
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº ETAPA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   11
      Left            =   0
      TabIndex        =   28
      Top             =   6555
      Width           =   915
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações sobre as etapas dos processos de fabricação:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   0
      TabIndex        =   27
      Top             =   5760
      Width           =   5115
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESSOS DE FABRICAÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   42
      Left            =   4320
      TabIndex        =   26
      Top             =   6195
      Width           =   2895
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7: CNC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   41
      Left            =   10560
      TabIndex        =   25
      Top             =   5115
      Width           =   630
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6: ROSQUEADEIRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   40
      Left            =   8400
      TabIndex        =   24
      Top             =   5115
      Width           =   1785
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5: FURADEIRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   39
      Left            =   6600
      TabIndex        =   23
      Top             =   5115
      Width           =   1335
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "4: T. ROSCA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   38
      Left            =   5040
      TabIndex        =   22
      Top             =   5115
      Width           =   1140
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3: T. USINAGEM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   37
      Left            =   3120
      TabIndex        =   21
      Top             =   5115
      Width           =   1485
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2: T. FURAÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   34
      Left            =   1320
      TabIndex        =   20
      Top             =   5115
      Width           =   1395
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1: SERRA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   33
      Left            =   0
      TabIndex        =   19
      Top             =   5115
      Width           =   900
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Legenda dos Processos de Fabricação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   35
      Left            =   4080
      TabIndex        =   18
      Top             =   4740
      Width           =   3390
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MATERIAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   28
      Left            =   7560
      TabIndex        =   17
      Top             =   2300
      UseMnemonic     =   0   'False
      Width           =   585
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-105"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7560
      TabIndex        =   16
      Top             =   2640
      Width           =   510
   End
   Begin VB.Line LV 
      Index           =   22
      X1              =   7440
      X2              =   7440
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BITOLA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   9
      Left            =   5520
      TabIndex        =   15
      Top             =   2300
      UseMnemonic     =   0   'False
      Width           =   405
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1/2"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5520
      TabIndex        =   14
      Top             =   2640
      Width           =   345
   End
   Begin VB.Line LV 
      Index           =   21
      X1              =   5400
      X2              =   5400
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIGURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   8
      Left            =   3240
      TabIndex        =   13
      Top             =   2300
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CP-COR-N8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   12
      Top             =   2640
      Width           =   1050
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE PRODUZIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   7
      Left            =   9600
      TabIndex        =   11
      Top             =   2300
      UseMnemonic     =   0   'False
      Width           =   1515
   End
   Begin VB.Line LV 
      Index           =   20
      X1              =   9480
      X2              =   9480
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   5
      Left            =   0
      TabIndex        =   10
      Top             =   3150
      UseMnemonic     =   0   'False
      Width           =   705
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   3480
      Width           =   5400
   End
   Begin VB.Line LV 
      Index           =   4
      X1              =   3120
      X2              =   3120
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE ESTIPULADA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   4
      Left            =   1320
      TabIndex        =   8
      Top             =   2300
      UseMnemonic     =   0   'False
      Width           =   1530
   End
   Begin VB.Label LB_QuantEst 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1 PÇ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1320
      TabIndex        =   7
      Top             =   2640
      Width           =   435
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   7680
      TabIndex        =   6
      Top             =   900
      Width           =   330
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEM DE FABRICAÇÃO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6840
      TabIndex        =   4
      Top             =   360
      Width           =   2970
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label LB_Num 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0123"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8160
      TabIndex        =   2
      Top             =   850
      Width           =   720
   End
   Begin VB.Label LB_Data 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   960
   End
   Begin VB.Line LV 
      Index           =   0
      X1              =   1200
      X2              =   1200
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   6
      Left            =   0
      TabIndex        =   0
      Top             =   2300
      UseMnemonic     =   0   'False
      Width           =   315
   End
End
Attribute VB_Name = "IT_OrdemFabricacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
