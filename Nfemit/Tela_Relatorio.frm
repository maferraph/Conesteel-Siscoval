VERSION 5.00
Begin VB.Form Tela_Relatorio 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   15495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   15495
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PIC_LOGO 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   240
      Picture         =   "Tela_Relatorio.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   367
      Top             =   480
      Width           =   600
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR / C:"
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
      Index           =   79
      Left            =   5880
      TabIndex        =   366
      Top             =   14640
      Width           =   600
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENCIMENTO / C:"
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
      Index           =   78
      Left            =   6960
      TabIndex        =   365
      Top             =   14640
      Width           =   1005
   End
   Begin VB.Label LB_LC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   5880
      TabIndex        =   364
      Top             =   14760
      Width           =   690
   End
   Begin VB.Label LB_CC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
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
      Left            =   6960
      TabIndex        =   363
      Top             =   14760
      Width           =   960
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR / D:"
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
      Index           =   72
      Left            =   8760
      TabIndex        =   362
      Top             =   14640
      Width           =   600
   End
   Begin VB.Label LB_LD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   8760
      TabIndex        =   361
      Top             =   14760
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENCIMENTO / D:"
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
      Index           =   71
      Left            =   9840
      TabIndex        =   360
      Top             =   14640
      Width           =   1005
   End
   Begin VB.Label LB_CD 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
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
      Left            =   9840
      TabIndex        =   359
      Top             =   14760
      Width           =   960
   End
   Begin VB.Line L2 
      Index           =   7
      X1              =   120
      X2              =   11520
      Y1              =   15000
      Y2              =   15000
   End
   Begin VB.Line L1 
      Index           =   9
      X1              =   120
      X2              =   11520
      Y1              =   14640
      Y2              =   14640
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESDOBRAMENTO DE DUPLICATAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   77
      Left            =   120
      TabIndex        =   358
      Top             =   14520
      Width           =   2130
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR / A:"
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
      Index           =   76
      Left            =   120
      TabIndex        =   357
      Top             =   14640
      Width           =   585
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENCIMENTO / A:"
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
      Index           =   75
      Left            =   1200
      TabIndex        =   356
      Top             =   14640
      Width           =   990
   End
   Begin VB.Label LB_LA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   120
      TabIndex        =   355
      Top             =   14760
      Width           =   690
   End
   Begin VB.Label LB_CA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
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
      Left            =   1200
      TabIndex        =   354
      Top             =   14760
      Width           =   960
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR / B:"
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
      Index           =   74
      Left            =   3000
      TabIndex        =   353
      Top             =   14640
      Width           =   585
   End
   Begin VB.Label LB_LB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   3000
      TabIndex        =   352
      Top             =   14760
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENCIMENTO / B:"
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
      Index           =   73
      Left            =   4080
      TabIndex        =   351
      Top             =   14640
      Width           =   990
   End
   Begin VB.Label LB_CB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
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
      Left            =   4080
      TabIndex        =   350
      Top             =   14760
      Width           =   960
   End
   Begin VB.Line L2 
      Index           =   6
      X1              =   120
      X2              =   11520
      Y1              =   14280
      Y2              =   14280
   End
   Begin VB.Line L1 
      Index           =   8
      X1              =   120
      X2              =   11520
      Y1              =   13920
      Y2              =   13920
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VOLUMES TRANSPORTADOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   70
      Left            =   120
      TabIndex        =   349
      Top             =   13800
      Width           =   1710
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE:"
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
      Index           =   69
      Left            =   120
      TabIndex        =   348
      Top             =   13920
      Width           =   810
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESPÉCIE:"
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
      Index           =   68
      Left            =   1800
      TabIndex        =   347
      Top             =   13920
      Width           =   525
   End
   Begin VB.Label LB_QT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   120
      TabIndex        =   346
      Top             =   14040
      Width           =   630
   End
   Begin VB.Label LB_ES 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pacote (s)"
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
      Left            =   1800
      TabIndex        =   345
      Top             =   14040
      Width           =   900
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MARCA:"
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
      Index           =   67
      Left            =   4080
      TabIndex        =   344
      Top             =   13920
      Width           =   465
   End
   Begin VB.Label LB_MA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel"
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
      Left            =   4080
      TabIndex        =   343
      Top             =   14040
      Width           =   870
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO:"
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
      Index           =   66
      Left            =   5880
      TabIndex        =   342
      Top             =   13920
      Width           =   570
   End
   Begin VB.Label LB_NU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   5880
      TabIndex        =   341
      Top             =   14040
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESO BRUTO:"
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
      Index           =   65
      Left            =   7800
      TabIndex        =   340
      Top             =   13920
      Width           =   795
   End
   Begin VB.Label LB_PB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00,00"
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
      Left            =   7800
      TabIndex        =   339
      Top             =   14040
      Width           =   480
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PESO LÍQUIDO:"
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
      Index           =   62
      Left            =   9840
      TabIndex        =   338
      Top             =   13920
      Width           =   855
   End
   Begin VB.Label LB_PQ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00,00"
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
      Left            =   9840
      TabIndex        =   337
      Top             =   14040
      Width           =   480
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENDEREÇO:"
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
      Index           =   64
      Left            =   120
      TabIndex        =   336
      Top             =   13200
      Width           =   705
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUNICÍPIO:"
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
      Index           =   63
      Left            =   5040
      TabIndex        =   335
      Top             =   13200
      Width           =   660
   End
   Begin VB.Label LB_ET 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avenida Montemagno, 2454"
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
      Left            =   120
      TabIndex        =   334
      Top             =   13320
      Width           =   2400
   End
   Begin VB.Label LB_MT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "São Paulo"
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
      Left            =   5040
      TabIndex        =   333
      Top             =   13320
      Width           =   900
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF"
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
      Index           =   61
      Left            =   8760
      TabIndex        =   332
      Top             =   13200
      Width           =   165
   End
   Begin VB.Label LB_UT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
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
      Left            =   8760
      TabIndex        =   331
      Top             =   13320
      Width           =   270
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INSCRIÇÃO ESTADUAL:"
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
      Index           =   60
      Left            =   9600
      TabIndex        =   330
      Top             =   13200
      Width           =   1335
   End
   Begin VB.Label LB_IT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "111.502.963.110"
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
      TabIndex        =   329
      Top             =   13320
      Width           =   1440
   End
   Begin VB.Line L2 
      Index           =   5
      X1              =   120
      X2              =   11520
      Y1              =   13560
      Y2              =   13560
   End
   Begin VB.Line L1 
      Index           =   7
      X1              =   120
      X2              =   11520
      Y1              =   12720
      Y2              =   12720
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSPORTADOR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   59
      Left            =   120
      TabIndex        =   328
      Top             =   12600
      Width           =   1110
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOME / RAZÃO SOCIAL:"
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
      Index           =   58
      Left            =   120
      TabIndex        =   327
      Top             =   12720
      Width           =   1335
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FRETE POR CONTA:"
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
      Index           =   57
      Left            =   5040
      TabIndex        =   326
      Top             =   12720
      Width           =   1275
   End
   Begin VB.Label LB_TR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
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
      Left            =   120
      TabIndex        =   325
      Top             =   12840
      Width           =   2970
   End
   Begin VB.Label LB_FT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remetente"
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
      Left            =   5040
      TabIndex        =   324
      Top             =   12840
      Width           =   945
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PLACA DO VEÍCULO:"
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
      Index           =   56
      Left            =   6600
      TabIndex        =   323
      Top             =   12720
      Width           =   1275
   End
   Begin VB.Label LB_PL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CIB 8192"
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
      Left            =   6600
      TabIndex        =   322
      Top             =   12840
      Width           =   795
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF:"
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
      Index           =   55
      Left            =   8760
      TabIndex        =   321
      Top             =   12720
      Width           =   195
   End
   Begin VB.Label PB_VP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
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
      Left            =   8760
      TabIndex        =   320
      Top             =   12840
      Width           =   270
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.N.P.J. / C.P.F.:"
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
      Index           =   54
      Left            =   9600
      TabIndex        =   319
      Top             =   12720
      Width           =   885
   End
   Begin VB.Label LB_CT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "55.783.427/0001-03"
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
      TabIndex        =   318
      Top             =   12840
      Width           =   1710
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR DO FRETE:"
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
      Index           =   53
      Left            =   120
      TabIndex        =   317
      Top             =   12000
      Width           =   1050
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR DO SEGURO:"
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
      Index           =   52
      Left            =   2400
      TabIndex        =   316
      Top             =   12000
      Width           =   1170
   End
   Begin VB.Label LB_VF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   120
      TabIndex        =   315
      Top             =   12120
      Width           =   690
   End
   Begin VB.Label LB_VG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   2400
      TabIndex        =   314
      Top             =   12120
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OUTRAS DESPESAS ACESSÓRIAS:"
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
      Index           =   51
      Left            =   3960
      TabIndex        =   313
      Top             =   12000
      Width           =   1950
   End
   Begin VB.Label LB_VO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   3960
      TabIndex        =   312
      Top             =   12120
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR TOTAL DO I.P.I.:"
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
      Index           =   50
      Left            =   7080
      TabIndex        =   311
      Top             =   12000
      Width           =   1290
   End
   Begin VB.Label LB_VP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      TabIndex        =   310
      Top             =   12120
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR TOTAL DA NOTA:"
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
      Index           =   43
      Left            =   9480
      TabIndex        =   309
      Top             =   12000
      Width           =   1380
   End
   Begin VB.Label LB_VN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   9480
      TabIndex        =   308
      Top             =   12120
      Width           =   690
   End
   Begin VB.Line L2 
      Index           =   4
      X1              =   120
      X2              =   11520
      Y1              =   12360
      Y2              =   12360
   End
   Begin VB.Line L1 
      Index           =   6
      X1              =   120
      X2              =   11520
      Y1              =   11520
      Y2              =   11520
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CÁLCULO DO IMPOSTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   49
      Left            =   120
      TabIndex        =   307
      Top             =   11400
      Width           =   1380
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BASE DE CÁLCULO I.C.M.S.:"
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
      Index           =   48
      Left            =   120
      TabIndex        =   306
      Top             =   11520
      Width           =   1560
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR I.C.M.S.:"
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
      Index           =   47
      Left            =   2400
      TabIndex        =   305
      Top             =   11520
      Width           =   870
   End
   Begin VB.Label LB_BI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   120
      TabIndex        =   304
      Top             =   11640
      Width           =   690
   End
   Begin VB.Label LB_VM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   2400
      TabIndex        =   303
      Top             =   11640
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BASE DE CÁLCULO I.C.M.S. SUBSTITUIÇÃO:"
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
      Index           =   46
      Left            =   3960
      TabIndex        =   302
      Top             =   11520
      Width           =   2460
   End
   Begin VB.Label LB_BS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   3960
      TabIndex        =   301
      Top             =   11640
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR I.C.M.S. SUBSTITUIÇÃO:"
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
      Index           =   45
      Left            =   7080
      TabIndex        =   300
      Top             =   11520
      Width           =   1770
   End
   Begin VB.Label LB_VS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      TabIndex        =   299
      Top             =   11640
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR TOTAL DOS PRODUTOS:"
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
      Index           =   44
      Left            =   9480
      TabIndex        =   298
      Top             =   11520
      Width           =   1815
   End
   Begin VB.Label LB_VT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   9480
      TabIndex        =   297
      Top             =   11640
      Width           =   690
   End
   Begin VB.Line L1 
      Index           =   5
      X1              =   120
      X2              =   11520
      Y1              =   11160
      Y2              =   11160
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   219
      Left            =   10680
      TabIndex        =   296
      Top             =   10920
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   218
      Left            =   10680
      TabIndex        =   295
      Top             =   10680
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   217
      Left            =   10680
      TabIndex        =   294
      Top             =   10440
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   216
      Left            =   10680
      TabIndex        =   293
      Top             =   10200
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   215
      Left            =   10680
      TabIndex        =   292
      Top             =   9960
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   214
      Left            =   10680
      TabIndex        =   291
      Top             =   9720
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   213
      Left            =   10680
      TabIndex        =   290
      Top             =   9480
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   212
      Left            =   10680
      TabIndex        =   289
      Top             =   9240
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   211
      Left            =   10680
      TabIndex        =   288
      Top             =   9000
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   210
      Left            =   10680
      TabIndex        =   287
      Top             =   8760
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   209
      Left            =   10680
      TabIndex        =   286
      Top             =   8520
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   208
      Left            =   10680
      TabIndex        =   285
      Top             =   8280
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   207
      Left            =   10680
      TabIndex        =   284
      Top             =   8040
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   206
      Left            =   10680
      TabIndex        =   283
      Top             =   7800
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   205
      Left            =   10680
      TabIndex        =   282
      Top             =   7560
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   204
      Left            =   10680
      TabIndex        =   281
      Top             =   7320
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   203
      Left            =   10680
      TabIndex        =   280
      Top             =   7080
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   202
      Left            =   10680
      TabIndex        =   279
      Top             =   6840
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   201
      Left            =   10680
      TabIndex        =   278
      Top             =   6600
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   200
      Left            =   10680
      TabIndex        =   277
      Top             =   6360
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR IPI:"
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
      Index           =   42
      Left            =   10680
      TabIndex        =   276
      Top             =   6120
      Width           =   585
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   199
      Left            =   10320
      TabIndex        =   275
      Top             =   10920
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   198
      Left            =   10320
      TabIndex        =   274
      Top             =   10680
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   197
      Left            =   10320
      TabIndex        =   273
      Top             =   10440
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   196
      Left            =   10320
      TabIndex        =   272
      Top             =   10200
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   195
      Left            =   10320
      TabIndex        =   271
      Top             =   9960
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   194
      Left            =   10320
      TabIndex        =   270
      Top             =   9720
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   193
      Left            =   10320
      TabIndex        =   269
      Top             =   9480
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   192
      Left            =   10320
      TabIndex        =   268
      Top             =   9240
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   191
      Left            =   10320
      TabIndex        =   267
      Top             =   9000
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   190
      Left            =   10320
      TabIndex        =   266
      Top             =   8760
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   189
      Left            =   10320
      TabIndex        =   265
      Top             =   8520
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   188
      Left            =   10320
      TabIndex        =   264
      Top             =   8280
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   187
      Left            =   10320
      TabIndex        =   263
      Top             =   8040
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   186
      Left            =   10320
      TabIndex        =   262
      Top             =   7800
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   185
      Left            =   10320
      TabIndex        =   261
      Top             =   7560
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   184
      Left            =   10320
      TabIndex        =   260
      Top             =   7320
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   183
      Left            =   10320
      TabIndex        =   259
      Top             =   7080
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   182
      Left            =   10320
      TabIndex        =   258
      Top             =   6840
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   181
      Left            =   10320
      TabIndex        =   257
      Top             =   6600
      Width           =   105
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
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
      Index           =   180
      Left            =   10320
      TabIndex        =   256
      Top             =   6360
      Width           =   105
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%IPI"
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
      Index           =   41
      Left            =   10200
      TabIndex        =   255
      Top             =   6120
      Width           =   240
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   179
      Left            =   9720
      TabIndex        =   254
      Top             =   10920
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   178
      Left            =   9720
      TabIndex        =   253
      Top             =   10680
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   177
      Left            =   9720
      TabIndex        =   252
      Top             =   10440
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   176
      Left            =   9720
      TabIndex        =   251
      Top             =   10200
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   175
      Left            =   9720
      TabIndex        =   250
      Top             =   9960
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   174
      Left            =   9720
      TabIndex        =   249
      Top             =   9720
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   173
      Left            =   9720
      TabIndex        =   248
      Top             =   9480
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   172
      Left            =   9720
      TabIndex        =   247
      Top             =   9240
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   171
      Left            =   9720
      TabIndex        =   246
      Top             =   9000
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   170
      Left            =   9720
      TabIndex        =   245
      Top             =   8760
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   169
      Left            =   9720
      TabIndex        =   244
      Top             =   8520
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   168
      Left            =   9720
      TabIndex        =   243
      Top             =   8280
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   167
      Left            =   9720
      TabIndex        =   242
      Top             =   8040
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   166
      Left            =   9720
      TabIndex        =   241
      Top             =   7800
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   165
      Left            =   9720
      TabIndex        =   240
      Top             =   7560
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   164
      Left            =   9720
      TabIndex        =   239
      Top             =   7320
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   163
      Left            =   9720
      TabIndex        =   238
      Top             =   7080
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   162
      Left            =   9720
      TabIndex        =   237
      Top             =   6840
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   161
      Left            =   9720
      TabIndex        =   236
      Top             =   6600
      Width           =   210
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
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
      Index           =   160
      Left            =   9720
      TabIndex        =   235
      Top             =   6360
      Width           =   210
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%ICMS"
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
      Index           =   40
      Left            =   9600
      TabIndex        =   234
      Top             =   6120
      Width           =   405
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   159
      Left            =   8520
      TabIndex        =   233
      Top             =   10920
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   158
      Left            =   8520
      TabIndex        =   232
      Top             =   10680
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   157
      Left            =   8520
      TabIndex        =   231
      Top             =   10440
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   156
      Left            =   8520
      TabIndex        =   230
      Top             =   10200
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   155
      Left            =   8520
      TabIndex        =   229
      Top             =   9960
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   154
      Left            =   8520
      TabIndex        =   228
      Top             =   9720
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   153
      Left            =   8520
      TabIndex        =   227
      Top             =   9480
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   152
      Left            =   8520
      TabIndex        =   226
      Top             =   9240
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   151
      Left            =   8520
      TabIndex        =   225
      Top             =   9000
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   150
      Left            =   8520
      TabIndex        =   224
      Top             =   8760
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   149
      Left            =   8520
      TabIndex        =   223
      Top             =   8520
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   148
      Left            =   8520
      TabIndex        =   222
      Top             =   8280
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   147
      Left            =   8520
      TabIndex        =   221
      Top             =   8040
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   146
      Left            =   8520
      TabIndex        =   220
      Top             =   7800
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   145
      Left            =   8520
      TabIndex        =   219
      Top             =   7560
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   144
      Left            =   8520
      TabIndex        =   218
      Top             =   7320
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   143
      Left            =   8520
      TabIndex        =   217
      Top             =   7080
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   142
      Left            =   8520
      TabIndex        =   216
      Top             =   6840
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   141
      Left            =   8520
      TabIndex        =   215
      Top             =   6600
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   140
      Left            =   8520
      TabIndex        =   214
      Top             =   6360
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PREÇO TOTAL:"
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
      Index           =   39
      Left            =   8520
      TabIndex        =   213
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   139
      Left            =   7320
      TabIndex        =   212
      Top             =   10920
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   138
      Left            =   7320
      TabIndex        =   211
      Top             =   10680
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   137
      Left            =   7320
      TabIndex        =   210
      Top             =   10440
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   136
      Left            =   7320
      TabIndex        =   209
      Top             =   10200
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   135
      Left            =   7320
      TabIndex        =   208
      Top             =   9960
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   134
      Left            =   7320
      TabIndex        =   207
      Top             =   9720
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   133
      Left            =   7320
      TabIndex        =   206
      Top             =   9480
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   132
      Left            =   7320
      TabIndex        =   205
      Top             =   9240
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   131
      Left            =   7320
      TabIndex        =   204
      Top             =   9000
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   130
      Left            =   7320
      TabIndex        =   203
      Top             =   8760
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   129
      Left            =   7320
      TabIndex        =   202
      Top             =   8520
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   128
      Left            =   7320
      TabIndex        =   201
      Top             =   8280
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   127
      Left            =   7320
      TabIndex        =   200
      Top             =   8040
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   126
      Left            =   7320
      TabIndex        =   199
      Top             =   7800
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   125
      Left            =   7320
      TabIndex        =   198
      Top             =   7560
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   124
      Left            =   7320
      TabIndex        =   197
      Top             =   7320
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   123
      Left            =   7320
      TabIndex        =   196
      Top             =   7080
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   122
      Left            =   7320
      TabIndex        =   195
      Top             =   6840
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   121
      Left            =   7320
      TabIndex        =   194
      Top             =   6600
      Width           =   690
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Index           =   120
      Left            =   7320
      TabIndex        =   193
      Top             =   6360
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PREÇO UNITÁRIO:"
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
      Index           =   38
      Left            =   7320
      TabIndex        =   192
      Top             =   6120
      Width           =   1050
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   119
      Left            =   6720
      TabIndex        =   191
      Top             =   10920
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   118
      Left            =   6720
      TabIndex        =   190
      Top             =   10680
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   117
      Left            =   6720
      TabIndex        =   189
      Top             =   10440
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   116
      Left            =   6720
      TabIndex        =   188
      Top             =   10200
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   115
      Left            =   6720
      TabIndex        =   187
      Top             =   9960
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   114
      Left            =   6720
      TabIndex        =   186
      Top             =   9720
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   113
      Left            =   6720
      TabIndex        =   185
      Top             =   9480
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   112
      Left            =   6720
      TabIndex        =   184
      Top             =   9240
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   111
      Left            =   6720
      TabIndex        =   183
      Top             =   9000
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   110
      Left            =   6720
      TabIndex        =   182
      Top             =   8760
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   109
      Left            =   6720
      TabIndex        =   181
      Top             =   8520
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   108
      Left            =   6720
      TabIndex        =   180
      Top             =   8280
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   107
      Left            =   6720
      TabIndex        =   179
      Top             =   8040
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   106
      Left            =   6720
      TabIndex        =   178
      Top             =   7800
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   105
      Left            =   6720
      TabIndex        =   177
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   104
      Left            =   6720
      TabIndex        =   176
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   103
      Left            =   6720
      TabIndex        =   175
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   102
      Left            =   6720
      TabIndex        =   174
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   101
      Left            =   6720
      TabIndex        =   173
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Index           =   100
      Left            =   6720
      TabIndex        =   172
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANT.:"
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
      Index           =   37
      Left            =   6720
      TabIndex        =   171
      Top             =   6120
      Width           =   480
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   99
      Left            =   6120
      TabIndex        =   170
      Top             =   10920
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   98
      Left            =   6120
      TabIndex        =   169
      Top             =   10680
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   97
      Left            =   6120
      TabIndex        =   168
      Top             =   10440
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   96
      Left            =   6120
      TabIndex        =   167
      Top             =   10200
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   95
      Left            =   6120
      TabIndex        =   166
      Top             =   9960
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   94
      Left            =   6120
      TabIndex        =   165
      Top             =   9720
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   93
      Left            =   6120
      TabIndex        =   164
      Top             =   9480
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   92
      Left            =   6120
      TabIndex        =   163
      Top             =   9240
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   91
      Left            =   6120
      TabIndex        =   162
      Top             =   9000
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   90
      Left            =   6120
      TabIndex        =   161
      Top             =   8760
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   89
      Left            =   6120
      TabIndex        =   160
      Top             =   8520
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   88
      Left            =   6120
      TabIndex        =   159
      Top             =   8280
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   87
      Left            =   6120
      TabIndex        =   158
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   86
      Left            =   6120
      TabIndex        =   157
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   85
      Left            =   6120
      TabIndex        =   156
      Top             =   7560
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   84
      Left            =   6120
      TabIndex        =   155
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   83
      Left            =   6120
      TabIndex        =   154
      Top             =   7080
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   82
      Left            =   6120
      TabIndex        =   153
      Top             =   6840
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   81
      Left            =   6120
      TabIndex        =   152
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pçs."
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
      Index           =   80
      Left            =   6120
      TabIndex        =   151
      Top             =   6360
      Width           =   375
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UNID.:"
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
      Index           =   36
      Left            =   6120
      TabIndex        =   150
      Top             =   6120
      Width           =   360
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   79
      Left            =   5520
      TabIndex        =   149
      Top             =   10920
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   78
      Left            =   5520
      TabIndex        =   148
      Top             =   10680
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   77
      Left            =   5520
      TabIndex        =   147
      Top             =   10440
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   76
      Left            =   5520
      TabIndex        =   146
      Top             =   10200
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   75
      Left            =   5520
      TabIndex        =   145
      Top             =   9960
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   74
      Left            =   5520
      TabIndex        =   144
      Top             =   9720
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   73
      Left            =   5520
      TabIndex        =   143
      Top             =   9480
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   72
      Left            =   5520
      TabIndex        =   142
      Top             =   9240
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   71
      Left            =   5520
      TabIndex        =   141
      Top             =   9000
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   70
      Left            =   5520
      TabIndex        =   140
      Top             =   8760
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   69
      Left            =   5520
      TabIndex        =   139
      Top             =   8520
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   68
      Left            =   5520
      TabIndex        =   138
      Top             =   8280
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   67
      Left            =   5520
      TabIndex        =   137
      Top             =   8040
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   66
      Left            =   5520
      TabIndex        =   136
      Top             =   7800
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   65
      Left            =   5520
      TabIndex        =   135
      Top             =   7560
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   64
      Left            =   5520
      TabIndex        =   134
      Top             =   7320
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   63
      Left            =   5520
      TabIndex        =   133
      Top             =   7080
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   62
      Left            =   5520
      TabIndex        =   132
      Top             =   6840
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   61
      Left            =   5520
      TabIndex        =   131
      Top             =   6600
      Width           =   270
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-0"
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
      Index           =   60
      Left            =   5520
      TabIndex        =   130
      Top             =   6360
      Width           =   270
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S.T.:"
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
      Index           =   35
      Left            =   5520
      TabIndex        =   129
      Top             =   6120
      Width           =   240
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   59
      Left            =   5040
      TabIndex        =   128
      Top             =   10920
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   58
      Left            =   5040
      TabIndex        =   127
      Top             =   10680
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   57
      Left            =   5040
      TabIndex        =   126
      Top             =   10440
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   56
      Left            =   5040
      TabIndex        =   125
      Top             =   10200
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   55
      Left            =   5040
      TabIndex        =   124
      Top             =   9960
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   54
      Left            =   5040
      TabIndex        =   123
      Top             =   9720
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   53
      Left            =   5040
      TabIndex        =   122
      Top             =   9480
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   52
      Left            =   5040
      TabIndex        =   121
      Top             =   9240
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   51
      Left            =   5040
      TabIndex        =   120
      Top             =   9000
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   50
      Left            =   5040
      TabIndex        =   119
      Top             =   8760
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   49
      Left            =   5040
      TabIndex        =   118
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   48
      Left            =   5040
      TabIndex        =   117
      Top             =   8280
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   47
      Left            =   5040
      TabIndex        =   116
      Top             =   8040
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   46
      Left            =   5040
      TabIndex        =   115
      Top             =   7800
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   45
      Left            =   5040
      TabIndex        =   114
      Top             =   7560
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   44
      Left            =   5040
      TabIndex        =   113
      Top             =   7320
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Index           =   43
      Left            =   5040
      TabIndex        =   112
      Top             =   7080
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Left            =   5040
      TabIndex        =   111
      Top             =   6840
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Left            =   5040
      TabIndex        =   110
      Top             =   6600
      Width           =   135
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Left            =   5040
      TabIndex        =   109
      Top             =   6360
      Width           =   135
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.F.:"
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
      Index           =   34
      Left            =   5040
      TabIndex        =   108
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Left            =   1320
      TabIndex        =   107
      Top             =   10920
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Left            =   1320
      TabIndex        =   106
      Top             =   10680
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Left            =   1320
      TabIndex        =   105
      Top             =   10440
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Left            =   1320
      TabIndex        =   104
      Top             =   10200
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Left            =   1320
      TabIndex        =   103
      Top             =   9960
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      TabIndex        =   102
      Top             =   9720
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Left            =   1320
      TabIndex        =   101
      Top             =   9480
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   32
      Left            =   1320
      TabIndex        =   100
      Top             =   9240
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   31
      Left            =   1320
      TabIndex        =   99
      Top             =   9000
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   30
      Left            =   1320
      TabIndex        =   98
      Top             =   8760
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   29
      Left            =   1320
      TabIndex        =   97
      Top             =   8520
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   28
      Left            =   1320
      TabIndex        =   96
      Top             =   8280
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   27
      Left            =   1320
      TabIndex        =   95
      Top             =   8040
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   26
      Left            =   1320
      TabIndex        =   94
      Top             =   7800
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   25
      Left            =   1320
      TabIndex        =   93
      Top             =   7560
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   24
      Left            =   1320
      TabIndex        =   92
      Top             =   7320
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   23
      Left            =   1320
      TabIndex        =   91
      Top             =   7080
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   22
      Left            =   1320
      TabIndex        =   90
      Top             =   6840
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   21
      Left            =   1320
      TabIndex        =   89
      Top             =   6600
      Width           =   1890
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AAAAAAAAAAAAAA"
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
      Index           =   20
      Left            =   1320
      TabIndex        =   88
      Top             =   6360
      Width           =   1890
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIÇÃO DOS PRODUTOS:"
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
      Index           =   33
      Left            =   1320
      TabIndex        =   87
      Top             =   6120
      Width           =   1725
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Index           =   19
      Left            =   120
      TabIndex        =   86
      Top             =   10920
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   85
      Top             =   10680
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   84
      Top             =   10440
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   83
      Top             =   10200
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   82
      Top             =   9960
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   81
      Top             =   9720
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   80
      Top             =   9480
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   79
      Top             =   9240
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   78
      Top             =   9000
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   77
      Top             =   8760
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   76
      Top             =   8520
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   75
      Top             =   8280
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   74
      Top             =   8040
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   73
      Top             =   7800
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   72
      Top             =   7560
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   71
      Top             =   7320
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   70
      Top             =   7080
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   69
      Top             =   6840
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   68
      Top             =   6600
      Width           =   885
   End
   Begin VB.Label LB_CP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00-ABC00"
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
      Left            =   120
      TabIndex        =   67
      Top             =   6360
      Width           =   885
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C. PROD."
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
      Index           =   32
      Left            =   120
      TabIndex        =   66
      Top             =   6120
      Width           =   525
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DADOS DO PRODUTO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   31
      Left            =   120
      TabIndex        =   65
      Top             =   6000
      Width           =   1275
   End
   Begin VB.Line L1 
      Index           =   4
      X1              =   120
      X2              =   11520
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label LB_SE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   9840
      TabIndex        =   64
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SETOR:"
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
      Index           =   30
      Left            =   9840
      TabIndex        =   63
      Top             =   5400
      Width           =   435
   End
   Begin VB.Label LB_VX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   7800
      TabIndex        =   62
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VEND EXT.:"
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
      Index           =   29
      Left            =   7800
      TabIndex        =   61
      Top             =   5400
      Width           =   645
   End
   Begin VB.Label LB_VI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   5760
      TabIndex        =   60
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VEND INT.:"
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
      Left            =   5760
      TabIndex        =   59
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label LB_OP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   3840
      TabIndex        =   58
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OP.:"
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
      Index           =   27
      Left            =   3840
      TabIndex        =   57
      Top             =   5400
      Width           =   225
   End
   Begin VB.Label LB_SP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   1920
      TabIndex        =   56
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label LB_PI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   120
      TabIndex        =   55
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S/PED.:"
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
      Index           =   26
      Left            =   1920
      TabIndex        =   54
      Top             =   5400
      Width           =   405
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.I.:"
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
      Index           =   25
      Left            =   120
      TabIndex        =   53
      Top             =   5400
      Width           =   195
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DADOS GERAIS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   24
      Left            =   120
      TabIndex        =   52
      Top             =   5280
      Width           =   915
   End
   Begin VB.Line L1 
      Index           =   3
      X1              =   120
      X2              =   11520
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line L2 
      Index           =   3
      X1              =   120
      X2              =   11520
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label LB_EX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hum mil, um real, um centavo"
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
      Left            =   5880
      TabIndex        =   51
      Top             =   4800
      Width           =   2595
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR POR EXTENSO"
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
      Index           =   23
      Left            =   5880
      TabIndex        =   50
      Top             =   4680
      Width           =   1260
   End
   Begin VB.Label LB_PR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O mesmo"
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
      Left            =   120
      TabIndex        =   49
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENDEREÇO DE COBRANÇA / PRAÇA DE PAGAMENTO:"
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
      Index           =   22
      Left            =   120
      TabIndex        =   48
      Top             =   4680
      Width           =   3075
   End
   Begin VB.Label LB_VE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
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
      Left            =   7320
      TabIndex        =   47
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VENCIMENTO:"
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
      Index           =   21
      Left            =   7320
      TabIndex        =   46
      Top             =   4200
      Width           =   825
   End
   Begin VB.Label LB_DP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   5400
      TabIndex        =   45
      Top             =   4320
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº DA DUPLICATA"
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
      Index           =   20
      Left            =   5400
      TabIndex        =   44
      Top             =   4200
      Width           =   1020
   End
   Begin VB.Label LB_FA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
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
      Left            =   1560
      TabIndex        =   43
      Top             =   4320
      Width           =   630
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº DA NOTA FISCAL-FATURA:"
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
      Index           =   19
      Left            =   1560
      TabIndex        =   42
      Top             =   4200
      Width           =   1680
   End
   Begin VB.Label LB_VA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000,00"
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
      Left            =   4080
      TabIndex        =   41
      Top             =   4320
      Width           =   690
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR:"
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
      Index           =   18
      Left            =   4080
      TabIndex        =   40
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label LB_FE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/0000"
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
      Left            =   120
      TabIndex        =   39
      Top             =   4320
      Width           =   960
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA EMISSÃO:"
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
      Index           =   17
      Left            =   120
      TabIndex        =   38
      Top             =   4200
      Width           =   900
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FATURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   16
      Left            =   120
      TabIndex        =   37
      Top             =   4080
      Width           =   510
   End
   Begin VB.Line L1 
      Index           =   2
      X1              =   120
      X2              =   11520
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line L2 
      Index           =   2
      X1              =   120
      X2              =   11520
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label LB_HS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
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
      Left            =   10560
      TabIndex        =   36
      Top             =   2400
      Width           =   750
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA SAÍDA:"
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
      Index           =   15
      Left            =   10560
      TabIndex        =   35
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label LB_DS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
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
      Left            =   8520
      TabIndex        =   34
      Top             =   2400
      Width           =   960
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA SAÍDA / ENTRADA:"
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
      Index           =   14
      Left            =   8520
      TabIndex        =   33
      Top             =   2280
      Width           =   1380
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
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
      Left            =   7200
      TabIndex        =   32
      Top             =   2400
      Width           =   960
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA EMISSÃO:"
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
      Index           =   13
      Left            =   7200
      TabIndex        =   31
      Top             =   2280
      Width           =   900
   End
   Begin VB.Label LB_UF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SP"
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
      Left            =   10920
      TabIndex        =   30
      Top             =   3600
      Width           =   270
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UF:"
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
      Index           =   12
      Left            =   10920
      TabIndex        =   29
      Top             =   3480
      Width           =   195
   End
   Begin VB.Label LB_MU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "São Paulo"
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
      Left            =   9000
      TabIndex        =   28
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUNICÍPIO:"
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
      Index           =   11
      Left            =   9000
      TabIndex        =   27
      Top             =   3480
      Width           =   660
   End
   Begin VB.Label LB_CE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "03371-000"
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
      TabIndex        =   26
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP:"
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
      Index           =   10
      Left            =   7560
      TabIndex        =   25
      Top             =   3480
      Width           =   270
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vila Formosa"
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
      Left            =   5400
      TabIndex        =   24
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BAIRRO:"
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
      Left            =   5400
      TabIndex        =   23
      Top             =   3480
      Width           =   480
   End
   Begin VB.Label LB_EN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avenida Montemagno, 2454"
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
      Left            =   120
      TabIndex        =   22
      Top             =   3600
      Width           =   2400
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ENDEREÇO:"
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
      Left            =   120
      TabIndex        =   21
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label LB_IE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "111.502.963.110"
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
      Left            =   9240
      TabIndex        =   20
      Top             =   3120
      Width           =   1440
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INSCR. ESTADUAL:"
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
      Left            =   9240
      TabIndex        =   19
      Top             =   3000
      Width           =   1080
   End
   Begin VB.Label LB_CN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "55.783.427/0001-03"
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
      Left            =   6720
      TabIndex        =   18
      Top             =   3120
      Width           =   1710
   End
   Begin VB.Label LB_RS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
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
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   2970
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.N.P.J. / C.P.F.:"
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
      Left            =   6720
      TabIndex        =   16
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZÃO SOCIAL:"
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
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESTINATÁRIO / REMETENTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   1725
   End
   Begin VB.Line L1 
      Index           =   1
      X1              =   120
      X2              =   11520
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line L2 
      Index           =   1
      X1              =   120
      X2              =   11520
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label LB_CF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "5.11"
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
      Left            =   6000
      TabIndex        =   13
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label LB_NO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Venda para Comercialização"
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
      Left            =   1680
      TabIndex        =   12
      Top             =   2400
      Width           =   2490
   End
   Begin VB.Label LB_TP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada"
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
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CFOP"
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
      Index           =   3
      Left            =   6000
      TabIndex        =   10
      Top             =   2280
      Width           =   330
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NATUREZA DA OPERAÇÃO:"
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
      Index           =   2
      Left            =   1680
      TabIndex        =   9
      Top             =   2280
      Width           =   1560
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   300
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOTA FISCAL - FATURA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1425
   End
   Begin VB.Line L1 
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line L2 
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label LB_NNF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8400
      TabIndex        =   6
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "São Paulo - (SP) - Fone: (011) 6910-1444"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   165
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   4080
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Av.Montemagno, 2.454 - Vila Formosa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   164
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   3825
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   3060
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL"
      BeginProperty Font 
         Name            =   "Futura Md BT"
         Size            =   19.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   2265
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   11520
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   1
      X1              =   6000
      X2              =   6000
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório Completo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   7560
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "da Nota Fiscal Emitida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   7320
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "Tela_Relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
