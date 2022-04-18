VERSION 5.00
Begin VB.Form Tela_RelSimp 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   15195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   15195
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
      Picture         =   "Tela_RelSimp.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   412
      Top             =   600
      Width           =   600
   End
   Begin VB.Label LB_ME1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Critério:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6120
      TabIndex        =   411
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label LB_ME2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Todas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   6840
      TabIndex        =   410
      Top             =   1440
      Width           =   465
   End
   Begin VB.Label LB_DA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9120
      TabIndex        =   409
      Top             =   1080
      Width           =   510
   End
   Begin VB.Label LB_DA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00/00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9720
      TabIndex        =   408
      Top             =   1080
      Width           =   840
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   3
      X1              =   120
      X2              =   11520
      Y1              =   15120
      Y2              =   15120
   End
   Begin VB.Label LB_FO2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00/00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8040
      TabIndex        =   407
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label LB_FO1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folhas:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7200
      TabIndex        =   406
      Top             =   1080
      Width           =   705
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   2160
      TabIndex        =   405
      Top             =   13800
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   3000
      TabIndex        =   404
      Top             =   13800
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   7080
      TabIndex        =   403
      Top             =   13800
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   8280
      TabIndex        =   402
      Top             =   13800
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   9720
      TabIndex        =   401
      Top             =   13800
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   10320
      TabIndex        =   400
      Top             =   13800
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   1200
      TabIndex        =   399
      Top             =   13800
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   120
      TabIndex        =   398
      Top             =   14040
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   1320
      TabIndex        =   397
      Top             =   14040
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   2520
      TabIndex        =   396
      Top             =   14040
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   3240
      TabIndex        =   395
      Top             =   14040
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   120
      TabIndex        =   394
      Top             =   14280
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   1320
      TabIndex        =   393
      Top             =   14280
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   2520
      TabIndex        =   392
      Top             =   14280
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   3240
      TabIndex        =   391
      Top             =   14280
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   120
      TabIndex        =   390
      Top             =   14520
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   1320
      TabIndex        =   389
      Top             =   14520
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   2520
      TabIndex        =   388
      Top             =   14520
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   3240
      TabIndex        =   387
      Top             =   14520
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   120
      TabIndex        =   386
      Top             =   14760
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   1320
      TabIndex        =   385
      Top             =   14760
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   2520
      TabIndex        =   384
      Top             =   14760
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   3240
      TabIndex        =   383
      Top             =   14760
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   4560
      TabIndex        =   382
      Top             =   14280
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   4560
      TabIndex        =   381
      Top             =   14760
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   4560
      TabIndex        =   380
      Top             =   14520
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   4560
      TabIndex        =   379
      Top             =   14040
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   6480
      TabIndex        =   378
      Top             =   14040
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   6480
      TabIndex        =   377
      Top             =   14280
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   6480
      TabIndex        =   376
      Top             =   14520
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   6480
      TabIndex        =   375
      Top             =   14760
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   9
      X1              =   120
      X2              =   11520
      Y1              =   13800
      Y2              =   13800
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   120
      TabIndex        =   374
      Top             =   13800
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   9600
      TabIndex        =   373
      Top             =   14760
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   9600
      TabIndex        =   372
      Top             =   14520
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   9600
      TabIndex        =   371
      Top             =   14040
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   8160
      TabIndex        =   370
      Top             =   14040
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   8160
      TabIndex        =   369
      Top             =   14520
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   8160
      TabIndex        =   368
      Top             =   14760
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   8160
      TabIndex        =   367
      Top             =   14280
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   9
      X1              =   120
      X2              =   11520
      Y1              =   14040
      Y2              =   14040
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   9
      Left            =   9600
      TabIndex        =   366
      Top             =   14280
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   2160
      TabIndex        =   365
      Top             =   12480
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   3000
      TabIndex        =   364
      Top             =   12480
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   7080
      TabIndex        =   363
      Top             =   12480
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   8280
      TabIndex        =   362
      Top             =   12480
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   9720
      TabIndex        =   361
      Top             =   12480
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   10320
      TabIndex        =   360
      Top             =   12480
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   1200
      TabIndex        =   359
      Top             =   12480
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   358
      Top             =   12720
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   1320
      TabIndex        =   357
      Top             =   12720
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   2520
      TabIndex        =   356
      Top             =   12720
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   3240
      TabIndex        =   355
      Top             =   12720
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   354
      Top             =   12960
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   1320
      TabIndex        =   353
      Top             =   12960
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   2520
      TabIndex        =   352
      Top             =   12960
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   3240
      TabIndex        =   351
      Top             =   12960
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   350
      Top             =   13200
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   1320
      TabIndex        =   349
      Top             =   13200
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   2520
      TabIndex        =   348
      Top             =   13200
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   3240
      TabIndex        =   347
      Top             =   13200
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   346
      Top             =   13440
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   1320
      TabIndex        =   345
      Top             =   13440
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   2520
      TabIndex        =   344
      Top             =   13440
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   3240
      TabIndex        =   343
      Top             =   13440
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   4560
      TabIndex        =   342
      Top             =   12960
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   4560
      TabIndex        =   341
      Top             =   13440
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   4560
      TabIndex        =   340
      Top             =   13200
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   4560
      TabIndex        =   339
      Top             =   12720
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   6480
      TabIndex        =   338
      Top             =   12720
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   6480
      TabIndex        =   337
      Top             =   12960
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   6480
      TabIndex        =   336
      Top             =   13200
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   6480
      TabIndex        =   335
      Top             =   13440
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   8
      X1              =   120
      X2              =   11520
      Y1              =   12480
      Y2              =   12480
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   120
      TabIndex        =   334
      Top             =   12480
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   9600
      TabIndex        =   333
      Top             =   13440
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   9600
      TabIndex        =   332
      Top             =   13200
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   9600
      TabIndex        =   331
      Top             =   12720
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   8160
      TabIndex        =   330
      Top             =   12720
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   8160
      TabIndex        =   329
      Top             =   13200
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   8160
      TabIndex        =   328
      Top             =   13440
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   8160
      TabIndex        =   327
      Top             =   12960
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   8
      X1              =   120
      X2              =   11520
      Y1              =   12720
      Y2              =   12720
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   8
      Left            =   9600
      TabIndex        =   326
      Top             =   12960
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   2160
      TabIndex        =   325
      Top             =   11160
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   3000
      TabIndex        =   324
      Top             =   11160
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   7080
      TabIndex        =   323
      Top             =   11160
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   8280
      TabIndex        =   322
      Top             =   11160
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   9720
      TabIndex        =   321
      Top             =   11160
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   10320
      TabIndex        =   320
      Top             =   11160
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   1200
      TabIndex        =   319
      Top             =   11160
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   120
      TabIndex        =   318
      Top             =   11400
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   1320
      TabIndex        =   317
      Top             =   11400
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   2520
      TabIndex        =   316
      Top             =   11400
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   3240
      TabIndex        =   315
      Top             =   11400
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   120
      TabIndex        =   314
      Top             =   11640
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   1320
      TabIndex        =   313
      Top             =   11640
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   2520
      TabIndex        =   312
      Top             =   11640
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   3240
      TabIndex        =   311
      Top             =   11640
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   120
      TabIndex        =   310
      Top             =   11880
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   1320
      TabIndex        =   309
      Top             =   11880
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   2520
      TabIndex        =   308
      Top             =   11880
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   3240
      TabIndex        =   307
      Top             =   11880
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   120
      TabIndex        =   306
      Top             =   12120
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   1320
      TabIndex        =   305
      Top             =   12120
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   2520
      TabIndex        =   304
      Top             =   12120
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   3240
      TabIndex        =   303
      Top             =   12120
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   4560
      TabIndex        =   302
      Top             =   11640
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   4560
      TabIndex        =   301
      Top             =   12120
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   4560
      TabIndex        =   300
      Top             =   11880
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   4560
      TabIndex        =   299
      Top             =   11400
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   6480
      TabIndex        =   298
      Top             =   11400
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   6480
      TabIndex        =   297
      Top             =   11640
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   6480
      TabIndex        =   296
      Top             =   11880
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   6480
      TabIndex        =   295
      Top             =   12120
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   7
      X1              =   120
      X2              =   11520
      Y1              =   11160
      Y2              =   11160
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   120
      TabIndex        =   294
      Top             =   11160
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   9600
      TabIndex        =   293
      Top             =   12120
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   9600
      TabIndex        =   292
      Top             =   11880
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   9600
      TabIndex        =   291
      Top             =   11400
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   8160
      TabIndex        =   290
      Top             =   11400
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   8160
      TabIndex        =   289
      Top             =   11880
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   8160
      TabIndex        =   288
      Top             =   12120
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   8160
      TabIndex        =   287
      Top             =   11640
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   7
      X1              =   120
      X2              =   11520
      Y1              =   11400
      Y2              =   11400
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   7
      Left            =   9600
      TabIndex        =   286
      Top             =   11640
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   2160
      TabIndex        =   285
      Top             =   9840
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   3000
      TabIndex        =   284
      Top             =   9840
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   7080
      TabIndex        =   283
      Top             =   9840
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   8280
      TabIndex        =   282
      Top             =   9840
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   9720
      TabIndex        =   281
      Top             =   9840
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   10320
      TabIndex        =   280
      Top             =   9840
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1200
      TabIndex        =   279
      Top             =   9840
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   278
      Top             =   10080
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1320
      TabIndex        =   277
      Top             =   10080
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   2520
      TabIndex        =   276
      Top             =   10080
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   3240
      TabIndex        =   275
      Top             =   10080
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   274
      Top             =   10320
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1320
      TabIndex        =   273
      Top             =   10320
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   2520
      TabIndex        =   272
      Top             =   10320
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   3240
      TabIndex        =   271
      Top             =   10320
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   270
      Top             =   10560
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1320
      TabIndex        =   269
      Top             =   10560
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   2520
      TabIndex        =   268
      Top             =   10560
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   3240
      TabIndex        =   267
      Top             =   10560
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   266
      Top             =   10800
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   1320
      TabIndex        =   265
      Top             =   10800
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   2520
      TabIndex        =   264
      Top             =   10800
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   3240
      TabIndex        =   263
      Top             =   10800
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   4560
      TabIndex        =   262
      Top             =   10320
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   4560
      TabIndex        =   261
      Top             =   10800
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   4560
      TabIndex        =   260
      Top             =   10560
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   4560
      TabIndex        =   259
      Top             =   10080
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   6480
      TabIndex        =   258
      Top             =   10080
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   6480
      TabIndex        =   257
      Top             =   10320
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   6480
      TabIndex        =   256
      Top             =   10560
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   6480
      TabIndex        =   255
      Top             =   10800
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   6
      X1              =   120
      X2              =   11520
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   120
      TabIndex        =   254
      Top             =   9840
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   9600
      TabIndex        =   253
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   9600
      TabIndex        =   252
      Top             =   10560
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   9600
      TabIndex        =   251
      Top             =   10080
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   8160
      TabIndex        =   250
      Top             =   10080
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   8160
      TabIndex        =   249
      Top             =   10560
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   8160
      TabIndex        =   248
      Top             =   10800
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   8160
      TabIndex        =   247
      Top             =   10320
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   6
      X1              =   120
      X2              =   11520
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   6
      Left            =   9600
      TabIndex        =   246
      Top             =   10320
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   2160
      TabIndex        =   245
      Top             =   8520
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   3000
      TabIndex        =   244
      Top             =   8520
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   7080
      TabIndex        =   243
      Top             =   8520
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   8280
      TabIndex        =   242
      Top             =   8520
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   9720
      TabIndex        =   241
      Top             =   8520
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   10320
      TabIndex        =   240
      Top             =   8520
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   1200
      TabIndex        =   239
      Top             =   8520
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   120
      TabIndex        =   238
      Top             =   8760
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   1320
      TabIndex        =   237
      Top             =   8760
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   2520
      TabIndex        =   236
      Top             =   8760
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   3240
      TabIndex        =   235
      Top             =   8760
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   120
      TabIndex        =   234
      Top             =   9000
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   1320
      TabIndex        =   233
      Top             =   9000
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   2520
      TabIndex        =   232
      Top             =   9000
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   3240
      TabIndex        =   231
      Top             =   9000
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   120
      TabIndex        =   230
      Top             =   9240
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   1320
      TabIndex        =   229
      Top             =   9240
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   2520
      TabIndex        =   228
      Top             =   9240
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   3240
      TabIndex        =   227
      Top             =   9240
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   120
      TabIndex        =   226
      Top             =   9480
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   1320
      TabIndex        =   225
      Top             =   9480
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   2520
      TabIndex        =   224
      Top             =   9480
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   3240
      TabIndex        =   223
      Top             =   9480
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   4560
      TabIndex        =   222
      Top             =   9000
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   4560
      TabIndex        =   221
      Top             =   9480
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   4560
      TabIndex        =   220
      Top             =   9240
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   4560
      TabIndex        =   219
      Top             =   8760
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6480
      TabIndex        =   218
      Top             =   8760
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6480
      TabIndex        =   217
      Top             =   9000
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6480
      TabIndex        =   216
      Top             =   9240
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6480
      TabIndex        =   215
      Top             =   9480
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   5
      X1              =   120
      X2              =   11520
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   120
      TabIndex        =   214
      Top             =   8520
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   9600
      TabIndex        =   213
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   9600
      TabIndex        =   212
      Top             =   9240
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   9600
      TabIndex        =   211
      Top             =   8760
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   8160
      TabIndex        =   210
      Top             =   8760
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   8160
      TabIndex        =   209
      Top             =   9240
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   8160
      TabIndex        =   208
      Top             =   9480
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   8160
      TabIndex        =   207
      Top             =   9000
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   5
      X1              =   120
      X2              =   11520
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   9600
      TabIndex        =   206
      Top             =   9000
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   2160
      TabIndex        =   205
      Top             =   7200
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   3000
      TabIndex        =   204
      Top             =   7200
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   7080
      TabIndex        =   203
      Top             =   7200
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   8280
      TabIndex        =   202
      Top             =   7200
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   9720
      TabIndex        =   201
      Top             =   7200
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   10320
      TabIndex        =   200
      Top             =   7200
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1200
      TabIndex        =   199
      Top             =   7200
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   198
      Top             =   7440
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1320
      TabIndex        =   197
      Top             =   7440
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   2520
      TabIndex        =   196
      Top             =   7440
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   3240
      TabIndex        =   195
      Top             =   7440
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   194
      Top             =   7680
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1320
      TabIndex        =   193
      Top             =   7680
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   2520
      TabIndex        =   192
      Top             =   7680
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   3240
      TabIndex        =   191
      Top             =   7680
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   190
      Top             =   7920
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1320
      TabIndex        =   189
      Top             =   7920
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   2520
      TabIndex        =   188
      Top             =   7920
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   3240
      TabIndex        =   187
      Top             =   7920
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   186
      Top             =   8160
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   1320
      TabIndex        =   185
      Top             =   8160
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   2520
      TabIndex        =   184
      Top             =   8160
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   3240
      TabIndex        =   183
      Top             =   8160
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   4560
      TabIndex        =   182
      Top             =   7680
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   4560
      TabIndex        =   181
      Top             =   8160
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   4560
      TabIndex        =   180
      Top             =   7920
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   4560
      TabIndex        =   179
      Top             =   7440
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   6480
      TabIndex        =   178
      Top             =   7440
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   6480
      TabIndex        =   177
      Top             =   7680
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   6480
      TabIndex        =   176
      Top             =   7920
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   6480
      TabIndex        =   175
      Top             =   8160
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   4
      X1              =   120
      X2              =   11520
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   174
      Top             =   7200
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   9600
      TabIndex        =   173
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   9600
      TabIndex        =   172
      Top             =   7920
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   9600
      TabIndex        =   171
      Top             =   7440
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   8160
      TabIndex        =   170
      Top             =   7440
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   8160
      TabIndex        =   169
      Top             =   7920
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   8160
      TabIndex        =   168
      Top             =   8160
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   8160
      TabIndex        =   167
      Top             =   7680
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   4
      X1              =   120
      X2              =   11520
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   9600
      TabIndex        =   166
      Top             =   7680
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2160
      TabIndex        =   165
      Top             =   5880
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3000
      TabIndex        =   164
      Top             =   5880
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   7080
      TabIndex        =   163
      Top             =   5880
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   8280
      TabIndex        =   162
      Top             =   5880
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   9720
      TabIndex        =   161
      Top             =   5880
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   10320
      TabIndex        =   160
      Top             =   5880
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1200
      TabIndex        =   159
      Top             =   5880
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   158
      Top             =   6120
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1320
      TabIndex        =   157
      Top             =   6120
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   156
      Top             =   6120
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3240
      TabIndex        =   155
      Top             =   6120
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   154
      Top             =   6360
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1320
      TabIndex        =   153
      Top             =   6360
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   152
      Top             =   6360
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3240
      TabIndex        =   151
      Top             =   6360
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   150
      Top             =   6600
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1320
      TabIndex        =   149
      Top             =   6600
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   148
      Top             =   6600
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3240
      TabIndex        =   147
      Top             =   6600
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   146
      Top             =   6840
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   1320
      TabIndex        =   145
      Top             =   6840
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   144
      Top             =   6840
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   3240
      TabIndex        =   143
      Top             =   6840
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   4560
      TabIndex        =   142
      Top             =   6360
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   4560
      TabIndex        =   141
      Top             =   6840
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   4560
      TabIndex        =   140
      Top             =   6600
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   4560
      TabIndex        =   139
      Top             =   6120
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   6480
      TabIndex        =   138
      Top             =   6120
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   6480
      TabIndex        =   137
      Top             =   6360
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   6480
      TabIndex        =   136
      Top             =   6600
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   6480
      TabIndex        =   135
      Top             =   6840
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   3
      X1              =   120
      X2              =   11520
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   120
      TabIndex        =   134
      Top             =   5880
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   9600
      TabIndex        =   133
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   9600
      TabIndex        =   132
      Top             =   6600
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   9600
      TabIndex        =   131
      Top             =   6120
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   8160
      TabIndex        =   130
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   8160
      TabIndex        =   129
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   8160
      TabIndex        =   128
      Top             =   6840
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   8160
      TabIndex        =   127
      Top             =   6360
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   3
      X1              =   120
      X2              =   11520
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   9600
      TabIndex        =   126
      Top             =   6360
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2160
      TabIndex        =   125
      Top             =   4560
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3000
      TabIndex        =   124
      Top             =   4560
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   7080
      TabIndex        =   123
      Top             =   4560
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   8280
      TabIndex        =   122
      Top             =   4560
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   9720
      TabIndex        =   121
      Top             =   4560
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   10320
      TabIndex        =   120
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1200
      TabIndex        =   119
      Top             =   4560
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   118
      Top             =   4800
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1320
      TabIndex        =   117
      Top             =   4800
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2520
      TabIndex        =   116
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3240
      TabIndex        =   115
      Top             =   4800
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   114
      Top             =   5040
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1320
      TabIndex        =   113
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2520
      TabIndex        =   112
      Top             =   5040
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3240
      TabIndex        =   111
      Top             =   5040
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   110
      Top             =   5280
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1320
      TabIndex        =   109
      Top             =   5280
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2520
      TabIndex        =   108
      Top             =   5280
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3240
      TabIndex        =   107
      Top             =   5280
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   106
      Top             =   5520
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   1320
      TabIndex        =   105
      Top             =   5520
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   2520
      TabIndex        =   104
      Top             =   5520
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   3240
      TabIndex        =   103
      Top             =   5520
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   102
      Top             =   5040
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   101
      Top             =   5520
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   100
      Top             =   5280
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   4560
      TabIndex        =   99
      Top             =   4800
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   6480
      TabIndex        =   98
      Top             =   4800
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   6480
      TabIndex        =   97
      Top             =   5040
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   6480
      TabIndex        =   96
      Top             =   5280
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   6480
      TabIndex        =   95
      Top             =   5520
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   2
      X1              =   120
      X2              =   11520
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   94
      Top             =   4560
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   9600
      TabIndex        =   93
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   9600
      TabIndex        =   92
      Top             =   5280
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   9600
      TabIndex        =   91
      Top             =   4800
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   8160
      TabIndex        =   90
      Top             =   4800
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   8160
      TabIndex        =   89
      Top             =   5280
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   8160
      TabIndex        =   88
      Top             =   5520
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   8160
      TabIndex        =   87
      Top             =   5040
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   2
      X1              =   120
      X2              =   11520
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   9600
      TabIndex        =   86
      Top             =   5040
      Width           =   540
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2160
      TabIndex        =   85
      Top             =   3240
      Width           =   675
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   3000
      TabIndex        =   84
      Top             =   3240
      Width           =   2475
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   7080
      TabIndex        =   83
      Top             =   3240
      Width           =   1020
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   8280
      TabIndex        =   82
      Top             =   3240
      Width           =   810
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   9720
      TabIndex        =   81
      Top             =   3240
      Width           =   435
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   10320
      TabIndex        =   80
      Top             =   3240
      Width           =   945
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1200
      TabIndex        =   79
      Top             =   3240
      Width           =   540
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   78
      Top             =   3480
      Width           =   1065
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1320
      TabIndex        =   77
      Top             =   3480
      Width           =   810
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2520
      TabIndex        =   76
      Top             =   3480
      Width           =   600
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   3240
      TabIndex        =   75
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   74
      Top             =   3720
      Width           =   1050
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1320
      TabIndex        =   73
      Top             =   3720
      Width           =   810
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2520
      TabIndex        =   72
      Top             =   3720
      Width           =   585
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   3240
      TabIndex        =   71
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   70
      Top             =   3960
      Width           =   1050
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1320
      TabIndex        =   69
      Top             =   3960
      Width           =   810
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2520
      TabIndex        =   68
      Top             =   3960
      Width           =   585
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   3240
      TabIndex        =   67
      Top             =   3960
      Width           =   945
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   66
      Top             =   4200
      Width           =   1050
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   1320
      TabIndex        =   65
      Top             =   4200
      Width           =   810
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   2520
      TabIndex        =   64
      Top             =   4200
      Width           =   585
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   3240
      TabIndex        =   63
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   62
      Top             =   3720
      Width           =   840
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   61
      Top             =   4200
      Width           =   1425
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   60
      Top             =   3960
      Width           =   645
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   4560
      TabIndex        =   59
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   6480
      TabIndex        =   58
      Top             =   3480
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   6480
      TabIndex        =   57
      Top             =   3720
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   6480
      TabIndex        =   56
      Top             =   3960
      Width           =   945
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   6480
      TabIndex        =   55
      Top             =   4200
      Width           =   945
   End
   Begin VB.Line L1 
      Index           =   1
      X1              =   120
      X2              =   11520
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   54
      Top             =   3240
      Width           =   840
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   9600
      TabIndex        =   53
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   9600
      TabIndex        =   52
      Top             =   3960
      Width           =   1500
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   9600
      TabIndex        =   51
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   8160
      TabIndex        =   50
      Top             =   3480
      Width           =   1290
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   8160
      TabIndex        =   49
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   8160
      TabIndex        =   48
      Top             =   4200
      Width           =   825
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   8160
      TabIndex        =   47
      Top             =   3720
      Width           =   870
   End
   Begin VB.Line L2 
      Index           =   1
      X1              =   120
      X2              =   11520
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   9600
      TabIndex        =   46
      Top             =   3720
      Width           =   540
   End
   Begin VB.Label LB_SP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9600
      TabIndex        =   45
      Top             =   2400
      Width           =   540
   End
   Begin VB.Line L2 
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label LB_SP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Seu Pedido:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   8160
      TabIndex        =   44
      Top             =   2400
      Width           =   870
   End
   Begin VB.Label LB_PB1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Peso Bruto:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   8160
      TabIndex        =   43
      Top             =   2880
      Width           =   825
   End
   Begin VB.Label LB_TR1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Transportadora:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   8160
      TabIndex        =   42
      Top             =   2640
      Width           =   1125
   End
   Begin VB.Label LB_PI1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Pedido Conesteel:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   8160
      TabIndex        =   41
      Top             =   2160
      Width           =   1290
   End
   Begin VB.Label LB_PI2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9600
      TabIndex        =   40
      Top             =   2160
      Width           =   540
   End
   Begin VB.Label LB_TR2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9600
      TabIndex        =   39
      Top             =   2640
      Width           =   1500
   End
   Begin VB.Label LB_PB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9600
      TabIndex        =   38
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label LB_NF1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nota Fiscal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "de Notas Fiscais Emitidas"
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
      Left            =   7200
      TabIndex        =   37
      Top             =   600
      Width           =   3360
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relatório Simplificado"
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
      Left            =   7440
      TabIndex        =   36
      Top             =   240
      Width           =   2880
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   1
      X1              =   6000
      X2              =   6000
      Y1              =   120
      Y2              =   1680
   End
   Begin VB.Line L1 
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label LB_VN2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   6480
      TabIndex        =   35
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label LB_IP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   6480
      TabIndex        =   34
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label LB_IC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   6480
      TabIndex        =   33
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label LB_TP2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   6480
      TabIndex        =   32
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label LB_TP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total dos Produtos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   31
      Top             =   2160
      Width           =   1740
   End
   Begin VB.Label LB_IP1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor IPI:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   30
      Top             =   2640
      Width           =   645
   End
   Begin VB.Label LB_VN1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total da Nota:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   29
      Top             =   2880
      Width           =   1425
   End
   Begin VB.Label LB_IC1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Valor ICMS:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   4560
      TabIndex        =   28
      Top             =   2400
      Width           =   840
   End
   Begin VB.Label LB_VD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   3240
      TabIndex        =   27
      Top             =   2880
      Width           =   945
   End
   Begin VB.Label LB_VD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2520
      TabIndex        =   26
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label LB_CD2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   25
      Top             =   2880
      Width           =   810
   End
   Begin VB.Label LB_CD1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/D:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   2880
      Width           =   1050
   End
   Begin VB.Label LB_VC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   3240
      TabIndex        =   23
      Top             =   2640
      Width           =   945
   End
   Begin VB.Label LB_VC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2520
      TabIndex        =   22
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label LB_CC2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   21
      Top             =   2640
      Width           =   810
   End
   Begin VB.Label LB_CC1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/C:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1050
   End
   Begin VB.Label LB_VB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   3240
      TabIndex        =   19
      Top             =   2400
      Width           =   945
   End
   Begin VB.Label LB_VB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2520
      TabIndex        =   18
      Top             =   2400
      Width           =   585
   End
   Begin VB.Label LB_CB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   17
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label LB_CB1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/B:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   16
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Label LB_VA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   3240
      TabIndex        =   15
      Top             =   2160
      Width           =   945
   End
   Begin VB.Label LB_VA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2520
      TabIndex        =   14
      Top             =   2160
      Width           =   600
   End
   Begin VB.Label LB_CA2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1320
      TabIndex        =   13
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label LB_CA1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vencimento/A:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   2160
      Width           =   1065
   End
   Begin VB.Label LB_NF2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "005000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   1200
      TabIndex        =   11
      Top             =   1920
      Width           =   540
   End
   Begin VB.Label LB_V2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   10320
      TabIndex        =   9
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label LB_V1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   9720
      TabIndex        =   8
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label LB_DE2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   8280
      TabIndex        =   7
      Top             =   1920
      Width           =   810
   End
   Begin VB.Label LB_DE1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Emissão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   7080
      TabIndex        =   6
      Top             =   1920
      Width           =   1020
   End
   Begin VB.Label LB_EM2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   2475
   End
   Begin VB.Label LB_EM1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   1920
      Width           =   675
   End
   Begin VB.Line LT 
      BorderWidth     =   2
      Index           =   2
      X1              =   120
      X2              =   11520
      Y1              =   1680
      Y2              =   1680
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
      TabIndex        =   3
      Top             =   240
      Width           =   2265
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
      TabIndex        =   2
      Top             =   720
      Width           =   3060
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
      TabIndex        =   1
      Top             =   1080
      Width           =   3825
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
      TabIndex        =   0
      Top             =   1320
      Width           =   4080
   End
End
Attribute VB_Name = "Tela_RelSimp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
