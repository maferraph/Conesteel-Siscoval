VERSION 5.00
Begin VB.Form Tela_Expedicao_EtiquetaSaco_Relatorio 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   17235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   17235
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Line Line1 
      Index           =   31
      X1              =   480
      X2              =   5400
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line2 
      Index           =   31
      X1              =   480
      X2              =   480
      Y1              =   8640
      Y2              =   10560
   End
   Begin VB.Line Line1 
      Index           =   30
      X1              =   480
      X2              =   5400
      Y1              =   10560
      Y2              =   10560
   End
   Begin VB.Line Line2 
      Index           =   30
      X1              =   5400
      X2              =   5400
      Y1              =   8640
      Y2              =   10560
   End
   Begin VB.Line Line2 
      Index           =   29
      X1              =   10560
      X2              =   10560
      Y1              =   8640
      Y2              =   10560
   End
   Begin VB.Line Line1 
      Index           =   29
      X1              =   5640
      X2              =   10560
      Y1              =   10560
      Y2              =   10560
   End
   Begin VB.Line Line2 
      Index           =   28
      X1              =   5640
      X2              =   5640
      Y1              =   8640
      Y2              =   10560
   End
   Begin VB.Line Line1 
      Index           =   28
      X1              =   5640
      X2              =   10560
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   143
      Left            =   600
      TabIndex        =   287
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label Label143 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   286
      Top             =   9000
      Width           =   435
   End
   Begin VB.Label Label142 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   285
      Top             =   10200
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   142
      Left            =   600
      TabIndex        =   284
      Top             =   9960
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   141
      Left            =   1440
      TabIndex        =   283
      Top             =   8760
      Width           =   660
   End
   Begin VB.Label Label141 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   282
      Top             =   9000
      Width           =   2790
   End
   Begin VB.Label Label140 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   281
      Top             =   9600
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   140
      Left            =   4440
      TabIndex        =   280
      Top             =   8760
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   139
      Left            =   3840
      TabIndex        =   279
      Top             =   9360
      Width           =   405
   End
   Begin VB.Label Label139 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   278
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label Label138 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   277
      Top             =   10200
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   138
      Left            =   3480
      TabIndex        =   276
      Top             =   9960
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   137
      Left            =   1920
      TabIndex        =   275
      Top             =   9360
      Width           =   630
   End
   Begin VB.Label Label137 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   274
      Top             =   9600
      Width           =   1755
   End
   Begin VB.Label Label136 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   273
      Top             =   10200
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   136
      Left            =   4680
      TabIndex        =   272
      Top             =   9960
      Width           =   540
   End
   Begin VB.Label Label135 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   271
      Top             =   9600
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   135
      Left            =   600
      TabIndex        =   270
      Top             =   9360
      Width           =   1065
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   134
      Left            =   5760
      TabIndex        =   269
      Top             =   8760
      Width           =   735
   End
   Begin VB.Label Label134 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   268
      Top             =   9000
      Width           =   435
   End
   Begin VB.Label Label133 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   267
      Top             =   10200
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   133
      Left            =   5760
      TabIndex        =   266
      Top             =   9960
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   132
      Left            =   6600
      TabIndex        =   265
      Top             =   8760
      Width           =   660
   End
   Begin VB.Label Label132 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   264
      Top             =   9000
      Width           =   2790
   End
   Begin VB.Label Label131 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   263
      Top             =   9600
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   131
      Left            =   9600
      TabIndex        =   262
      Top             =   8760
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   130
      Left            =   9000
      TabIndex        =   261
      Top             =   9360
      Width           =   405
   End
   Begin VB.Label Label130 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   260
      Top             =   9000
      Width           =   735
   End
   Begin VB.Label Label129 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   259
      Top             =   10200
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   129
      Left            =   8640
      TabIndex        =   258
      Top             =   9960
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   128
      Left            =   7080
      TabIndex        =   257
      Top             =   9360
      Width           =   630
   End
   Begin VB.Label Label128 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   256
      Top             =   9600
      Width           =   1755
   End
   Begin VB.Label Label127 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   255
      Top             =   10200
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   127
      Left            =   9840
      TabIndex        =   254
      Top             =   9960
      Width           =   540
   End
   Begin VB.Label Label126 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   253
      Top             =   9600
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   126
      Left            =   5760
      TabIndex        =   252
      Top             =   9360
      Width           =   1065
   End
   Begin VB.Line Line1 
      Index           =   27
      X1              =   5640
      X2              =   10560
      Y1              =   10800
      Y2              =   10800
   End
   Begin VB.Line Line2 
      Index           =   27
      X1              =   5640
      X2              =   5640
      Y1              =   10800
      Y2              =   12720
   End
   Begin VB.Line Line1 
      Index           =   26
      X1              =   5640
      X2              =   10560
      Y1              =   12720
      Y2              =   12720
   End
   Begin VB.Line Line2 
      Index           =   26
      X1              =   10560
      X2              =   10560
      Y1              =   10800
      Y2              =   12720
   End
   Begin VB.Line Line2 
      Index           =   25
      X1              =   5400
      X2              =   5400
      Y1              =   10800
      Y2              =   12720
   End
   Begin VB.Line Line1 
      Index           =   25
      X1              =   480
      X2              =   5400
      Y1              =   12720
      Y2              =   12720
   End
   Begin VB.Line Line1 
      Index           =   24
      X1              =   480
      X2              =   5400
      Y1              =   10800
      Y2              =   10800
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   125
      Left            =   600
      TabIndex        =   251
      Top             =   10920
      Width           =   735
   End
   Begin VB.Label Label125 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   250
      Top             =   11160
      Width           =   435
   End
   Begin VB.Label Label124 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   249
      Top             =   12360
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   124
      Left            =   600
      TabIndex        =   248
      Top             =   12120
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   123
      Left            =   1440
      TabIndex        =   247
      Top             =   10920
      Width           =   660
   End
   Begin VB.Label Label123 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   246
      Top             =   11160
      Width           =   2790
   End
   Begin VB.Label Label122 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   245
      Top             =   11760
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   122
      Left            =   4440
      TabIndex        =   244
      Top             =   10920
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   121
      Left            =   3840
      TabIndex        =   243
      Top             =   11520
      Width           =   405
   End
   Begin VB.Label Label121 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   242
      Top             =   11160
      Width           =   735
   End
   Begin VB.Label Label120 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   241
      Top             =   12360
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   120
      Left            =   3480
      TabIndex        =   240
      Top             =   12120
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   119
      Left            =   1920
      TabIndex        =   239
      Top             =   11520
      Width           =   630
   End
   Begin VB.Label Label119 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   238
      Top             =   11760
      Width           =   1755
   End
   Begin VB.Label Label118 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   237
      Top             =   12360
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   118
      Left            =   4680
      TabIndex        =   236
      Top             =   12120
      Width           =   540
   End
   Begin VB.Label Label117 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   235
      Top             =   11760
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   117
      Left            =   600
      TabIndex        =   234
      Top             =   11520
      Width           =   1065
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   116
      Left            =   5760
      TabIndex        =   233
      Top             =   10920
      Width           =   735
   End
   Begin VB.Label Label116 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   232
      Top             =   11160
      Width           =   435
   End
   Begin VB.Label Label115 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   231
      Top             =   12360
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   115
      Left            =   5760
      TabIndex        =   230
      Top             =   12120
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   114
      Left            =   6600
      TabIndex        =   229
      Top             =   10920
      Width           =   660
   End
   Begin VB.Label Label114 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   228
      Top             =   11160
      Width           =   2790
   End
   Begin VB.Label Label113 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   227
      Top             =   11760
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   113
      Left            =   9600
      TabIndex        =   226
      Top             =   10920
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   112
      Left            =   9000
      TabIndex        =   225
      Top             =   11520
      Width           =   405
   End
   Begin VB.Label Label112 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   224
      Top             =   11160
      Width           =   735
   End
   Begin VB.Label Label111 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   223
      Top             =   12360
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   111
      Left            =   8640
      TabIndex        =   222
      Top             =   12120
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   110
      Left            =   7080
      TabIndex        =   221
      Top             =   11520
      Width           =   630
   End
   Begin VB.Label Label110 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   220
      Top             =   11760
      Width           =   1755
   End
   Begin VB.Label Label109 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   219
      Top             =   12360
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   109
      Left            =   9840
      TabIndex        =   218
      Top             =   12120
      Width           =   540
   End
   Begin VB.Label Label108 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   217
      Top             =   11760
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   108
      Left            =   5760
      TabIndex        =   216
      Top             =   11520
      Width           =   1065
   End
   Begin VB.Line Line2 
      Index           =   24
      X1              =   480
      X2              =   480
      Y1              =   10800
      Y2              =   12720
   End
   Begin VB.Line Line1 
      Index           =   23
      X1              =   5640
      X2              =   10560
      Y1              =   12960
      Y2              =   12960
   End
   Begin VB.Line Line2 
      Index           =   23
      X1              =   5640
      X2              =   5640
      Y1              =   12960
      Y2              =   14880
   End
   Begin VB.Line Line1 
      Index           =   22
      X1              =   5640
      X2              =   10560
      Y1              =   14880
      Y2              =   14880
   End
   Begin VB.Line Line2 
      Index           =   22
      X1              =   10560
      X2              =   10560
      Y1              =   12960
      Y2              =   14880
   End
   Begin VB.Line Line2 
      Index           =   21
      X1              =   5400
      X2              =   5400
      Y1              =   12960
      Y2              =   14880
   End
   Begin VB.Line Line1 
      Index           =   21
      X1              =   480
      X2              =   5400
      Y1              =   14880
      Y2              =   14880
   End
   Begin VB.Line Line2 
      Index           =   20
      X1              =   480
      X2              =   480
      Y1              =   12960
      Y2              =   14880
   End
   Begin VB.Line Line1 
      Index           =   20
      X1              =   480
      X2              =   5400
      Y1              =   12960
      Y2              =   12960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   107
      Left            =   600
      TabIndex        =   215
      Top             =   13080
      Width           =   735
   End
   Begin VB.Label Label107 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   214
      Top             =   13320
      Width           =   435
   End
   Begin VB.Label Label106 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   213
      Top             =   14520
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   106
      Left            =   600
      TabIndex        =   212
      Top             =   14280
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   105
      Left            =   1440
      TabIndex        =   211
      Top             =   13080
      Width           =   660
   End
   Begin VB.Label Label105 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   210
      Top             =   13320
      Width           =   2790
   End
   Begin VB.Label Label104 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   209
      Top             =   13920
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   104
      Left            =   4440
      TabIndex        =   208
      Top             =   13080
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   103
      Left            =   3840
      TabIndex        =   207
      Top             =   13680
      Width           =   405
   End
   Begin VB.Label Label103 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   206
      Top             =   13320
      Width           =   735
   End
   Begin VB.Label Label102 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   205
      Top             =   14520
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   102
      Left            =   3480
      TabIndex        =   204
      Top             =   14280
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   101
      Left            =   1920
      TabIndex        =   203
      Top             =   13680
      Width           =   630
   End
   Begin VB.Label Label101 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   202
      Top             =   13920
      Width           =   1755
   End
   Begin VB.Label Label100 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   201
      Top             =   14520
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   100
      Left            =   4680
      TabIndex        =   200
      Top             =   14280
      Width           =   540
   End
   Begin VB.Label Label99 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   199
      Top             =   13920
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   99
      Left            =   600
      TabIndex        =   198
      Top             =   13680
      Width           =   1065
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   98
      Left            =   5760
      TabIndex        =   197
      Top             =   13080
      Width           =   735
   End
   Begin VB.Label Label98 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   196
      Top             =   13320
      Width           =   435
   End
   Begin VB.Label Label97 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   195
      Top             =   14520
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   97
      Left            =   5760
      TabIndex        =   194
      Top             =   14280
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   96
      Left            =   6600
      TabIndex        =   193
      Top             =   13080
      Width           =   660
   End
   Begin VB.Label Label96 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   192
      Top             =   13320
      Width           =   2790
   End
   Begin VB.Label Label95 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   191
      Top             =   13920
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   95
      Left            =   9600
      TabIndex        =   190
      Top             =   13080
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   94
      Left            =   9000
      TabIndex        =   189
      Top             =   13680
      Width           =   405
   End
   Begin VB.Label Label94 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   188
      Top             =   13320
      Width           =   735
   End
   Begin VB.Label Label93 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   187
      Top             =   14520
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   93
      Left            =   8640
      TabIndex        =   186
      Top             =   14280
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   92
      Left            =   7080
      TabIndex        =   185
      Top             =   13680
      Width           =   630
   End
   Begin VB.Label Label92 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   184
      Top             =   13920
      Width           =   1755
   End
   Begin VB.Label Label91 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   183
      Top             =   14520
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   91
      Left            =   9840
      TabIndex        =   182
      Top             =   14280
      Width           =   540
   End
   Begin VB.Label Label90 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   181
      Top             =   13920
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   90
      Left            =   5760
      TabIndex        =   180
      Top             =   13680
      Width           =   1065
   End
   Begin VB.Line Line1 
      Index           =   19
      X1              =   5640
      X2              =   10560
      Y1              =   15120
      Y2              =   15120
   End
   Begin VB.Line Line2 
      Index           =   19
      X1              =   5640
      X2              =   5640
      Y1              =   15120
      Y2              =   17040
   End
   Begin VB.Line Line1 
      Index           =   18
      X1              =   5640
      X2              =   10560
      Y1              =   17040
      Y2              =   17040
   End
   Begin VB.Line Line2 
      Index           =   18
      X1              =   10560
      X2              =   10560
      Y1              =   15120
      Y2              =   17040
   End
   Begin VB.Line Line2 
      Index           =   17
      X1              =   5400
      X2              =   5400
      Y1              =   15120
      Y2              =   17040
   End
   Begin VB.Line Line1 
      Index           =   17
      X1              =   480
      X2              =   5400
      Y1              =   17040
      Y2              =   17040
   End
   Begin VB.Line Line1 
      Index           =   16
      X1              =   480
      X2              =   5400
      Y1              =   15120
      Y2              =   15120
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   89
      Left            =   600
      TabIndex        =   179
      Top             =   15240
      Width           =   735
   End
   Begin VB.Label Label89 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   178
      Top             =   15480
      Width           =   435
   End
   Begin VB.Label Label88 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   177
      Top             =   16680
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   88
      Left            =   600
      TabIndex        =   176
      Top             =   16440
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   87
      Left            =   1440
      TabIndex        =   175
      Top             =   15240
      Width           =   660
   End
   Begin VB.Label Label87 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   174
      Top             =   15480
      Width           =   2790
   End
   Begin VB.Label Label86 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   173
      Top             =   16080
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   86
      Left            =   4440
      TabIndex        =   172
      Top             =   15240
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   85
      Left            =   3840
      TabIndex        =   171
      Top             =   15840
      Width           =   405
   End
   Begin VB.Label Label85 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   170
      Top             =   15480
      Width           =   735
   End
   Begin VB.Label Label84 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   169
      Top             =   16680
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   84
      Left            =   3480
      TabIndex        =   168
      Top             =   16440
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   83
      Left            =   1920
      TabIndex        =   167
      Top             =   15840
      Width           =   630
   End
   Begin VB.Label Label83 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   166
      Top             =   16080
      Width           =   1755
   End
   Begin VB.Label Label82 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   165
      Top             =   16680
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   82
      Left            =   4680
      TabIndex        =   164
      Top             =   16440
      Width           =   540
   End
   Begin VB.Label Label81 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   163
      Top             =   16080
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   81
      Left            =   600
      TabIndex        =   162
      Top             =   15840
      Width           =   1065
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   80
      Left            =   5760
      TabIndex        =   161
      Top             =   15240
      Width           =   735
   End
   Begin VB.Label Label80 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   160
      Top             =   15480
      Width           =   435
   End
   Begin VB.Label Label79 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   159
      Top             =   16680
      Width           =   2385
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   79
      Left            =   5760
      TabIndex        =   158
      Top             =   16440
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   78
      Left            =   6600
      TabIndex        =   157
      Top             =   15240
      Width           =   660
   End
   Begin VB.Label Label78 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   156
      Top             =   15480
      Width           =   2790
   End
   Begin VB.Label Label77 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   155
      Top             =   16080
      Width           =   1410
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   77
      Left            =   9600
      TabIndex        =   154
      Top             =   15240
      Width           =   600
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   76
      Left            =   9000
      TabIndex        =   153
      Top             =   15840
      Width           =   405
   End
   Begin VB.Label Label76 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   152
      Top             =   15480
      Width           =   735
   End
   Begin VB.Label Label75 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   151
      Top             =   16680
      Width           =   885
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   75
      Left            =   8640
      TabIndex        =   150
      Top             =   16440
      Width           =   540
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   74
      Left            =   7080
      TabIndex        =   149
      Top             =   15840
      Width           =   630
   End
   Begin VB.Label Label74 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   148
      Top             =   16080
      Width           =   1755
   End
   Begin VB.Label Label73 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   147
      Top             =   16680
      Width           =   255
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   73
      Left            =   9840
      TabIndex        =   146
      Top             =   16440
      Width           =   540
   End
   Begin VB.Label Label72 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   145
      Top             =   16080
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   72
      Left            =   5760
      TabIndex        =   144
      Top             =   15840
      Width           =   1065
   End
   Begin VB.Line Line2 
      Index           =   16
      X1              =   480
      X2              =   480
      Y1              =   15120
      Y2              =   17040
   End
   Begin VB.Line Line2 
      Index           =   15
      X1              =   480
      X2              =   480
      Y1              =   6480
      Y2              =   8400
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   71
      Left            =   5760
      TabIndex        =   143
      Top             =   7200
      Width           =   1065
   End
   Begin VB.Label Label71 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   142
      Top             =   7440
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   70
      Left            =   9840
      TabIndex        =   141
      Top             =   7800
      Width           =   540
   End
   Begin VB.Label Label70 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   140
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label69 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   139
      Top             =   7440
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   69
      Left            =   7080
      TabIndex        =   138
      Top             =   7200
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   68
      Left            =   8640
      TabIndex        =   137
      Top             =   7800
      Width           =   540
   End
   Begin VB.Label Label68 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   136
      Top             =   8040
      Width           =   885
   End
   Begin VB.Label Label67 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   135
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   67
      Left            =   9000
      TabIndex        =   134
      Top             =   7200
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   66
      Left            =   9600
      TabIndex        =   133
      Top             =   6600
      Width           =   600
   End
   Begin VB.Label Label66 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   132
      Top             =   7440
      Width           =   1410
   End
   Begin VB.Label Label65 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   131
      Top             =   6840
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   65
      Left            =   6600
      TabIndex        =   130
      Top             =   6600
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   64
      Left            =   5760
      TabIndex        =   129
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label64 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   128
      Top             =   8040
      Width           =   2385
   End
   Begin VB.Label Label63 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   127
      Top             =   6840
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   63
      Left            =   5760
      TabIndex        =   126
      Top             =   6600
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   62
      Left            =   600
      TabIndex        =   125
      Top             =   7200
      Width           =   1065
   End
   Begin VB.Label Label62 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   124
      Top             =   7440
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   61
      Left            =   4680
      TabIndex        =   123
      Top             =   7800
      Width           =   540
   End
   Begin VB.Label Label61 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   122
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label60 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   121
      Top             =   7440
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   60
      Left            =   1920
      TabIndex        =   120
      Top             =   7200
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   59
      Left            =   3480
      TabIndex        =   119
      Top             =   7800
      Width           =   540
   End
   Begin VB.Label Label59 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   118
      Top             =   8040
      Width           =   885
   End
   Begin VB.Label Label58 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   117
      Top             =   6840
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   58
      Left            =   3840
      TabIndex        =   116
      Top             =   7200
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   57
      Left            =   4440
      TabIndex        =   115
      Top             =   6600
      Width           =   600
   End
   Begin VB.Label Label57 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   114
      Top             =   7440
      Width           =   1410
   End
   Begin VB.Label Label56 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   113
      Top             =   6840
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   56
      Left            =   1440
      TabIndex        =   112
      Top             =   6600
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   55
      Left            =   600
      TabIndex        =   111
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label55 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   110
      Top             =   8040
      Width           =   2385
   End
   Begin VB.Label Label54 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   109
      Top             =   6840
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   54
      Left            =   600
      TabIndex        =   108
      Top             =   6600
      Width           =   735
   End
   Begin VB.Line Line1 
      Index           =   15
      X1              =   480
      X2              =   5400
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      Index           =   14
      X1              =   480
      X2              =   5400
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line2 
      Index           =   14
      X1              =   5400
      X2              =   5400
      Y1              =   6480
      Y2              =   8400
   End
   Begin VB.Line Line2 
      Index           =   13
      X1              =   10560
      X2              =   10560
      Y1              =   6480
      Y2              =   8400
   End
   Begin VB.Line Line1 
      Index           =   13
      X1              =   5640
      X2              =   10560
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line2 
      Index           =   12
      X1              =   5640
      X2              =   5640
      Y1              =   6480
      Y2              =   8400
   End
   Begin VB.Line Line1 
      Index           =   12
      X1              =   5640
      X2              =   10560
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   53
      Left            =   5760
      TabIndex        =   107
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label53 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   106
      Top             =   5280
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   52
      Left            =   9840
      TabIndex        =   105
      Top             =   5640
      Width           =   540
   End
   Begin VB.Label Label52 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   104
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label51 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   103
      Top             =   5280
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   51
      Left            =   7080
      TabIndex        =   102
      Top             =   5040
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   50
      Left            =   8640
      TabIndex        =   101
      Top             =   5640
      Width           =   540
   End
   Begin VB.Label Label50 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   100
      Top             =   5880
      Width           =   885
   End
   Begin VB.Label Label49 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   99
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   49
      Left            =   9000
      TabIndex        =   98
      Top             =   5040
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   48
      Left            =   9600
      TabIndex        =   97
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label Label48 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   96
      Top             =   5280
      Width           =   1410
   End
   Begin VB.Label Label47 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   95
      Top             =   4680
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   47
      Left            =   6600
      TabIndex        =   94
      Top             =   4440
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   46
      Left            =   5760
      TabIndex        =   93
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label46 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   92
      Top             =   5880
      Width           =   2385
   End
   Begin VB.Label Label45 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   91
      Top             =   4680
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   45
      Left            =   5760
      TabIndex        =   90
      Top             =   4440
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   44
      Left            =   600
      TabIndex        =   89
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label44 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   88
      Top             =   5280
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   43
      Left            =   4680
      TabIndex        =   87
      Top             =   5640
      Width           =   540
   End
   Begin VB.Label Label43 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   86
      Top             =   5880
      Width           =   255
   End
   Begin VB.Label Label42 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   85
      Top             =   5280
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   42
      Left            =   1920
      TabIndex        =   84
      Top             =   5040
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   41
      Left            =   3480
      TabIndex        =   83
      Top             =   5640
      Width           =   540
   End
   Begin VB.Label Label41 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   82
      Top             =   5880
      Width           =   885
   End
   Begin VB.Label Label40 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   81
      Top             =   4680
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   40
      Left            =   3840
      TabIndex        =   80
      Top             =   5040
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   39
      Left            =   4440
      TabIndex        =   79
      Top             =   4440
      Width           =   600
   End
   Begin VB.Label Label39 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   78
      Top             =   5280
      Width           =   1410
   End
   Begin VB.Label Label38 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   77
      Top             =   4680
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   38
      Left            =   1440
      TabIndex        =   76
      Top             =   4440
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   37
      Left            =   600
      TabIndex        =   75
      Top             =   5640
      Width           =   735
   End
   Begin VB.Label Label37 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   74
      Top             =   5880
      Width           =   2385
   End
   Begin VB.Label Label36 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   73
      Top             =   4680
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   36
      Left            =   600
      TabIndex        =   72
      Top             =   4440
      Width           =   735
   End
   Begin VB.Line Line1 
      Index           =   11
      X1              =   480
      X2              =   5400
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      Index           =   11
      X1              =   480
      X2              =   480
      Y1              =   4320
      Y2              =   6240
   End
   Begin VB.Line Line1 
      Index           =   10
      X1              =   480
      X2              =   5400
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      Index           =   10
      X1              =   5400
      X2              =   5400
      Y1              =   4320
      Y2              =   6240
   End
   Begin VB.Line Line2 
      Index           =   9
      X1              =   10560
      X2              =   10560
      Y1              =   4320
      Y2              =   6240
   End
   Begin VB.Line Line1 
      Index           =   9
      X1              =   5640
      X2              =   10560
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line2 
      Index           =   8
      X1              =   5640
      X2              =   5640
      Y1              =   4320
      Y2              =   6240
   End
   Begin VB.Line Line1 
      Index           =   8
      X1              =   5640
      X2              =   10560
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      Index           =   7
      X1              =   480
      X2              =   480
      Y1              =   2160
      Y2              =   4080
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   35
      Left            =   5760
      TabIndex        =   71
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Label Label35 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   70
      Top             =   3120
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   34
      Left            =   9840
      TabIndex        =   69
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label Label34 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   68
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label33 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   67
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   33
      Left            =   7080
      TabIndex        =   66
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   32
      Left            =   8640
      TabIndex        =   65
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label Label32 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   64
      Top             =   3720
      Width           =   885
   End
   Begin VB.Label Label31 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   63
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   31
      Left            =   9000
      TabIndex        =   62
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   9600
      TabIndex        =   61
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label Label30 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   60
      Top             =   3120
      Width           =   1410
   End
   Begin VB.Label Label29 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   59
      Top             =   2520
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   29
      Left            =   6600
      TabIndex        =   58
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   28
      Left            =   5760
      TabIndex        =   57
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label28 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   56
      Top             =   3720
      Width           =   2385
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   55
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   27
      Left            =   5760
      TabIndex        =   54
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   26
      Left            =   600
      TabIndex        =   53
      Top             =   2880
      Width           =   1065
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   52
      Top             =   3120
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   25
      Left            =   4680
      TabIndex        =   51
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   50
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   49
      Top             =   3120
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   1920
      TabIndex        =   48
      Top             =   2880
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   3480
      TabIndex        =   47
      Top             =   3480
      Width           =   540
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   46
      Top             =   3720
      Width           =   885
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   45
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   22
      Left            =   3840
      TabIndex        =   44
      Top             =   2880
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   21
      Left            =   4440
      TabIndex        =   43
      Top             =   2280
      Width           =   600
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   42
      Top             =   3120
      Width           =   1410
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   41
      Top             =   2520
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   20
      Left            =   1440
      TabIndex        =   40
      Top             =   2280
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   19
      Left            =   600
      TabIndex        =   39
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   38
      Top             =   3720
      Width           =   2385
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   37
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   18
      Left            =   600
      TabIndex        =   36
      Top             =   2280
      Width           =   735
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   480
      X2              =   5400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   6
      X1              =   480
      X2              =   5400
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      Index           =   6
      X1              =   5400
      X2              =   5400
      Y1              =   2160
      Y2              =   4080
   End
   Begin VB.Line Line2 
      Index           =   5
      X1              =   10560
      X2              =   10560
      Y1              =   2160
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   5
      X1              =   5640
      X2              =   10560
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      Index           =   4
      X1              =   5640
      X2              =   5640
      Y1              =   2160
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   4
      X1              =   5640
      X2              =   10560
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   5760
      TabIndex        =   35
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   34
      Top             =   960
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   9840
      TabIndex        =   33
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9840
      TabIndex        =   32
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7080
      TabIndex        =   31
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   7080
      TabIndex        =   30
      Top             =   720
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   8640
      TabIndex        =   29
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8640
      TabIndex        =   28
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9600
      TabIndex        =   27
      Top             =   360
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   9000
      TabIndex        =   26
      Top             =   720
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   9600
      TabIndex        =   25
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   9000
      TabIndex        =   24
      Top             =   960
      Width           =   1410
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6600
      TabIndex        =   23
      Top             =   360
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   6600
      TabIndex        =   22
      Top             =   120
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   5760
      TabIndex        =   21
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   20
      Top             =   1560
      Width           =   2385
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5760
      TabIndex        =   19
      Top             =   360
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   5760
      TabIndex        =   18
      Top             =   120
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Seu Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   600
      TabIndex        =   17
      Top             =   720
      Width           =   1065
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "123456890"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   16
      Top             =   960
      Width           =   960
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Estado:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   4680
      TabIndex        =   15
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4680
      TabIndex        =   14
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "55.783.427/0001-03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   960
      Width           =   1755
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "C.N.P.J.:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   1920
      TabIndex        =   12
      Top             =   720
      Width           =   630
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Cidade:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   3480
      TabIndex        =   11
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "São Paulo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3480
      TabIndex        =   10
      Top             =   1560
      Width           =   885
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Mauricio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4440
      TabIndex        =   9
      Top             =   360
      Width           =   735
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Fone:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   3840
      TabIndex        =   8
      Top             =   720
      Width           =   405
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Contato:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   4440
      TabIndex        =   7
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3840
      TabIndex        =   6
      Top             =   960
      Width           =   1410
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Conesteel Válv.Conex.Inds.Ltda."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1440
      TabIndex        =   5
      Top             =   360
      Width           =   2790
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Empresa:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   660
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Endereço:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Avenida Montemagno, 2454"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   2
      Top             =   1560
      Width           =   2385
   End
   Begin VB.Label LB_PE 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   435
   End
   Begin VB.Label LB 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Pedido nº:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   5640
      X2              =   10560
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line2 
      Index           =   3
      X1              =   5640
      X2              =   5640
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   5640
      X2              =   10560
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   10560
      X2              =   10560
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   5400
      X2              =   5400
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   480
      X2              =   5400
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   480
      X2              =   480
      Y1              =   0
      Y2              =   1920
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   480
      X2              =   5400
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "Tela_Expedicao_EtiquetaSaco_Relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
