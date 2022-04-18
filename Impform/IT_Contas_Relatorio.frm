VERSION 5.00
Begin VB.Form IT_Contas_Relatorio 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   14700
   ClientLeft      =   0
   ClientTop       =   0
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
      Picture         =   "IT_Contas_Relatorio.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   12
      Top             =   240
      Width           =   600
   End
   Begin VB.Label LB_EB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM COBRANÇA:"
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
      Left            =   1440
      TabIndex        =   194
      Top             =   12975
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label LB_DataEmissao 
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
      Index           =   6
      Left            =   120
      TabIndex        =   193
      Top             =   12135
      Width           =   960
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE EMISSÃO:"
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
      Left            =   120
      TabIndex        =   192
      Top             =   12015
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Line LHIC 
      Index           =   6
      X1              =   0
      X2              =   11160
      Y1              =   12000
      Y2              =   12000
   End
   Begin VB.Label LB_DataVencimento 
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
      Index           =   6
      Left            =   1560
      TabIndex        =   191
      Top             =   12135
      Width           =   960
   End
   Begin VB.Label LB_DV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VENCIMENTO:"
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
      Left            =   1560
      TabIndex        =   190
      Top             =   12015
      UseMnemonic     =   0   'False
      Width           =   1365
   End
   Begin VB.Label LB_Movimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À Pagar"
      BeginProperty Font 
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
      Left            =   3480
      TabIndex        =   189
      Top             =   12135
      Width           =   705
   End
   Begin VB.Label LB_MO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO:"
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
      Left            =   3480
      TabIndex        =   188
      Top             =   12015
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
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
      TabIndex        =   187
      Top             =   12135
      Width           =   1125
   End
   Begin VB.Label LB_VA 
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
      Index           =   6
      Left            =   4560
      TabIndex        =   186
      Top             =   12015
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Origem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL CONEXÕES DE AÇO LTDA."
      BeginProperty Font 
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
      Left            =   6000
      TabIndex        =   185
      Top             =   12135
      Width           =   3720
   End
   Begin VB.Label LB_OR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEM OU DESTINO:"
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
      Left            =   6000
      TabIndex        =   184
      Top             =   12015
      UseMnemonic     =   0   'False
      Width           =   1275
   End
   Begin VB.Label LB_NumeroDocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      TabIndex        =   183
      Top             =   12615
      Width           =   525
   End
   Begin VB.Label LB_ND 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO DOCUMENTO:"
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
      Left            =   120
      TabIndex        =   182
      Top             =   12480
      UseMnemonic     =   0   'False
      Width           =   1605
   End
   Begin VB.Label LB_SeuNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   181
      Top             =   12615
      Width           =   525
   End
   Begin VB.Label LB_NN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO NÚMERO:"
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
      Left            =   3360
      TabIndex        =   180
      Top             =   12495
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label LB_NossoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   3360
      TabIndex        =   179
      Top             =   12615
      Width           =   525
   End
   Begin VB.Label LB_SN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEU NÚMERO:"
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
      Left            =   2040
      TabIndex        =   178
      Top             =   12495
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATA"
      BeginProperty Font 
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
      Left            =   4920
      TabIndex        =   177
      Top             =   12615
      Width           =   1065
   End
   Begin VB.Label LB_TI 
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
      Index           =   6
      Left            =   4920
      TabIndex        =   176
      Top             =   12495
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Banco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRADESCO"
      BeginProperty Font 
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
      TabIndex        =   175
      Top             =   12615
      Width           =   1095
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO:"
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
      Left            =   7920
      TabIndex        =   174
      Top             =   12495
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_EmCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      TabIndex        =   173
      Top             =   13095
      Width           =   345
   End
   Begin VB.Label LB_EC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM CARTEIRA:"
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
      Left            =   120
      TabIndex        =   172
      Top             =   12975
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_EmCobranca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   171
      Top             =   13095
      Width           =   345
   End
   Begin VB.Label LB_Observacoes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO TEM"
      BeginProperty Font 
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
      TabIndex        =   170
      Top             =   13095
      Width           =   885
   End
   Begin VB.Label LB_OB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES:"
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
      Left            =   2760
      TabIndex        =   169
      Top             =   12975
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Line LHIB 
      Index           =   6
      X1              =   0
      X2              =   11160
      Y1              =   13335
      Y2              =   13335
   End
   Begin VB.Label LB_EB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM COBRANÇA:"
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
      Left            =   1440
      TabIndex        =   168
      Top             =   11295
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label LB_DataEmissao 
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
      Index           =   5
      Left            =   120
      TabIndex        =   167
      Top             =   10455
      Width           =   960
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE EMISSÃO:"
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
      TabIndex        =   166
      Top             =   10335
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Line LHIC 
      Index           =   5
      X1              =   0
      X2              =   11160
      Y1              =   10320
      Y2              =   10320
   End
   Begin VB.Label LB_DataVencimento 
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
      Index           =   5
      Left            =   1560
      TabIndex        =   165
      Top             =   10455
      Width           =   960
   End
   Begin VB.Label LB_DV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VENCIMENTO:"
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
      Left            =   1560
      TabIndex        =   164
      Top             =   10335
      UseMnemonic     =   0   'False
      Width           =   1365
   End
   Begin VB.Label LB_Movimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À Pagar"
      BeginProperty Font 
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
      Left            =   3480
      TabIndex        =   163
      Top             =   10455
      Width           =   705
   End
   Begin VB.Label LB_MO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO:"
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
      Left            =   3480
      TabIndex        =   162
      Top             =   10335
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
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
      TabIndex        =   161
      Top             =   10455
      Width           =   1125
   End
   Begin VB.Label LB_VA 
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
      Index           =   5
      Left            =   4560
      TabIndex        =   160
      Top             =   10335
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Origem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL CONEXÕES DE AÇO LTDA."
      BeginProperty Font 
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
      Left            =   6000
      TabIndex        =   159
      Top             =   10455
      Width           =   3720
   End
   Begin VB.Label LB_OR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEM OU DESTINO:"
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
      Left            =   6000
      TabIndex        =   158
      Top             =   10335
      UseMnemonic     =   0   'False
      Width           =   1275
   End
   Begin VB.Label LB_NumeroDocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      TabIndex        =   157
      Top             =   10935
      Width           =   525
   End
   Begin VB.Label LB_ND 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO DOCUMENTO:"
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
      TabIndex        =   156
      Top             =   10815
      UseMnemonic     =   0   'False
      Width           =   1605
   End
   Begin VB.Label LB_SeuNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   155
      Top             =   10935
      Width           =   525
   End
   Begin VB.Label LB_NN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO NÚMERO:"
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
      Left            =   3360
      TabIndex        =   154
      Top             =   10815
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label LB_NossoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   3360
      TabIndex        =   153
      Top             =   10935
      Width           =   525
   End
   Begin VB.Label LB_SN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEU NÚMERO:"
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
      Left            =   2040
      TabIndex        =   152
      Top             =   10815
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATA"
      BeginProperty Font 
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
      Left            =   4920
      TabIndex        =   151
      Top             =   10935
      Width           =   1065
   End
   Begin VB.Label LB_TI 
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
      Index           =   5
      Left            =   4920
      TabIndex        =   150
      Top             =   10815
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Banco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRADESCO"
      BeginProperty Font 
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
      TabIndex        =   149
      Top             =   10935
      Width           =   1095
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO:"
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
      Left            =   7920
      TabIndex        =   148
      Top             =   10815
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_EmCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      TabIndex        =   147
      Top             =   11415
      Width           =   345
   End
   Begin VB.Label LB_EC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM CARTEIRA:"
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
      TabIndex        =   146
      Top             =   11295
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_EmCobranca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   145
      Top             =   11415
      Width           =   345
   End
   Begin VB.Label LB_Observacoes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO TEM"
      BeginProperty Font 
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
      TabIndex        =   144
      Top             =   11415
      Width           =   885
   End
   Begin VB.Label LB_OB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES:"
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
      Left            =   2760
      TabIndex        =   143
      Top             =   11295
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Line LHIB 
      Index           =   5
      X1              =   0
      X2              =   11160
      Y1              =   11655
      Y2              =   11655
   End
   Begin VB.Label LB_EB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM COBRANÇA:"
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
      Left            =   1440
      TabIndex        =   142
      Top             =   9615
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label LB_DataEmissao 
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
      Index           =   4
      Left            =   120
      TabIndex        =   141
      Top             =   8775
      Width           =   960
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE EMISSÃO:"
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
      Left            =   120
      TabIndex        =   140
      Top             =   8655
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Line LHIC 
      Index           =   4
      X1              =   0
      X2              =   11160
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Label LB_DataVencimento 
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
      Index           =   4
      Left            =   1560
      TabIndex        =   139
      Top             =   8775
      Width           =   960
   End
   Begin VB.Label LB_DV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VENCIMENTO:"
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
      Left            =   1560
      TabIndex        =   138
      Top             =   8655
      UseMnemonic     =   0   'False
      Width           =   1365
   End
   Begin VB.Label LB_Movimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À Pagar"
      BeginProperty Font 
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
      Left            =   3480
      TabIndex        =   137
      Top             =   8775
      Width           =   705
   End
   Begin VB.Label LB_MO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO:"
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
      Left            =   3480
      TabIndex        =   136
      Top             =   8655
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
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
      TabIndex        =   135
      Top             =   8775
      Width           =   1125
   End
   Begin VB.Label LB_VA 
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
      Index           =   4
      Left            =   4560
      TabIndex        =   134
      Top             =   8655
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Origem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL CONEXÕES DE AÇO LTDA."
      BeginProperty Font 
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
      Left            =   6000
      TabIndex        =   133
      Top             =   8775
      Width           =   3720
   End
   Begin VB.Label LB_OR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEM OU DESTINO:"
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
      Left            =   6000
      TabIndex        =   132
      Top             =   8655
      UseMnemonic     =   0   'False
      Width           =   1275
   End
   Begin VB.Label LB_NumeroDocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      TabIndex        =   131
      Top             =   9255
      Width           =   525
   End
   Begin VB.Label LB_ND 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO DOCUMENTO:"
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
      Left            =   120
      TabIndex        =   130
      Top             =   9135
      UseMnemonic     =   0   'False
      Width           =   1605
   End
   Begin VB.Label LB_SeuNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   129
      Top             =   9255
      Width           =   525
   End
   Begin VB.Label LB_NN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO NÚMERO:"
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
      Left            =   3360
      TabIndex        =   128
      Top             =   9135
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label LB_NossoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   3360
      TabIndex        =   127
      Top             =   9255
      Width           =   525
   End
   Begin VB.Label LB_SN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEU NÚMERO:"
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
      Left            =   2040
      TabIndex        =   126
      Top             =   9135
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATA"
      BeginProperty Font 
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
      Left            =   4920
      TabIndex        =   125
      Top             =   9255
      Width           =   1065
   End
   Begin VB.Label LB_TI 
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
      Index           =   4
      Left            =   4920
      TabIndex        =   124
      Top             =   9135
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Banco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRADESCO"
      BeginProperty Font 
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
      TabIndex        =   123
      Top             =   9255
      Width           =   1095
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO:"
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
      Left            =   7920
      TabIndex        =   122
      Top             =   9135
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_EmCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      TabIndex        =   121
      Top             =   9735
      Width           =   345
   End
   Begin VB.Label LB_EC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM CARTEIRA:"
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
      Left            =   120
      TabIndex        =   120
      Top             =   9615
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_EmCobranca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   119
      Top             =   9735
      Width           =   345
   End
   Begin VB.Label LB_Observacoes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO TEM"
      BeginProperty Font 
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
      TabIndex        =   118
      Top             =   9735
      Width           =   885
   End
   Begin VB.Label LB_OB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES:"
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
      Left            =   2760
      TabIndex        =   117
      Top             =   9615
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Line LHIB 
      Index           =   4
      X1              =   0
      X2              =   11160
      Y1              =   9975
      Y2              =   9975
   End
   Begin VB.Label LB_EB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM COBRANÇA:"
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
      Left            =   1440
      TabIndex        =   116
      Top             =   7935
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label LB_DataEmissao 
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
      Index           =   3
      Left            =   120
      TabIndex        =   115
      Top             =   7095
      Width           =   960
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE EMISSÃO:"
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
      Left            =   120
      TabIndex        =   114
      Top             =   6975
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Line LHIC 
      Index           =   3
      X1              =   0
      X2              =   11160
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Label LB_DataVencimento 
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
      Index           =   3
      Left            =   1560
      TabIndex        =   113
      Top             =   7095
      Width           =   960
   End
   Begin VB.Label LB_DV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VENCIMENTO:"
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
      Left            =   1560
      TabIndex        =   112
      Top             =   6975
      UseMnemonic     =   0   'False
      Width           =   1365
   End
   Begin VB.Label LB_Movimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À Pagar"
      BeginProperty Font 
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
      Left            =   3480
      TabIndex        =   111
      Top             =   7095
      Width           =   705
   End
   Begin VB.Label LB_MO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO:"
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
      Left            =   3480
      TabIndex        =   110
      Top             =   6975
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
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
      TabIndex        =   109
      Top             =   7095
      Width           =   1125
   End
   Begin VB.Label LB_VA 
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
      Index           =   3
      Left            =   4560
      TabIndex        =   108
      Top             =   6975
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Origem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL CONEXÕES DE AÇO LTDA."
      BeginProperty Font 
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
      Left            =   6000
      TabIndex        =   107
      Top             =   7095
      Width           =   3720
   End
   Begin VB.Label LB_OR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEM OU DESTINO:"
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
      TabIndex        =   106
      Top             =   6975
      UseMnemonic     =   0   'False
      Width           =   1275
   End
   Begin VB.Label LB_NumeroDocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      TabIndex        =   105
      Top             =   7575
      Width           =   525
   End
   Begin VB.Label LB_ND 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO DOCUMENTO:"
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
      Left            =   120
      TabIndex        =   104
      Top             =   7455
      UseMnemonic     =   0   'False
      Width           =   1605
   End
   Begin VB.Label LB_SeuNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   103
      Top             =   7575
      Width           =   525
   End
   Begin VB.Label LB_NN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO NÚMERO:"
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
      Left            =   3360
      TabIndex        =   102
      Top             =   7455
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label LB_NossoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   3360
      TabIndex        =   101
      Top             =   7575
      Width           =   525
   End
   Begin VB.Label LB_SN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEU NÚMERO:"
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
      Left            =   2040
      TabIndex        =   100
      Top             =   7455
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATA"
      BeginProperty Font 
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
      Left            =   4920
      TabIndex        =   99
      Top             =   7575
      Width           =   1065
   End
   Begin VB.Label LB_TI 
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
      Index           =   3
      Left            =   4920
      TabIndex        =   98
      Top             =   7455
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Banco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRADESCO"
      BeginProperty Font 
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
      TabIndex        =   97
      Top             =   7575
      Width           =   1095
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO:"
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
      Left            =   7920
      TabIndex        =   96
      Top             =   7455
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_EmCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      TabIndex        =   95
      Top             =   8055
      Width           =   345
   End
   Begin VB.Label LB_EC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM CARTEIRA:"
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
      Left            =   120
      TabIndex        =   94
      Top             =   7935
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_EmCobranca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   93
      Top             =   8055
      Width           =   345
   End
   Begin VB.Label LB_Observacoes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO TEM"
      BeginProperty Font 
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
      TabIndex        =   92
      Top             =   8055
      Width           =   885
   End
   Begin VB.Label LB_OB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES:"
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
      Left            =   2760
      TabIndex        =   91
      Top             =   7935
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Line LHIB 
      Index           =   3
      X1              =   0
      X2              =   11160
      Y1              =   8295
      Y2              =   8295
   End
   Begin VB.Label LB_EB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM COBRANÇA:"
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
      Left            =   1440
      TabIndex        =   90
      Top             =   6255
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label LB_DataEmissao 
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
      Index           =   2
      Left            =   120
      TabIndex        =   89
      Top             =   5415
      Width           =   960
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE EMISSÃO:"
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
      Left            =   120
      TabIndex        =   88
      Top             =   5295
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Line LHIC 
      Index           =   2
      X1              =   0
      X2              =   11160
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label LB_DataVencimento 
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
      Index           =   2
      Left            =   1560
      TabIndex        =   87
      Top             =   5415
      Width           =   960
   End
   Begin VB.Label LB_DV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VENCIMENTO:"
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
      Left            =   1560
      TabIndex        =   86
      Top             =   5295
      UseMnemonic     =   0   'False
      Width           =   1365
   End
   Begin VB.Label LB_Movimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À Pagar"
      BeginProperty Font 
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
      Left            =   3480
      TabIndex        =   85
      Top             =   5415
      Width           =   705
   End
   Begin VB.Label LB_MO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO:"
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
      Left            =   3480
      TabIndex        =   84
      Top             =   5295
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
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
      TabIndex        =   83
      Top             =   5415
      Width           =   1125
   End
   Begin VB.Label LB_VA 
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
      Index           =   2
      Left            =   4560
      TabIndex        =   82
      Top             =   5295
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Origem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL CONEXÕES DE AÇO LTDA."
      BeginProperty Font 
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
      Left            =   6000
      TabIndex        =   81
      Top             =   5415
      Width           =   3720
   End
   Begin VB.Label LB_OR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEM OU DESTINO:"
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
      Left            =   6000
      TabIndex        =   80
      Top             =   5295
      UseMnemonic     =   0   'False
      Width           =   1275
   End
   Begin VB.Label LB_NumeroDocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      TabIndex        =   79
      Top             =   5895
      Width           =   525
   End
   Begin VB.Label LB_ND 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO DOCUMENTO:"
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
      Left            =   120
      TabIndex        =   78
      Top             =   5775
      UseMnemonic     =   0   'False
      Width           =   1605
   End
   Begin VB.Label LB_SeuNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   77
      Top             =   5895
      Width           =   525
   End
   Begin VB.Label LB_NN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO NÚMERO:"
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
      Left            =   3360
      TabIndex        =   76
      Top             =   5775
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label LB_NossoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   3360
      TabIndex        =   75
      Top             =   5895
      Width           =   525
   End
   Begin VB.Label LB_SN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEU NÚMERO:"
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
      Left            =   2040
      TabIndex        =   74
      Top             =   5775
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATA"
      BeginProperty Font 
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
      Left            =   4920
      TabIndex        =   73
      Top             =   5895
      Width           =   1065
   End
   Begin VB.Label LB_TI 
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
      Index           =   2
      Left            =   4920
      TabIndex        =   72
      Top             =   5775
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Banco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRADESCO"
      BeginProperty Font 
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
      TabIndex        =   71
      Top             =   5895
      Width           =   1095
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO:"
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
      Left            =   7920
      TabIndex        =   70
      Top             =   5775
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_EmCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Top             =   6375
      Width           =   345
   End
   Begin VB.Label LB_EC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM CARTEIRA:"
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
      Left            =   120
      TabIndex        =   68
      Top             =   6255
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_EmCobranca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   67
      Top             =   6375
      Width           =   345
   End
   Begin VB.Label LB_Observacoes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO TEM"
      BeginProperty Font 
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
      TabIndex        =   66
      Top             =   6375
      Width           =   885
   End
   Begin VB.Label LB_OB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES:"
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
      Left            =   2760
      TabIndex        =   65
      Top             =   6255
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Line LHIB 
      Index           =   2
      X1              =   0
      X2              =   11160
      Y1              =   6615
      Y2              =   6615
   End
   Begin VB.Label LB_EB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM COBRANÇA:"
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
      Left            =   1440
      TabIndex        =   64
      Top             =   4575
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label LB_DataEmissao 
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
      Index           =   1
      Left            =   120
      TabIndex        =   63
      Top             =   3735
      Width           =   960
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE EMISSÃO:"
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
      TabIndex        =   62
      Top             =   3615
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Line LHIC 
      Index           =   1
      X1              =   0
      X2              =   11160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label LB_DataVencimento 
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
      Index           =   1
      Left            =   1560
      TabIndex        =   61
      Top             =   3735
      Width           =   960
   End
   Begin VB.Label LB_DV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VENCIMENTO:"
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
      Left            =   1560
      TabIndex        =   60
      Top             =   3615
      UseMnemonic     =   0   'False
      Width           =   1365
   End
   Begin VB.Label LB_Movimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À Pagar"
      BeginProperty Font 
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
      Left            =   3480
      TabIndex        =   59
      Top             =   3735
      Width           =   705
   End
   Begin VB.Label LB_MO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO:"
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
      Left            =   3480
      TabIndex        =   58
      Top             =   3615
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
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
      TabIndex        =   57
      Top             =   3735
      Width           =   1125
   End
   Begin VB.Label LB_VA 
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
      Index           =   1
      Left            =   4560
      TabIndex        =   56
      Top             =   3615
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Origem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL CONEXÕES DE AÇO LTDA."
      BeginProperty Font 
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
      Left            =   6000
      TabIndex        =   55
      Top             =   3735
      Width           =   3720
   End
   Begin VB.Label LB_OR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEM OU DESTINO:"
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
      Left            =   6000
      TabIndex        =   54
      Top             =   3615
      UseMnemonic     =   0   'False
      Width           =   1275
   End
   Begin VB.Label LB_NumeroDocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      TabIndex        =   53
      Top             =   4215
      Width           =   525
   End
   Begin VB.Label LB_ND 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO DOCUMENTO:"
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
      TabIndex        =   52
      Top             =   4095
      UseMnemonic     =   0   'False
      Width           =   1605
   End
   Begin VB.Label LB_SeuNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   51
      Top             =   4215
      Width           =   525
   End
   Begin VB.Label LB_NN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO NÚMERO:"
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
      Left            =   3360
      TabIndex        =   50
      Top             =   4095
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label LB_NossoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   3360
      TabIndex        =   49
      Top             =   4215
      Width           =   525
   End
   Begin VB.Label LB_SN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEU NÚMERO:"
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
      Left            =   2040
      TabIndex        =   48
      Top             =   4095
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATA"
      BeginProperty Font 
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
      Left            =   4920
      TabIndex        =   47
      Top             =   4215
      Width           =   1065
   End
   Begin VB.Label LB_TI 
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
      Left            =   4920
      TabIndex        =   46
      Top             =   4095
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Banco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRADESCO"
      BeginProperty Font 
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
      TabIndex        =   45
      Top             =   4215
      Width           =   1095
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO:"
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
      Left            =   7920
      TabIndex        =   44
      Top             =   4095
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_EmCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      TabIndex        =   43
      Top             =   4695
      Width           =   345
   End
   Begin VB.Label LB_EC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM CARTEIRA:"
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
      TabIndex        =   42
      Top             =   4575
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_EmCobranca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   41
      Top             =   4695
      Width           =   345
   End
   Begin VB.Label LB_Observacoes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO TEM"
      BeginProperty Font 
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
      TabIndex        =   40
      Top             =   4695
      Width           =   885
   End
   Begin VB.Label LB_OB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES:"
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
      Left            =   2760
      TabIndex        =   39
      Top             =   4575
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Line LHIB 
      Index           =   1
      X1              =   0
      X2              =   11160
      Y1              =   4935
      Y2              =   4935
   End
   Begin VB.Label LB_EB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM COBRANÇA:"
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
      Index           =   0
      Left            =   1440
      TabIndex        =   38
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Line LHIB 
      Index           =   0
      X1              =   0
      X2              =   11160
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label LB_OB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES:"
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
      Index           =   0
      Left            =   2760
      TabIndex        =   37
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label LB_Observacoes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÃO TEM"
      BeginProperty Font 
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
      Top             =   3000
      Width           =   885
   End
   Begin VB.Label LB_EmCobranca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      Left            =   1440
      TabIndex        =   35
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label LB_EC 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EM CARTEIRA:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_EmCarteira 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SIM"
      BeginProperty Font 
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
      TabIndex        =   33
      Top             =   3000
      Width           =   345
   End
   Begin VB.Label LB_BA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BANCO:"
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
      Index           =   0
      Left            =   7920
      TabIndex        =   32
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_Banco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BRADESCO"
      BeginProperty Font 
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
      TabIndex        =   31
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label LB_TI 
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
      Index           =   0
      Left            =   4920
      TabIndex        =   30
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUPLICATA"
      BeginProperty Font 
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
      Left            =   4920
      TabIndex        =   29
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label LB_SN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEU NÚMERO:"
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
      Index           =   0
      Left            =   2040
      TabIndex        =   28
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_NossoNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   3360
      TabIndex        =   27
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label LB_NN 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO NÚMERO:"
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
      Index           =   0
      Left            =   3360
      TabIndex        =   26
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label LB_SeuNumero 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   25
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label LB_ND 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO DOCUMENTO:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   2400
      UseMnemonic     =   0   'False
      Width           =   1605
   End
   Begin VB.Label LB_NumeroDocumento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
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
      TabIndex        =   23
      Top             =   2520
      Width           =   525
   End
   Begin VB.Label LB_OR 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORIGEM OU DESTINO:"
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
      Index           =   0
      Left            =   6000
      TabIndex        =   22
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   1275
   End
   Begin VB.Label LB_Origem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL CONEXÕES DE AÇO LTDA."
      BeginProperty Font 
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
      Left            =   6000
      TabIndex        =   21
      Top             =   2040
      Width           =   3720
   End
   Begin VB.Label LB_VA 
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
      Index           =   0
      Left            =   4560
      TabIndex        =   20
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.000.000,00"
      BeginProperty Font 
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
      TabIndex        =   19
      Top             =   2040
      Width           =   1125
   End
   Begin VB.Label LB_MO 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOVIMENTO:"
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
      Index           =   0
      Left            =   3480
      TabIndex        =   18
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Movimento 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "À Pagar"
      BeginProperty Font 
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
      Left            =   3480
      TabIndex        =   17
      Top             =   2040
      Width           =   705
   End
   Begin VB.Label LB_DV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE VENCIMENTO:"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   16
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   1365
   End
   Begin VB.Label LB_DataVencimento 
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
      Index           =   0
      Left            =   1560
      TabIndex        =   15
      Top             =   2040
      Width           =   960
   End
   Begin VB.Line LHIC 
      Index           =   0
      X1              =   0
      X2              =   11160
      Y1              =   1900
      Y2              =   1900
   End
   Begin VB.Label LB_DE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE EMISSÃO:"
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      UseMnemonic     =   0   'False
      Width           =   1095
   End
   Begin VB.Label LB_DataEmissao 
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
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   960
   End
   Begin VB.Line LV 
      Index           =   4
      X1              =   2760
      X2              =   2760
      Y1              =   13680
      Y2              =   14280
   End
   Begin VB.Label LB_QuantDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pacote"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2880
      TabIndex        =   11
      Top             =   13920
      Width           =   615
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE DE DOCUMENTOS:"
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
      Left            =   2880
      TabIndex        =   10
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1905
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DESTA FOLHA:"
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
      Left            =   5640
      TabIndex        =   9
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1245
   End
   Begin VB.Label LB_ValFol 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1234"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   8
      Top             =   13920
      Width           =   420
   End
   Begin VB.Line LV 
      Index           =   1
      X1              =   5520
      X2              =   5520
      Y1              =   13680
      Y2              =   14280
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL ACUMULADO:"
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
      Left            =   8400
      TabIndex        =   7
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1200
   End
   Begin VB.Label LB_ValAcu 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8400
      TabIndex        =   6
      Top             =   13920
      Width           =   705
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DESTA PÁGINA:"
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
      Left            =   0
      TabIndex        =   5
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label LB_NumPag 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOSSO MOTORISTA"
      BeginProperty Font 
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
      TabIndex        =   4
      Top             =   13920
      Width           =   1890
   End
   Begin VB.Line LV 
      Index           =   21
      X1              =   8280
      X2              =   8280
      Y1              =   13680
      Y2              =   14280
   End
   Begin VB.Line LH 
      Index           =   7
      X1              =   0
      X2              =   11160
      Y1              =   14310
      Y2              =   14310
   End
   Begin VB.Line LH 
      Index           =   6
      X1              =   0
      X2              =   11160
      Y1              =   14280
      Y2              =   14280
   End
   Begin VB.Line LH 
      Index           =   17
      X1              =   0
      X2              =   11160
      Y1              =   13650
      Y2              =   13650
   End
   Begin VB.Line LH 
      Index           =   18
      X1              =   0
      X2              =   11160
      Y1              =   13680
      Y2              =   13680
   End
   Begin VB.Line LV 
      Index           =   2
      X1              =   5520
      X2              =   5520
      Y1              =   150
      Y2              =   1590
   End
   Begin VB.Line LH 
      Index           =   1
      X1              =   0
      X2              =   11160
      Y1              =   1590
      Y2              =   1590
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
      Top             =   990
      Width           =   1785
   End
   Begin VB.Line LH 
      Index           =   0
      X1              =   0
      X2              =   11160
      Y1              =   150
      Y2              =   150
   End
   Begin VB.Label LB_Titulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RELATÓRIO DE CONTAS À PAGAR"
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
      Left            =   6360
      TabIndex        =   2
      Top             =   390
      Width           =   4065
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
      Left            =   120
      TabIndex        =   1
      Top             =   1250
      Width           =   5220
   End
   Begin VB.Line LH 
      Index           =   9
      X1              =   0
      X2              =   11160
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line LH 
      Index           =   10
      X1              =   0
      X2              =   11160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label LB_Data 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data: 01/06/2001"
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
      Left            =   7440
      TabIndex        =   0
      Top             =   990
      Width           =   1905
   End
End
Attribute VB_Name = "IT_Contas_Relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
