VERSION 5.00
Begin VB.Form IT_Equipamentos 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   11295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16005
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
   ScaleHeight     =   11295
   ScaleWidth      =   16005
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
      Picture         =   "IT_Equipamentos.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   3
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
      TabIndex        =   198
      Top             =   1250
      Width           =   5220
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data:"
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
      Left            =   9600
      TabIndex        =   197
      Top             =   960
      Width           =   615
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   196
      Top             =   10560
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   17
      Left            =   3360
      TabIndex        =   195
      Top             =   10560
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   194
      Top             =   10560
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   193
      Top             =   10560
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   192
      Top             =   10560
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   191
      Top             =   10560
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   190
      Top             =   10560
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   17
      Left            =   13680
      TabIndex        =   189
      Top             =   10560
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   188
      Top             =   10560
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   187
      Top             =   10200
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   16
      Left            =   3360
      TabIndex        =   186
      Top             =   10200
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   185
      Top             =   10200
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   184
      Top             =   10200
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   183
      Top             =   10200
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   182
      Top             =   10200
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   181
      Top             =   10200
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   16
      Left            =   13680
      TabIndex        =   180
      Top             =   10200
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   179
      Top             =   10200
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   178
      Top             =   9840
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   15
      Left            =   3360
      TabIndex        =   177
      Top             =   9840
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   176
      Top             =   9840
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   175
      Top             =   9840
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   174
      Top             =   9840
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   173
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   172
      Top             =   9840
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   15
      Left            =   13680
      TabIndex        =   171
      Top             =   9840
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   170
      Top             =   9840
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   169
      Top             =   9480
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   14
      Left            =   3360
      TabIndex        =   168
      Top             =   9480
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   167
      Top             =   9480
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   166
      Top             =   9480
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   165
      Top             =   9480
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   164
      Top             =   9480
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   163
      Top             =   9480
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   14
      Left            =   13680
      TabIndex        =   162
      Top             =   9480
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   161
      Top             =   9480
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   160
      Top             =   9120
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   13
      Left            =   3360
      TabIndex        =   159
      Top             =   9120
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   158
      Top             =   9120
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   157
      Top             =   9120
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   156
      Top             =   9120
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   155
      Top             =   9120
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   154
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   13
      Left            =   13680
      TabIndex        =   153
      Top             =   9120
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   152
      Top             =   9120
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   151
      Top             =   8760
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   12
      Left            =   3360
      TabIndex        =   150
      Top             =   8760
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   149
      Top             =   8760
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   148
      Top             =   8760
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   147
      Top             =   8760
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   146
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   145
      Top             =   8760
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   12
      Left            =   13680
      TabIndex        =   144
      Top             =   8760
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   143
      Top             =   8760
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   142
      Top             =   8400
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   11
      Left            =   3360
      TabIndex        =   141
      Top             =   8400
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   140
      Top             =   8400
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   139
      Top             =   8400
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   138
      Top             =   8400
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   137
      Top             =   8400
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   136
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   11
      Left            =   13680
      TabIndex        =   135
      Top             =   8400
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   134
      Top             =   8400
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   133
      Top             =   8040
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   10
      Left            =   3360
      TabIndex        =   132
      Top             =   8040
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   131
      Top             =   8040
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      Left            =   6120
      TabIndex        =   130
      Top             =   8040
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   129
      Top             =   8040
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   128
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   127
      Top             =   8040
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   10
      Left            =   13680
      TabIndex        =   126
      Top             =   8040
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   125
      Top             =   8040
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   124
      Top             =   7680
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   9
      Left            =   3360
      TabIndex        =   123
      Top             =   7680
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   122
      Top             =   7680
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   121
      Top             =   7680
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   120
      Top             =   7680
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   119
      Top             =   7680
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   118
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   9
      Left            =   13680
      TabIndex        =   117
      Top             =   7680
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   116
      Top             =   7680
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   115
      Top             =   7320
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   8
      Left            =   3360
      TabIndex        =   114
      Top             =   7320
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   113
      Top             =   7320
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   112
      Top             =   7320
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   111
      Top             =   7320
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   110
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   109
      Top             =   7320
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   8
      Left            =   13680
      TabIndex        =   108
      Top             =   7320
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   107
      Top             =   7320
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   106
      Top             =   6960
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   7
      Left            =   3360
      TabIndex        =   105
      Top             =   6960
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   104
      Top             =   6960
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   103
      Top             =   6960
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   102
      Top             =   6960
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   101
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   100
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Index           =   7
      Left            =   13680
      TabIndex        =   99
      Top             =   6960
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   98
      Top             =   6960
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   97
      Top             =   6600
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   6
      Left            =   3360
      TabIndex        =   96
      Top             =   6600
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   95
      Top             =   6600
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   94
      Top             =   6600
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   93
      Top             =   6600
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   92
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   91
      Top             =   6600
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Left            =   13680
      TabIndex        =   90
      Top             =   6600
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   89
      Top             =   6600
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   88
      Top             =   6240
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   5
      Left            =   3360
      TabIndex        =   87
      Top             =   6240
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   86
      Top             =   6240
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   85
      Top             =   6240
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   84
      Top             =   6240
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   83
      Top             =   6240
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   82
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Left            =   13680
      TabIndex        =   81
      Top             =   6240
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   80
      Top             =   6240
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   79
      Top             =   5880
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   4
      Left            =   3360
      TabIndex        =   78
      Top             =   5880
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   77
      Top             =   5880
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   76
      Top             =   5880
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   75
      Top             =   5880
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   74
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   73
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Left            =   13680
      TabIndex        =   72
      Top             =   5880
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   71
      Top             =   5880
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      Top             =   5520
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   3
      Left            =   3360
      TabIndex        =   69
      Top             =   5520
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   68
      Top             =   5520
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   67
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   66
      Top             =   5520
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   65
      Top             =   5520
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   64
      Top             =   5520
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Left            =   13680
      TabIndex        =   63
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   62
      Top             =   5520
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   61
      Top             =   5160
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   2
      Left            =   3360
      TabIndex        =   60
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   59
      Top             =   5160
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   58
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   57
      Top             =   5160
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   56
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   55
      Top             =   5160
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Left            =   13680
      TabIndex        =   54
      Top             =   5160
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   53
      Top             =   5160
      Width           =   555
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   52
      Top             =   4800
      Width           =   2265
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   1
      Left            =   3360
      TabIndex        =   51
      Top             =   4800
      Width           =   960
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   50
      Top             =   4800
      Width           =   960
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   49
      Top             =   4800
      Width           =   375
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   48
      Top             =   4800
      Width           =   1035
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   47
      Top             =   4800
      Width           =   975
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   46
      Top             =   4800
      Width           =   495
   End
   Begin VB.Label LB_Aviso 
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
      Left            =   13680
      TabIndex        =   45
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   44
      Top             =   4800
      Width           =   555
   End
   Begin VB.Label LB_Usuario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admin"
      BeginProperty Font 
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
      Left            =   14280
      TabIndex        =   43
      Top             =   4440
      Width           =   555
   End
   Begin VB.Label LB_Aviso 
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
      Left            =   13680
      TabIndex        =   42
      Top             =   4440
      Width           =   345
   End
   Begin VB.Label LB_TipMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO:"
      BeginProperty Font 
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
      Left            =   10440
      TabIndex        =   41
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label LB_ConMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
      BeginProperty Font 
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
      Left            =   9000
      TabIndex        =   40
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label LB_EmpMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
      BeginProperty Font 
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
      Left            =   7200
      TabIndex        =   39
      Top             =   4440
      Width           =   1035
   End
   Begin VB.Label LB_ValorMan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      BeginProperty Font 
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
      TabIndex        =   38
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label LB_DataProxima 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0000"
      BeginProperty Font 
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
      Left            =   4680
      TabIndex        =   37
      Top             =   4440
      Width           =   960
   End
   Begin VB.Label LB_DataUltima 
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
      Index           =   0
      Left            =   3360
      TabIndex        =   36
      Top             =   4440
      Width           =   960
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USUÁRIO:"
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
      Left            =   14280
      TabIndex        =   35
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   570
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AVISO:"
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
      Left            =   13680
      TabIndex        =   34
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   375
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   28
      Left            =   10440
      TabIndex        =   33
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTATO:"
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
      Left            =   9000
      TabIndex        =   32
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA:"
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
      Left            =   7200
      TabIndex        =   31
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   600
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   25
      Left            =   6120
      TabIndex        =   30
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   420
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DA PRÓXIMA:"
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
      Index           =   24
      Left            =   4680
      TabIndex        =   29
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   1110
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DA ÚLTIMA:"
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
      Left            =   3360
      TabIndex        =   28
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   1005
   End
   Begin VB.Line LHIC 
      Index           =   3
      X1              =   0
      X2              =   15800
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line LHIC 
      Index           =   2
      X1              =   0
      X2              =   15800
      Y1              =   11160
      Y2              =   11160
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MANUTENÇÕES CONFIGURADAS PARA ESTE EQUIPAMENTO:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   27
      Top             =   3720
      Width           =   5820
   End
   Begin VB.Line LHIC 
      Index           =   1
      X1              =   0
      X2              =   15800
      Y1              =   3945
      Y2              =   3945
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMAÇÕES SOBRE O EQUIPAMENTO:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   26
      Top             =   1920
      Width           =   3945
   End
   Begin VB.Line LHIB 
      Index           =   0
      X1              =   0
      X2              =   15800
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FUNCIONÁRIO QUE ESTÁ OPERANDO O EQUIPAMENTO ATUALMENTE:"
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
      Left            =   9000
      TabIndex        =   25
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   4035
   End
   Begin VB.Label LB_Funcionario 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHIFRUDO DA SILVA"
      BeginProperty Font 
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
      TabIndex        =   24
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
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
      TabIndex        =   23
      Top             =   4080
      UseMnemonic     =   0   'False
      Width           =   1395
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÍTEM DE MANUTENÇÃO:"
      BeginProperty Font 
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
      TabIndex        =   22
      Top             =   4440
      Width           =   2265
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROCESSO:"
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
      Left            =   5040
      TabIndex        =   21
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   690
   End
   Begin VB.Label LB_Processo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FURAÇÃO"
      BeginProperty Font 
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
      TabIndex        =   20
      Top             =   2760
      Width           =   945
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOME DA FOTO:"
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
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   945
   End
   Begin VB.Label LB_Custo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$0,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   18
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTO/HORA DO EQUIPAMENTO:"
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
      Index           =   16
      Left            =   2640
      TabIndex        =   17
      Top             =   2640
      UseMnemonic     =   0   'False
      Width           =   1935
   End
   Begin VB.Label LB_Foto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XUXU"
      BeginProperty Font 
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
      TabIndex        =   16
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VALOR DO EQUIPAMENTO:"
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
      Left            =   12840
      TabIndex        =   15
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1530
   End
   Begin VB.Label LB_Valor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "$1000,00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   12840
      TabIndex        =   14
      Top             =   2280
      Width           =   795
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FABRICANTE:"
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
      TabIndex        =   13
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   780
   End
   Begin VB.Label LB_Fabricante 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONESTEEL"
      BeginProperty Font 
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
      TabIndex        =   12
      Top             =   2280
      Width           =   1170
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MODELO:"
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
      Left            =   7320
      TabIndex        =   11
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   540
   End
   Begin VB.Label LB_Modelo 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   7320
      TabIndex        =   10
      Top             =   2280
      Width           =   405
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO EQUIPAMENTO:"
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
      Left            =   5040
      TabIndex        =   9
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1200
   End
   Begin VB.Label LB_Tipo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FERRAMENTA"
      BeginProperty Font 
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
      TabIndex        =   8
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOME DO EQUIPAMENTO:"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1500
   End
   Begin VB.Label LB_Nome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FER_ABC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   900
   End
   Begin VB.Line LHIC 
      Index           =   0
      X1              =   0
      X2              =   15800
      Y1              =   2145
      Y2              =   2145
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÚMERO DO EQUIPAMENTO NE:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1875
   End
   Begin VB.Label LB_NE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
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
      TabIndex        =   4
      Top             =   2280
      Width           =   315
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
      X2              =   15800
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
      TabIndex        =   2
      Top             =   990
      Width           =   1785
   End
   Begin VB.Line LH 
      Index           =   0
      X1              =   0
      X2              =   15800
      Y1              =   150
      Y2              =   150
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RELATÓRIO DE CONFIGURAÇÕES DE EQUIPAMENTOS"
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
      Left            =   7440
      TabIndex        =   1
      Top             =   390
      Width           =   6435
   End
   Begin VB.Line LH 
      Index           =   9
      X1              =   0
      X2              =   15800
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line LH 
      Index           =   10
      X1              =   0
      X2              =   15800
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Label LB_Data 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/06/2001"
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
      Left            =   10320
      TabIndex        =   0
      Top             =   960
      Width           =   1200
   End
End
Attribute VB_Name = "IT_Equipamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
