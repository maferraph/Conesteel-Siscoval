VERSION 5.00
Begin VB.Form Tela_Fabrica_OM_Relatorio 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "ITEM"
   ClientHeight    =   16005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   ScaleHeight     =   16005
   ScaleWidth      =   11595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   19
      Left            =   10080
      TabIndex        =   236
      Top             =   8520
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   19
      Left            =   8640
      TabIndex        =   235
      Top             =   8520
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   19
      Left            =   7080
      TabIndex        =   234
      Top             =   8520
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   19
      Left            =   5640
      TabIndex        =   233
      Top             =   8520
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   19
      Left            =   4200
      TabIndex        =   232
      Top             =   8520
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   19
      Left            =   1560
      TabIndex        =   231
      Top             =   8520
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   19
      Left            =   600
      TabIndex        =   230
      Top             =   8520
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   19
      Left            =   120
      TabIndex        =   229
      Top             =   8520
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   18
      Left            =   10080
      TabIndex        =   228
      Top             =   8280
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   18
      Left            =   8640
      TabIndex        =   227
      Top             =   8280
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   18
      Left            =   7080
      TabIndex        =   226
      Top             =   8280
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   18
      Left            =   5640
      TabIndex        =   225
      Top             =   8280
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   18
      Left            =   4200
      TabIndex        =   224
      Top             =   8280
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   18
      Left            =   1560
      TabIndex        =   223
      Top             =   8280
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   18
      Left            =   600
      TabIndex        =   222
      Top             =   8280
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   18
      Left            =   120
      TabIndex        =   221
      Top             =   8280
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   17
      Left            =   10080
      TabIndex        =   220
      Top             =   8040
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   17
      Left            =   8640
      TabIndex        =   219
      Top             =   8040
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   17
      Left            =   7080
      TabIndex        =   218
      Top             =   8040
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   17
      Left            =   5640
      TabIndex        =   217
      Top             =   8040
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   17
      Left            =   4200
      TabIndex        =   216
      Top             =   8040
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   17
      Left            =   1560
      TabIndex        =   215
      Top             =   8040
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   17
      Left            =   600
      TabIndex        =   214
      Top             =   8040
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   17
      Left            =   120
      TabIndex        =   213
      Top             =   8040
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   16
      Left            =   10080
      TabIndex        =   212
      Top             =   7800
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   16
      Left            =   8640
      TabIndex        =   211
      Top             =   7800
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   16
      Left            =   7080
      TabIndex        =   210
      Top             =   7800
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   16
      Left            =   5640
      TabIndex        =   209
      Top             =   7800
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   16
      Left            =   4200
      TabIndex        =   208
      Top             =   7800
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   16
      Left            =   1560
      TabIndex        =   207
      Top             =   7800
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   16
      Left            =   600
      TabIndex        =   206
      Top             =   7800
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   16
      Left            =   120
      TabIndex        =   205
      Top             =   7800
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   15
      Left            =   10080
      TabIndex        =   204
      Top             =   7560
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   15
      Left            =   8640
      TabIndex        =   203
      Top             =   7560
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   15
      Left            =   7080
      TabIndex        =   202
      Top             =   7560
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   15
      Left            =   5640
      TabIndex        =   201
      Top             =   7560
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   15
      Left            =   4200
      TabIndex        =   200
      Top             =   7560
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   15
      Left            =   1560
      TabIndex        =   199
      Top             =   7560
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   15
      Left            =   600
      TabIndex        =   198
      Top             =   7560
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   15
      Left            =   120
      TabIndex        =   197
      Top             =   7560
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   14
      Left            =   10080
      TabIndex        =   196
      Top             =   7320
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   14
      Left            =   8640
      TabIndex        =   195
      Top             =   7320
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   14
      Left            =   7080
      TabIndex        =   194
      Top             =   7320
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   14
      Left            =   5640
      TabIndex        =   193
      Top             =   7320
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   14
      Left            =   4200
      TabIndex        =   192
      Top             =   7320
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   14
      Left            =   1560
      TabIndex        =   191
      Top             =   7320
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   14
      Left            =   600
      TabIndex        =   190
      Top             =   7320
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   14
      Left            =   120
      TabIndex        =   189
      Top             =   7320
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   13
      Left            =   10080
      TabIndex        =   188
      Top             =   7080
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   13
      Left            =   8640
      TabIndex        =   187
      Top             =   7080
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   13
      Left            =   7080
      TabIndex        =   186
      Top             =   7080
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   13
      Left            =   5640
      TabIndex        =   185
      Top             =   7080
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   13
      Left            =   4200
      TabIndex        =   184
      Top             =   7080
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   13
      Left            =   1560
      TabIndex        =   183
      Top             =   7080
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   13
      Left            =   600
      TabIndex        =   182
      Top             =   7080
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   13
      Left            =   120
      TabIndex        =   181
      Top             =   7080
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   12
      Left            =   10080
      TabIndex        =   180
      Top             =   6840
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   12
      Left            =   8640
      TabIndex        =   179
      Top             =   6840
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   12
      Left            =   7080
      TabIndex        =   178
      Top             =   6840
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   12
      Left            =   5640
      TabIndex        =   177
      Top             =   6840
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   12
      Left            =   4200
      TabIndex        =   176
      Top             =   6840
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   12
      Left            =   1560
      TabIndex        =   175
      Top             =   6840
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   12
      Left            =   600
      TabIndex        =   174
      Top             =   6840
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   12
      Left            =   120
      TabIndex        =   173
      Top             =   6840
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   11
      Left            =   10080
      TabIndex        =   172
      Top             =   6600
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   11
      Left            =   8640
      TabIndex        =   171
      Top             =   6600
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   11
      Left            =   7080
      TabIndex        =   170
      Top             =   6600
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   11
      Left            =   5640
      TabIndex        =   169
      Top             =   6600
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   11
      Left            =   4200
      TabIndex        =   168
      Top             =   6600
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   11
      Left            =   1560
      TabIndex        =   167
      Top             =   6600
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   11
      Left            =   600
      TabIndex        =   166
      Top             =   6600
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   11
      Left            =   120
      TabIndex        =   165
      Top             =   6600
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Index           =   10
      Left            =   10080
      TabIndex        =   164
      Top             =   6360
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Index           =   10
      Left            =   8640
      TabIndex        =   163
      Top             =   6360
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Index           =   10
      Left            =   7080
      TabIndex        =   162
      Top             =   6360
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Index           =   10
      Left            =   5640
      TabIndex        =   161
      Top             =   6360
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Index           =   10
      Left            =   4200
      TabIndex        =   160
      Top             =   6360
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Index           =   10
      Left            =   1560
      TabIndex        =   159
      Top             =   6360
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Index           =   10
      Left            =   600
      TabIndex        =   158
      Top             =   6360
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Index           =   10
      Left            =   120
      TabIndex        =   157
      Top             =   6360
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   156
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   155
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   154
      Top             =   6120
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   153
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   152
      Top             =   6120
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   151
      Top             =   6120
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   150
      Top             =   6120
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   149
      Top             =   6120
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   148
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   147
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   146
      Top             =   5880
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   145
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   144
      Top             =   5880
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   143
      Top             =   5880
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   142
      Top             =   5880
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   141
      Top             =   5880
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   140
      Top             =   5640
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   139
      Top             =   5640
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   138
      Top             =   5640
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   137
      Top             =   5640
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   136
      Top             =   5640
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   135
      Top             =   5640
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   134
      Top             =   5640
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   133
      Top             =   5640
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   132
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   131
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   130
      Top             =   5400
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   129
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   128
      Top             =   5400
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   127
      Top             =   5400
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   126
      Top             =   5400
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   125
      Top             =   5400
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   124
      Top             =   5160
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   123
      Top             =   5160
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   122
      Top             =   5160
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   121
      Top             =   5160
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   120
      Top             =   5160
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   119
      Top             =   5160
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   118
      Top             =   5160
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   117
      Top             =   5160
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   116
      Top             =   4920
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   115
      Top             =   4920
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   114
      Top             =   4920
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   113
      Top             =   4920
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   112
      Top             =   4920
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   111
      Top             =   4920
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   110
      Top             =   4920
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   109
      Top             =   4920
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   108
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   107
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   106
      Top             =   4680
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   105
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   104
      Top             =   4680
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   103
      Top             =   4680
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   102
      Top             =   4680
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   101
      Top             =   4680
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   100
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   99
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   98
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   97
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   96
      Top             =   4440
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   95
      Top             =   4440
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   94
      Top             =   4440
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   93
      Top             =   4440
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   92
      Top             =   4200
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   91
      Top             =   4200
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   90
      Top             =   4200
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   89
      Top             =   4200
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   88
      Top             =   4200
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   87
      Top             =   4200
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   86
      Top             =   4200
      Width           =   810
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   85
      Top             =   4200
      Width           =   330
   End
   Begin VB.Label LB_CCC 
      Alignment       =   2  'Center
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
      Left            =   10080
      TabIndex        =   84
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label LB_ORIC 
      Alignment       =   2  'Center
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
      Left            =   8640
      TabIndex        =   83
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label LB_OFC 
      Alignment       =   2  'Center
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
      Left            =   7080
      TabIndex        =   82
      Top             =   3960
      Width           =   1410
   End
   Begin VB.Label LB_BC 
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
      Left            =   5640
      TabIndex        =   81
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label LB_MC 
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
      Left            =   4200
      TabIndex        =   80
      Top             =   3960
      Width           =   1290
   End
   Begin VB.Label LB_NC 
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
      Left            =   1560
      TabIndex        =   79
      Top             =   3960
      Width           =   2490
   End
   Begin VB.Label LB_QC 
      Alignment       =   2  'Center
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
      Left            =   600
      TabIndex        =   78
      Top             =   3960
      Width           =   810
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dados sobre a válvula:"
      BeginProperty Font 
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
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Line LV 
      Index           =   24
      X1              =   2160
      X2              =   2160
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aparafusado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   76
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CASTELO"
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
      Left            =   2280
      TabIndex        =   75
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   540
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Não aplicado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10200
      TabIndex        =   74
      Top             =   1710
      Width           =   1125
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REVESTIMENTO"
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
      Index           =   59
      Left            =   10200
      TabIndex        =   73
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   930
   End
   Begin VB.Line LV 
      Index           =   23
      X1              =   10080
      X2              =   10080
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "410"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8880
      TabIndex        =   72
      Top             =   1710
      Width           =   315
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INTERNOS"
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
      Left            =   8880
      TabIndex        =   71
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Line LV 
      Index           =   21
      X1              =   8760
      X2              =   8760
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A182 F304L"
      BeginProperty Font 
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
      TabIndex        =   70
      Top             =   1710
      Width           =   1050
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
      Index           =   4
      Left            =   7560
      TabIndex        =   69
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   585
   End
   Begin VB.Line LV 
      Index           =   20
      X1              =   7440
      X2              =   7440
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   68
      Top             =   1680
      Width           =   510
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
      Index           =   2
      Left            =   6240
      TabIndex        =   67
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   405
   End
   Begin VB.Line LV 
      Index           =   14
      X1              =   6120
      X2              =   6120
      Y1              =   1440
      Y2              =   2040
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
      Index           =   1
      Left            =   9000
      TabIndex        =   66
      Top             =   720
      Width           =   615
   End
   Begin VB.Label LB_Fixo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEM DE MONTAGEM"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   4320
      TabIndex        =   65
      Top             =   315
      Width           =   4545
   End
   Begin VB.Line LV 
      Index           =   13
      X1              =   8880
      X2              =   8880
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   120
      Picture         =   "Tela_Fabrica_OM_Relatorio.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   4080
   End
   Begin VB.Label LB_DI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "___________________________"
      BeginProperty Font 
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
      Left            =   8520
      TabIndex        =   64
      Top             =   12240
      Width           =   2835
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      Left            =   7560
      TabIndex        =   63
      Top             =   12240
      Width           =   840
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      Left            =   6600
      TabIndex        =   62
      Top             =   12240
      Width           =   840
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
      Index           =   14
      Left            =   3840
      TabIndex        =   61
      Top             =   12240
      Width           =   1380
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
      Left            =   5400
      TabIndex        =   60
      Top             =   12240
      Width           =   1110
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      TabIndex        =   59
      Top             =   12240
      Width           =   840
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
      Index           =   12
      Left            =   1080
      TabIndex        =   58
      Top             =   12240
      Width           =   1380
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
      Left            =   2640
      TabIndex        =   57
      Top             =   12240
      Width           =   1110
   End
   Begin VB.Label LB_DI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "___________________________"
      BeginProperty Font 
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
      Left            =   8520
      TabIndex        =   56
      Top             =   11880
      Width           =   2835
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      Left            =   7560
      TabIndex        =   55
      Top             =   11880
      Width           =   840
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      Left            =   6600
      TabIndex        =   54
      Top             =   11880
      Width           =   840
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
      Left            =   3840
      TabIndex        =   53
      Top             =   11880
      Width           =   1380
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
      Left            =   5400
      TabIndex        =   52
      Top             =   11880
      Width           =   1110
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      TabIndex        =   51
      Top             =   11880
      Width           =   840
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
      Left            =   1080
      TabIndex        =   50
      Top             =   11880
      Width           =   1380
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
      Left            =   2640
      TabIndex        =   49
      Top             =   11880
      Width           =   1110
   End
   Begin VB.Label LB_DI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "___________________________"
      BeginProperty Font 
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
      Left            =   8520
      TabIndex        =   48
      Top             =   11520
      Width           =   2835
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MOTIVO:"
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
      Left            =   8520
      TabIndex        =   47
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   495
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      Left            =   7560
      TabIndex        =   46
      Top             =   11520
      Width           =   840
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      Left            =   6600
      TabIndex        =   45
      Top             =   11520
      Width           =   840
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "REPROVADAS:"
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
      Left            =   7560
      TabIndex        =   44
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   840
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "APROVADAS:"
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
      TabIndex        =   43
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   750
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA FINAL:"
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
      Left            =   5400
      TabIndex        =   42
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA FINAL:"
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
      Left            =   3840
      TabIndex        =   41
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   705
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
      Left            =   3840
      TabIndex        =   40
      Top             =   11520
      Width           =   1380
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
      Left            =   5400
      TabIndex        =   39
      Top             =   11520
      Width           =   1110
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "HORA INÍCIO:"
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
      Left            =   2640
      TabIndex        =   38
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   765
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA INÍCIO:"
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
      Left            =   1080
      TabIndex        =   37
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label LB_DI 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "________"
      BeginProperty Font 
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
      TabIndex        =   36
      Top             =   11520
      Width           =   840
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TESTADAS:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   11160
      UseMnemonic     =   0   'False
      Width           =   645
   End
   Begin VB.Line LH 
      Index           =   38
      X1              =   120
      X2              =   11420
      Y1              =   11400
      Y2              =   11400
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
      Left            =   1080
      TabIndex        =   34
      Top             =   11520
      Width           =   1380
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
      Left            =   2640
      TabIndex        =   33
      Top             =   11520
      Width           =   1110
   End
   Begin VB.Line LH 
      Index           =   37
      X1              =   120
      X2              =   11420
      Y1              =   11040
      Y2              =   11040
   End
   Begin VB.Line LH 
      Index           =   36
      X1              =   120
      X2              =   11420
      Y1              =   11070
      Y2              =   11070
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ensaio Hidrostático / Pneumático:"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   32
      Top             =   10800
      Width           =   2970
   End
   Begin VB.Line LV 
      Index           =   12
      X1              =   4080
      X2              =   4080
      Y1              =   3600
      Y2              =   8760
   End
   Begin VB.Line LV 
      Index           =   11
      X1              =   5520
      X2              =   5520
      Y1              =   3600
      Y2              =   8760
   End
   Begin VB.Line LV 
      Index           =   10
      X1              =   6960
      X2              =   6960
      Y1              =   3600
      Y2              =   8760
   End
   Begin VB.Line LV 
      Index           =   9
      X1              =   8520
      X2              =   8520
      Y1              =   3600
      Y2              =   8760
   End
   Begin VB.Line LV 
      Index           =   8
      X1              =   9960
      X2              =   9960
      Y1              =   3600
      Y2              =   8760
   End
   Begin VB.Line LV 
      Index           =   7
      X1              =   1440
      X2              =   1440
      Y1              =   3600
      Y2              =   8760
   End
   Begin VB.Line LV 
      Index           =   6
      X1              =   480
      X2              =   480
      Y1              =   3600
      Y2              =   8760
   End
   Begin VB.Line LH 
      Index           =   35
      X1              =   120
      X2              =   11420
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Line LH 
      Index           =   34
      X1              =   120
      X2              =   11420
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line LH 
      Index           =   33
      X1              =   120
      X2              =   11420
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Line LH 
      Index           =   32
      X1              =   120
      X2              =   11420
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Line LH 
      Index           =   31
      X1              =   120
      X2              =   11420
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line LH 
      Index           =   30
      X1              =   120
      X2              =   11420
      Y1              =   7560
      Y2              =   7560
   End
   Begin VB.Line LH 
      Index           =   29
      X1              =   120
      X2              =   11420
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line LH 
      Index           =   28
      X1              =   120
      X2              =   11420
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line LH 
      Index           =   27
      X1              =   120
      X2              =   11420
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line LH 
      Index           =   26
      X1              =   120
      X2              =   11420
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line LH 
      Index           =   25
      X1              =   120
      X2              =   11420
      Y1              =   6360
      Y2              =   6360
   End
   Begin VB.Line LH 
      Index           =   24
      X1              =   120
      X2              =   11420
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line LH 
      Index           =   23
      X1              =   120
      X2              =   11420
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line LH 
      Index           =   22
      X1              =   120
      X2              =   11420
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line LH 
      Index           =   21
      X1              =   120
      X2              =   11420
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line LH 
      Index           =   20
      X1              =   120
      X2              =   11420
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line LH 
      Index           =   19
      X1              =   120
      X2              =   11420
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line LH 
      Index           =   18
      X1              =   120
      X2              =   11420
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line LH 
      Index           =   17
      X1              =   120
      X2              =   11420
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line LH 
      Index           =   16
      X1              =   120
      X2              =   11420
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line LH 
      Index           =   15
      X1              =   120
      X2              =   11420
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label LB_IC 
      Alignment       =   2  'Center
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
      Left            =   120
      TabIndex        =   31
      Top             =   3960
      Width           =   330
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CÓDIGO DE CORRIDA"
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
      Left            =   10080
      TabIndex        =   30
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   1260
   End
   Begin VB.Label LB_Fixo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ORI nº"
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
      Left            =   8640
      TabIndex        =   29
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   1305
   End
   Begin VB.Label LB_Fixo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OF nº"
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
      Left            =   7080
      TabIndex        =   28
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   1380
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BITOLA:"
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
      Left            =   5640
      TabIndex        =   27
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   435
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MATERIAL:"
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
      Left            =   4200
      TabIndex        =   26
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPONENTE:"
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
      Left            =   1560
      TabIndex        =   25
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   19
      Left            =   600
      TabIndex        =   24
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   810
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM"
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
      TabIndex        =   23
      Top             =   3720
      UseMnemonic     =   0   'False
      Width           =   285
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de Componentes:"
      BeginProperty Font 
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
      TabIndex        =   22
      Top             =   3360
      Width           =   2010
   End
   Begin VB.Line LH 
      Index           =   7
      X1              =   120
      X2              =   11420
      Y1              =   3630
      Y2              =   3630
   End
   Begin VB.Line LH 
      Index           =   6
      X1              =   120
      X2              =   11420
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line LH 
      Index           =   5
      X1              =   120
      X2              =   11420
      Y1              =   15870
      Y2              =   15870
   End
   Begin VB.Line LH 
      Index           =   4
      X1              =   120
      X2              =   11420
      Y1              =   15840
      Y2              =   15840
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FONE"
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
      Left            =   0
      TabIndex        =   21
      Top             =   15720
      UseMnemonic     =   0   'False
      Width           =   330
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPLEMENTO"
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
      Left            =   6240
      TabIndex        =   20
      Top             =   2070
      UseMnemonic     =   0   'False
      Width           =   930
   End
   Begin VB.Label LB_Bairro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Complemento"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   19
      Top             =   2310
      Width           =   1200
   End
   Begin VB.Line LV 
      Index           =   4
      X1              =   4800
      X2              =   4800
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESPECIAL"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2070
      UseMnemonic     =   0   'False
      Width           =   555
   End
   Begin VB.Label LB_Endereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Gaxeta PTFE"
      BeginProperty Font 
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
      Top             =   2310
      Width           =   1185
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CLASSE"
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
      Left            =   4920
      TabIndex        =   16
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_IE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1500"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4920
      TabIndex        =   15
      Top             =   1710
      Width           =   420
   End
   Begin VB.Line LV 
      Index           =   1
      X1              =   6120
      X2              =   6120
      Y1              =   2040
      Y2              =   2640
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EXTREMIDADE"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   855
   End
   Begin VB.Label LB_CNPJ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Flange"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3600
      TabIndex        =   13
      Top             =   1710
      Width           =   585
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DA VÁLVULA"
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
      TabIndex        =   12
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   1005
   End
   Begin VB.Line LH 
      Index           =   3
      X1              =   120
      X2              =   11420
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line LH 
      Index           =   2
      X1              =   120
      X2              =   11420
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line LV 
      Index           =   0
      X1              =   3480
      X2              =   3480
      Y1              =   1440
      Y2              =   2040
   End
   Begin VB.Label LB_Empresa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Retenção Portinhola"
      BeginProperty Font 
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
      Top             =   1710
      Width           =   1755
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
      Left            =   10320
      TabIndex        =   10
      Top             =   240
      Width           =   720
   End
   Begin VB.Line LV 
      Index           =   2
      X1              =   4320
      X2              =   4320
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Line LH 
      Index           =   1
      X1              =   120
      X2              =   11420
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line LH 
      Index           =   0
      X1              =   120
      X2              =   11420
      Y1              =   30
      Y2              =   30
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O.M. nº:"
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
      Left            =   9000
      TabIndex        =   9
      Top             =   240
      Width           =   885
   End
   Begin VB.Label LB_Data 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/2001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10200
      TabIndex        =   8
      Top             =   720
      Width           =   1200
   End
   Begin VB.Line LH 
      Index           =   9
      X1              =   120
      X2              =   11420
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Line LH 
      Index           =   10
      X1              =   120
      X2              =   11420
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line LH 
      Index           =   11
      X1              =   120
      X2              =   11420
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line LH 
      Index           =   12
      X1              =   120
      X2              =   11420
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line LH 
      Index           =   13
      X1              =   120
      X2              =   11420
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Label LB_Fone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1000"
      BeginProperty Font 
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
      TabIndex        =   7
      Top             =   2880
      Width           =   420
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2670
      UseMnemonic     =   0   'False
      Width           =   780
   End
   Begin VB.Line LV 
      Index           =   16
      X1              =   2160
      X2              =   2160
      Y1              =   2640
      Y2              =   3240
   End
   Begin VB.Label LB_Fax 
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
      Left            =   2280
      TabIndex        =   5
      Top             =   2880
      Width           =   315
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PEDIDO Nº"
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
      Left            =   2280
      TabIndex        =   4
      Top             =   2670
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Line LV 
      Index           =   17
      X1              =   4200
      X2              =   4200
      Y1              =   2640
      Y2              =   3240
   End
   Begin VB.Label LB_Depto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01/01/0001"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   3
      Top             =   2880
      Width           =   960
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA DE ENTREGA:"
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
      Left            =   4320
      TabIndex        =   2
      Top             =   2670
      UseMnemonic     =   0   'False
      Width           =   1140
   End
   Begin VB.Line LV 
      Index           =   18
      X1              =   6120
      X2              =   6120
      Y1              =   2640
      Y2              =   3240
   End
   Begin VB.Label LB_Contato 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observação"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6240
      TabIndex        =   1
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÃO"
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
      Index           =   31
      Left            =   6240
      TabIndex        =   0
      Top             =   2670
      UseMnemonic     =   0   'False
      Width           =   810
   End
   Begin VB.Line LH 
      Index           =   14
      X1              =   120
      X2              =   11420
      Y1              =   3240
      Y2              =   3240
   End
End
Attribute VB_Name = "Tela_Fabrica_OM_Relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

