VERSION 5.00
Begin VB.Form Tela_Cfg_MateriaPrima_Relatorio 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Tela_Cfg_MateriaPrima_Relatorio.frm"
   ClientHeight    =   16005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   16005
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
      Left            =   360
      Picture         =   "Tela_Cfg_MateriaPrima_Relatorio.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   178
      Top             =   480
      Width           =   600
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   177
      Top             =   15000
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   23
      Left            =   6960
      TabIndex        =   176
      Top             =   15000
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   175
      Top             =   15000
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   174
      Top             =   15000
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   173
      Top             =   15000
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      Left            =   120
      TabIndex        =   172
      Top             =   14640
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   22
      Left            =   6960
      TabIndex        =   171
      Top             =   14640
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   170
      Top             =   14640
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   169
      Top             =   14640
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   168
      Top             =   14640
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   167
      Top             =   14280
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   21
      Left            =   6960
      TabIndex        =   166
      Top             =   14280
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   165
      Top             =   14280
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   164
      Top             =   14280
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   163
      Top             =   14280
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   162
      Top             =   13920
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   20
      Left            =   6960
      TabIndex        =   161
      Top             =   13920
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   160
      Top             =   13920
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   159
      Top             =   13920
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   158
      Top             =   13920
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   157
      Top             =   13560
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   19
      Left            =   6960
      TabIndex        =   156
      Top             =   13560
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   155
      Top             =   13560
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   154
      Top             =   13560
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   153
      Top             =   13560
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   152
      Top             =   13200
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   18
      Left            =   6960
      TabIndex        =   151
      Top             =   13200
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   150
      Top             =   13200
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   149
      Top             =   13200
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   148
      Top             =   13200
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   147
      Top             =   12840
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   17
      Left            =   6960
      TabIndex        =   146
      Top             =   12840
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   145
      Top             =   12840
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   144
      Top             =   12840
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   143
      Top             =   12840
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   142
      Top             =   12480
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   16
      Left            =   6960
      TabIndex        =   141
      Top             =   12480
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   140
      Top             =   12480
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   139
      Top             =   12480
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   138
      Top             =   12480
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   137
      Top             =   12120
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   15
      Left            =   6960
      TabIndex        =   136
      Top             =   12120
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   135
      Top             =   12120
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   134
      Top             =   12120
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   133
      Top             =   12120
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   132
      Top             =   11760
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   14
      Left            =   6960
      TabIndex        =   131
      Top             =   11760
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   130
      Top             =   11760
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   129
      Top             =   11760
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   128
      Top             =   11760
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   127
      Top             =   11400
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   13
      Left            =   6960
      TabIndex        =   126
      Top             =   11400
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   125
      Top             =   11400
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   124
      Top             =   11400
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   123
      Top             =   11400
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   122
      Top             =   11040
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   12
      Left            =   6960
      TabIndex        =   121
      Top             =   11040
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   120
      Top             =   11040
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   119
      Top             =   11040
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   118
      Top             =   11040
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   117
      Top             =   10680
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   11
      Left            =   6960
      TabIndex        =   116
      Top             =   10680
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   115
      Top             =   10680
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   114
      Top             =   10680
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   113
      Top             =   10680
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   112
      Top             =   10320
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   10
      Left            =   6960
      TabIndex        =   111
      Top             =   10320
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   110
      Top             =   10320
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   109
      Top             =   10320
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   108
      Top             =   10320
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   107
      Top             =   9960
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   9
      Left            =   6960
      TabIndex        =   106
      Top             =   9960
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   105
      Top             =   9960
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   104
      Top             =   9960
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   103
      Top             =   9960
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   102
      Top             =   9600
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   8
      Left            =   6960
      TabIndex        =   101
      Top             =   9600
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   100
      Top             =   9600
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   99
      Top             =   9600
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   98
      Top             =   9600
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   97
      Top             =   9240
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   7
      Left            =   6960
      TabIndex        =   96
      Top             =   9240
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   95
      Top             =   9240
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      Left            =   2040
      TabIndex        =   94
      Top             =   9240
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   93
      Top             =   9240
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   92
      Top             =   8880
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   6
      Left            =   6960
      TabIndex        =   91
      Top             =   8880
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   90
      Top             =   8880
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   89
      Top             =   8880
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   88
      Top             =   8880
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   87
      Top             =   8520
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   5
      Left            =   6960
      TabIndex        =   86
      Top             =   8520
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   85
      Top             =   8520
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   84
      Top             =   8520
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   83
      Top             =   8520
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   82
      Top             =   8160
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   4
      Left            =   6960
      TabIndex        =   81
      Top             =   8160
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   80
      Top             =   8160
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   79
      Top             =   8160
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   78
      Top             =   8160
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   77
      Top             =   7800
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   3
      Left            =   6960
      TabIndex        =   76
      Top             =   7800
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   75
      Top             =   7800
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   74
      Top             =   7800
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   73
      Top             =   7800
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   72
      Top             =   7440
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   2
      Left            =   6960
      TabIndex        =   71
      Top             =   7440
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   70
      Top             =   7440
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   69
      Top             =   7440
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   68
      Top             =   7440
      Width           =   615
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   67
      Top             =   7080
      Width           =   525
   End
   Begin VB.Label LB_BIT 
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
      Index           =   1
      Left            =   6960
      TabIndex        =   66
      Top             =   7080
      Width           =   510
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   65
      Top             =   7080
      Width           =   1110
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   64
      Top             =   7080
      Width           =   510
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   63
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   62
      Top             =   5760
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   61
      Top             =   5760
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   60
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   59
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   58
      Top             =   5400
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   57
      Top             =   5400
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   56
      Top             =   5400
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   55
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   54
      Top             =   5040
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   53
      Top             =   5040
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   52
      Top             =   5040
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   51
      Top             =   5040
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   50
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   49
      Top             =   4680
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   48
      Top             =   4680
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   47
      Top             =   4680
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   46
      Top             =   4320
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   45
      Top             =   4320
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   44
      Top             =   4320
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   43
      Top             =   4320
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   42
      Top             =   3960
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   41
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   40
      Top             =   3960
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   39
      Top             =   3960
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   38
      Top             =   3600
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   37
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   36
      Top             =   3600
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   35
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   34
      Top             =   3240
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   33
      Top             =   3240
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   32
      Top             =   3240
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   31
      Top             =   3240
      Width           =   615
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   30
      Top             =   2880
      Width           =   510
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   29
      Top             =   2880
      Width           =   1245
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   28
      Top             =   2880
      Width           =   1110
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   27
      Top             =   2880
      Width           =   615
   End
   Begin VB.Line L1 
      Index           =   2
      X1              =   120
      X2              =   11520
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line L2 
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   5760
      TabIndex        =   26
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIÇÃO:"
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
      Left            =   5760
      TabIndex        =   25
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label LB_Relatorio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Relatorio"
      BeginProperty Font 
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
      TabIndex        =   24
      Top             =   15360
      Width           =   765
   End
   Begin VB.Label LB_NOM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula"
      BeginProperty Font 
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
      Left            =   3720
      TabIndex        =   23
      Top             =   6720
      Width           =   615
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOME:"
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
      Left            =   3720
      TabIndex        =   22
      Top             =   6480
      Width           =   390
   End
   Begin VB.Label LB_FIG 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   21
      Top             =   6720
      Width           =   510
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIGURA:"
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
      Left            =   2040
      TabIndex        =   20
      Top             =   6480
      Width           =   480
   End
   Begin VB.Label LB_MAT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   9240
      TabIndex        =   19
      Top             =   6720
      Width           =   1110
   End
   Begin VB.Label LB_NF 
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
      Index           =   7
      Left            =   9240
      TabIndex        =   18
      Top             =   6480
      Width           =   615
   End
   Begin VB.Label LB_BIT 
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
      Index           =   0
      Left            =   6960
      TabIndex        =   17
      Top             =   6720
      Width           =   510
   End
   Begin VB.Label LB_QUA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10000"
      BeginProperty Font 
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
      TabIndex        =   16
      Top             =   6720
      Width           =   525
   End
   Begin VB.Label LB_NF 
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
      Index           =   6
      Left            =   6960
      TabIndex        =   15
      Top             =   6480
      Width           =   435
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
      Index           =   5
      Left            =   120
      TabIndex        =   14
      Top             =   6480
      Width           =   810
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIGURAÇÃO DA MATÉRIA-PRIMA:"
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
      TabIndex        =   13
      Top             =   6360
      Width           =   2280
   End
   Begin VB.Line L1 
      Index           =   1
      X1              =   120
      X2              =   11520
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line L2 
      Index           =   1
      X1              =   120
      X2              =   11520
      Y1              =   15360
      Y2              =   15360
   End
   Begin VB.Label LB_Material 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ASTM A-105"
      BeginProperty Font 
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
      Left            =   3600
      TabIndex        =   12
      Top             =   2520
      Width           =   1110
   End
   Begin VB.Label LB_Bitola 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X 1.1/4"""
      BeginProperty Font 
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
      Left            =   1680
      TabIndex        =   11
      Top             =   2520
      Width           =   1245
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-N"
      BeginProperty Font 
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
      TabIndex        =   10
      Top             =   2520
      Width           =   510
   End
   Begin VB.Label LB_NF 
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
      Index           =   3
      Left            =   3600
      TabIndex        =   9
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label LB_NF 
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
      Index           =   2
      Left            =   1680
      TabIndex        =   8
      Top             =   2280
      Width           =   435
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FIGURA:"
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
      TabIndex        =   7
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label LB_NF 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "INFORMAÇÕES SOBRE O ÍTEM:"
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
      TabIndex        =   6
      Top             =   2160
      Width           =   1830
   End
   Begin VB.Line L1 
      Index           =   0
      X1              =   120
      X2              =   11520
      Y1              =   2280
      Y2              =   2280
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
      Caption         =   "Relatório de Configuração"
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
      Left            =   7320
      TabIndex        =   1
      Top             =   480
      Width           =   3405
   End
   Begin VB.Label LB_T 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "de Matéria-Prima"
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
      Left            =   7800
      TabIndex        =   0
      Top             =   960
      Width           =   2205
   End
End
Attribute VB_Name = "Tela_Cfg_MateriaPrima_Relatorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
