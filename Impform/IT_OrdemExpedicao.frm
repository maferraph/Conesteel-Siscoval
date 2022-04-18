VERSION 5.00
Begin VB.Form IT_OrdemExpedicao 
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
      Picture         =   "IT_OrdemExpedicao.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   151
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
      TabIndex        =   152
      Top             =   1250
      Width           =   5220
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
      Index           =   14
      Left            =   0
      TabIndex        =   150
      Top             =   2340
      UseMnemonic     =   0   'False
      Width           =   315
   End
   Begin VB.Line LV 
      Index           =   8
      X1              =   1080
      X2              =   1080
      Y1              =   2280
      Y2              =   2880
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
      TabIndex        =   149
      Top             =   2550
      Width           =   960
   End
   Begin VB.Label LB_OBS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBS"
      BeginProperty Font 
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
      TabIndex        =   148
      Top             =   13320
      Width           =   420
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   23
      Left            =   1440
      TabIndex        =   147
      Top             =   12600
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   23
      Left            =   75
      TabIndex        =   146
      Top             =   12600
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   23
      Left            =   480
      TabIndex        =   145
      Top             =   12600
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   23
      Left            =   10200
      TabIndex        =   144
      Top             =   12600
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   23
      Left            =   2760
      TabIndex        =   143
      Top             =   12600
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   22
      Left            =   1440
      TabIndex        =   142
      Top             =   12240
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   22
      Left            =   75
      TabIndex        =   141
      Top             =   12240
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   22
      Left            =   480
      TabIndex        =   140
      Top             =   12240
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   22
      Left            =   10200
      TabIndex        =   139
      Top             =   12240
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   22
      Left            =   2760
      TabIndex        =   138
      Top             =   12240
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   21
      Left            =   1440
      TabIndex        =   137
      Top             =   11880
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   21
      Left            =   75
      TabIndex        =   136
      Top             =   11880
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   21
      Left            =   480
      TabIndex        =   135
      Top             =   11880
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   21
      Left            =   10200
      TabIndex        =   134
      Top             =   11880
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   21
      Left            =   2760
      TabIndex        =   133
      Top             =   11880
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   20
      Left            =   1440
      TabIndex        =   132
      Top             =   11520
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   20
      Left            =   75
      TabIndex        =   131
      Top             =   11520
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   20
      Left            =   480
      TabIndex        =   130
      Top             =   11520
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   20
      Left            =   10200
      TabIndex        =   129
      Top             =   11520
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   20
      Left            =   2760
      TabIndex        =   128
      Top             =   11520
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   19
      Left            =   1440
      TabIndex        =   127
      Top             =   11160
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   19
      Left            =   75
      TabIndex        =   126
      Top             =   11160
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   19
      Left            =   480
      TabIndex        =   125
      Top             =   11160
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   19
      Left            =   10200
      TabIndex        =   124
      Top             =   11160
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   19
      Left            =   2760
      TabIndex        =   123
      Top             =   11160
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   18
      Left            =   1440
      TabIndex        =   122
      Top             =   10800
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   18
      Left            =   75
      TabIndex        =   121
      Top             =   10800
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   18
      Left            =   480
      TabIndex        =   120
      Top             =   10800
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   18
      Left            =   10200
      TabIndex        =   119
      Top             =   10800
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   18
      Left            =   2760
      TabIndex        =   118
      Top             =   10800
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   17
      Left            =   1440
      TabIndex        =   117
      Top             =   10440
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   17
      Left            =   75
      TabIndex        =   116
      Top             =   10440
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   17
      Left            =   480
      TabIndex        =   115
      Top             =   10440
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   17
      Left            =   10200
      TabIndex        =   114
      Top             =   10440
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   17
      Left            =   2760
      TabIndex        =   113
      Top             =   10440
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   16
      Left            =   1440
      TabIndex        =   112
      Top             =   10080
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   16
      Left            =   75
      TabIndex        =   111
      Top             =   10080
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   16
      Left            =   480
      TabIndex        =   110
      Top             =   10080
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   16
      Left            =   10200
      TabIndex        =   109
      Top             =   10080
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   16
      Left            =   2760
      TabIndex        =   108
      Top             =   10080
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   15
      Left            =   1440
      TabIndex        =   107
      Top             =   9720
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   15
      Left            =   75
      TabIndex        =   106
      Top             =   9720
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   15
      Left            =   480
      TabIndex        =   105
      Top             =   9720
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   15
      Left            =   10200
      TabIndex        =   104
      Top             =   9720
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   15
      Left            =   2760
      TabIndex        =   103
      Top             =   9720
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   14
      Left            =   1440
      TabIndex        =   102
      Top             =   9360
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   14
      Left            =   75
      TabIndex        =   101
      Top             =   9360
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   14
      Left            =   480
      TabIndex        =   100
      Top             =   9360
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   14
      Left            =   10200
      TabIndex        =   99
      Top             =   9360
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   14
      Left            =   2760
      TabIndex        =   98
      Top             =   9360
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   13
      Left            =   1440
      TabIndex        =   97
      Top             =   9000
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   13
      Left            =   75
      TabIndex        =   96
      Top             =   9000
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   13
      Left            =   480
      TabIndex        =   95
      Top             =   9000
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   13
      Left            =   10200
      TabIndex        =   94
      Top             =   9000
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   13
      Left            =   2760
      TabIndex        =   93
      Top             =   9000
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   12
      Left            =   1440
      TabIndex        =   92
      Top             =   8640
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   12
      Left            =   75
      TabIndex        =   91
      Top             =   8640
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   12
      Left            =   480
      TabIndex        =   90
      Top             =   8640
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   12
      Left            =   10200
      TabIndex        =   89
      Top             =   8640
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   12
      Left            =   2760
      TabIndex        =   88
      Top             =   8640
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   11
      Left            =   1440
      TabIndex        =   87
      Top             =   8280
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   11
      Left            =   75
      TabIndex        =   86
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   11
      Left            =   480
      TabIndex        =   85
      Top             =   8280
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   11
      Left            =   10200
      TabIndex        =   84
      Top             =   8280
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   11
      Left            =   2760
      TabIndex        =   83
      Top             =   8280
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   10
      Left            =   1440
      TabIndex        =   82
      Top             =   7920
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   10
      Left            =   75
      TabIndex        =   81
      Top             =   7920
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   10
      Left            =   480
      TabIndex        =   80
      Top             =   7920
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   10
      Left            =   10200
      TabIndex        =   79
      Top             =   7920
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   10
      Left            =   2760
      TabIndex        =   78
      Top             =   7920
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   9
      Left            =   1440
      TabIndex        =   77
      Top             =   7560
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   9
      Left            =   75
      TabIndex        =   76
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   9
      Left            =   480
      TabIndex        =   75
      Top             =   7560
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   9
      Left            =   10200
      TabIndex        =   74
      Top             =   7560
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   9
      Left            =   2760
      TabIndex        =   73
      Top             =   7560
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   8
      Left            =   1440
      TabIndex        =   72
      Top             =   7200
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   8
      Left            =   75
      TabIndex        =   71
      Top             =   7200
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   8
      Left            =   480
      TabIndex        =   70
      Top             =   7200
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   8
      Left            =   10200
      TabIndex        =   69
      Top             =   7200
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   8
      Left            =   2760
      TabIndex        =   68
      Top             =   7200
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   7
      Left            =   1440
      TabIndex        =   67
      Top             =   6840
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   7
      Left            =   75
      TabIndex        =   66
      Top             =   6840
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   7
      Left            =   480
      TabIndex        =   65
      Top             =   6840
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   7
      Left            =   10200
      TabIndex        =   64
      Top             =   6840
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   7
      Left            =   2760
      TabIndex        =   63
      Top             =   6840
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   6
      Left            =   1440
      TabIndex        =   62
      Top             =   6480
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   6
      Left            =   75
      TabIndex        =   61
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   6
      Left            =   480
      TabIndex        =   60
      Top             =   6480
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   6
      Left            =   10200
      TabIndex        =   59
      Top             =   6480
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   6
      Left            =   2760
      TabIndex        =   58
      Top             =   6480
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   5
      Left            =   1440
      TabIndex        =   57
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   5
      Left            =   75
      TabIndex        =   56
      Top             =   6120
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   55
      Top             =   6120
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   5
      Left            =   10200
      TabIndex        =   54
      Top             =   6120
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   5
      Left            =   2760
      TabIndex        =   53
      Top             =   6120
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   4
      Left            =   1440
      TabIndex        =   52
      Top             =   5760
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   4
      Left            =   75
      TabIndex        =   51
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   50
      Top             =   5760
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   4
      Left            =   10200
      TabIndex        =   49
      Top             =   5760
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   4
      Left            =   2760
      TabIndex        =   48
      Top             =   5760
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   3
      Left            =   1440
      TabIndex        =   47
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   3
      Left            =   75
      TabIndex        =   46
      Top             =   5400
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   45
      Top             =   5400
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   3
      Left            =   10200
      TabIndex        =   44
      Top             =   5400
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   3
      Left            =   2760
      TabIndex        =   43
      Top             =   5400
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   2
      Left            =   1440
      TabIndex        =   42
      Top             =   5040
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   2
      Left            =   75
      TabIndex        =   41
      Top             =   5040
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   40
      Top             =   5040
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   2
      Left            =   10200
      TabIndex        =   39
      Top             =   5040
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   2
      Left            =   2760
      TabIndex        =   38
      Top             =   5040
      Width           =   4470
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   1
      Left            =   1440
      TabIndex        =   37
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   1
      Left            =   75
      TabIndex        =   36
      Top             =   4680
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   35
      Top             =   4680
      Width           =   540
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   1
      Left            =   10200
      TabIndex        =   34
      Top             =   4680
      Width           =   585
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   1
      Left            =   2760
      TabIndex        =   33
      Top             =   4680
      Width           =   4470
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OBSERVAÇÕES"
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
      Left            =   0
      TabIndex        =   32
      Top             =   13080
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Line LH 
      Index           =   2
      X1              =   0
      X2              =   11160
      Y1              =   13680
      Y2              =   13680
   End
   Begin VB.Line LV 
      Index           =   4
      X1              =   2760
      X2              =   2760
      Y1              =   13680
      Y2              =   14280
   End
   Begin VB.Label LB_TIP 
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
      TabIndex        =   31
      Top             =   13920
      Width           =   615
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO DE EMBALAGEM"
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
      TabIndex        =   30
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1230
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº DO PEDIDO DE ESTOQUE"
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
      TabIndex        =   29
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1620
   End
   Begin VB.Label LB_PE 
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
      TabIndex        =   28
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
      Caption         =   "PESO ESTIMADO"
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
      TabIndex        =   27
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1080
   End
   Begin VB.Label LB_PL 
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
      TabIndex        =   26
      Top             =   13920
      Width           =   705
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSPORTADORA"
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
      TabIndex        =   25
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   1155
   End
   Begin VB.Label LB_TRA 
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
      TabIndex        =   24
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
      Y1              =   13050
      Y2              =   13050
   End
   Begin VB.Line LH 
      Index           =   18
      X1              =   0
      X2              =   11160
      Y1              =   13080
      Y2              =   13080
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   0
      Left            =   2760
      TabIndex        =   23
      Top             =   4320
      Width           =   4470
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   210
      Index           =   24
      Left            =   2760
      TabIndex        =   22
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label LB_Con 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   0
      Left            =   10200
      TabIndex        =   21
      Top             =   4320
      Width           =   585
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Condição"
      Height          =   210
      Index           =   8
      Left            =   10200
      TabIndex        =   20
      Top             =   3735
      Width           =   675
   End
   Begin VB.Line LV 
      Index           =   5
      X1              =   10080
      X2              =   10080
      Y1              =   3600
      Y2              =   4080
   End
   Begin VB.Line LH 
      Index           =   25
      X1              =   0
      X2              =   11160
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line LH 
      Index           =   26
      X1              =   0
      X2              =   11160
      Y1              =   3570
      Y2              =   3570
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de peças para fazer embalagem:"
      BeginProperty Font 
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
      Left            =   0
      TabIndex        =   19
      Top             =   3240
      Width           =   3315
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   18
      Top             =   4320
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   0
      Left            =   75
      TabIndex        =   17
      Top             =   4320
      Width           =   180
   End
   Begin VB.Line LV 
      Index           =   13
      X1              =   2640
      X2              =   2640
      Y1              =   3600
      Y2              =   4080
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
      Height          =   210
      Index           =   30
      Left            =   435
      TabIndex        =   16
      Top             =   3735
      Width           =   825
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   210
      Index           =   31
      Left            =   15
      TabIndex        =   15
      Top             =   3735
      Width           =   285
   End
   Begin VB.Line LV 
      Index           =   7
      X1              =   360
      X2              =   360
      Y1              =   3600
      Y2              =   4080
   End
   Begin VB.Line LH 
      Index           =   15
      X1              =   0
      X2              =   11160
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line LV 
      Index           =   6
      X1              =   1320
      X2              =   1320
      Y1              =   3600
      Y2              =   4080
   End
   Begin VB.Label LB_Fig 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   0
      Left            =   1440
      TabIndex        =   14
      Top             =   4320
      Width           =   1005
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Figura"
      Height          =   210
      Index           =   32
      Left            =   1440
      TabIndex        =   13
      Top             =   3720
      Width           =   450
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informações sobre o destinatário:"
      BeginProperty Font 
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
      TabIndex        =   12
      Top             =   1920
      Width           =   2880
   End
   Begin VB.Line LV 
      Index           =   3
      X1              =   8280
      X2              =   8280
      Y1              =   2310
      Y2              =   2910
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.N.P.J."
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
      Left            =   6480
      TabIndex        =   11
      Top             =   2340
      UseMnemonic     =   0   'False
      Width           =   435
   End
   Begin VB.Label LB_CNPJ 
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
      Left            =   6480
      TabIndex        =   10
      Top             =   2550
      Width           =   1710
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "EMPRESA"
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
      Left            =   1200
      TabIndex        =   9
      Top             =   2340
      UseMnemonic     =   0   'False
      Width           =   570
   End
   Begin VB.Line LH 
      Index           =   3
      X1              =   0
      X2              =   11160
      Y1              =   2310
      Y2              =   2310
   End
   Begin VB.Line LV 
      Index           =   0
      X1              =   6360
      X2              =   6360
      Y1              =   2310
      Y2              =   2910
   End
   Begin VB.Label LB_Empresa 
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
      Left            =   1200
      TabIndex        =   8
      Top             =   2550
      Width           =   2970
   End
   Begin VB.Line LH 
      Index           =   11
      X1              =   0
      X2              =   11160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label LB_Cidade 
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
      Left            =   8400
      TabIndex        =   7
      Top             =   2550
      Width           =   900
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MUNICÍPIO"
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
      Left            =   8400
      TabIndex        =   6
      Top             =   2340
      UseMnemonic     =   0   'False
      Width           =   630
   End
   Begin VB.Line LV 
      Index           =   15
      X1              =   9960
      X2              =   9960
      Y1              =   2310
      Y2              =   2910
   End
   Begin VB.Label LB_Estado 
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
      Left            =   10080
      TabIndex        =   5
      Top             =   2550
      Width           =   270
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO"
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
      Left            =   10080
      TabIndex        =   4
      Top             =   2340
      UseMnemonic     =   0   'False
      Width           =   480
   End
   Begin VB.Line LH 
      Index           =   13
      X1              =   0
      X2              =   11160
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Line LH 
      Index           =   14
      X1              =   0
      X2              =   11160
      Y1              =   2940
      Y2              =   2940
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
      TabIndex        =   3
      Top             =   960
      Width           =   720
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
      TabIndex        =   2
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
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEM DE EXPEDIÇÃO"
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
      TabIndex        =   1
      Top             =   390
      Width           =   2790
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
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OE"
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
      TabIndex        =   0
      Top             =   990
      Width           =   345
   End
End
Attribute VB_Name = "IT_OrdemExpedicao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
