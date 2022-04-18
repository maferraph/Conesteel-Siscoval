VERSION 5.00
Begin VB.Form Tela_Cotacao_IT 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
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
   Begin VB.TextBox TXT_Cond 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   61
      Text            =   "Tela_Cotacao_IT.frx":0000
      Top             =   12600
      Width           =   5055
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   27
      Left            =   10140
      TabIndex        =   277
      Top             =   12240
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   27
      Left            =   9645
      TabIndex        =   276
      Top             =   12240
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   27
      Left            =   9090
      TabIndex        =   275
      Top             =   12240
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   27
      Left            =   8160
      TabIndex        =   274
      Top             =   12240
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   27
      Left            =   2520
      TabIndex        =   273
      Top             =   12240
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   27
      Left            =   1440
      TabIndex        =   272
      Top             =   12240
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   27
      Left            =   480
      TabIndex        =   271
      Top             =   12240
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   27
      Left            =   60
      TabIndex        =   270
      Top             =   12240
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   26
      Left            =   10140
      TabIndex        =   269
      Top             =   12000
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   26
      Left            =   9645
      TabIndex        =   268
      Top             =   12000
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   26
      Left            =   9090
      TabIndex        =   267
      Top             =   12000
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   26
      Left            =   8160
      TabIndex        =   266
      Top             =   12000
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   26
      Left            =   2520
      TabIndex        =   265
      Top             =   12000
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   26
      Left            =   1440
      TabIndex        =   264
      Top             =   12000
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   26
      Left            =   480
      TabIndex        =   263
      Top             =   12000
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   26
      Left            =   60
      TabIndex        =   262
      Top             =   12000
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   25
      Left            =   10140
      TabIndex        =   261
      Top             =   11760
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   25
      Left            =   9645
      TabIndex        =   260
      Top             =   11760
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   25
      Left            =   9090
      TabIndex        =   259
      Top             =   11760
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   25
      Left            =   8160
      TabIndex        =   258
      Top             =   11760
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   25
      Left            =   2520
      TabIndex        =   257
      Top             =   11760
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   25
      Left            =   1440
      TabIndex        =   256
      Top             =   11760
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   25
      Left            =   480
      TabIndex        =   255
      Top             =   11760
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   25
      Left            =   60
      TabIndex        =   254
      Top             =   11760
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   24
      Left            =   10140
      TabIndex        =   253
      Top             =   11520
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   24
      Left            =   9645
      TabIndex        =   252
      Top             =   11520
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   24
      Left            =   9090
      TabIndex        =   251
      Top             =   11520
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   24
      Left            =   8160
      TabIndex        =   250
      Top             =   11520
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   24
      Left            =   2520
      TabIndex        =   249
      Top             =   11520
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   24
      Left            =   1440
      TabIndex        =   248
      Top             =   11520
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   24
      Left            =   480
      TabIndex        =   247
      Top             =   11520
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   24
      Left            =   60
      TabIndex        =   246
      Top             =   11520
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   23
      Left            =   10140
      TabIndex        =   245
      Top             =   11280
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   23
      Left            =   9645
      TabIndex        =   244
      Top             =   11280
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   23
      Left            =   9090
      TabIndex        =   243
      Top             =   11280
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   23
      Left            =   8160
      TabIndex        =   242
      Top             =   11280
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   23
      Left            =   2520
      TabIndex        =   241
      Top             =   11280
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   23
      Left            =   1440
      TabIndex        =   240
      Top             =   11280
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   23
      Left            =   480
      TabIndex        =   239
      Top             =   11280
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   23
      Left            =   60
      TabIndex        =   238
      Top             =   11280
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   22
      Left            =   10140
      TabIndex        =   237
      Top             =   11040
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   22
      Left            =   9645
      TabIndex        =   236
      Top             =   11040
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   22
      Left            =   9090
      TabIndex        =   235
      Top             =   11040
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   22
      Left            =   8160
      TabIndex        =   234
      Top             =   11040
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   22
      Left            =   2520
      TabIndex        =   233
      Top             =   11040
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   22
      Left            =   1440
      TabIndex        =   232
      Top             =   11040
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   22
      Left            =   480
      TabIndex        =   231
      Top             =   11040
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   22
      Left            =   60
      TabIndex        =   230
      Top             =   11040
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   21
      Left            =   10140
      TabIndex        =   229
      Top             =   10800
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   21
      Left            =   9645
      TabIndex        =   228
      Top             =   10800
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   21
      Left            =   9090
      TabIndex        =   227
      Top             =   10800
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   21
      Left            =   8160
      TabIndex        =   226
      Top             =   10800
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   21
      Left            =   2520
      TabIndex        =   225
      Top             =   10800
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   21
      Left            =   1440
      TabIndex        =   224
      Top             =   10800
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   21
      Left            =   480
      TabIndex        =   223
      Top             =   10800
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   21
      Left            =   60
      TabIndex        =   222
      Top             =   10800
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   20
      Left            =   10140
      TabIndex        =   221
      Top             =   10560
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   20
      Left            =   9645
      TabIndex        =   220
      Top             =   10560
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   20
      Left            =   9090
      TabIndex        =   219
      Top             =   10560
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   20
      Left            =   8160
      TabIndex        =   218
      Top             =   10560
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   20
      Left            =   2520
      TabIndex        =   217
      Top             =   10560
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   20
      Left            =   1440
      TabIndex        =   216
      Top             =   10560
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   20
      Left            =   480
      TabIndex        =   215
      Top             =   10560
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   20
      Left            =   60
      TabIndex        =   214
      Top             =   10560
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   19
      Left            =   10140
      TabIndex        =   213
      Top             =   10320
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   19
      Left            =   9645
      TabIndex        =   212
      Top             =   10320
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   19
      Left            =   9090
      TabIndex        =   211
      Top             =   10320
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   19
      Left            =   8160
      TabIndex        =   210
      Top             =   10320
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   19
      Left            =   2520
      TabIndex        =   209
      Top             =   10320
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   19
      Left            =   1440
      TabIndex        =   208
      Top             =   10320
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   19
      Left            =   480
      TabIndex        =   207
      Top             =   10320
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   19
      Left            =   60
      TabIndex        =   206
      Top             =   10320
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   18
      Left            =   10140
      TabIndex        =   205
      Top             =   10080
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   18
      Left            =   9645
      TabIndex        =   204
      Top             =   10080
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   18
      Left            =   9090
      TabIndex        =   203
      Top             =   10080
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   18
      Left            =   8160
      TabIndex        =   202
      Top             =   10080
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   18
      Left            =   2520
      TabIndex        =   201
      Top             =   10080
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   18
      Left            =   1440
      TabIndex        =   200
      Top             =   10080
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   18
      Left            =   480
      TabIndex        =   199
      Top             =   10080
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   18
      Left            =   60
      TabIndex        =   198
      Top             =   10080
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   17
      Left            =   10140
      TabIndex        =   197
      Top             =   9840
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   17
      Left            =   9645
      TabIndex        =   196
      Top             =   9840
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   17
      Left            =   9090
      TabIndex        =   195
      Top             =   9840
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   17
      Left            =   8160
      TabIndex        =   194
      Top             =   9840
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   17
      Left            =   2520
      TabIndex        =   193
      Top             =   9840
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   17
      Left            =   1440
      TabIndex        =   192
      Top             =   9840
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   17
      Left            =   480
      TabIndex        =   191
      Top             =   9840
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   17
      Left            =   60
      TabIndex        =   190
      Top             =   9840
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   16
      Left            =   10140
      TabIndex        =   189
      Top             =   9600
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   16
      Left            =   9645
      TabIndex        =   188
      Top             =   9600
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   16
      Left            =   9090
      TabIndex        =   187
      Top             =   9600
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   16
      Left            =   8160
      TabIndex        =   186
      Top             =   9600
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   16
      Left            =   2520
      TabIndex        =   185
      Top             =   9600
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   16
      Left            =   1440
      TabIndex        =   184
      Top             =   9600
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   16
      Left            =   480
      TabIndex        =   183
      Top             =   9600
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   16
      Left            =   60
      TabIndex        =   182
      Top             =   9600
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   15
      Left            =   10140
      TabIndex        =   181
      Top             =   9360
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   15
      Left            =   9645
      TabIndex        =   180
      Top             =   9360
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   15
      Left            =   9090
      TabIndex        =   179
      Top             =   9360
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   15
      Left            =   8160
      TabIndex        =   178
      Top             =   9360
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   15
      Left            =   2520
      TabIndex        =   177
      Top             =   9360
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   15
      Left            =   1440
      TabIndex        =   176
      Top             =   9360
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   15
      Left            =   480
      TabIndex        =   175
      Top             =   9360
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   15
      Left            =   60
      TabIndex        =   174
      Top             =   9360
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   14
      Left            =   10140
      TabIndex        =   173
      Top             =   9120
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   14
      Left            =   9645
      TabIndex        =   172
      Top             =   9120
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   14
      Left            =   9090
      TabIndex        =   171
      Top             =   9120
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   14
      Left            =   8160
      TabIndex        =   170
      Top             =   9120
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   14
      Left            =   2520
      TabIndex        =   169
      Top             =   9120
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   14
      Left            =   1440
      TabIndex        =   168
      Top             =   9120
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   14
      Left            =   480
      TabIndex        =   167
      Top             =   9120
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   14
      Left            =   60
      TabIndex        =   166
      Top             =   9120
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   13
      Left            =   10140
      TabIndex        =   165
      Top             =   8880
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   13
      Left            =   9645
      TabIndex        =   164
      Top             =   8880
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   13
      Left            =   9090
      TabIndex        =   163
      Top             =   8880
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   13
      Left            =   8160
      TabIndex        =   162
      Top             =   8880
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   13
      Left            =   2520
      TabIndex        =   161
      Top             =   8880
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   13
      Left            =   1440
      TabIndex        =   160
      Top             =   8880
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   13
      Left            =   480
      TabIndex        =   159
      Top             =   8880
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   13
      Left            =   60
      TabIndex        =   158
      Top             =   8880
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   12
      Left            =   10140
      TabIndex        =   157
      Top             =   8640
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   12
      Left            =   9645
      TabIndex        =   156
      Top             =   8640
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   12
      Left            =   9090
      TabIndex        =   155
      Top             =   8640
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   12
      Left            =   8160
      TabIndex        =   154
      Top             =   8640
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   12
      Left            =   2520
      TabIndex        =   153
      Top             =   8640
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   12
      Left            =   1440
      TabIndex        =   152
      Top             =   8640
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   12
      Left            =   480
      TabIndex        =   151
      Top             =   8640
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   12
      Left            =   60
      TabIndex        =   150
      Top             =   8640
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   11
      Left            =   10140
      TabIndex        =   149
      Top             =   8400
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   11
      Left            =   9645
      TabIndex        =   148
      Top             =   8400
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   11
      Left            =   9090
      TabIndex        =   147
      Top             =   8400
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   11
      Left            =   8160
      TabIndex        =   146
      Top             =   8400
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   11
      Left            =   2520
      TabIndex        =   145
      Top             =   8400
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   11
      Left            =   1440
      TabIndex        =   144
      Top             =   8400
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   11
      Left            =   480
      TabIndex        =   143
      Top             =   8400
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   11
      Left            =   60
      TabIndex        =   142
      Top             =   8400
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   10
      Left            =   10140
      TabIndex        =   141
      Top             =   8160
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   10
      Left            =   9645
      TabIndex        =   140
      Top             =   8160
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   10
      Left            =   9090
      TabIndex        =   139
      Top             =   8160
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   10
      Left            =   8160
      TabIndex        =   138
      Top             =   8160
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   10
      Left            =   2520
      TabIndex        =   137
      Top             =   8160
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   10
      Left            =   1440
      TabIndex        =   136
      Top             =   8160
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   10
      Left            =   480
      TabIndex        =   135
      Top             =   8160
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   10
      Left            =   60
      TabIndex        =   134
      Top             =   8160
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   9
      Left            =   10140
      TabIndex        =   133
      Top             =   7920
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   9
      Left            =   9645
      TabIndex        =   132
      Top             =   7920
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   9
      Left            =   9090
      TabIndex        =   131
      Top             =   7920
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   9
      Left            =   8160
      TabIndex        =   130
      Top             =   7920
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   9
      Left            =   2520
      TabIndex        =   129
      Top             =   7920
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   9
      Left            =   1440
      TabIndex        =   128
      Top             =   7920
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   9
      Left            =   480
      TabIndex        =   127
      Top             =   7920
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   9
      Left            =   60
      TabIndex        =   126
      Top             =   7920
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   8
      Left            =   10140
      TabIndex        =   125
      Top             =   7680
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   8
      Left            =   9645
      TabIndex        =   124
      Top             =   7680
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   8
      Left            =   9090
      TabIndex        =   123
      Top             =   7680
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   8
      Left            =   8160
      TabIndex        =   122
      Top             =   7680
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   8
      Left            =   2520
      TabIndex        =   121
      Top             =   7680
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   8
      Left            =   1440
      TabIndex        =   120
      Top             =   7680
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   8
      Left            =   480
      TabIndex        =   119
      Top             =   7680
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   8
      Left            =   60
      TabIndex        =   118
      Top             =   7680
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   7
      Left            =   10140
      TabIndex        =   117
      Top             =   7440
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   7
      Left            =   9645
      TabIndex        =   116
      Top             =   7440
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   7
      Left            =   9090
      TabIndex        =   115
      Top             =   7440
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   7
      Left            =   8160
      TabIndex        =   114
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   7
      Left            =   2520
      TabIndex        =   113
      Top             =   7440
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   7
      Left            =   1440
      TabIndex        =   112
      Top             =   7440
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   7
      Left            =   480
      TabIndex        =   111
      Top             =   7440
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   7
      Left            =   60
      TabIndex        =   110
      Top             =   7440
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   6
      Left            =   10140
      TabIndex        =   109
      Top             =   7200
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   6
      Left            =   9645
      TabIndex        =   108
      Top             =   7200
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   6
      Left            =   9090
      TabIndex        =   107
      Top             =   7200
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   6
      Left            =   8160
      TabIndex        =   106
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   6
      Left            =   2520
      TabIndex        =   105
      Top             =   7200
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   6
      Left            =   1440
      TabIndex        =   104
      Top             =   7200
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   6
      Left            =   480
      TabIndex        =   103
      Top             =   7200
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   6
      Left            =   60
      TabIndex        =   102
      Top             =   7200
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   5
      Left            =   10140
      TabIndex        =   101
      Top             =   6960
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   5
      Left            =   9645
      TabIndex        =   100
      Top             =   6960
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   5
      Left            =   9090
      TabIndex        =   99
      Top             =   6960
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   5
      Left            =   8160
      TabIndex        =   98
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   5
      Left            =   2520
      TabIndex        =   97
      Top             =   6960
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   5
      Left            =   1440
      TabIndex        =   96
      Top             =   6960
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   95
      Top             =   6960
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   5
      Left            =   60
      TabIndex        =   94
      Top             =   6960
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   4
      Left            =   10140
      TabIndex        =   93
      Top             =   6720
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   4
      Left            =   9645
      TabIndex        =   92
      Top             =   6720
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   4
      Left            =   9090
      TabIndex        =   91
      Top             =   6720
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   4
      Left            =   8160
      TabIndex        =   90
      Top             =   6720
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   4
      Left            =   2520
      TabIndex        =   89
      Top             =   6720
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   4
      Left            =   1440
      TabIndex        =   88
      Top             =   6720
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   87
      Top             =   6720
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   4
      Left            =   60
      TabIndex        =   86
      Top             =   6720
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   3
      Left            =   10140
      TabIndex        =   85
      Top             =   6480
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   3
      Left            =   9645
      TabIndex        =   84
      Top             =   6480
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   3
      Left            =   9090
      TabIndex        =   83
      Top             =   6480
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   3
      Left            =   8160
      TabIndex        =   82
      Top             =   6480
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   81
      Top             =   6480
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   3
      Left            =   1440
      TabIndex        =   80
      Top             =   6480
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   79
      Top             =   6480
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   3
      Left            =   60
      TabIndex        =   78
      Top             =   6480
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   2
      Left            =   10140
      TabIndex        =   77
      Top             =   6240
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   2
      Left            =   9645
      TabIndex        =   76
      Top             =   6240
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   2
      Left            =   9090
      TabIndex        =   75
      Top             =   6240
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   2
      Left            =   8160
      TabIndex        =   74
      Top             =   6240
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   2
      Left            =   2520
      TabIndex        =   73
      Top             =   6240
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   2
      Left            =   1440
      TabIndex        =   72
      Top             =   6240
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   71
      Top             =   6240
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   2
      Left            =   60
      TabIndex        =   70
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   1
      Left            =   10140
      TabIndex        =   69
      Top             =   6000
      Width           =   585
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   1
      Left            =   9645
      TabIndex        =   68
      Top             =   6000
      Width           =   240
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   1
      Left            =   9090
      TabIndex        =   67
      Top             =   6000
      Width           =   330
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   1
      Left            =   8160
      TabIndex        =   66
      Top             =   6000
      Width           =   495
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   1
      Left            =   2520
      TabIndex        =   65
      Top             =   6000
      Width           =   4470
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   1
      Left            =   1440
      TabIndex        =   64
      Top             =   6000
      Width           =   630
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   63
      Top             =   6000
      Width           =   750
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   1
      Left            =   60
      TabIndex        =   62
      Top             =   6000
      Width           =   180
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEMAIS CONDIÇÕES:"
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
      Left            =   2400
      TabIndex        =   60
      Top             =   12480
      UseMnemonic     =   0   'False
      Width           =   1230
   End
   Begin VB.Line LH 
      Index           =   19
      X1              =   7680
      X2              =   11040
      Y1              =   13800
      Y2              =   13800
   End
   Begin VB.Label LB_Vend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Departamento de Engenharia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7680
      TabIndex        =   59
      Top             =   14040
      Width           =   2520
   End
   Begin VB.Label LB_Nome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maurício Fernandes Raphael"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7680
      TabIndex        =   58
      Top             =   13800
      Width           =   2475
   End
   Begin VB.Line LV 
      Index           =   12
      X1              =   7560
      X2              =   7560
      Y1              =   12480
      Y2              =   14280
   End
   Begin VB.Line LH 
      Index           =   18
      X1              =   0
      X2              =   11160
      Y1              =   12450
      Y2              =   12450
   End
   Begin VB.Line LH 
      Index           =   17
      X1              =   0
      X2              =   11160
      Y1              =   12480
      Y2              =   12480
   End
   Begin VB.Line LH 
      Index           =   16
      X1              =   0
      X2              =   2280
      Y1              =   13080
      Y2              =   13080
   End
   Begin VB.Label LB_Frete 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FOB (São Paulo)"
      BeginProperty Font 
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
      TabIndex        =   57
      Top             =   13920
      Width           =   1485
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frete:"
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
      Left            =   0
      TabIndex        =   56
      Top             =   13680
      UseMnemonic     =   0   'False
      Width           =   300
   End
   Begin VB.Line LH 
      Index           =   15
      X1              =   0
      X2              =   2280
      Y1              =   13680
      Y2              =   13680
   End
   Begin VB.Line LV 
      Index           =   11
      X1              =   2280
      X2              =   2280
      Y1              =   12480
      Y2              =   14280
   End
   Begin VB.Label LB_CondPagto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "28 dd / 35 dd"
      BeginProperty Font 
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
      TabIndex        =   55
      Top             =   12720
      Width           =   1140
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONDIÇÕES DE PAGAMENTO:"
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
      Left            =   0
      TabIndex        =   54
      Top             =   12480
      UseMnemonic     =   0   'False
      Width           =   1725
   End
   Begin VB.Label LB_Trans 
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
      TabIndex        =   53
      Top             =   13320
      Width           =   1890
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSPORTADORA:"
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
      TabIndex        =   52
      Top             =   13080
      UseMnemonic     =   0   'False
      Width           =   1185
   End
   Begin VB.Line LH 
      Index           =   6
      X1              =   0
      X2              =   11160
      Y1              =   14310
      Y2              =   14310
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   210
      Index           =   24
      Left            =   2520
      TabIndex        =   51
      Top             =   5295
      Width           =   780
   End
   Begin VB.Line LV 
      Index           =   10
      X1              =   2400
      X2              =   2400
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preço Unitário:"
      Height          =   210
      Index           =   23
      Left            =   7860
      TabIndex        =   50
      Top             =   5295
      Width           =   1050
   End
   Begin VB.Line LV 
      Index           =   9
      X1              =   7680
      X2              =   7680
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Line LV 
      Index           =   8
      X1              =   9000
      X2              =   9000
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ICMS:"
      Height          =   210
      Index           =   22
      Left            =   9075
      TabIndex        =   49
      Top             =   5295
      Width           =   405
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informamo-lhes preços e demais condições para os seguintes materiais:"
      BeginProperty Font 
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
      TabIndex        =   48
      Top             =   4800
      Width           =   6285
   End
   Begin VB.Line LH 
      Index           =   14
      X1              =   0
      X2              =   11160
      Y1              =   4590
      Y2              =   4590
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAMAL:"
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
      Left            =   8400
      TabIndex        =   47
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   435
   End
   Begin VB.Label LB_Ramal 
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
      Left            =   8400
      TabIndex        =   46
      Top             =   4200
      Width           =   315
   End
   Begin VB.Line LV 
      Index           =   19
      X1              =   8280
      X2              =   8280
      Y1              =   3960
      Y2              =   4560
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
      Index           =   31
      Left            =   6240
      TabIndex        =   45
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label LB_Contato 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maurício"
      BeginProperty Font 
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
      TabIndex        =   44
      Top             =   4200
      Width           =   735
   End
   Begin VB.Line LV 
      Index           =   18
      X1              =   6120
      X2              =   6120
      Y1              =   3960
      Y2              =   4560
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ATT. DEPARTAMENTO:"
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
      Left            =   4080
      TabIndex        =   43
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   1305
   End
   Begin VB.Label LB_Depto 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Compras"
      BeginProperty Font 
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
      TabIndex        =   42
      Top             =   4200
      Width           =   780
   End
   Begin VB.Line LV 
      Index           =   17
      X1              =   3960
      X2              =   3960
      Y1              =   3960
      Y2              =   4560
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FAX:"
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
      Left            =   2040
      TabIndex        =   41
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   255
   End
   Begin VB.Label LB_Fax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(011) 6107-6667"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   40
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Line LV 
      Index           =   16
      X1              =   1920
      X2              =   1920
      Y1              =   3960
      Y2              =   4560
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FONE:"
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
      TabIndex        =   39
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   360
   End
   Begin VB.Label LB_Fone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(011) 6910-1444"
      BeginProperty Font 
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
      TabIndex        =   38
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Line LH 
      Index           =   13
      X1              =   0
      X2              =   11160
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO:"
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
      TabIndex        =   37
      Top             =   3360
      UseMnemonic     =   0   'False
      Width           =   510
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
      TabIndex        =   36
      Top             =   3600
      Width           =   270
   End
   Begin VB.Line LV 
      Index           =   15
      X1              =   9960
      X2              =   9960
      Y1              =   3360
      Y2              =   3960
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   10
      Left            =   8400
      TabIndex        =   35
      Top             =   3360
      UseMnemonic     =   0   'False
      Width           =   660
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
      TabIndex        =   34
      Top             =   3600
      Width           =   900
   End
   Begin VB.Line LH 
      Index           =   12
      X1              =   0
      X2              =   11160
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line LH 
      Index           =   11
      X1              =   0
      X2              =   11160
      Y1              =   2730
      Y2              =   2730
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
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Label LB_Data 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "São Paulo, 30 de Novembro de 2000"
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
      Left            =   6600
      TabIndex        =   33
      Top             =   1800
      Width           =   3885
   End
   Begin VB.Line LH 
      Index           =   8
      X1              =   6480
      X2              =   11160
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "site: http://www.conesteel.ind.br   -   e-mail: conesteel.valves@hipernet.com.br"
      Height          =   210
      Index           =   8
      Left            =   360
      TabIndex        =   32
      Top             =   2280
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   5760
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fone (0xx11) 6910-1444   -   Fax: (0xx11) 6107-6667"
      Height          =   210
      Index           =   5
      Left            =   1440
      TabIndex        =   31
      Top             =   2040
      UseMnemonic     =   0   'False
      Width           =   3885
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.N.P.J. 55.783.427/0001-03   -   Inscrição Estadual 111.502.963.110"
      Height          =   210
      Index           =   4
      Left            =   720
      TabIndex        =   30
      Top             =   1560
      UseMnemonic     =   0   'False
      Width           =   4965
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avenida Montemagno, 2.454 - Vila Formosa - São Paulo - (SP) - Brasil - CEP 03371-000"
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   29
      Top             =   1800
      UseMnemonic     =   0   'False
      Width           =   6300
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conexões de Aço Ltda."
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
      Left            =   1680
      TabIndex        =   28
      Top             =   1200
      Width           =   3210
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COTAÇÃO DE PREÇOS Nº"
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
      Left            =   6600
      TabIndex        =   27
      Top             =   600
      Width           =   3030
   End
   Begin VB.Line LH 
      Index           =   0
      X1              =   0
      X2              =   11160
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Image IMG 
      Height          =   735
      Left            =   2880
      Picture         =   "Tela_Cotacao_IT.frx":0150
      Stretch         =   -1  'True
      Top             =   180
      Width           =   720
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
      Left            =   2400
      TabIndex        =   26
      Top             =   960
      Width           =   1785
   End
   Begin VB.Line LH 
      Index           =   1
      X1              =   0
      X2              =   11160
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line LV 
      Index           =   2
      X1              =   6480
      X2              =   6480
      Y1              =   120
      Y2              =   2520
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
      Left            =   9840
      TabIndex        =   25
      Top             =   550
      Width           =   720
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
      Left            =   0
      TabIndex        =   24
      Top             =   3000
      Width           =   2970
   End
   Begin VB.Line LV 
      Index           =   0
      X1              =   6720
      X2              =   6720
      Y1              =   2760
      Y2              =   3360
   End
   Begin VB.Line LH 
      Index           =   2
      X1              =   0
      X2              =   11160
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line LH 
      Index           =   3
      X1              =   0
      X2              =   11160
      Y1              =   2760
      Y2              =   2760
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
      Index           =   6
      Left            =   0
      TabIndex        =   23
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   600
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
      Left            =   6840
      TabIndex        =   22
      Top             =   3000
      Width           =   1710
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.N.P.J.:"
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
      Left            =   6840
      TabIndex        =   21
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   465
   End
   Begin VB.Line LV 
      Index           =   1
      X1              =   5160
      X2              =   5160
      Y1              =   3360
      Y2              =   3960
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
      Left            =   8880
      TabIndex        =   20
      Top             =   3000
      Width           =   1440
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   9
      Left            =   8880
      TabIndex        =   19
      Top             =   2760
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Line LV 
      Index           =   3
      X1              =   8280
      X2              =   8280
      Y1              =   3360
      Y2              =   3960
   End
   Begin VB.Label LB_Endereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avenida Montemagno, 2.454"
      BeginProperty Font 
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
      TabIndex        =   18
      Top             =   3600
      Width           =   2460
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   11
      Left            =   0
      TabIndex        =   17
      Top             =   3360
      UseMnemonic     =   0   'False
      Width           =   705
   End
   Begin VB.Line LV 
      Index           =   4
      X1              =   8760
      X2              =   8760
      Y1              =   2760
      Y2              =   3360
   End
   Begin VB.Label LB_CEP 
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
      Left            =   7200
      TabIndex        =   16
      Top             =   3600
      Width           =   900
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   13
      Left            =   7200
      TabIndex        =   15
      Top             =   3360
      UseMnemonic     =   0   'False
      Width           =   270
   End
   Begin VB.Line LV 
      Index           =   5
      X1              =   7080
      X2              =   7080
      Y1              =   3360
      Y2              =   3960
   End
   Begin VB.Label LB_Bairro 
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
      Left            =   5280
      TabIndex        =   14
      Top             =   3600
      Width           =   1155
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   15
      Left            =   5280
      TabIndex        =   13
      Top             =   3360
      UseMnemonic     =   0   'False
      Width           =   480
   End
   Begin VB.Line LV 
      Index           =   6
      X1              =   1320
      X2              =   1320
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Line LH 
      Index           =   4
      X1              =   0
      X2              =   11160
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line LH 
      Index           =   5
      X1              =   0
      X2              =   11160
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line LV 
      Index           =   7
      X1              =   360
      X2              =   360
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   210
      Index           =   16
      Left            =   15
      TabIndex        =   12
      Top             =   5295
      Width           =   285
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Figura:"
      Height          =   210
      Index           =   17
      Left            =   620
      TabIndex        =   11
      Top             =   5295
      Width           =   495
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade:"
      Height          =   210
      Index           =   18
      Left            =   1440
      TabIndex        =   10
      Top             =   5295
      Width           =   870
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IPI:"
      Height          =   210
      Index           =   19
      Left            =   9675
      TabIndex        =   9
      Top             =   5295
      Width           =   195
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prazo Entrega:"
      Height          =   210
      Index           =   20
      Left            =   10095
      TabIndex        =   8
      Top             =   5295
      Width           =   1065
   End
   Begin VB.Line LV 
      Index           =   13
      X1              =   9480
      X2              =   9480
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Line LV 
      Index           =   14
      X1              =   9960
      X2              =   9960
      Y1              =   5160
      Y2              =   5640
   End
   Begin VB.Label LB_Item 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   7
      Top             =   5760
      Width           =   180
   End
   Begin VB.Label LB_Figura 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100-W160"
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   6
      Top             =   5760
      Width           =   750
   End
   Begin VB.Label LB_Quant 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100 pçs."
      Height          =   210
      Index           =   0
      Left            =   1440
      TabIndex        =   5
      Top             =   5760
      Width           =   630
   End
   Begin VB.Label LB_Descricao 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta Cast. Aparaf. Int.F6 300PSI WN RF A-105 1/2"""
      Height          =   210
      Index           =   0
      Left            =   2520
      TabIndex        =   4
      Top             =   5760
      Width           =   4470
   End
   Begin VB.Label LB_Preco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100,00"
      Height          =   210
      Index           =   0
      Left            =   8160
      TabIndex        =   3
      Top             =   5760
      Width           =   495
   End
   Begin VB.Label LB_ICMS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18%"
      Height          =   210
      Index           =   0
      Left            =   9090
      TabIndex        =   2
      Top             =   5760
      Width           =   330
   End
   Begin VB.Label LB_IPI 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8%"
      Height          =   210
      Index           =   0
      Left            =   9645
      TabIndex        =   1
      Top             =   5760
      Width           =   240
   End
   Begin VB.Label LB_Prazo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imediato"
      Height          =   210
      Index           =   0
      Left            =   10140
      TabIndex        =   0
      Top             =   5760
      Width           =   585
   End
   Begin VB.Line LH 
      Index           =   7
      X1              =   0
      X2              =   11160
      Y1              =   14280
      Y2              =   14280
   End
End
Attribute VB_Name = "Tela_Cotacao_IT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label5_Click()

End Sub

Private Sub LB_Qne_Click(Index As Integer)

End Sub

