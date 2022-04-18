VERSION 5.00
Begin VB.Form IT_OrdemMontagem 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
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
      Picture         =   "IT_OrdemMontagem.frx":0000
      ScaleHeight     =   675
      ScaleWidth      =   600
      TabIndex        =   314
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
      TabIndex        =   315
      Top             =   1250
      Width           =   5220
   End
   Begin VB.Label LB_OMA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """E X I S T E     O M     A B E R T A"""
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
      TabIndex        =   313
      Top             =   1250
      Width           =   3750
   End
   Begin VB.Label LB_QuantMon 
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
      Left            =   9720
      TabIndex        =   312
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   18
      Left            =   10440
      TabIndex        =   311
      Top             =   13800
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   18
      Left            =   9720
      TabIndex        =   310
      Top             =   13800
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   18
      Left            =   9000
      TabIndex        =   309
      Top             =   13800
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   18
      Left            =   8280
      TabIndex        =   308
      Top             =   13800
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   18
      Left            =   7560
      TabIndex        =   307
      Top             =   13800
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   17
      Left            =   10440
      TabIndex        =   306
      Top             =   13320
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   17
      Left            =   9720
      TabIndex        =   305
      Top             =   13320
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   17
      Left            =   9000
      TabIndex        =   304
      Top             =   13320
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   17
      Left            =   8280
      TabIndex        =   303
      Top             =   13320
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   17
      Left            =   7560
      TabIndex        =   302
      Top             =   13320
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   16
      Left            =   10440
      TabIndex        =   301
      Top             =   12840
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   16
      Left            =   9720
      TabIndex        =   300
      Top             =   12840
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   16
      Left            =   9000
      TabIndex        =   299
      Top             =   12840
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   16
      Left            =   8280
      TabIndex        =   298
      Top             =   12840
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   16
      Left            =   7560
      TabIndex        =   297
      Top             =   12840
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   15
      Left            =   10440
      TabIndex        =   296
      Top             =   12360
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   15
      Left            =   9720
      TabIndex        =   295
      Top             =   12360
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   15
      Left            =   9000
      TabIndex        =   294
      Top             =   12360
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   15
      Left            =   8280
      TabIndex        =   293
      Top             =   12360
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   15
      Left            =   7560
      TabIndex        =   292
      Top             =   12360
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   14
      Left            =   10440
      TabIndex        =   291
      Top             =   11880
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   14
      Left            =   9720
      TabIndex        =   290
      Top             =   11880
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   14
      Left            =   9000
      TabIndex        =   289
      Top             =   11880
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   14
      Left            =   8280
      TabIndex        =   288
      Top             =   11880
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   14
      Left            =   7560
      TabIndex        =   287
      Top             =   11880
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   13
      Left            =   10440
      TabIndex        =   286
      Top             =   11400
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   13
      Left            =   9720
      TabIndex        =   285
      Top             =   11400
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   13
      Left            =   9000
      TabIndex        =   284
      Top             =   11400
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   13
      Left            =   8280
      TabIndex        =   283
      Top             =   11400
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   13
      Left            =   7560
      TabIndex        =   282
      Top             =   11400
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   12
      Left            =   10440
      TabIndex        =   281
      Top             =   10920
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   12
      Left            =   9720
      TabIndex        =   280
      Top             =   10920
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   12
      Left            =   9000
      TabIndex        =   279
      Top             =   10920
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   12
      Left            =   8280
      TabIndex        =   278
      Top             =   10920
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   12
      Left            =   7560
      TabIndex        =   277
      Top             =   10920
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   11
      Left            =   10440
      TabIndex        =   276
      Top             =   10440
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   11
      Left            =   9720
      TabIndex        =   275
      Top             =   10440
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   11
      Left            =   9000
      TabIndex        =   274
      Top             =   10440
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   11
      Left            =   8280
      TabIndex        =   273
      Top             =   10440
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   11
      Left            =   7560
      TabIndex        =   272
      Top             =   10440
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   10
      Left            =   10440
      TabIndex        =   271
      Top             =   9960
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   10
      Left            =   9720
      TabIndex        =   270
      Top             =   9960
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   10
      Left            =   9000
      TabIndex        =   269
      Top             =   9960
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   10
      Left            =   8280
      TabIndex        =   268
      Top             =   9960
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   10
      Left            =   7560
      TabIndex        =   267
      Top             =   9960
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   9
      Left            =   10440
      TabIndex        =   266
      Top             =   9480
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   9
      Left            =   9720
      TabIndex        =   265
      Top             =   9480
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   9
      Left            =   9000
      TabIndex        =   264
      Top             =   9480
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   9
      Left            =   8280
      TabIndex        =   263
      Top             =   9480
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   9
      Left            =   7560
      TabIndex        =   262
      Top             =   9480
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   8
      Left            =   10440
      TabIndex        =   261
      Top             =   9000
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   8
      Left            =   9720
      TabIndex        =   260
      Top             =   9000
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   8
      Left            =   9000
      TabIndex        =   259
      Top             =   9000
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   8
      Left            =   8280
      TabIndex        =   258
      Top             =   9000
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   8
      Left            =   7560
      TabIndex        =   257
      Top             =   9000
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   7
      Left            =   10440
      TabIndex        =   256
      Top             =   8520
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   7
      Left            =   9720
      TabIndex        =   255
      Top             =   8520
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   7
      Left            =   9000
      TabIndex        =   254
      Top             =   8520
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   7
      Left            =   8280
      TabIndex        =   253
      Top             =   8520
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   7
      Left            =   7560
      TabIndex        =   252
      Top             =   8520
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   6
      Left            =   10440
      TabIndex        =   251
      Top             =   8040
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   6
      Left            =   9720
      TabIndex        =   250
      Top             =   8040
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   6
      Left            =   9000
      TabIndex        =   249
      Top             =   8040
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   6
      Left            =   8280
      TabIndex        =   248
      Top             =   8040
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   6
      Left            =   7560
      TabIndex        =   247
      Top             =   8040
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   5
      Left            =   10440
      TabIndex        =   246
      Top             =   7560
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   5
      Left            =   9720
      TabIndex        =   245
      Top             =   7560
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   5
      Left            =   9000
      TabIndex        =   244
      Top             =   7560
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   5
      Left            =   8280
      TabIndex        =   243
      Top             =   7560
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   5
      Left            =   7560
      TabIndex        =   242
      Top             =   7560
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   4
      Left            =   10440
      TabIndex        =   241
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   4
      Left            =   9720
      TabIndex        =   240
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   4
      Left            =   9000
      TabIndex        =   239
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   4
      Left            =   8280
      TabIndex        =   238
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   4
      Left            =   7560
      TabIndex        =   237
      Top             =   7080
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   3
      Left            =   10440
      TabIndex        =   236
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   3
      Left            =   9720
      TabIndex        =   235
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   3
      Left            =   9000
      TabIndex        =   234
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   3
      Left            =   8280
      TabIndex        =   233
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   3
      Left            =   7560
      TabIndex        =   232
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   2
      Left            =   10440
      TabIndex        =   231
      Top             =   6120
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   2
      Left            =   9720
      TabIndex        =   230
      Top             =   6120
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   2
      Left            =   9000
      TabIndex        =   229
      Top             =   6120
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   2
      Left            =   8280
      TabIndex        =   228
      Top             =   6120
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   2
      Left            =   7560
      TabIndex        =   227
      Top             =   6120
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   1
      Left            =   10440
      TabIndex        =   226
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   1
      Left            =   9720
      TabIndex        =   225
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   1
      Left            =   9000
      TabIndex        =   224
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   1
      Left            =   8280
      TabIndex        =   223
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   1
      Left            =   7560
      TabIndex        =   222
      Top             =   5640
      Width           =   720
   End
   Begin VB.Label LB_CE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   0
      Left            =   10440
      TabIndex        =   221
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label LB_CD 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   0
      Left            =   9720
      TabIndex        =   220
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label LB_CC 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   0
      Left            =   9000
      TabIndex        =   219
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label LB_CB 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   0
      Left            =   8280
      TabIndex        =   218
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label LB_CA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ABC123"
      Height          =   210
      Index           =   0
      Left            =   7560
      TabIndex        =   217
      Top             =   5160
      Width           =   720
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "E"
      Height          =   210
      Index           =   18
      Left            =   10800
      TabIndex        =   216
      Top             =   4710
      Width           =   240
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      Height          =   210
      Index           =   17
      Left            =   10080
      TabIndex        =   215
      Top             =   4710
      Width           =   120
   End
   Begin VB.Line LV 
      Index           =   15
      X1              =   10440
      X2              =   10440
      Y1              =   4680
      Y2              =   14280
   End
   Begin VB.Line LV 
      Index           =   12
      X1              =   9720
      X2              =   9720
      Y1              =   4680
      Y2              =   14280
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   6840
      TabIndex        =   214
      Top             =   13785
      Width           =   700
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   6120
      TabIndex        =   213
      Top             =   13785
      Width           =   700
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   18
      Left            =   1140
      TabIndex        =   212
      Top             =   13875
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   18
      Left            =   80
      TabIndex        =   211
      Top             =   13785
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   18
      Left            =   480
      TabIndex        =   210
      Top             =   13785
      Width           =   540
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   18
      Left            =   1140
      TabIndex        =   209
      Top             =   13680
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   18
      Left            =   2760
      TabIndex        =   208
      Top             =   13785
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   18
      Left            =   4080
      TabIndex        =   207
      Top             =   13785
      Width           =   1125
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   5400
      TabIndex        =   206
      Top             =   13785
      Width           =   700
   End
   Begin VB.Line LV 
      Index           =   11
      X1              =   7920
      X2              =   7920
      Y1              =   2160
      Y2              =   3000
   End
   Begin VB.Label LB_DC 
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
      Left            =   8040
      TabIndex        =   205
      Top             =   2520
      Width           =   510
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
      Index           =   16
      Left            =   8040
      TabIndex        =   204
      Top             =   2190
      UseMnemonic     =   0   'False
      Width           =   1320
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   5400
      TabIndex        =   203
      Top             =   13290
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   17
      Left            =   4080
      TabIndex        =   202
      Top             =   13200
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   17
      Left            =   2760
      TabIndex        =   201
      Top             =   13290
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   17
      Left            =   1140
      TabIndex        =   200
      Top             =   13185
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   17
      Left            =   480
      TabIndex        =   199
      Top             =   13290
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   17
      Left            =   80
      TabIndex        =   198
      Top             =   13290
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   17
      Left            =   1140
      TabIndex        =   197
      Top             =   13395
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   6120
      TabIndex        =   196
      Top             =   13290
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   6840
      TabIndex        =   195
      Top             =   13290
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   5400
      TabIndex        =   194
      Top             =   12810
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   16
      Left            =   4080
      TabIndex        =   193
      Top             =   12810
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   16
      Left            =   2760
      TabIndex        =   192
      Top             =   12810
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   16
      Left            =   1140
      TabIndex        =   191
      Top             =   12705
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   16
      Left            =   480
      TabIndex        =   190
      Top             =   12810
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   16
      Left            =   80
      TabIndex        =   189
      Top             =   12810
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   16
      Left            =   1140
      TabIndex        =   188
      Top             =   12915
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   6120
      TabIndex        =   187
      Top             =   12810
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   6840
      TabIndex        =   186
      Top             =   12810
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   5400
      TabIndex        =   185
      Top             =   12345
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   15
      Left            =   4080
      TabIndex        =   184
      Top             =   12345
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   15
      Left            =   2760
      TabIndex        =   183
      Top             =   12345
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   15
      Left            =   1140
      TabIndex        =   182
      Top             =   12240
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   15
      Left            =   480
      TabIndex        =   181
      Top             =   12345
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   15
      Left            =   80
      TabIndex        =   180
      Top             =   12345
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   15
      Left            =   1140
      TabIndex        =   179
      Top             =   12450
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   6120
      TabIndex        =   178
      Top             =   12345
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   6840
      TabIndex        =   177
      Top             =   12345
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   5400
      TabIndex        =   176
      Top             =   11850
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   14
      Left            =   4080
      TabIndex        =   175
      Top             =   11850
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   14
      Left            =   2760
      TabIndex        =   174
      Top             =   11850
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   14
      Left            =   1140
      TabIndex        =   173
      Top             =   11745
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   14
      Left            =   480
      TabIndex        =   172
      Top             =   11850
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   14
      Left            =   80
      TabIndex        =   171
      Top             =   11850
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   14
      Left            =   1140
      TabIndex        =   170
      Top             =   11955
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   6120
      TabIndex        =   169
      Top             =   11850
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   6840
      TabIndex        =   168
      Top             =   11850
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   5400
      TabIndex        =   167
      Top             =   11370
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   13
      Left            =   4080
      TabIndex        =   166
      Top             =   11370
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   13
      Left            =   2760
      TabIndex        =   165
      Top             =   11370
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   13
      Left            =   1140
      TabIndex        =   164
      Top             =   11265
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   13
      Left            =   480
      TabIndex        =   163
      Top             =   11370
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   13
      Left            =   80
      TabIndex        =   162
      Top             =   11370
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   13
      Left            =   1140
      TabIndex        =   161
      Top             =   11475
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   6120
      TabIndex        =   160
      Top             =   11370
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   6840
      TabIndex        =   159
      Top             =   11370
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   5400
      TabIndex        =   158
      Top             =   10905
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   12
      Left            =   4080
      TabIndex        =   157
      Top             =   10905
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   12
      Left            =   2760
      TabIndex        =   156
      Top             =   10905
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   12
      Left            =   1140
      TabIndex        =   155
      Top             =   10800
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   12
      Left            =   480
      TabIndex        =   154
      Top             =   10905
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   12
      Left            =   80
      TabIndex        =   153
      Top             =   10905
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   12
      Left            =   1140
      TabIndex        =   152
      Top             =   11010
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   6120
      TabIndex        =   151
      Top             =   10905
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   6840
      TabIndex        =   150
      Top             =   10905
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   5400
      TabIndex        =   149
      Top             =   10410
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   11
      Left            =   4080
      TabIndex        =   148
      Top             =   10410
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   11
      Left            =   2760
      TabIndex        =   147
      Top             =   10410
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   11
      Left            =   1140
      TabIndex        =   146
      Top             =   10305
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   11
      Left            =   480
      TabIndex        =   145
      Top             =   10410
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   11
      Left            =   80
      TabIndex        =   144
      Top             =   10410
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   11
      Left            =   1140
      TabIndex        =   143
      Top             =   10515
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   6120
      TabIndex        =   142
      Top             =   10410
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   6840
      TabIndex        =   141
      Top             =   10410
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   5400
      TabIndex        =   140
      Top             =   9930
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   10
      Left            =   4080
      TabIndex        =   139
      Top             =   9930
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   10
      Left            =   2760
      TabIndex        =   138
      Top             =   9930
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   10
      Left            =   1140
      TabIndex        =   137
      Top             =   9825
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   10
      Left            =   480
      TabIndex        =   136
      Top             =   9930
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   10
      Left            =   80
      TabIndex        =   135
      Top             =   9930
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   10
      Left            =   1140
      TabIndex        =   134
      Top             =   10035
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   6120
      TabIndex        =   133
      Top             =   9930
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   6840
      TabIndex        =   132
      Top             =   9930
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   5400
      TabIndex        =   131
      Top             =   9465
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   9
      Left            =   4080
      TabIndex        =   130
      Top             =   9465
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   9
      Left            =   2760
      TabIndex        =   129
      Top             =   9465
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   9
      Left            =   1140
      TabIndex        =   128
      Top             =   9360
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   9
      Left            =   480
      TabIndex        =   127
      Top             =   9465
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   9
      Left            =   80
      TabIndex        =   126
      Top             =   9465
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   9
      Left            =   1140
      TabIndex        =   125
      Top             =   9570
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   6120
      TabIndex        =   124
      Top             =   9465
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   6840
      TabIndex        =   123
      Top             =   9465
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   5400
      TabIndex        =   122
      Top             =   8970
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   8
      Left            =   4080
      TabIndex        =   121
      Top             =   8970
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   8
      Left            =   2760
      TabIndex        =   120
      Top             =   8970
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   8
      Left            =   1140
      TabIndex        =   119
      Top             =   8865
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   8
      Left            =   480
      TabIndex        =   118
      Top             =   8970
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   8
      Left            =   80
      TabIndex        =   117
      Top             =   8970
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   8
      Left            =   1140
      TabIndex        =   116
      Top             =   9075
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   6120
      TabIndex        =   115
      Top             =   8970
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   6840
      TabIndex        =   114
      Top             =   8970
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   5400
      TabIndex        =   113
      Top             =   8490
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   7
      Left            =   4080
      TabIndex        =   112
      Top             =   8490
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   7
      Left            =   2760
      TabIndex        =   111
      Top             =   8490
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   7
      Left            =   1140
      TabIndex        =   110
      Top             =   8385
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   7
      Left            =   480
      TabIndex        =   109
      Top             =   8490
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   7
      Left            =   80
      TabIndex        =   108
      Top             =   8490
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   7
      Left            =   1140
      TabIndex        =   107
      Top             =   8595
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   6120
      TabIndex        =   106
      Top             =   8490
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   6840
      TabIndex        =   105
      Top             =   8490
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   5400
      TabIndex        =   104
      Top             =   8025
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   6
      Left            =   4080
      TabIndex        =   103
      Top             =   8025
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   6
      Left            =   2760
      TabIndex        =   102
      Top             =   8025
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   6
      Left            =   1140
      TabIndex        =   101
      Top             =   7920
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   6
      Left            =   480
      TabIndex        =   100
      Top             =   8025
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   6
      Left            =   80
      TabIndex        =   99
      Top             =   8025
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   6
      Left            =   1140
      TabIndex        =   98
      Top             =   8130
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   6120
      TabIndex        =   97
      Top             =   8025
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   6840
      TabIndex        =   96
      Top             =   8025
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   5400
      TabIndex        =   95
      Top             =   7530
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   5
      Left            =   4080
      TabIndex        =   94
      Top             =   7530
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   5
      Left            =   2760
      TabIndex        =   93
      Top             =   7530
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   5
      Left            =   1140
      TabIndex        =   92
      Top             =   7425
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   91
      Top             =   7530
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   5
      Left            =   80
      TabIndex        =   90
      Top             =   7530
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   5
      Left            =   1140
      TabIndex        =   89
      Top             =   7635
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   6120
      TabIndex        =   88
      Top             =   7530
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   6840
      TabIndex        =   87
      Top             =   7530
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   5400
      TabIndex        =   86
      Top             =   7050
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   4
      Left            =   4080
      TabIndex        =   85
      Top             =   7050
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   4
      Left            =   2760
      TabIndex        =   84
      Top             =   7050
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   4
      Left            =   1140
      TabIndex        =   83
      Top             =   6945
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   82
      Top             =   7050
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   4
      Left            =   80
      TabIndex        =   81
      Top             =   7050
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   4
      Left            =   1140
      TabIndex        =   80
      Top             =   7155
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   6120
      TabIndex        =   79
      Top             =   7050
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   6840
      TabIndex        =   78
      Top             =   7050
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   5400
      TabIndex        =   77
      Top             =   6585
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   3
      Left            =   4080
      TabIndex        =   76
      Top             =   6585
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   3
      Left            =   2760
      TabIndex        =   75
      Top             =   6585
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   3
      Left            =   1140
      TabIndex        =   74
      Top             =   6480
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   73
      Top             =   6585
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   3
      Left            =   80
      TabIndex        =   72
      Top             =   6585
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   3
      Left            =   1140
      TabIndex        =   71
      Top             =   6690
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   6120
      TabIndex        =   70
      Top             =   6585
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   6840
      TabIndex        =   69
      Top             =   6585
      Width           =   705
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   5400
      TabIndex        =   68
      Top             =   6105
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   2
      Left            =   4080
      TabIndex        =   67
      Top             =   6105
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   2
      Left            =   2760
      TabIndex        =   66
      Top             =   6105
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   2
      Left            =   1140
      TabIndex        =   65
      Top             =   6000
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   64
      Top             =   6105
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   2
      Left            =   80
      TabIndex        =   63
      Top             =   6105
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   2
      Left            =   1140
      TabIndex        =   62
      Top             =   6210
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   6120
      TabIndex        =   61
      Top             =   6105
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   6840
      TabIndex        =   60
      Top             =   6105
      Width           =   700
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   5400
      TabIndex        =   59
      Top             =   5625
      Width           =   700
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   1
      Left            =   4080
      TabIndex        =   58
      Top             =   5625
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   1
      Left            =   2760
      TabIndex        =   57
      Top             =   5625
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   1
      Left            =   1140
      TabIndex        =   56
      Top             =   5520
      Width           =   1260
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   55
      Top             =   5625
      Width           =   540
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   1
      Left            =   80
      TabIndex        =   54
      Top             =   5625
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   1
      Left            =   1140
      TabIndex        =   53
      Top             =   5730
      Width           =   1005
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   6120
      TabIndex        =   52
      Top             =   5625
      Width           =   700
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   6840
      TabIndex        =   51
      Top             =   5625
      Width           =   700
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      Height          =   210
      Index           =   14
      Left            =   9360
      TabIndex        =   50
      Top             =   4710
      Width           =   120
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      Height          =   210
      Index           =   13
      Left            =   8640
      TabIndex        =   49
      Top             =   4710
      Width           =   120
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      Height          =   210
      Index           =   12
      Left            =   7800
      TabIndex        =   48
      Top             =   4710
      Width           =   240
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configurações dos Números de Corrida"
      Height          =   210
      Index           =   11
      Left            =   7920
      TabIndex        =   47
      Top             =   4455
      Width           =   2865
   End
   Begin VB.Line LV 
      Index           =   10
      X1              =   8280
      X2              =   8280
      Y1              =   4680
      Y2              =   14280
   End
   Begin VB.Line LV 
      Index           =   5
      X1              =   9000
      X2              =   9000
      Y1              =   4680
      Y2              =   14280
   End
   Begin VB.Line LH 
      Index           =   2
      X1              =   8640
      X2              =   11160
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label LB_Eme 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   6840
      TabIndex        =   46
      Top             =   5160
      Width           =   700
   End
   Begin VB.Label LB_Epe 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   6120
      TabIndex        =   45
      Top             =   5160
      Width           =   700
   End
   Begin VB.Line LV 
      Index           =   3
      X1              =   7560
      X2              =   7560
      Y1              =   4440
      Y2              =   14280
   End
   Begin VB.Line LV 
      Index           =   1
      X1              =   3960
      X2              =   3960
      Y1              =   4440
      Y2              =   4920
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peça"
      Height          =   210
      Index           =   32
      Left            =   1200
      TabIndex        =   44
      Top             =   4680
      Width           =   360
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   0
      Left            =   1140
      TabIndex        =   43
      Top             =   5265
      Width           =   1005
   End
   Begin VB.Line LV 
      Index           =   6
      X1              =   1080
      X2              =   1080
      Y1              =   4440
      Y2              =   4920
   End
   Begin VB.Line LH 
      Index           =   15
      X1              =   0
      X2              =   11160
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line LV 
      Index           =   7
      X1              =   360
      X2              =   360
      Y1              =   4440
      Y2              =   4920
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   210
      Index           =   31
      Left            =   15
      TabIndex        =   42
      Top             =   4575
      Width           =   285
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quant."
      Height          =   210
      Index           =   30
      Left            =   495
      TabIndex        =   41
      Top             =   4575
      Width           =   480
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      Height          =   210
      Index           =   29
      Left            =   1200
      TabIndex        =   40
      Top             =   4485
      Width           =   735
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bitola"
      Height          =   210
      Index           =   21
      Left            =   2760
      TabIndex        =   39
      Top             =   4560
      Width           =   390
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
      Height          =   210
      Index           =   20
      Left            =   4080
      TabIndex        =   38
      Top             =   4575
      Width           =   555
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "COM"
      Height          =   210
      Index           =   24
      Left            =   5520
      TabIndex        =   37
      Top             =   4710
      Width           =   465
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "PA"
      Height          =   210
      Index           =   25
      Left            =   6360
      TabIndex        =   36
      Top             =   4710
      Width           =   345
   End
   Begin VB.Label LB_Fixo 
      BackStyle       =   0  'Transparent
      Caption         =   "MP"
      Height          =   210
      Index           =   26
      Left            =   7080
      TabIndex        =   35
      Top             =   4710
      Width           =   345
   End
   Begin VB.Line LH 
      Index           =   12
      X1              =   5400
      X2              =   8640
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo de Estoque Hoje"
      Height          =   210
      Index           =   19
      Left            =   5700
      TabIndex        =   34
      Top             =   4455
      Width           =   1620
   End
   Begin VB.Line LV 
      Index           =   8
      X1              =   6840
      X2              =   6840
      Y1              =   4680
      Y2              =   4920
   End
   Begin VB.Line LV 
      Index           =   9
      X1              =   6120
      X2              =   6120
      Y1              =   4680
      Y2              =   4920
   End
   Begin VB.Line LV 
      Index           =   13
      X1              =   2640
      X2              =   2640
      Y1              =   4440
      Y2              =   4920
   End
   Begin VB.Line LV 
      Index           =   14
      X1              =   5400
      X2              =   5400
      Y1              =   4440
      Y2              =   4920
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   0
      Left            =   80
      TabIndex        =   33
      Top             =   5160
      Width           =   180
   End
   Begin VB.Label LB_Qua 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10 pçs."
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   32
      Top             =   5160
      Width           =   540
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   0
      Left            =   1140
      TabIndex        =   31
      Top             =   5055
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   0
      Left            =   2760
      TabIndex        =   30
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   0
      Left            =   4080
      TabIndex        =   29
      Top             =   5160
      Width           =   1125
   End
   Begin VB.Label LB_Ece 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   5400
      TabIndex        =   28
      Top             =   5160
      Width           =   700
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
      Index           =   57
      Left            =   5760
      TabIndex        =   27
      Top             =   3120
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
      Index           =   56
      Left            =   7185
      TabIndex        =   26
      Top             =   3120
      Width           =   1170
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
      Index           =   55
      Left            =   8760
      TabIndex        =   25
      Top             =   3120
      Width           =   900
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
      Index           =   54
      Left            =   10140
      TabIndex        =   24
      Top             =   3120
      Width           =   945
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
      Left            =   5640
      TabIndex        =   23
      Top             =   3525
      Width           =   1380
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
      Left            =   8520
      TabIndex        =   22
      Top             =   3525
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
      Left            =   7200
      TabIndex        =   21
      Top             =   3525
      Width           =   1110
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
      Left            =   10080
      TabIndex        =   20
      Top             =   3525
      Width           =   1110
   End
   Begin VB.Line LV 
      Index           =   46
      X1              =   5520
      X2              =   5520
      Y1              =   3000
      Y2              =   3840
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
      TabIndex        =   19
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista de componentes necessários para montagem:"
      BeginProperty Font 
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
      TabIndex        =   18
      Top             =   4080
      Width           =   4500
   End
   Begin VB.Line LH 
      Index           =   26
      X1              =   0
      X2              =   11160
      Y1              =   4410
      Y2              =   4410
   End
   Begin VB.Line LH 
      Index           =   25
      X1              =   0
      X2              =   11160
      Y1              =   4440
      Y2              =   4440
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
      Left            =   6240
      TabIndex        =   17
      Top             =   2190
      UseMnemonic     =   0   'False
      Width           =   585
   End
   Begin VB.Label LB_Material 
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
      Left            =   6240
      TabIndex        =   16
      Top             =   2520
      Width           =   510
   End
   Begin VB.Line LV 
      Index           =   22
      X1              =   6120
      X2              =   6120
      Y1              =   2160
      Y2              =   3000
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
      Left            =   4440
      TabIndex        =   15
      Top             =   2190
      UseMnemonic     =   0   'False
      Width           =   405
   End
   Begin VB.Label LB_Bitola 
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
      Left            =   4440
      TabIndex        =   14
      Top             =   2520
      Width           =   345
   End
   Begin VB.Line LV 
      Index           =   21
      X1              =   4320
      X2              =   4320
      Y1              =   2160
      Y2              =   3000
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
      Left            =   2520
      TabIndex        =   13
      Top             =   2190
      UseMnemonic     =   0   'False
      Width           =   450
   End
   Begin VB.Label LB_Figura 
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
      Left            =   2520
      TabIndex        =   12
      Top             =   2520
      Width           =   1050
   End
   Begin VB.Line LH 
      Index           =   8
      X1              =   0
      X2              =   11160
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTIDADE MONTADA"
      BeginProperty Font 
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
      Left            =   9720
      TabIndex        =   11
      Top             =   2190
      UseMnemonic     =   0   'False
      Width           =   1410
   End
   Begin VB.Line LV 
      Index           =   20
      X1              =   9600
      X2              =   9600
      Y1              =   2160
      Y2              =   3000
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
      Top             =   3030
      UseMnemonic     =   0   'False
      Width           =   705
   End
   Begin VB.Label LB_Descricao 
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
      Height          =   480
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   5400
   End
   Begin VB.Line LV 
      Index           =   4
      X1              =   2400
      X2              =   2400
      Y1              =   2160
      Y2              =   3000
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANT.ESTIPULADA"
      BeginProperty Font 
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
      Left            =   1100
      TabIndex        =   8
      Top             =   2190
      UseMnemonic     =   0   'False
      Width           =   1170
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
      Left            =   1200
      TabIndex        =   7
      Top             =   2520
      Width           =   435
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "OM"
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
      Left            =   7800
      TabIndex        =   6
      Top             =   900
      Width           =   375
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
      Caption         =   "Informações sobre a peça para ser montada:"
      BeginProperty Font 
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
      Top             =   1800
      Width           =   3870
   End
   Begin VB.Line LH 
      Index           =   14
      X1              =   0
      X2              =   11160
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line LH 
      Index           =   13
      X1              =   0
      X2              =   11160
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line LH 
      Index           =   11
      X1              =   0
      X2              =   11160
      Y1              =   2130
      Y2              =   2130
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
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ORDEM DE MONTAGEM"
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
      Left            =   6960
      TabIndex        =   4
      Top             =   360
      Width           =   2775
   End
   Begin VB.Line LH 
      Index           =   0
      X1              =   0
      X2              =   11160
      Y1              =   120
      Y2              =   120
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
      Left            =   8280
      TabIndex        =   2
      Top             =   855
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
      Top             =   2520
      Width           =   960
   End
   Begin VB.Line LV 
      Index           =   0
      X1              =   1080
      X2              =   1080
      Y1              =   2160
      Y2              =   3000
   End
   Begin VB.Line LH 
      Index           =   3
      X1              =   0
      X2              =   11160
      Y1              =   2160
      Y2              =   2160
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
      Top             =   2190
      UseMnemonic     =   0   'False
      Width           =   315
   End
   Begin VB.Line LH 
      Index           =   7
      X1              =   0
      X2              =   11160
      Y1              =   14280
      Y2              =   14280
   End
End
Attribute VB_Name = "IT_OrdemMontagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
