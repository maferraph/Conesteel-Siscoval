VERSION 5.00
Begin VB.Form Tela_Cotacao_MP_IT 
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
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   14700
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   43
      Left            =   10500
      TabIndex        =   468
      Top             =   12840
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   43
      Left            =   9405
      TabIndex        =   467
      Top             =   12840
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   43
      Left            =   8325
      TabIndex        =   466
      Top             =   12840
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   43
      Left            =   7320
      TabIndex        =   465
      Top             =   12840
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   43
      Left            =   6360
      TabIndex        =   464
      Top             =   12840
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   43
      Left            =   4800
      TabIndex        =   463
      Top             =   12840
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   43
      Left            =   3600
      TabIndex        =   462
      Top             =   12840
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   43
      Left            =   2040
      TabIndex        =   461
      Top             =   12840
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   43
      Left            =   480
      TabIndex        =   460
      Top             =   12840
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "44"
      Height          =   210
      Index           =   43
      Left            =   45
      TabIndex        =   459
      Top             =   12840
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   42
      Left            =   10500
      TabIndex        =   458
      Top             =   12600
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   42
      Left            =   9405
      TabIndex        =   457
      Top             =   12600
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   42
      Left            =   8325
      TabIndex        =   456
      Top             =   12600
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   42
      Left            =   7320
      TabIndex        =   455
      Top             =   12600
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   42
      Left            =   6360
      TabIndex        =   454
      Top             =   12600
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   42
      Left            =   4800
      TabIndex        =   453
      Top             =   12600
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   42
      Left            =   3600
      TabIndex        =   452
      Top             =   12600
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   42
      Left            =   2040
      TabIndex        =   451
      Top             =   12600
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   42
      Left            =   480
      TabIndex        =   450
      Top             =   12600
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "43"
      Height          =   210
      Index           =   42
      Left            =   45
      TabIndex        =   449
      Top             =   12600
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   41
      Left            =   10500
      TabIndex        =   448
      Top             =   12360
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   41
      Left            =   9405
      TabIndex        =   447
      Top             =   12360
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   41
      Left            =   8325
      TabIndex        =   446
      Top             =   12360
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   41
      Left            =   7320
      TabIndex        =   445
      Top             =   12360
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   41
      Left            =   6360
      TabIndex        =   444
      Top             =   12360
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   41
      Left            =   4800
      TabIndex        =   443
      Top             =   12360
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   41
      Left            =   3600
      TabIndex        =   442
      Top             =   12360
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   41
      Left            =   2040
      TabIndex        =   441
      Top             =   12360
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   41
      Left            =   480
      TabIndex        =   440
      Top             =   12360
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "42"
      Height          =   210
      Index           =   41
      Left            =   45
      TabIndex        =   439
      Top             =   12360
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   40
      Left            =   10500
      TabIndex        =   438
      Top             =   12120
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   40
      Left            =   9405
      TabIndex        =   437
      Top             =   12120
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   40
      Left            =   8325
      TabIndex        =   436
      Top             =   12120
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   40
      Left            =   7320
      TabIndex        =   435
      Top             =   12120
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   40
      Left            =   6360
      TabIndex        =   434
      Top             =   12120
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   40
      Left            =   4800
      TabIndex        =   433
      Top             =   12120
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   40
      Left            =   3600
      TabIndex        =   432
      Top             =   12120
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   40
      Left            =   2040
      TabIndex        =   431
      Top             =   12120
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   40
      Left            =   480
      TabIndex        =   430
      Top             =   12120
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "41"
      Height          =   210
      Index           =   40
      Left            =   45
      TabIndex        =   429
      Top             =   12120
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   39
      Left            =   10500
      TabIndex        =   428
      Top             =   11880
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   39
      Left            =   9405
      TabIndex        =   427
      Top             =   11880
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   39
      Left            =   8325
      TabIndex        =   426
      Top             =   11880
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   39
      Left            =   7320
      TabIndex        =   425
      Top             =   11880
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   39
      Left            =   6360
      TabIndex        =   424
      Top             =   11880
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   39
      Left            =   4800
      TabIndex        =   423
      Top             =   11880
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   39
      Left            =   3600
      TabIndex        =   422
      Top             =   11880
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   39
      Left            =   2040
      TabIndex        =   421
      Top             =   11880
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   39
      Left            =   480
      TabIndex        =   420
      Top             =   11880
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      Height          =   210
      Index           =   39
      Left            =   45
      TabIndex        =   419
      Top             =   11880
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   38
      Left            =   10500
      TabIndex        =   418
      Top             =   11640
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   38
      Left            =   9405
      TabIndex        =   417
      Top             =   11640
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   38
      Left            =   8325
      TabIndex        =   416
      Top             =   11640
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   38
      Left            =   7320
      TabIndex        =   415
      Top             =   11640
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   38
      Left            =   6360
      TabIndex        =   414
      Top             =   11640
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   38
      Left            =   4800
      TabIndex        =   413
      Top             =   11640
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   38
      Left            =   3600
      TabIndex        =   412
      Top             =   11640
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   38
      Left            =   2040
      TabIndex        =   411
      Top             =   11640
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   38
      Left            =   480
      TabIndex        =   410
      Top             =   11640
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "39"
      Height          =   210
      Index           =   38
      Left            =   45
      TabIndex        =   409
      Top             =   11640
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   37
      Left            =   10500
      TabIndex        =   408
      Top             =   11400
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   37
      Left            =   9405
      TabIndex        =   407
      Top             =   11400
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   37
      Left            =   8325
      TabIndex        =   406
      Top             =   11400
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   37
      Left            =   7320
      TabIndex        =   405
      Top             =   11400
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   37
      Left            =   6360
      TabIndex        =   404
      Top             =   11400
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   37
      Left            =   4800
      TabIndex        =   403
      Top             =   11400
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   37
      Left            =   3600
      TabIndex        =   402
      Top             =   11400
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   37
      Left            =   2040
      TabIndex        =   401
      Top             =   11400
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   37
      Left            =   480
      TabIndex        =   400
      Top             =   11400
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "38"
      Height          =   210
      Index           =   37
      Left            =   45
      TabIndex        =   399
      Top             =   11400
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   36
      Left            =   10500
      TabIndex        =   398
      Top             =   11160
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   36
      Left            =   9405
      TabIndex        =   397
      Top             =   11160
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   36
      Left            =   8325
      TabIndex        =   396
      Top             =   11160
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   36
      Left            =   7320
      TabIndex        =   395
      Top             =   11160
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   36
      Left            =   6360
      TabIndex        =   394
      Top             =   11160
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   36
      Left            =   4800
      TabIndex        =   393
      Top             =   11160
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   36
      Left            =   3600
      TabIndex        =   392
      Top             =   11160
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   36
      Left            =   2040
      TabIndex        =   391
      Top             =   11160
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   36
      Left            =   480
      TabIndex        =   390
      Top             =   11160
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "37"
      Height          =   210
      Index           =   36
      Left            =   45
      TabIndex        =   389
      Top             =   11160
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   35
      Left            =   10500
      TabIndex        =   388
      Top             =   10920
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   35
      Left            =   9405
      TabIndex        =   387
      Top             =   10920
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   35
      Left            =   8325
      TabIndex        =   386
      Top             =   10920
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   35
      Left            =   7320
      TabIndex        =   385
      Top             =   10920
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   35
      Left            =   6360
      TabIndex        =   384
      Top             =   10920
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   35
      Left            =   4800
      TabIndex        =   383
      Top             =   10920
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   35
      Left            =   3600
      TabIndex        =   382
      Top             =   10920
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   35
      Left            =   2040
      TabIndex        =   381
      Top             =   10920
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   35
      Left            =   480
      TabIndex        =   380
      Top             =   10920
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "36"
      Height          =   210
      Index           =   35
      Left            =   45
      TabIndex        =   379
      Top             =   10920
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   34
      Left            =   10500
      TabIndex        =   378
      Top             =   10680
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   34
      Left            =   9405
      TabIndex        =   377
      Top             =   10680
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   34
      Left            =   8325
      TabIndex        =   376
      Top             =   10680
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   34
      Left            =   7320
      TabIndex        =   375
      Top             =   10680
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   34
      Left            =   6360
      TabIndex        =   374
      Top             =   10680
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   34
      Left            =   4800
      TabIndex        =   373
      Top             =   10680
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   34
      Left            =   3600
      TabIndex        =   372
      Top             =   10680
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   34
      Left            =   2040
      TabIndex        =   371
      Top             =   10680
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   34
      Left            =   480
      TabIndex        =   370
      Top             =   10680
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "35"
      Height          =   210
      Index           =   34
      Left            =   45
      TabIndex        =   369
      Top             =   10680
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   33
      Left            =   10500
      TabIndex        =   368
      Top             =   10440
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   33
      Left            =   9405
      TabIndex        =   367
      Top             =   10440
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   33
      Left            =   8325
      TabIndex        =   366
      Top             =   10440
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   33
      Left            =   7320
      TabIndex        =   365
      Top             =   10440
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   33
      Left            =   6360
      TabIndex        =   364
      Top             =   10440
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   33
      Left            =   4800
      TabIndex        =   363
      Top             =   10440
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   33
      Left            =   3600
      TabIndex        =   362
      Top             =   10440
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   33
      Left            =   2040
      TabIndex        =   361
      Top             =   10440
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   33
      Left            =   480
      TabIndex        =   360
      Top             =   10440
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "34"
      Height          =   210
      Index           =   33
      Left            =   45
      TabIndex        =   359
      Top             =   10440
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   32
      Left            =   10500
      TabIndex        =   358
      Top             =   10200
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   32
      Left            =   9405
      TabIndex        =   357
      Top             =   10200
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   32
      Left            =   8325
      TabIndex        =   356
      Top             =   10200
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   32
      Left            =   7320
      TabIndex        =   355
      Top             =   10200
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   32
      Left            =   6360
      TabIndex        =   354
      Top             =   10200
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   32
      Left            =   4800
      TabIndex        =   353
      Top             =   10200
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   32
      Left            =   3600
      TabIndex        =   352
      Top             =   10200
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   32
      Left            =   2040
      TabIndex        =   351
      Top             =   10200
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   32
      Left            =   480
      TabIndex        =   350
      Top             =   10200
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "33"
      Height          =   210
      Index           =   32
      Left            =   45
      TabIndex        =   349
      Top             =   10200
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   31
      Left            =   10500
      TabIndex        =   348
      Top             =   9960
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   31
      Left            =   9405
      TabIndex        =   347
      Top             =   9960
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   31
      Left            =   8325
      TabIndex        =   346
      Top             =   9960
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   31
      Left            =   7320
      TabIndex        =   345
      Top             =   9960
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   31
      Left            =   6360
      TabIndex        =   344
      Top             =   9960
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   31
      Left            =   4800
      TabIndex        =   343
      Top             =   9960
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   31
      Left            =   3600
      TabIndex        =   342
      Top             =   9960
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   31
      Left            =   2040
      TabIndex        =   341
      Top             =   9960
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   31
      Left            =   480
      TabIndex        =   340
      Top             =   9960
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      Height          =   210
      Index           =   31
      Left            =   45
      TabIndex        =   339
      Top             =   9960
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   30
      Left            =   10500
      TabIndex        =   338
      Top             =   9720
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   30
      Left            =   9405
      TabIndex        =   337
      Top             =   9720
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   30
      Left            =   8325
      TabIndex        =   336
      Top             =   9720
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   30
      Left            =   7320
      TabIndex        =   335
      Top             =   9720
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   30
      Left            =   6360
      TabIndex        =   334
      Top             =   9720
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   30
      Left            =   4800
      TabIndex        =   333
      Top             =   9720
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   30
      Left            =   3600
      TabIndex        =   332
      Top             =   9720
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   30
      Left            =   2040
      TabIndex        =   331
      Top             =   9720
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   30
      Left            =   480
      TabIndex        =   330
      Top             =   9720
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "31"
      Height          =   210
      Index           =   30
      Left            =   45
      TabIndex        =   329
      Top             =   9720
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   29
      Left            =   10500
      TabIndex        =   328
      Top             =   9480
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   29
      Left            =   9405
      TabIndex        =   327
      Top             =   9480
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   29
      Left            =   8325
      TabIndex        =   326
      Top             =   9480
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   29
      Left            =   7320
      TabIndex        =   325
      Top             =   9480
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   29
      Left            =   6360
      TabIndex        =   324
      Top             =   9480
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   29
      Left            =   4800
      TabIndex        =   323
      Top             =   9480
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   29
      Left            =   3600
      TabIndex        =   322
      Top             =   9480
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   29
      Left            =   2040
      TabIndex        =   321
      Top             =   9480
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   29
      Left            =   480
      TabIndex        =   320
      Top             =   9480
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30"
      Height          =   210
      Index           =   29
      Left            =   45
      TabIndex        =   319
      Top             =   9480
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   28
      Left            =   10500
      TabIndex        =   318
      Top             =   9240
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   28
      Left            =   9405
      TabIndex        =   317
      Top             =   9240
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   28
      Left            =   8325
      TabIndex        =   316
      Top             =   9240
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   28
      Left            =   7320
      TabIndex        =   315
      Top             =   9240
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   28
      Left            =   6360
      TabIndex        =   314
      Top             =   9240
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   28
      Left            =   4800
      TabIndex        =   313
      Top             =   9240
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   28
      Left            =   3600
      TabIndex        =   312
      Top             =   9240
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   28
      Left            =   2040
      TabIndex        =   311
      Top             =   9240
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   28
      Left            =   480
      TabIndex        =   310
      Top             =   9240
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "29"
      Height          =   210
      Index           =   28
      Left            =   45
      TabIndex        =   309
      Top             =   9240
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   27
      Left            =   10500
      TabIndex        =   308
      Top             =   9000
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   27
      Left            =   9405
      TabIndex        =   307
      Top             =   9000
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   27
      Left            =   8325
      TabIndex        =   306
      Top             =   9000
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   27
      Left            =   7320
      TabIndex        =   305
      Top             =   9000
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   27
      Left            =   6360
      TabIndex        =   304
      Top             =   9000
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   27
      Left            =   4800
      TabIndex        =   303
      Top             =   9000
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   27
      Left            =   3600
      TabIndex        =   302
      Top             =   9000
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   27
      Left            =   2040
      TabIndex        =   301
      Top             =   9000
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   27
      Left            =   480
      TabIndex        =   300
      Top             =   9000
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "28"
      Height          =   210
      Index           =   27
      Left            =   45
      TabIndex        =   299
      Top             =   9000
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   26
      Left            =   10500
      TabIndex        =   298
      Top             =   8760
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   26
      Left            =   9405
      TabIndex        =   297
      Top             =   8760
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   26
      Left            =   8325
      TabIndex        =   296
      Top             =   8760
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   26
      Left            =   7320
      TabIndex        =   295
      Top             =   8760
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   26
      Left            =   6360
      TabIndex        =   294
      Top             =   8760
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   26
      Left            =   4800
      TabIndex        =   293
      Top             =   8760
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   26
      Left            =   3600
      TabIndex        =   292
      Top             =   8760
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   26
      Left            =   2040
      TabIndex        =   291
      Top             =   8760
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   26
      Left            =   480
      TabIndex        =   290
      Top             =   8760
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "27"
      Height          =   210
      Index           =   26
      Left            =   45
      TabIndex        =   289
      Top             =   8760
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   25
      Left            =   10500
      TabIndex        =   288
      Top             =   8520
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   25
      Left            =   9405
      TabIndex        =   287
      Top             =   8520
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   25
      Left            =   8325
      TabIndex        =   286
      Top             =   8520
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   25
      Left            =   7320
      TabIndex        =   285
      Top             =   8520
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   25
      Left            =   6360
      TabIndex        =   284
      Top             =   8520
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   25
      Left            =   4800
      TabIndex        =   283
      Top             =   8520
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   25
      Left            =   3600
      TabIndex        =   282
      Top             =   8520
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   25
      Left            =   2040
      TabIndex        =   281
      Top             =   8520
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   25
      Left            =   480
      TabIndex        =   280
      Top             =   8520
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "26"
      Height          =   210
      Index           =   25
      Left            =   45
      TabIndex        =   279
      Top             =   8520
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   24
      Left            =   10500
      TabIndex        =   278
      Top             =   8280
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   24
      Left            =   9405
      TabIndex        =   277
      Top             =   8280
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   24
      Left            =   8325
      TabIndex        =   276
      Top             =   8280
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   24
      Left            =   7320
      TabIndex        =   275
      Top             =   8280
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   24
      Left            =   6360
      TabIndex        =   274
      Top             =   8280
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   24
      Left            =   4800
      TabIndex        =   273
      Top             =   8280
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   24
      Left            =   3600
      TabIndex        =   272
      Top             =   8280
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   24
      Left            =   2040
      TabIndex        =   271
      Top             =   8280
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   24
      Left            =   480
      TabIndex        =   270
      Top             =   8280
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "25"
      Height          =   210
      Index           =   24
      Left            =   45
      TabIndex        =   269
      Top             =   8280
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   23
      Left            =   10500
      TabIndex        =   268
      Top             =   8040
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   23
      Left            =   9400
      TabIndex        =   267
      Top             =   8040
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   23
      Left            =   8320
      TabIndex        =   266
      Top             =   8040
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   23
      Left            =   7320
      TabIndex        =   265
      Top             =   8040
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   23
      Left            =   6360
      TabIndex        =   264
      Top             =   8040
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   23
      Left            =   4800
      TabIndex        =   263
      Top             =   8040
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   23
      Left            =   3600
      TabIndex        =   262
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   23
      Left            =   2040
      TabIndex        =   261
      Top             =   8040
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   23
      Left            =   480
      TabIndex        =   260
      Top             =   8040
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      Height          =   210
      Index           =   23
      Left            =   45
      TabIndex        =   259
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   22
      Left            =   10500
      TabIndex        =   258
      Top             =   7800
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   22
      Left            =   9400
      TabIndex        =   257
      Top             =   7800
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   22
      Left            =   8320
      TabIndex        =   256
      Top             =   7800
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   22
      Left            =   7320
      TabIndex        =   255
      Top             =   7800
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   22
      Left            =   6360
      TabIndex        =   254
      Top             =   7800
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   22
      Left            =   4800
      TabIndex        =   253
      Top             =   7800
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   22
      Left            =   3600
      TabIndex        =   252
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   22
      Left            =   2040
      TabIndex        =   251
      Top             =   7800
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   22
      Left            =   480
      TabIndex        =   250
      Top             =   7800
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "23"
      Height          =   210
      Index           =   22
      Left            =   45
      TabIndex        =   249
      Top             =   7800
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   21
      Left            =   10500
      TabIndex        =   248
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   21
      Left            =   9400
      TabIndex        =   247
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   21
      Left            =   8320
      TabIndex        =   246
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   21
      Left            =   7320
      TabIndex        =   245
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   21
      Left            =   6360
      TabIndex        =   244
      Top             =   7560
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   21
      Left            =   4800
      TabIndex        =   243
      Top             =   7560
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   21
      Left            =   3600
      TabIndex        =   242
      Top             =   7560
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   21
      Left            =   2040
      TabIndex        =   241
      Top             =   7560
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   21
      Left            =   480
      TabIndex        =   240
      Top             =   7560
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "22"
      Height          =   210
      Index           =   21
      Left            =   45
      TabIndex        =   239
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   20
      Left            =   10500
      TabIndex        =   238
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   20
      Left            =   9400
      TabIndex        =   237
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   20
      Left            =   8320
      TabIndex        =   236
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   20
      Left            =   7320
      TabIndex        =   235
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   20
      Left            =   6360
      TabIndex        =   234
      Top             =   7320
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   20
      Left            =   4800
      TabIndex        =   233
      Top             =   7320
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   20
      Left            =   3600
      TabIndex        =   232
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   20
      Left            =   2040
      TabIndex        =   231
      Top             =   7320
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   20
      Left            =   480
      TabIndex        =   230
      Top             =   7320
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "21"
      Height          =   210
      Index           =   20
      Left            =   45
      TabIndex        =   229
      Top             =   7320
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   19
      Left            =   10500
      TabIndex        =   228
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   19
      Left            =   9400
      TabIndex        =   227
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   19
      Left            =   8320
      TabIndex        =   226
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   19
      Left            =   7320
      TabIndex        =   225
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   19
      Left            =   6360
      TabIndex        =   224
      Top             =   7080
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   19
      Left            =   4800
      TabIndex        =   223
      Top             =   7080
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   19
      Left            =   3600
      TabIndex        =   222
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   19
      Left            =   2040
      TabIndex        =   221
      Top             =   7080
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   19
      Left            =   480
      TabIndex        =   220
      Top             =   7080
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      Height          =   210
      Index           =   19
      Left            =   45
      TabIndex        =   219
      Top             =   7080
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   10500
      TabIndex        =   218
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   9400
      TabIndex        =   217
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   8320
      TabIndex        =   216
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   7320
      TabIndex        =   215
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   18
      Left            =   6360
      TabIndex        =   214
      Top             =   6840
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   18
      Left            =   4800
      TabIndex        =   213
      Top             =   6840
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   18
      Left            =   3600
      TabIndex        =   212
      Top             =   6840
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   18
      Left            =   2040
      TabIndex        =   211
      Top             =   6840
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   18
      Left            =   480
      TabIndex        =   210
      Top             =   6840
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "19"
      Height          =   210
      Index           =   18
      Left            =   45
      TabIndex        =   209
      Top             =   6840
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   10500
      TabIndex        =   208
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   9400
      TabIndex        =   207
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   8320
      TabIndex        =   206
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   7320
      TabIndex        =   205
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   17
      Left            =   6360
      TabIndex        =   204
      Top             =   6600
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   17
      Left            =   4800
      TabIndex        =   203
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   17
      Left            =   3600
      TabIndex        =   202
      Top             =   6600
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   17
      Left            =   2040
      TabIndex        =   201
      Top             =   6600
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   17
      Left            =   480
      TabIndex        =   200
      Top             =   6600
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      Height          =   210
      Index           =   17
      Left            =   45
      TabIndex        =   199
      Top             =   6600
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   10500
      TabIndex        =   198
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   9400
      TabIndex        =   197
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   8320
      TabIndex        =   196
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   7320
      TabIndex        =   195
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   16
      Left            =   6360
      TabIndex        =   194
      Top             =   6360
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   16
      Left            =   4800
      TabIndex        =   193
      Top             =   6360
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   16
      Left            =   3600
      TabIndex        =   192
      Top             =   6360
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   16
      Left            =   2040
      TabIndex        =   191
      Top             =   6360
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   16
      Left            =   480
      TabIndex        =   190
      Top             =   6360
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "17"
      Height          =   210
      Index           =   16
      Left            =   45
      TabIndex        =   189
      Top             =   6360
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   10500
      TabIndex        =   188
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   9400
      TabIndex        =   187
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   8320
      TabIndex        =   186
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   7320
      TabIndex        =   185
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   15
      Left            =   6360
      TabIndex        =   184
      Top             =   6120
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   15
      Left            =   4800
      TabIndex        =   183
      Top             =   6120
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   15
      Left            =   3600
      TabIndex        =   182
      Top             =   6120
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   15
      Left            =   2040
      TabIndex        =   181
      Top             =   6120
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   15
      Left            =   480
      TabIndex        =   180
      Top             =   6120
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "16"
      Height          =   210
      Index           =   15
      Left            =   45
      TabIndex        =   179
      Top             =   6120
      Width           =   180
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   10500
      TabIndex        =   178
      Top             =   5880
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   9400
      TabIndex        =   177
      Top             =   5880
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   8320
      TabIndex        =   176
      Top             =   5880
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   7320
      TabIndex        =   175
      Top             =   5880
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   14
      Left            =   6360
      TabIndex        =   174
      Top             =   5880
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   14
      Left            =   4800
      TabIndex        =   173
      Top             =   5880
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   14
      Left            =   3600
      TabIndex        =   172
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   14
      Left            =   2040
      TabIndex        =   171
      Top             =   5880
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   14
      Left            =   480
      TabIndex        =   170
      Top             =   5880
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      Height          =   210
      Index           =   14
      Left            =   45
      TabIndex        =   169
      Top             =   5880
      Width           =   180
   End
   Begin VB.Line LH 
      Index           =   7
      X1              =   0
      X2              =   11160
      Y1              =   14280
      Y2              =   14280
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avenida Montemagno, 2.454 - Vila Formosa - São Paulo - (SP) - Brasil - Fone (011) 6910-1444 - Fax: (011) 6107-6667"
      Height          =   210
      Index           =   21
      Left            =   1320
      TabIndex        =   168
      Top             =   14400
      UseMnemonic     =   0   'False
      Width           =   8535
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      Height          =   210
      Index           =   13
      Left            =   50
      TabIndex        =   167
      Top             =   5640
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   13
      Left            =   480
      TabIndex        =   166
      Top             =   5640
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   13
      Left            =   2040
      TabIndex        =   165
      Top             =   5640
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   13
      Left            =   3600
      TabIndex        =   164
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   13
      Left            =   4800
      TabIndex        =   163
      Top             =   5640
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   6360
      TabIndex        =   162
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   7320
      TabIndex        =   161
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   8320
      TabIndex        =   160
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   9400
      TabIndex        =   159
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   13
      Left            =   10500
      TabIndex        =   158
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "13"
      Height          =   210
      Index           =   12
      Left            =   50
      TabIndex        =   157
      Top             =   5400
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   12
      Left            =   480
      TabIndex        =   156
      Top             =   5400
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   12
      Left            =   2040
      TabIndex        =   155
      Top             =   5400
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   12
      Left            =   3600
      TabIndex        =   154
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   12
      Left            =   4800
      TabIndex        =   153
      Top             =   5400
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   6360
      TabIndex        =   152
      Top             =   5400
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   7320
      TabIndex        =   151
      Top             =   5400
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   8320
      TabIndex        =   150
      Top             =   5400
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   9400
      TabIndex        =   149
      Top             =   5400
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   12
      Left            =   10500
      TabIndex        =   148
      Top             =   5400
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   210
      Index           =   11
      Left            =   50
      TabIndex        =   147
      Top             =   5160
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   11
      Left            =   480
      TabIndex        =   146
      Top             =   5160
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   11
      Left            =   2040
      TabIndex        =   145
      Top             =   5160
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   11
      Left            =   3600
      TabIndex        =   144
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   11
      Left            =   4800
      TabIndex        =   143
      Top             =   5160
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   6360
      TabIndex        =   142
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   7320
      TabIndex        =   141
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   8320
      TabIndex        =   140
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   9400
      TabIndex        =   139
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   11
      Left            =   10500
      TabIndex        =   138
      Top             =   5160
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   210
      Index           =   10
      Left            =   50
      TabIndex        =   137
      Top             =   4920
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   10
      Left            =   480
      TabIndex        =   136
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   10
      Left            =   2040
      TabIndex        =   135
      Top             =   4920
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   10
      Left            =   3600
      TabIndex        =   134
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   10
      Left            =   4800
      TabIndex        =   133
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   6360
      TabIndex        =   132
      Top             =   4920
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   7320
      TabIndex        =   131
      Top             =   4920
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   8320
      TabIndex        =   130
      Top             =   4920
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   9400
      TabIndex        =   129
      Top             =   4920
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   10
      Left            =   10500
      TabIndex        =   128
      Top             =   4920
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   210
      Index           =   9
      Left            =   50
      TabIndex        =   127
      Top             =   4680
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   9
      Left            =   480
      TabIndex        =   126
      Top             =   4680
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   9
      Left            =   2040
      TabIndex        =   125
      Top             =   4680
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   9
      Left            =   3600
      TabIndex        =   124
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   9
      Left            =   4800
      TabIndex        =   123
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   6360
      TabIndex        =   122
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   7320
      TabIndex        =   121
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   8320
      TabIndex        =   120
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   9400
      TabIndex        =   119
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   9
      Left            =   10500
      TabIndex        =   118
      Top             =   4680
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "09"
      Height          =   210
      Index           =   8
      Left            =   50
      TabIndex        =   117
      Top             =   4440
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   8
      Left            =   480
      TabIndex        =   116
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   8
      Left            =   2040
      TabIndex        =   115
      Top             =   4440
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   8
      Left            =   3600
      TabIndex        =   114
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   8
      Left            =   4800
      TabIndex        =   113
      Top             =   4440
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   6360
      TabIndex        =   112
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   7320
      TabIndex        =   111
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   8320
      TabIndex        =   110
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   9400
      TabIndex        =   109
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   8
      Left            =   10500
      TabIndex        =   108
      Top             =   4440
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "08"
      Height          =   210
      Index           =   7
      Left            =   50
      TabIndex        =   107
      Top             =   4200
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   7
      Left            =   480
      TabIndex        =   106
      Top             =   4200
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   7
      Left            =   2040
      TabIndex        =   105
      Top             =   4200
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   7
      Left            =   3600
      TabIndex        =   104
      Top             =   4200
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   7
      Left            =   4800
      TabIndex        =   103
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   6360
      TabIndex        =   102
      Top             =   4200
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   7320
      TabIndex        =   101
      Top             =   4200
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   8320
      TabIndex        =   100
      Top             =   4200
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   9400
      TabIndex        =   99
      Top             =   4200
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   7
      Left            =   10500
      TabIndex        =   98
      Top             =   4200
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "07"
      Height          =   210
      Index           =   6
      Left            =   50
      TabIndex        =   97
      Top             =   3960
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   6
      Left            =   480
      TabIndex        =   96
      Top             =   3960
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   6
      Left            =   2040
      TabIndex        =   95
      Top             =   3960
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   6
      Left            =   3600
      TabIndex        =   94
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   6
      Left            =   4800
      TabIndex        =   93
      Top             =   3960
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   6360
      TabIndex        =   92
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   7320
      TabIndex        =   91
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   8320
      TabIndex        =   90
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   9400
      TabIndex        =   89
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   6
      Left            =   10500
      TabIndex        =   88
      Top             =   3960
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "06"
      Height          =   210
      Index           =   5
      Left            =   50
      TabIndex        =   87
      Top             =   3720
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   5
      Left            =   480
      TabIndex        =   86
      Top             =   3720
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   5
      Left            =   2040
      TabIndex        =   85
      Top             =   3720
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   5
      Left            =   3600
      TabIndex        =   84
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   5
      Left            =   4800
      TabIndex        =   83
      Top             =   3720
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   6360
      TabIndex        =   82
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   7320
      TabIndex        =   81
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   8320
      TabIndex        =   80
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   9400
      TabIndex        =   79
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   5
      Left            =   10500
      TabIndex        =   78
      Top             =   3720
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "05"
      Height          =   210
      Index           =   4
      Left            =   50
      TabIndex        =   77
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   4
      Left            =   480
      TabIndex        =   76
      Top             =   3480
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   4
      Left            =   2040
      TabIndex        =   75
      Top             =   3480
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   4
      Left            =   3600
      TabIndex        =   74
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   4
      Left            =   4800
      TabIndex        =   73
      Top             =   3480
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   6360
      TabIndex        =   72
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   7320
      TabIndex        =   71
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   8320
      TabIndex        =   70
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   9400
      TabIndex        =   69
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   4
      Left            =   10500
      TabIndex        =   68
      Top             =   3480
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "04"
      Height          =   210
      Index           =   3
      Left            =   50
      TabIndex        =   67
      Top             =   3240
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   66
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   3
      Left            =   2040
      TabIndex        =   65
      Top             =   3240
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   3
      Left            =   3600
      TabIndex        =   64
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   3
      Left            =   4800
      TabIndex        =   63
      Top             =   3240
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   6360
      TabIndex        =   62
      Top             =   3240
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   7320
      TabIndex        =   61
      Top             =   3240
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   8320
      TabIndex        =   60
      Top             =   3240
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   9400
      TabIndex        =   59
      Top             =   3240
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   3
      Left            =   10500
      TabIndex        =   58
      Top             =   3240
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "03"
      Height          =   210
      Index           =   2
      Left            =   50
      TabIndex        =   57
      Top             =   3000
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   2
      Left            =   480
      TabIndex        =   56
      Top             =   3000
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   2
      Left            =   2040
      TabIndex        =   55
      Top             =   3000
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   2
      Left            =   3600
      TabIndex        =   54
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   2
      Left            =   4800
      TabIndex        =   53
      Top             =   3000
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   6360
      TabIndex        =   52
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   7320
      TabIndex        =   51
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   8320
      TabIndex        =   50
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   9400
      TabIndex        =   49
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   2
      Left            =   10500
      TabIndex        =   48
      Top             =   3000
      Width           =   315
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "02"
      Height          =   210
      Index           =   1
      Left            =   50
      TabIndex        =   47
      Top             =   2760
      Width           =   180
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   1
      Left            =   480
      TabIndex        =   46
      Top             =   2760
      Width           =   1005
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   1
      Left            =   2040
      TabIndex        =   45
      Top             =   2760
      Width           =   1260
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   1
      Left            =   3600
      TabIndex        =   44
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   1
      Left            =   4800
      TabIndex        =   43
      Top             =   2760
      Width           =   1125
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   6360
      TabIndex        =   42
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   7320
      TabIndex        =   41
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   8320
      TabIndex        =   40
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   9400
      TabIndex        =   39
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   1
      Left            =   10500
      TabIndex        =   38
      Top             =   2760
      Width           =   315
   End
   Begin VB.Label LB_Ema 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   10500
      TabIndex        =   37
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label LB_Epr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   9400
      TabIndex        =   36
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label LB_Eco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   8320
      TabIndex        =   35
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label LB_Qne 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   7320
      TabIndex        =   34
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label LB_Qun 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      Height          =   210
      Index           =   0
      Left            =   6360
      TabIndex        =   33
      Top             =   2520
      Width           =   315
   End
   Begin VB.Label LB_Mat 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A-217 Gr.CA15"
      Height          =   210
      Index           =   0
      Left            =   4800
      TabIndex        =   32
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Label LB_Bit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.1/2"" X !.1/4"""
      Height          =   210
      Index           =   0
      Left            =   3600
      TabIndex        =   31
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label LB_Des 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Junta Espirotálica"
      Height          =   210
      Index           =   0
      Left            =   2040
      TabIndex        =   30
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Label LB_Pec 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MP-CP-PA-SC"
      Height          =   210
      Index           =   0
      Left            =   480
      TabIndex        =   29
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Label LB_Ite 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "01"
      Height          =   210
      Index           =   0
      Left            =   50
      TabIndex        =   28
      Top             =   2520
      Width           =   180
   End
   Begin VB.Line LV 
      Index           =   14
      X1              =   4680
      X2              =   4680
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Line LV 
      Index           =   13
      X1              =   3480
      X2              =   3480
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Line LV 
      Index           =   12
      X1              =   6000
      X2              =   6000
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Line LV 
      Index           =   11
      X1              =   6960
      X2              =   6960
      Y1              =   2160
      Y2              =   2400
   End
   Begin VB.Line LV 
      Index           =   10
      X1              =   7920
      X2              =   7920
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Line LV 
      Index           =   9
      X1              =   9000
      X2              =   9000
      Y1              =   2160
      Y2              =   2400
   End
   Begin VB.Line LV 
      Index           =   8
      X1              =   10080
      X2              =   10080
      Y1              =   2160
      Y2              =   2400
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Saldo do Estoque"
      Height          =   210
      Index           =   28
      Left            =   8960
      TabIndex        =   27
      Top             =   1920
      Width           =   1260
   End
   Begin VB.Line LH 
      Index           =   6
      X1              =   6000
      X2              =   11160
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantidade"
      Height          =   210
      Index           =   27
      Left            =   6580
      TabIndex        =   26
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Matéria-Prima"
      Height          =   210
      Index           =   26
      Left            =   10200
      TabIndex        =   25
      Top             =   2190
      Width           =   975
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Produção"
      Height          =   210
      Index           =   25
      Left            =   9240
      TabIndex        =   24
      Top             =   2190
      Width           =   690
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Componente"
      Height          =   210
      Index           =   24
      Left            =   8040
      TabIndex        =   23
      Top             =   2190
      Width           =   900
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Necessária"
      Height          =   210
      Index           =   23
      Left            =   7060
      TabIndex        =   22
      Top             =   2190
      Width           =   825
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unitário"
      Height          =   210
      Index           =   22
      Left            =   6240
      TabIndex        =   21
      Top             =   2190
      Width           =   540
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Material"
      Height          =   210
      Index           =   20
      Left            =   5060
      TabIndex        =   20
      Top             =   2060
      Width           =   555
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bitola:"
      Height          =   210
      Index           =   19
      Left            =   3920
      TabIndex        =   19
      Top             =   2060
      Width           =   435
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição:"
      Height          =   210
      Index           =   18
      Left            =   2340
      TabIndex        =   18
      Top             =   2060
      Width           =   780
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Peça"
      Height          =   210
      Index           =   17
      Left            =   960
      TabIndex        =   17
      Top             =   2060
      Width           =   360
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item"
      Height          =   210
      Index           =   16
      Left            =   20
      TabIndex        =   16
      Top             =   2060
      Width           =   285
   End
   Begin VB.Line LV 
      Index           =   7
      X1              =   360
      X2              =   360
      Y1              =   1920
      Y2              =   2400
   End
   Begin VB.Line LH 
      Index           =   5
      X1              =   0
      X2              =   11160
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line LH 
      Index           =   4
      X1              =   0
      X2              =   11160
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line LV 
      Index           =   6
      X1              =   1920
      X2              =   1920
      Y1              =   1920
      Y2              =   2400
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
      Index           =   15
      Left            =   4800
      TabIndex        =   15
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   615
   End
   Begin VB.Label LB_Mate 
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
      Left            =   4800
      TabIndex        =   14
      Top             =   1320
      Width           =   510
   End
   Begin VB.Line LV 
      Index           =   5
      X1              =   4680
      X2              =   4680
      Y1              =   1080
      Y2              =   1680
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   13
      Left            =   6000
      TabIndex        =   13
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   735
   End
   Begin VB.Label LB_Desc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Válvula Gaveta 800PSI NPT A-105 1.1/2"""
      BeginProperty Font 
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
      TabIndex        =   12
      Top             =   1320
      Width           =   3540
   End
   Begin VB.Line LV 
      Index           =   4
      X1              =   5880
      X2              =   5880
      Y1              =   1080
      Y2              =   1680
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
      Index           =   11
      Left            =   3600
      TabIndex        =   11
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   435
   End
   Begin VB.Label LB_Bito 
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
      Left            =   3600
      TabIndex        =   10
      Top             =   1320
      Width           =   510
   End
   Begin VB.Line LV 
      Index           =   3
      X1              =   3480
      X2              =   3480
      Y1              =   1080
      Y2              =   1680
   End
   Begin VB.Label LB_Fixo 
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
      Index           =   9
      Left            =   2400
      TabIndex        =   9
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   480
   End
   Begin VB.Label LB_Figu 
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
      Left            =   2400
      TabIndex        =   8
      Top             =   1320
      Width           =   510
   End
   Begin VB.Line LV 
      Index           =   1
      X1              =   2280
      X2              =   2280
      Y1              =   1080
      Y2              =   1680
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
      Index           =   7
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   810
   End
   Begin VB.Label LB_Quan 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
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
      TabIndex        =   6
      Top             =   1320
      Width           =   315
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATA:"
      BeginProperty Font 
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
      TabIndex        =   5
      Top             =   1080
      UseMnemonic     =   0   'False
      Width           =   345
   End
   Begin VB.Line LH 
      Index           =   3
      X1              =   0
      X2              =   11160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line LH 
      Index           =   2
      X1              =   0
      X2              =   11160
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line LV 
      Index           =   0
      X1              =   1080
      X2              =   1080
      Y1              =   1080
      Y2              =   1680
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
      TabIndex        =   4
      Top             =   1320
      Width           =   960
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Matéria-Prima de Ítem de Estoque"
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
      Index           =   1
      Left            =   6840
      TabIndex        =   3
      Top             =   495
      Width           =   3585
   End
   Begin VB.Line LV 
      Index           =   2
      X1              =   6240
      X2              =   6240
      Y1              =   120
      Y2              =   840
   End
   Begin VB.Line LH 
      Index           =   1
      X1              =   0
      X2              =   11160
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conesteel Conexões de Aço Ltda."
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
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   4680
   End
   Begin VB.Image IMG 
      Height          =   495
      Left            =   120
      Picture         =   "Tela_Cotacao_MP_IT.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   520
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
      Caption         =   "Relatório de Configuração de"
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
      Index           =   3
      Left            =   7080
      TabIndex        =   1
      Top             =   180
      Width           =   3075
   End
   Begin VB.Label LB_Fixo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C.N.P.J. 55.783.427/0001-03"
      Height          =   210
      Index           =   5
      Left            =   2415
      TabIndex        =   0
      Top             =   555
      UseMnemonic     =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "Tela_Cotacao_MP_IT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

