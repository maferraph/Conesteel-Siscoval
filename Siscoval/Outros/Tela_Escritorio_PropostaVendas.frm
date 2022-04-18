VERSION 5.00
Begin VB.Form Tela_Escritorio_PropostaVendas 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proposta de Vendas"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "PV"
      Height          =   855
      Left            =   1440
      Picture         =   "Tela_Escritorio_PropostaVendas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ACC"
      Height          =   855
      Left            =   360
      Picture         =   "Tela_Escritorio_PropostaVendas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   5040
      Width           =   855
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   7080
      Picture         =   "Tela_Escritorio_PropostaVendas.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   5040
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10935
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   9
         Left            =   10560
         TabIndex        =   144
         Top             =   3480
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   9
         Left            =   10320
         TabIndex        =   143
         Top             =   3480
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   142
         Top             =   3480
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   9
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0CC6
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":0CD9
         TabIndex        =   141
         Top             =   3480
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   9
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0D0C
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":0D22
         TabIndex        =   140
         Top             =   3480
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   9
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0D45
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":0D58
         TabIndex        =   139
         Top             =   3480
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   9
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0D76
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":0D8F
         TabIndex        =   138
         Top             =   3480
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   9
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0DBD
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":0DD9
         TabIndex        =   137
         Top             =   3480
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   9
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0E2E
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":0E5C
         TabIndex        =   136
         Top             =   3480
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   9
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0EC1
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":0ECE
         TabIndex        =   135
         Top             =   3480
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   9
         Left            =   7440
         TabIndex        =   134
         Top             =   3480
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   9
         Left            =   8040
         TabIndex        =   133
         Top             =   3480
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   9
         Left            =   8880
         TabIndex        =   132
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   8
         Left            =   10560
         TabIndex        =   131
         Top             =   3120
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   8
         Left            =   10320
         TabIndex        =   130
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   129
         Top             =   3120
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   8
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0EEA
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":0EFD
         TabIndex        =   128
         Top             =   3120
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   8
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0F30
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":0F46
         TabIndex        =   127
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   8
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0F69
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":0F7C
         TabIndex        =   126
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   8
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0F9A
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":0FB3
         TabIndex        =   125
         Top             =   3120
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   8
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":0FE1
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":0FFD
         TabIndex        =   124
         Top             =   3120
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   8
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1052
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":1080
         TabIndex        =   123
         Top             =   3120
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   8
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":10E5
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":10F2
         TabIndex        =   122
         Top             =   3120
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   8
         Left            =   7440
         TabIndex        =   121
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   8
         Left            =   8040
         TabIndex        =   120
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   8
         Left            =   8880
         TabIndex        =   119
         Top             =   3120
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   7
         Left            =   10560
         TabIndex        =   118
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   7
         Left            =   10320
         TabIndex        =   117
         Top             =   2760
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   116
         Top             =   2760
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   7
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":110E
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":1121
         TabIndex        =   115
         Top             =   2760
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   7
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1154
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":116A
         TabIndex        =   114
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   7
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":118D
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":11A0
         TabIndex        =   113
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   7
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":11BE
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":11D7
         TabIndex        =   112
         Top             =   2760
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   7
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1205
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":1221
         TabIndex        =   111
         Top             =   2760
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   7
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1276
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":12A4
         TabIndex        =   110
         Top             =   2760
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   7
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1309
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1316
         TabIndex        =   109
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   7
         Left            =   7440
         TabIndex        =   108
         Top             =   2760
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   7
         Left            =   8040
         TabIndex        =   107
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   7
         Left            =   8880
         TabIndex        =   106
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   6
         Left            =   10560
         TabIndex        =   105
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   6
         Left            =   10320
         TabIndex        =   104
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   103
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   6
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1332
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":1345
         TabIndex        =   102
         Top             =   2400
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   6
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1378
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":138E
         TabIndex        =   101
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   6
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":13B1
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":13C4
         TabIndex        =   100
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   6
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":13E2
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":13FB
         TabIndex        =   99
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   6
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1429
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":1445
         TabIndex        =   98
         Top             =   2400
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   6
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":149A
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":14C8
         TabIndex        =   97
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   6
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":152D
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":153A
         TabIndex        =   96
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   6
         Left            =   7440
         TabIndex        =   95
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   6
         Left            =   8040
         TabIndex        =   94
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   6
         Left            =   8880
         TabIndex        =   93
         Top             =   2400
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   5
         Left            =   10560
         TabIndex        =   92
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   5
         Left            =   10320
         TabIndex        =   91
         Top             =   2040
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   5
         Left            =   120
         TabIndex        =   90
         Top             =   2040
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   5
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1556
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":1569
         TabIndex        =   89
         Top             =   2040
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   5
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":159C
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":15B2
         TabIndex        =   88
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   5
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":15D5
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":15E8
         TabIndex        =   87
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   5
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1606
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":161F
         TabIndex        =   86
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   5
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":164D
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":1669
         TabIndex        =   85
         Top             =   2040
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   5
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":16BE
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":16EC
         TabIndex        =   84
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   5
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1751
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":175E
         TabIndex        =   83
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   5
         Left            =   7440
         TabIndex        =   82
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   5
         Left            =   8040
         TabIndex        =   81
         Top             =   2040
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   5
         Left            =   8880
         TabIndex        =   80
         Top             =   2040
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   4
         Left            =   10560
         TabIndex        =   79
         Top             =   1680
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   4
         Left            =   10320
         TabIndex        =   78
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   77
         Top             =   1680
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   4
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":177A
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":178D
         TabIndex        =   76
         Top             =   1680
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   4
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":17C0
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":17D6
         TabIndex        =   75
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   4
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":17F9
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":180C
         TabIndex        =   74
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   4
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":182A
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1843
         TabIndex        =   73
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   4
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1871
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":188D
         TabIndex        =   72
         Top             =   1680
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   4
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":18E2
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":1910
         TabIndex        =   71
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   4
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1975
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1982
         TabIndex        =   70
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   4
         Left            =   7440
         TabIndex        =   69
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   4
         Left            =   8040
         TabIndex        =   68
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   4
         Left            =   8880
         TabIndex        =   67
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   3
         Left            =   10560
         TabIndex        =   66
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   3
         Left            =   10320
         TabIndex        =   65
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   64
         Top             =   1320
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   3
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":199E
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":19B1
         TabIndex        =   63
         Top             =   1320
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   3
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":19E4
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":19FA
         TabIndex        =   62
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   3
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1A1D
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":1A30
         TabIndex        =   61
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   3
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1A4E
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1A67
         TabIndex        =   60
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   3
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1A95
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":1AB1
         TabIndex        =   59
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   3
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1B06
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":1B34
         TabIndex        =   58
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   3
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1B99
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1BA6
         TabIndex        =   57
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   3
         Left            =   7440
         TabIndex        =   56
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   3
         Left            =   8040
         TabIndex        =   55
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   3
         Left            =   8880
         TabIndex        =   54
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   2
         Left            =   10560
         TabIndex        =   53
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   2
         Left            =   10320
         TabIndex        =   52
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   960
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   2
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1BC2
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":1BD5
         TabIndex        =   50
         Top             =   960
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   2
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1C08
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":1C1E
         TabIndex        =   49
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   2
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1C41
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":1C54
         TabIndex        =   48
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   2
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1C72
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1C8B
         TabIndex        =   47
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   2
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1CB9
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":1CD5
         TabIndex        =   46
         Top             =   960
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   2
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1D2A
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":1D58
         TabIndex        =   45
         Top             =   960
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   2
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1DBD
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1DCA
         TabIndex        =   44
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   2
         Left            =   7440
         TabIndex        =   43
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   2
         Left            =   8040
         TabIndex        =   42
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   2
         Left            =   8880
         TabIndex        =   41
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   1
         Left            =   10560
         TabIndex        =   40
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   1
         Left            =   10320
         TabIndex        =   39
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   615
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   1
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1DE6
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":1DF9
         TabIndex        =   37
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   1
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1E2C
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":1E42
         TabIndex        =   36
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   1
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1E65
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":1E78
         TabIndex        =   35
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   1
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1E96
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1EAF
         TabIndex        =   34
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   1
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1EDD
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":1EF9
         TabIndex        =   33
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   1
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1F4E
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":1F7C
         TabIndex        =   32
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   1
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":1FE1
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":1FEE
         TabIndex        =   31
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   1
         Left            =   7440
         TabIndex        =   30
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   1
         Left            =   8040
         TabIndex        =   29
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   1
         Left            =   8880
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton BT_C 
         Caption         =   "C"
         Height          =   285
         Index           =   0
         Left            =   10560
         TabIndex        =   25
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton BT_A 
         Caption         =   "A"
         Height          =   285
         Index           =   0
         Left            =   10320
         TabIndex        =   24
         Top             =   240
         Width           =   255
      End
      Begin VB.TextBox TXT_O 
         Height          =   285
         Index           =   0
         Left            =   8880
         TabIndex        =   23
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TXT_Z 
         Height          =   285
         Index           =   0
         Left            =   8040
         TabIndex        =   22
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TXT_P 
         Height          =   285
         Index           =   0
         Left            =   7440
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox CB_R 
         Height          =   315
         Index           =   0
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":200A
         Left            =   6480
         List            =   "Tela_Escritorio_PropostaVendas.frx":2017
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox CB_I 
         Height          =   315
         Index           =   0
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":2033
         Left            =   5520
         List            =   "Tela_Escritorio_PropostaVendas.frx":2061
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox CB_M 
         Height          =   315
         Index           =   0
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":20C6
         Left            =   4320
         List            =   "Tela_Escritorio_PropostaVendas.frx":20E2
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox CB_B 
         Height          =   315
         Index           =   0
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":2137
         Left            =   3480
         List            =   "Tela_Escritorio_PropostaVendas.frx":2150
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CB_E 
         Height          =   315
         Index           =   0
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":217E
         Left            =   2640
         List            =   "Tela_Escritorio_PropostaVendas.frx":2191
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CB_C 
         Height          =   315
         Index           =   0
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":21AF
         Left            =   1800
         List            =   "Tela_Escritorio_PropostaVendas.frx":21C5
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox CB_V 
         Height          =   315
         Index           =   0
         ItemData        =   "Tela_Escritorio_PropostaVendas.frx":21E8
         Left            =   720
         List            =   "Tela_Escritorio_PropostaVendas.frx":21FB
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TXT_Q 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Observações"
         Height          =   195
         Index           =   10
         Left            =   8880
         TabIndex        =   14
         Top             =   0
         Width           =   945
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Prazo"
         Height          =   195
         Index           =   9
         Left            =   8040
         TabIndex        =   13
         Top             =   0
         Width           =   405
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Preço"
         Height          =   195
         Index           =   8
         Left            =   7440
         TabIndex        =   12
         Top             =   0
         Width           =   420
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Revestimento"
         Height          =   195
         Index           =   7
         Left            =   6480
         TabIndex        =   11
         Top             =   0
         Width           =   975
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Internos"
         Height          =   195
         Index           =   6
         Left            =   5520
         TabIndex        =   10
         Top             =   0
         Width           =   570
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Material"
         Height          =   195
         Index           =   5
         Left            =   4320
         TabIndex        =   9
         Top             =   0
         Width           =   555
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Bitola"
         Height          =   195
         Index           =   4
         Left            =   3480
         TabIndex        =   8
         Top             =   0
         Width           =   390
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Extrem."
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   7
         Top             =   0
         Width           =   525
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Classe"
         Height          =   195
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   0
         Width           =   465
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Válvula"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   3
         Top             =   0
         Width           =   525
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Quant."
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   480
      End
   End
End
Attribute VB_Name = "Tela_Escritorio_PropostaVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Voltar_Click()
    Unload Me
End Sub

