VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_NotaFiscal 
   AutoRedraw      =   -1  'True
   Caption         =   "Assistente de elaboração da Nota Fiscal"
   ClientHeight    =   5160
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   8655
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   2  'Horizontal Line
   Icon            =   "Tela_NotaFiscal.frx":0000
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList LI 
      Left            =   2520
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483638
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_NotaFiscal.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_NotaFiscal.frx":075C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_NotaFiscal.frx":0A76
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog DIMP 
      Left            =   3720
      Top             =   4560
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Frame FR_Tela 
      BorderStyle     =   0  'None
      Height          =   5052
      Left            =   120
      TabIndex        =   96
      Top             =   120
      Width           =   8412
      Begin VB.CommandButton BT_Avancar 
         Caption         =   "Ava&nçar"
         Height          =   732
         Left            =   5520
         Picture         =   "Tela_NotaFiscal.frx":0EC8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Vai para próxima tela"
         Top             =   4200
         Width           =   732
      End
      Begin VB.CommandButton BT_Concluir 
         Caption         =   "C&oncluir"
         Enabled         =   0   'False
         Height          =   732
         Left            =   6600
         Picture         =   "Tela_NotaFiscal.frx":130A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimi nota fiscal"
         Top             =   4200
         Width           =   732
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   732
         Left            =   7680
         Picture         =   "Tela_NotaFiscal.frx":174C
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Cancela Assistente da nota fiscal"
         Top             =   4200
         Width           =   732
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Enabled         =   0   'False
         Height          =   732
         Left            =   4800
         Picture         =   "Tela_NotaFiscal.frx":1A56
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Volta para tela anterior"
         Top             =   4200
         Width           =   732
      End
      Begin VB.Frame FR_9 
         BackColor       =   &H8000000B&
         Caption         =   "Concluído"
         Height          =   4095
         Left            =   6840
         TabIndex        =   185
         ToolTipText     =   "Selecione uma das opções de operação"
         Top             =   2520
         Visible         =   0   'False
         Width           =   8412
         Begin VB.Frame FR_9_1 
            Caption         =   "Executando"
            Height          =   2532
            Left            =   240
            TabIndex        =   192
            Top             =   1320
            Width           =   7812
            Begin MSComctlLib.ProgressBar BP 
               Height          =   252
               Left            =   3000
               TabIndex        =   206
               Top             =   1920
               Width           =   1692
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               BorderStyle     =   1
               Appearance      =   1
               Scrolling       =   1
            End
            Begin VB.Label LB_12 
               AutoSize        =   -1  'True
               Caption         =   "12-) Finalizando banco de dados e variáveis"
               Height          =   192
               Left            =   4080
               TabIndex        =   204
               Top             =   1560
               Width           =   3180
            End
            Begin VB.Label LB_11 
               AutoSize        =   -1  'True
               Caption         =   "11-) Imprimindo Nota Fiscal"
               Height          =   192
               Left            =   4080
               TabIndex        =   203
               Top             =   1320
               Width           =   1944
            End
            Begin VB.Label LB_10 
               AutoSize        =   -1  'True
               Caption         =   "10-) Montando Nota Fiscal"
               Height          =   192
               Left            =   4080
               TabIndex        =   202
               Top             =   1080
               Width           =   1872
            End
            Begin VB.Label LB_9 
               AutoSize        =   -1  'True
               Caption         =   "9-) Imprimindo Certificado(s) de Qualidade"
               Height          =   192
               Left            =   4080
               TabIndex        =   201
               Top             =   840
               Width           =   3012
            End
            Begin VB.Image IMG_12 
               Height          =   252
               Left            =   3840
               Picture         =   "Tela_NotaFiscal.frx":1E98
               Stretch         =   -1  'True
               Top             =   1560
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_11 
               Height          =   252
               Left            =   3840
               Picture         =   "Tela_NotaFiscal.frx":21DA
               Stretch         =   -1  'True
               Top             =   1320
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_10 
               Height          =   252
               Left            =   3840
               Picture         =   "Tela_NotaFiscal.frx":251C
               Stretch         =   -1  'True
               Top             =   1080
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_9 
               Height          =   252
               Left            =   3840
               Picture         =   "Tela_NotaFiscal.frx":285E
               Stretch         =   -1  'True
               Top             =   840
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_8 
               Height          =   252
               Left            =   3840
               Picture         =   "Tela_NotaFiscal.frx":2BA0
               Stretch         =   -1  'True
               Top             =   600
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_7 
               Height          =   252
               Left            =   3840
               Picture         =   "Tela_NotaFiscal.frx":2EE2
               Stretch         =   -1  'True
               Top             =   360
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Label LB_7 
               AutoSize        =   -1  'True
               Caption         =   "7-) Adicionando nota fiscal mapa impostos"
               Height          =   192
               Left            =   4080
               TabIndex        =   200
               Top             =   360
               Width           =   3036
            End
            Begin VB.Label LB_8 
               AutoSize        =   -1  'True
               Caption         =   "8-) Montando Certificado(s) de Qualidade"
               Height          =   192
               Left            =   4080
               TabIndex        =   199
               Top             =   600
               Width           =   2940
            End
            Begin VB.Label LB_6 
               AutoSize        =   -1  'True
               Caption         =   "6-) Adicionando nota fiscal contas à receber"
               Height          =   192
               Left            =   360
               TabIndex        =   198
               Top             =   1560
               Width           =   3132
            End
            Begin VB.Label LB_5 
               AutoSize        =   -1  'True
               Caption         =   "5-) Baixando produtos no estoque"
               Height          =   192
               Left            =   360
               TabIndex        =   197
               Top             =   1320
               Width           =   2412
            End
            Begin VB.Label LB_4 
               AutoSize        =   -1  'True
               Caption         =   "4-) Lançando nota fiscal mapa faturamento"
               Height          =   192
               Left            =   360
               TabIndex        =   196
               Top             =   1080
               Width           =   3024
            End
            Begin VB.Label LB_3 
               AutoSize        =   -1  'True
               Caption         =   "3-) Baixando pedidos deste cliente"
               Height          =   192
               Left            =   360
               TabIndex        =   195
               Top             =   840
               Width           =   2484
            End
            Begin VB.Image IMG_6 
               Height          =   252
               Left            =   120
               Picture         =   "Tela_NotaFiscal.frx":3224
               Stretch         =   -1  'True
               Top             =   1560
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_5 
               Height          =   252
               Left            =   120
               Picture         =   "Tela_NotaFiscal.frx":3566
               Stretch         =   -1  'True
               Top             =   1320
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_4 
               Height          =   252
               Left            =   120
               Picture         =   "Tela_NotaFiscal.frx":38A8
               Stretch         =   -1  'True
               Top             =   1080
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_3 
               Height          =   252
               Left            =   120
               Picture         =   "Tela_NotaFiscal.frx":3BEA
               Stretch         =   -1  'True
               Top             =   840
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_2 
               Height          =   252
               Left            =   120
               Picture         =   "Tela_NotaFiscal.frx":3F2C
               Stretch         =   -1  'True
               Top             =   600
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Image IMG_1 
               Height          =   252
               Left            =   120
               Picture         =   "Tela_NotaFiscal.frx":426E
               Stretch         =   -1  'True
               Top             =   360
               Visible         =   0   'False
               Width           =   252
            End
            Begin VB.Label LB_2 
               AutoSize        =   -1  'True
               Caption         =   "2-) Salvando informações sobre a nota fiscal"
               Height          =   192
               Left            =   360
               TabIndex        =   194
               Top             =   600
               Width           =   3180
            End
            Begin VB.Label LB_1 
               AutoSize        =   -1  'True
               Caption         =   "1-) Conferindo numeração da nota fiscal"
               Height          =   192
               Left            =   360
               TabIndex        =   193
               Top             =   360
               Width           =   2832
            End
         End
         Begin VB.Label LB_TXT 
            AutoSize        =   -1  'True
            Caption         =   "dados do Sistema Siscoval, e então imprimi-lá."
            Height          =   192
            Index           =   4
            Left            =   240
            TabIndex        =   191
            Top             =   960
            Width           =   3360
         End
         Begin VB.Label LB_TXT 
            AutoSize        =   -1  'True
            Caption         =   "para o Assistente registrar esta nota fiscal no banco de"
            Height          =   192
            Index           =   3
            Left            =   4080
            TabIndex        =   190
            Top             =   720
            Width           =   3912
         End
         Begin VB.Label LB_TXT 
            AutoSize        =   -1  'True
            Caption         =   "CONCLUIR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   2
            Left            =   3000
            TabIndex        =   189
            Top             =   720
            Width           =   924
         End
         Begin VB.Label LB_TXT 
            AutoSize        =   -1  'True
            Caption         =   "Você deve agora pressionar o botão"
            Height          =   192
            Index           =   1
            Left            =   240
            TabIndex        =   188
            Top             =   720
            Width           =   2652
         End
         Begin VB.Label LB_TXT 
            AutoSize        =   -1  'True
            Caption         =   "Elaboração da Nota Fiscal concluída."
            Height          =   192
            Index           =   0
            Left            =   240
            TabIndex        =   187
            Top             =   360
            Width           =   2700
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Height          =   192
            Left            =   3960
            TabIndex        =   186
            Top             =   3840
            Width           =   36
         End
      End
      Begin VB.Frame FR_8 
         BackColor       =   &H8000000B&
         Caption         =   "Transportador / Volumes"
         Height          =   4095
         Left            =   6600
         TabIndex        =   163
         ToolTipText     =   "Selecione uma das opções de operação"
         Top             =   2280
         Visible         =   0   'False
         Width           =   8412
         Begin VB.Frame FR_8_1 
            Height          =   2172
            Left            =   120
            TabIndex        =   165
            Top             =   240
            Width           =   8172
            Begin VB.CheckBox CK_EditarTrans 
               Caption         =   "Editar dados da empresa"
               Height          =   192
               Left            =   5880
               TabIndex        =   80
               ToolTipText     =   "Se você deseja editar os dados da transportadora."
               Top             =   240
               Width           =   2172
            End
            Begin VB.ListBox LT_NomeTrans 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1020
               ItemData        =   "Tela_NotaFiscal.frx":45B0
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":45B2
               TabIndex        =   79
               ToolTipText     =   "Lista de empresas transportadoras."
               Top             =   1080
               Width           =   1812
            End
            Begin VB.ComboBox CB_TipoTrans 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":45B4
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":45B6
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   78
               ToolTipText     =   "Selecione o tipo de empresas para exibir."
               Top             =   480
               Width           =   1812
            End
            Begin VB.TextBox TXT_IETrans 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   3960
               MaxLength       =   20
               TabIndex        =   83
               ToolTipText     =   "Inscrição Estadual da transportadora."
               Top             =   1080
               Width           =   1692
            End
            Begin VB.ComboBox CB_EstTrans 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":45B8
               Left            =   7080
               List            =   "Tela_NotaFiscal.frx":45BA
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   86
               ToolTipText     =   "Estado da transportadora."
               Top             =   1680
               Width           =   972
            End
            Begin VB.TextBox TXT_CidTrans 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   5760
               MaxLength       =   20
               TabIndex        =   84
               ToolTipText     =   "Cidade da transportadora."
               Top             =   1080
               Width           =   2292
            End
            Begin VB.TextBox TXT_EndTrans 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   85
               ToolTipText     =   "Endereço da transportadora."
               Top             =   1680
               Width           =   4812
            End
            Begin VB.TextBox TXT_Trans 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   81
               ToolTipText     =   "Nome do transportador."
               Top             =   480
               Width           =   5892
            End
            Begin MSMask.MaskEdBox TXT_CGCTrans 
               Height          =   288
               Left            =   2160
               TabIndex        =   82
               ToolTipText     =   "C.N.P.J da transportadora (antigo C.G.C.)"
               Top             =   1080
               Width           =   1692
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               ForeColor       =   -2147483630
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   18
               Mask            =   "99.999.999/9999-99"
               PromptChar      =   "_"
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               Caption         =   "Tipo de Empresa:"
               Height          =   192
               Left            =   120
               TabIndex        =   173
               Top             =   240
               Width           =   1296
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               Caption         =   "Nome Fantasia:"
               Height          =   192
               Left            =   120
               TabIndex        =   172
               Top             =   840
               Width           =   1140
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   5760
               TabIndex        =   171
               Top             =   840
               Width           =   564
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "C.N.P.J.:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   2160
               TabIndex        =   170
               Top             =   840
               Width           =   600
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               Caption         =   "Inscrição Estadual:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   3960
               TabIndex        =   169
               Top             =   840
               Width           =   1356
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Estado:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   7080
               TabIndex        =   168
               Top             =   1440
               Width           =   552
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   2160
               TabIndex        =   167
               Top             =   1440
               Width           =   744
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               Caption         =   "Empresa:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   2160
               TabIndex        =   166
               Top             =   240
               Width           =   696
            End
         End
         Begin VB.Frame FR_8_2 
            Height          =   1572
            Left            =   120
            TabIndex        =   174
            Top             =   2400
            Width           =   2172
            Begin VB.ComboBox CB_Placa 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":45BC
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":45BE
               Sorted          =   -1  'True
               TabIndex        =   88
               ToolTipText     =   "Placa do veículo."
               Top             =   1080
               Width           =   1092
            End
            Begin VB.ComboBox CB_EstVei 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":45C0
               Left            =   1320
               List            =   "Tela_NotaFiscal.frx":45C2
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   89
               ToolTipText     =   "Estado do veículo."
               Top             =   1080
               Width           =   732
            End
            Begin VB.ComboBox CB_Frete 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":45C4
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":45C6
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   87
               ToolTipText     =   "Selecione quem irá pagar o frete."
               Top             =   480
               Width           =   1932
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "Estado:"
               Height          =   192
               Left            =   1320
               TabIndex        =   177
               Top             =   840
               Width           =   552
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               Caption         =   "Frete por conta de:"
               Height          =   192
               Left            =   120
               TabIndex        =   176
               Top             =   240
               Width           =   1344
            End
            Begin VB.Label Label37 
               AutoSize        =   -1  'True
               Caption         =   "Placa:"
               Height          =   192
               Left            =   120
               TabIndex        =   175
               Top             =   840
               Width           =   456
            End
         End
         Begin VB.Frame Frame5 
            Height          =   1572
            Left            =   2400
            TabIndex        =   178
            Top             =   2400
            Width           =   5892
            Begin VB.TextBox TXT_NumVol 
               Height          =   288
               Left            =   120
               MaxLength       =   15
               TabIndex        =   93
               ToolTipText     =   "Número dos volumes."
               Top             =   1200
               Width           =   1812
            End
            Begin VB.ComboBox CB_Marca 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":45C8
               Left            =   3840
               List            =   "Tela_NotaFiscal.frx":45CA
               Sorted          =   -1  'True
               TabIndex        =   92
               ToolTipText     =   "Marca dos produtos."
               Top             =   480
               Width           =   1932
            End
            Begin VB.TextBox TXT_QuantVol 
               Height          =   288
               Left            =   2280
               TabIndex        =   91
               ToolTipText     =   "Quantidade."
               Top             =   480
               Width           =   1452
            End
            Begin VB.ComboBox CB_Especie 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":45CC
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":45CE
               Sorted          =   -1  'True
               TabIndex        =   90
               ToolTipText     =   "Espécie de embalagens."
               Top             =   480
               Width           =   2052
            End
            Begin MSMask.MaskEdBox TXT_PesoBruto 
               Height          =   288
               Left            =   2160
               TabIndex        =   94
               ToolTipText     =   "Peso bruto dos produtos."
               Top             =   1200
               Width           =   1692
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_PesoLiquido 
               Height          =   288
               Left            =   4080
               TabIndex        =   95
               ToolTipText     =   "Peso líquido dos produtos."
               Top             =   1200
               Width           =   1692
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               Format          =   "#,##0.00;(#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               Caption         =   "Peso Líquido:"
               Height          =   192
               Left            =   4080
               TabIndex        =   184
               Top             =   960
               Width           =   984
            End
            Begin VB.Label Label36 
               AutoSize        =   -1  'True
               Caption         =   "Peso Bruto:"
               Height          =   192
               Left            =   2160
               TabIndex        =   183
               Top             =   960
               Width           =   828
            End
            Begin VB.Label Label26 
               AutoSize        =   -1  'True
               Caption         =   "Número Volumes:"
               Height          =   192
               Left            =   120
               TabIndex        =   182
               Top             =   960
               Width           =   1284
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               Caption         =   "Marca:"
               Height          =   192
               Left            =   3840
               TabIndex        =   181
               Top             =   240
               Width           =   492
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               Caption         =   "Quantidade:"
               Height          =   192
               Left            =   2280
               TabIndex        =   180
               Top             =   240
               Width           =   876
            End
            Begin VB.Label Label27 
               AutoSize        =   -1  'True
               Caption         =   "Espécie:"
               Height          =   192
               Left            =   120
               TabIndex        =   179
               Top             =   240
               Width           =   636
            End
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Height          =   192
            Left            =   3960
            TabIndex        =   164
            Top             =   3840
            Width           =   36
         End
      End
      Begin VB.Frame FR_7 
         BackColor       =   &H8000000B&
         Caption         =   "Certificados"
         Height          =   4095
         Left            =   6360
         TabIndex        =   160
         ToolTipText     =   "Certificados de qualidade."
         Top             =   2040
         Visible         =   0   'False
         Width           =   8412
         Begin VB.Frame Frame4 
            Height          =   852
            Left            =   240
            TabIndex        =   162
            Top             =   240
            Width           =   7932
            Begin VB.OptionButton RB_CF1 
               Caption         =   "Continuar sem elaboração do certificado de qualidade"
               Height          =   252
               Left            =   960
               TabIndex        =   76
               ToolTipText     =   "Continua com o assistente de nota fiscal sem elaborar certificado(s) de qualidade."
               Top             =   240
               Width           =   4332
            End
            Begin VB.OptionButton RB_CF2 
               Caption         =   "Elaborar certificado de qualidade dos produtos desta nota fiscal"
               Height          =   192
               Left            =   960
               TabIndex        =   77
               ToolTipText     =   "Exibe peças desta nota fiscal com seus respectivos números de corrida."
               Top             =   480
               Width           =   6252
            End
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            Height          =   192
            Left            =   3960
            TabIndex        =   161
            Top             =   3840
            Width           =   36
         End
      End
      Begin VB.Frame FR_6 
         BackColor       =   &H8000000B&
         Caption         =   "Informações complementares"
         Height          =   4095
         Left            =   6120
         TabIndex        =   136
         ToolTipText     =   "Informações complementares da nota fiscal."
         Top             =   1800
         Visible         =   0   'False
         Width           =   8412
         Begin VB.CheckBox CK_EditarValores 
            Caption         =   "Editar os valores de cálculo do imposto"
            Height          =   192
            Left            =   240
            TabIndex        =   58
            ToolTipText     =   "Se você deseja editar os valores da nota fiscal."
            Top             =   360
            Width           =   3372
         End
         Begin VB.Frame FR_5_1 
            Caption         =   "Cálculo do Imposto:"
            Height          =   1812
            Left            =   240
            TabIndex        =   137
            ToolTipText     =   "Valores da nota fiscal."
            Top             =   600
            Width           =   7932
            Begin MSMask.MaskEdBox TXT_BaseICMS 
               Height          =   288
               Left            =   120
               TabIndex        =   59
               ToolTipText     =   "Base de cálculo do I.C.M.S."
               Top             =   600
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_ValorICMS 
               Height          =   288
               Left            =   1680
               TabIndex        =   60
               ToolTipText     =   "Valor do I.C.M.S."
               Top             =   600
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_BaseICMSSub 
               Height          =   288
               Left            =   3240
               TabIndex        =   61
               ToolTipText     =   "Base de cálculo do I.C.M.S. substituição."
               Top             =   600
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_ValorICMSSub 
               Height          =   288
               Left            =   4800
               TabIndex        =   62
               ToolTipText     =   "Indica o valor do I.C.M.S substituição."
               Top             =   600
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_ValorTotalProdutos 
               Height          =   288
               Left            =   6360
               TabIndex        =   63
               ToolTipText     =   "Valor total dos produtos."
               Top             =   600
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_ValorFrete 
               Height          =   288
               Left            =   120
               TabIndex        =   64
               ToolTipText     =   "Valor do frete."
               Top             =   1320
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_ValorSeguro 
               Height          =   288
               Left            =   1680
               TabIndex        =   65
               ToolTipText     =   "Valor do seguro."
               Top             =   1320
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_Outras 
               Height          =   288
               Left            =   3240
               TabIndex        =   66
               ToolTipText     =   "Outras despesas acessórias."
               Top             =   1320
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_ValorTotalIPI 
               Height          =   288
               Left            =   4800
               TabIndex        =   67
               ToolTipText     =   "Valor total do I.P.I."
               Top             =   1320
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_ValorTotalNotaFiscal 
               Height          =   288
               Left            =   6360
               TabIndex        =   68
               ToolTipText     =   "Valor total da nota fiscal."
               Top             =   1320
               Width           =   1452
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   20
               Format          =   "$ #,##0.00;- $ #,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label LB_BaseICMS 
               AutoSize        =   -1  'True
               Caption         =   "Base Cálc. I.C.M.S.:"
               Height          =   192
               Left            =   120
               TabIndex        =   147
               Top             =   360
               Width           =   1380
            End
            Begin VB.Label LB_ValorICMS 
               AutoSize        =   -1  'True
               Caption         =   "Valor do I.C.M.S.:"
               Height          =   192
               Left            =   1680
               TabIndex        =   146
               Top             =   360
               Width           =   1212
            End
            Begin VB.Label LB_BaseICMSSub 
               AutoSize        =   -1  'True
               Caption         =   "B.Cálc. ICMS Subst.:"
               Height          =   192
               Left            =   3240
               TabIndex        =   145
               Top             =   360
               Width           =   1440
            End
            Begin VB.Label LB_ValorICMSSub 
               AutoSize        =   -1  'True
               Caption         =   "Valor I.C.M.S. Subst.:"
               Height          =   192
               Left            =   4800
               TabIndex        =   144
               Top             =   360
               Width           =   1464
            End
            Begin VB.Label LB_ValorTotalProdutos 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total Produtos:"
               Height          =   192
               Left            =   6360
               TabIndex        =   143
               Top             =   360
               Width           =   1512
            End
            Begin VB.Label LB_ValorFrete 
               AutoSize        =   -1  'True
               Caption         =   "Valor do Frete:"
               Height          =   192
               Left            =   120
               TabIndex        =   142
               Top             =   1080
               Width           =   1056
            End
            Begin VB.Label LB_ValorSeguro 
               AutoSize        =   -1  'True
               Caption         =   "Valor do Seguro:"
               Height          =   192
               Left            =   1680
               TabIndex        =   141
               Top             =   1080
               Width           =   1212
            End
            Begin VB.Label LB_Outras 
               AutoSize        =   -1  'True
               Caption         =   "Outras Desp.Acess.:"
               Height          =   192
               Left            =   3240
               TabIndex        =   140
               Top             =   1080
               Width           =   1464
            End
            Begin VB.Label LB_ValorTotalIPI 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total I.P.I.:"
               Height          =   192
               Left            =   4800
               TabIndex        =   139
               Top             =   1080
               Width           =   1152
            End
            Begin VB.Label LB_ValorTotalNota 
               AutoSize        =   -1  'True
               Caption         =   "Valor Total Nota:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Left            =   6360
               TabIndex        =   138
               Top             =   1080
               Width           =   1428
            End
         End
         Begin VB.Frame FR_5_2 
            Caption         =   "Dados Adicionais:"
            Height          =   1092
            Left            =   240
            TabIndex        =   149
            ToolTipText     =   "Dados adicionais da nota fiscal."
            Top             =   2640
            Width           =   7932
            Begin VB.TextBox TXT_Setor 
               Height          =   288
               Left            =   7080
               MaxLength       =   10
               TabIndex        =   75
               ToolTipText     =   "Setor."
               Top             =   600
               Width           =   732
            End
            Begin VB.TextBox TXT_VendExterno 
               Height          =   288
               Left            =   6000
               MaxLength       =   10
               TabIndex        =   74
               ToolTipText     =   "Nome e/ou código do vendedor externo"
               Top             =   600
               Width           =   972
            End
            Begin VB.TextBox TXT_VendInterno 
               Height          =   288
               Left            =   4920
               MaxLength       =   10
               TabIndex        =   73
               ToolTipText     =   "Nome e/ou código do vendedor interno."
               Top             =   600
               Width           =   972
            End
            Begin VB.TextBox TXT_Operacao 
               Height          =   288
               Left            =   3720
               MaxLength       =   15
               TabIndex        =   72
               ToolTipText     =   "Operação."
               Top             =   600
               Width           =   1092
            End
            Begin VB.TextBox TXT_SeuPedido 
               Height          =   288
               Left            =   2520
               MaxLength       =   10
               TabIndex        =   71
               ToolTipText     =   "Número de pedido do cliente."
               Top             =   600
               Width           =   1092
            End
            Begin VB.TextBox TXT_PedidoInterno 
               Height          =   288
               Left            =   1320
               MaxLength       =   10
               TabIndex        =   70
               ToolTipText     =   "Pedido interno da Conesteel."
               Top             =   600
               Width           =   1092
            End
            Begin VB.TextBox TXT_HoraSaida 
               Height          =   288
               Left            =   120
               MaxLength       =   8
               TabIndex        =   69
               ToolTipText     =   "Hora de saída da nota fiscal."
               Top             =   600
               Width           =   1092
            End
            Begin VB.Label LB_Setor 
               AutoSize        =   -1  'True
               Caption         =   "Setor:"
               Height          =   192
               Left            =   7080
               TabIndex        =   156
               Top             =   360
               Width           =   420
            End
            Begin VB.Label LB_VendExterno 
               AutoSize        =   -1  'True
               Caption         =   "Vend.Externo:"
               Height          =   192
               Left            =   6000
               TabIndex        =   155
               Top             =   360
               Width           =   996
            End
            Begin VB.Label LB_VendInterno 
               AutoSize        =   -1  'True
               Caption         =   "Vend.Interno:"
               Height          =   192
               Left            =   4920
               TabIndex        =   154
               Top             =   360
               Width           =   936
            End
            Begin VB.Label LB_Operacao 
               AutoSize        =   -1  'True
               Caption         =   "Operação:"
               Height          =   192
               Left            =   3720
               TabIndex        =   153
               Top             =   360
               Width           =   768
            End
            Begin VB.Label LB_SeuPedido 
               AutoSize        =   -1  'True
               Caption         =   "Seu Pedido:"
               Height          =   192
               Left            =   2520
               TabIndex        =   152
               Top             =   360
               Width           =   888
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Pedido Interno:"
               Height          =   192
               Left            =   1320
               TabIndex        =   151
               Top             =   360
               Width           =   1080
            End
            Begin VB.Label LB_HoraSaida 
               AutoSize        =   -1  'True
               Caption         =   "Hora da saída:"
               Height          =   192
               Left            =   120
               TabIndex        =   150
               Top             =   360
               Width           =   1068
            End
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Height          =   192
            Left            =   3960
            TabIndex        =   148
            Top             =   3840
            Width           =   36
         End
      End
      Begin VB.Frame FR_5 
         Caption         =   "Produtos"
         Height          =   4095
         Left            =   6000
         TabIndex        =   97
         ToolTipText     =   "Produtos da nota fiscal."
         Top             =   1560
         Visible         =   0   'False
         Width           =   8412
         Begin VB.CommandButton BT_Comentarios 
            Caption         =   "Incluir comentários"
            Height          =   852
            Left            =   7320
            Picture         =   "Tela_NotaFiscal.frx":45D0
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Abre assistente de comentários da nota fiscal."
            Top             =   240
            Width           =   972
         End
         Begin VB.CommandButton BT_DepositoBancario 
            Caption         =   "Incluir C/C Banco"
            Height          =   852
            Left            =   6360
            Picture         =   "Tela_NotaFiscal.frx":48DA
            Style           =   1  'Graphical
            TabIndex        =   56
            ToolTipText     =   "Abre assistente de bancos para incluir na nota fiscal."
            Top             =   240
            Width           =   972
         End
         Begin VB.CommandButton BT_DeclaracoesFiscais 
            Caption         =   "Incluir Dec.Fiscais"
            Height          =   852
            Left            =   5400
            Picture         =   "Tela_NotaFiscal.frx":4BE4
            Style           =   1  'Graphical
            TabIndex        =   55
            ToolTipText     =   "Inclui declarações fiscais dos produtos da nota fiscal se houver."
            Top             =   240
            Width           =   972
         End
         Begin VB.CommandButton BT_EditarItem 
            Caption         =   "&Editar os ítens da NF"
            Height          =   852
            Left            =   2040
            Picture         =   "Tela_NotaFiscal.frx":4EEE
            Style           =   1  'Graphical
            TabIndex        =   52
            ToolTipText     =   "Abre assistente para editar os produtos da nota fiscal."
            Top             =   240
            Width           =   972
         End
         Begin VB.CommandButton BT_ApagarTudo 
            Caption         =   "Apagar &tudo"
            Height          =   852
            Left            =   4200
            Picture         =   "Tela_NotaFiscal.frx":5330
            Style           =   1  'Graphical
            TabIndex        =   54
            ToolTipText     =   "Apaga tudo do quadro produtos."
            Top             =   240
            Width           =   972
         End
         Begin VB.CommandButton BT_ApagarItem 
            Caption         =   "Apa&gar ítens"
            Height          =   852
            Left            =   3240
            Picture         =   "Tela_NotaFiscal.frx":5772
            Style           =   1  'Graphical
            TabIndex        =   53
            ToolTipText     =   "Apaga um produto da nota fiscal."
            Top             =   240
            Width           =   972
         End
         Begin VB.CommandButton BT_InserirManual 
            Caption         =   "Inserir &manual"
            Height          =   852
            Left            =   1080
            Picture         =   "Tela_NotaFiscal.frx":5BB4
            Style           =   1  'Graphical
            TabIndex        =   51
            ToolTipText     =   "Abre assistente para inserir produtos na nota fiscal manualmente."
            Top             =   240
            Width           =   972
         End
         Begin VB.CommandButton BT_InserirItem 
            Caption         =   "Ass&istente de ítens"
            Height          =   852
            Left            =   120
            Picture         =   "Tela_NotaFiscal.frx":5EBE
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Abre o Assistente de Ítens de estoque para inserir produtos na nota fiscal."
            Top             =   240
            Width           =   972
         End
         Begin MSFlexGridLib.MSFlexGrid FG_2 
            Height          =   1692
            Left            =   5040
            TabIndex        =   134
            TabStop         =   0   'False
            ToolTipText     =   "Dados que serão impressos na nota fiscal."
            Top             =   1200
            Width           =   3012
            _ExtentX        =   5318
            _ExtentY        =   2990
            _Version        =   393216
            Rows            =   21
            Cols            =   6
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            TextStyleFixed  =   4
            FocusRect       =   0
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
         End
         Begin MSFlexGridLib.MSFlexGrid FG_1 
            Height          =   2772
            Left            =   120
            TabIndex        =   135
            TabStop         =   0   'False
            ToolTipText     =   "Produtos que serão impressos na nota fiscal."
            Top             =   1200
            Width           =   8172
            _ExtentX        =   14420
            _ExtentY        =   4895
            _Version        =   393216
            Rows            =   21
            Cols            =   12
            AllowBigSelection=   0   'False
            Enabled         =   -1  'True
            TextStyleFixed  =   4
            FocusRect       =   0
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
         End
      End
      Begin VB.Frame FR_4 
         BackColor       =   &H8000000B&
         Caption         =   "Informações gerais"
         Height          =   4095
         Left            =   5760
         TabIndex        =   98
         ToolTipText     =   "Informações gerais da nota fiscal."
         Top             =   1320
         Visible         =   0   'False
         Width           =   8412
         Begin VB.TextBox TXT_NF 
            BackColor       =   &H8000000A&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   612
            Left            =   1080
            MaxLength       =   6
            TabIndex        =   29
            ToolTipText     =   "Confirme o número da nota fiscal."
            Top             =   720
            Width           =   1692
         End
         Begin VB.Frame FR_3_1 
            Caption         =   "Data:"
            Height          =   1452
            Left            =   3960
            TabIndex        =   108
            Top             =   240
            Width           =   4092
            Begin VB.CheckBox CK_DataSaida 
               Caption         =   "Imprimir data de saída"
               Height          =   252
               Left            =   2040
               TabIndex        =   32
               ToolTipText     =   "Se você deseja imprimir a data de saída, ative este botão."
               Top             =   1080
               Width           =   1932
            End
            Begin MSMask.MaskEdBox TXT_DataEmissao 
               Height          =   492
               Left            =   240
               TabIndex        =   30
               ToolTipText     =   "Data de emissão da nota fiscal."
               Top             =   480
               Width           =   1812
               _ExtentX        =   3201
               _ExtentY        =   873
               _Version        =   393216
               BackColor       =   -2147483633
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "&&&&&&&&&&"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_DataSaida 
               Height          =   492
               Left            =   2160
               TabIndex        =   31
               ToolTipText     =   "Data de saída da nota fiscal."
               Top             =   480
               Width           =   1812
               _ExtentX        =   3201
               _ExtentY        =   873
               _Version        =   393216
               BackColor       =   -2147483633
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "&&&&&&&&&&"
               PromptChar      =   "_"
            End
            Begin VB.Label LB_IMP 
               AutoSize        =   -1  'True
               Height          =   195
               Left            =   480
               TabIndex        =   218
               Top             =   1080
               Visible         =   0   'False
               Width           =   45
            End
            Begin VB.Label LB_DataSaida 
               AutoSize        =   -1  'True
               Caption         =   "Data de Saída:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   2160
               TabIndex        =   110
               Top             =   240
               Width           =   1080
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               Caption         =   "Data de Emissão:"
               Height          =   192
               Left            =   240
               TabIndex        =   109
               Top             =   240
               Width           =   1284
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Data de Vencimento:"
            Height          =   1932
            Left            =   240
            TabIndex        =   99
            Top             =   1920
            Width           =   7932
            Begin VB.CheckBox CK_Desdobrar 
               Caption         =   "Desdobrar duplicatas"
               Height          =   252
               Left            =   120
               TabIndex        =   43
               ToolTipText     =   "Ative este botão se você deseja desdobrar a duplicata."
               Top             =   1440
               Width           =   1932
            End
            Begin VB.OptionButton RD_Escdd 
               Height          =   252
               Left            =   6840
               TabIndex        =   40
               ToolTipText     =   "Selecione aqui para vencimento à partir do número de dias digitadas na caixa de texto ao lado."
               Top             =   240
               Width           =   252
            End
            Begin VB.TextBox TXT_Escdd 
               Enabled         =   0   'False
               Height          =   288
               Left            =   7080
               MaxLength       =   3
               TabIndex        =   41
               ToolTipText     =   "Digite quantos dias para vencimento"
               Top             =   240
               Width           =   372
            End
            Begin VB.TextBox TXT_DVB 
               Enabled         =   0   'False
               Height          =   288
               Left            =   2040
               MaxLength       =   3
               TabIndex        =   45
               ToolTipText     =   "Digite quantos dias para vencimento"
               Top             =   1440
               Width           =   372
            End
            Begin VB.TextBox TXT_DVC 
               Enabled         =   0   'False
               Height          =   288
               Left            =   3960
               MaxLength       =   3
               TabIndex        =   47
               ToolTipText     =   "Digite quantos dias para vencimento"
               Top             =   1440
               Width           =   372
            End
            Begin VB.TextBox TXT_DVD 
               Enabled         =   0   'False
               Height          =   288
               Left            =   5880
               MaxLength       =   3
               TabIndex        =   49
               ToolTipText     =   "Digite quantos dias para vencimento"
               Top             =   1440
               Width           =   372
            End
            Begin VB.OptionButton RD_CApres 
               Caption         =   "C/Apres."
               Height          =   252
               Left            =   960
               TabIndex        =   33
               ToolTipText     =   "Selecione aqui para vencimento Contra Apresentação."
               Top             =   240
               Width           =   972
            End
            Begin MSMask.MaskEdBox TXT_DataVenc_A 
               Height          =   492
               Left            =   120
               TabIndex        =   42
               ToolTipText     =   "Confirme o vencimento."
               Top             =   840
               Width           =   1812
               _ExtentX        =   3201
               _ExtentY        =   873
               _Version        =   393216
               BackColor       =   -2147483633
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "&&&&&&&&&&"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_DataVenc_B 
               Height          =   492
               Left            =   2040
               TabIndex        =   44
               ToolTipText     =   "Segundo vencimento."
               Top             =   840
               Width           =   1812
               _ExtentX        =   3201
               _ExtentY        =   873
               _Version        =   393216
               BackColor       =   -2147483633
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "&&&&&&&&&&"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_DataVenc_C 
               Height          =   492
               Left            =   3960
               TabIndex        =   46
               ToolTipText     =   "Terceiro vencimento."
               Top             =   840
               Width           =   1812
               _ExtentX        =   3201
               _ExtentY        =   873
               _Version        =   393216
               BackColor       =   -2147483633
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "&&&&&&&&&&"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_DataVenc_D 
               Height          =   492
               Left            =   5880
               TabIndex        =   48
               ToolTipText     =   "Quarto vencimento."
               Top             =   840
               Width           =   1812
               _ExtentX        =   3201
               _ExtentY        =   873
               _Version        =   393216
               BackColor       =   -2147483633
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   "&&&&&&&&&&"
               PromptChar      =   "_"
            End
            Begin VB.OptionButton RB_SV 
               Caption         =   "S/Venc."
               Height          =   252
               Left            =   120
               TabIndex        =   205
               ToolTipText     =   "Selecione aqui para vencimento Contra Apresentação."
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton RD_45dd 
               Caption         =   "45 d.d."
               Height          =   252
               Left            =   6000
               TabIndex        =   39
               ToolTipText     =   "Selecione aqui para vencimento em 45dd."
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton RD_35dd 
               Caption         =   "35 d.d."
               Height          =   252
               Left            =   5160
               TabIndex        =   38
               ToolTipText     =   "Selecione aqui para vencimento em 35dd."
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton RD_30dd 
               Caption         =   "30 d.d."
               Height          =   252
               Left            =   4320
               TabIndex        =   37
               ToolTipText     =   "Selecione aqui para vencimento em 30dd."
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton RD_28dd 
               Caption         =   "28 d.d."
               Height          =   252
               Left            =   3480
               TabIndex        =   36
               ToolTipText     =   "Selecione aqui para vencimento em 28dd."
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton RD_21dd 
               Caption         =   "21 d.d."
               Height          =   252
               Left            =   2680
               TabIndex        =   35
               ToolTipText     =   "Selecione aqui para vencimento em 21dd"
               Top             =   240
               Width           =   975
            End
            Begin VB.OptionButton RD_AVista 
               Caption         =   "À Vista"
               Height          =   252
               Left            =   1900
               TabIndex        =   34
               ToolTipText     =   "Selecione aqui para vencimento À Vista."
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label21 
               AutoSize        =   -1  'True
               Caption         =   "Data de vencimento /A:"
               Height          =   192
               Left            =   120
               TabIndex        =   107
               Top             =   600
               Width           =   1692
            End
            Begin VB.Label LB_DataVenc_B 
               AutoSize        =   -1  'True
               Caption         =   "Data de vencimento /B:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   2040
               TabIndex        =   106
               Top             =   600
               Width           =   1692
            End
            Begin VB.Label LB_DataVenc_C 
               AutoSize        =   -1  'True
               Caption         =   "Data de vencimento /C:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   3960
               TabIndex        =   105
               Top             =   600
               Width           =   1692
            End
            Begin VB.Label LB_DataVenc_D 
               AutoSize        =   -1  'True
               Caption         =   "Data de vencimento /D:"
               Enabled         =   0   'False
               Height          =   192
               Left            =   5880
               TabIndex        =   104
               Top             =   600
               Width           =   1704
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "d.d."
               Height          =   192
               Left            =   7560
               TabIndex        =   103
               Top             =   240
               Width           =   264
            End
            Begin VB.Label LB_DVB 
               AutoSize        =   -1  'True
               Caption         =   "d.d."
               Enabled         =   0   'False
               Height          =   192
               Left            =   2520
               TabIndex        =   102
               Top             =   1440
               Width           =   264
            End
            Begin VB.Label LB_DVC 
               AutoSize        =   -1  'True
               Caption         =   "d.d."
               Enabled         =   0   'False
               Height          =   192
               Left            =   4440
               TabIndex        =   101
               Top             =   1440
               Width           =   264
            End
            Begin VB.Label LB_DVD 
               AutoSize        =   -1  'True
               Caption         =   "d.d."
               Enabled         =   0   'False
               Height          =   192
               Left            =   6360
               TabIndex        =   100
               Top             =   1440
               Width           =   264
            End
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Número da Nota Fiscal:"
            Height          =   192
            Left            =   1080
            TabIndex        =   111
            Top             =   480
            Width           =   1692
         End
      End
      Begin VB.Frame FR_3 
         Caption         =   "Pedidos"
         Height          =   4092
         Left            =   480
         TabIndex        =   157
         ToolTipText     =   "Pedidos da empresa."
         Top             =   0
         Visible         =   0   'False
         Width           =   8412
         Begin VB.Frame FR_4_2 
            Caption         =   "Pedidos pendentes deste cliente:"
            Height          =   2652
            Left            =   240
            TabIndex        =   159
            ToolTipText     =   "Pedidos pendentes deste cliente."
            Top             =   1200
            Width           =   7932
            Begin VB.CommandButton BT_LL 
               Caption         =   "Limpa Lista"
               Height          =   495
               Left            =   3600
               TabIndex        =   26
               ToolTipText     =   "Limpa a limpa de ítens importados"
               Top             =   1920
               Width           =   735
            End
            Begin VB.CommandButton BT_RI 
               Caption         =   "Remove Ítem"
               Height          =   495
               Left            =   3600
               TabIndex        =   25
               ToolTipText     =   "Remove um ítem da lista de ítens importados"
               Top             =   1440
               Width           =   735
            End
            Begin VB.CommandButton BT_II 
               Caption         =   "Importa Ítem"
               Height          =   495
               Left            =   3600
               TabIndex        =   24
               ToolTipText     =   "Importa um ítem da lista de ítens importados"
               Top             =   960
               Width           =   735
            End
            Begin VB.CommandButton BT_IT 
               Caption         =   "Importa Tudo"
               Height          =   495
               Left            =   3600
               TabIndex        =   23
               ToolTipText     =   "Importa todos ítens do pedido"
               Top             =   480
               Width           =   735
            End
            Begin VB.Frame FR_4_2_2 
               Caption         =   "Lista de ítens importados para a nota:"
               Height          =   2295
               Left            =   4440
               TabIndex        =   215
               Top             =   240
               Width           =   3375
               Begin MSFlexGridLib.MSFlexGrid FG_P2N2 
                  Height          =   855
                  Left            =   840
                  TabIndex        =   217
                  Top             =   1200
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   1508
                  _Version        =   393216
               End
               Begin MSFlexGridLib.MSFlexGrid FG_P2N 
                  Height          =   1935
                  Left            =   120
                  TabIndex        =   27
                  Top             =   240
                  Width           =   3135
                  _ExtentX        =   5530
                  _ExtentY        =   3413
                  _Version        =   393216
                  SelectionMode   =   1
               End
            End
            Begin VB.Frame FR_4_2_1 
               Height          =   2295
               Left            =   120
               TabIndex        =   214
               Top             =   240
               Width           =   3375
               Begin MSFlexGridLib.MSFlexGrid FG_PED2 
                  Height          =   615
                  Left            =   720
                  TabIndex        =   216
                  Top             =   1320
                  Width           =   1695
                  _ExtentX        =   2990
                  _ExtentY        =   1085
                  _Version        =   393216
               End
               Begin VB.ComboBox CB_Pedidos 
                  Height          =   315
                  Left            =   120
                  Style           =   2  'Dropdown List
                  TabIndex        =   21
                  Top             =   240
                  Width           =   3135
               End
               Begin MSFlexGridLib.MSFlexGrid FG_PED 
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   22
                  Top             =   600
                  Width           =   3135
                  _ExtentX        =   5530
                  _ExtentY        =   2778
                  _Version        =   393216
                  SelectionMode   =   1
               End
            End
         End
         Begin VB.Frame FR_4_1 
            Height          =   852
            Left            =   240
            TabIndex        =   158
            Top             =   240
            Width           =   7932
            Begin VB.OptionButton RD_Pedido 
               Caption         =   "Listar os pedidos pendentes deste cliente, e selecionar um para emitir a nota fiscal."
               Height          =   192
               Left            =   960
               TabIndex        =   28
               ToolTipText     =   "Lista todos pedidos pendentes deste cliente para que possam ser inseridos no assistente da nota fiscal."
               Top             =   480
               Width           =   6252
            End
            Begin VB.OptionButton RD_Inserir 
               Caption         =   "Elaborar a nota fiscal, digitando todos os ítens do pedido"
               Height          =   252
               Left            =   960
               TabIndex        =   20
               ToolTipText     =   "Elabora a nota fiscal digitando os ítens do pedido."
               Top             =   240
               Width           =   4332
            End
         End
      End
      Begin VB.Frame FR_2 
         Caption         =   "Destinatário"
         Height          =   4095
         Left            =   240
         TabIndex        =   112
         ToolTipText     =   "Dados sobre o destinatário."
         Top             =   0
         Visible         =   0   'False
         Width           =   8412
         Begin VB.Frame Frame3 
            BorderStyle     =   0  'None
            Height          =   3735
            Left            =   120
            TabIndex        =   132
            Top             =   240
            Width           =   2052
            Begin VB.CommandButton BT_CadEmp 
               Caption         =   "Cadastro de Empresas"
               Height          =   495
               Left            =   240
               TabIndex        =   210
               ToolTipText     =   "Abre a tela de cadastro de empresas"
               Top             =   3240
               Width           =   1575
            End
            Begin VB.ListBox LT_Apelido 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1500
               ItemData        =   "Tela_NotaFiscal.frx":61C8
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":61CA
               TabIndex        =   209
               ToolTipText     =   "Lista de empresas"
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox TXT_Apelido 
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   336
               Left            =   120
               MaxLength       =   20
               TabIndex        =   208
               ToolTipText     =   "Nome fantasia da empresa"
               Top             =   2760
               Width           =   1812
            End
            Begin VB.ComboBox CB_Exibir 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":61CC
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":61CE
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   207
               ToolTipText     =   "Selecione os tipos de empresas à serem exibidas na lista"
               Top             =   240
               Width           =   1812
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               Caption         =   "Nome Fantasia:"
               Height          =   195
               Left            =   120
               TabIndex        =   213
               Top             =   2520
               Width           =   1140
            End
            Begin VB.Label LB_Exibir 
               AutoSize        =   -1  'True
               Caption         =   "Exibir:"
               Height          =   195
               Left            =   120
               TabIndex        =   212
               Top             =   0
               Width           =   435
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Nome / Razão Social:"
               Height          =   195
               Left            =   120
               TabIndex        =   211
               Top             =   720
               Width           =   1590
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "Dados sobre a Empresa:"
            Enabled         =   0   'False
            Height          =   3852
            Left            =   2280
            TabIndex        =   120
            Top             =   120
            Width           =   6012
            Begin VB.TextBox TXT_PracaPagamento 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   3240
               MaxLength       =   100
               TabIndex        =   12
               ToolTipText     =   "Praça de pagamento (se for difetente do endereço)"
               Top             =   1680
               Width           =   2652
            End
            Begin VB.TextBox TXT_Observacao 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   3120
               MaxLength       =   50
               TabIndex        =   19
               ToolTipText     =   "Observações."
               Top             =   3480
               Width           =   2772
            End
            Begin VB.TextBox TXT_Comentario 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   120
               MaxLength       =   100
               TabIndex        =   18
               ToolTipText     =   "Comentários à serem impressos na nota fiscal somente desta empresa"
               Top             =   3480
               Width           =   2772
            End
            Begin VB.TextBox TXT_Empresa 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   120
               MaxLength       =   50
               TabIndex        =   8
               ToolTipText     =   "Razão Social da empresa."
               Top             =   480
               Width           =   5772
            End
            Begin VB.TextBox TXT_Endereco 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   120
               MaxLength       =   50
               TabIndex        =   11
               ToolTipText     =   "Endereço da empresa"
               Top             =   1680
               Width           =   2892
            End
            Begin VB.TextBox TXT_Bairro 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   120
               MaxLength       =   15
               TabIndex        =   13
               ToolTipText     =   "Bairro da empresa"
               Top             =   2280
               Width           =   2652
            End
            Begin VB.TextBox TXT_Cidade 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   3480
               MaxLength       =   20
               TabIndex        =   14
               ToolTipText     =   "Cidade da empresa"
               Top             =   2280
               Width           =   2412
            End
            Begin VB.ComboBox CB_Estado 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               ItemData        =   "Tela_NotaFiscal.frx":61D0
               Left            =   120
               List            =   "Tela_NotaFiscal.frx":61D2
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   15
               ToolTipText     =   "Selecione o estado"
               Top             =   2880
               Width           =   972
            End
            Begin VB.TextBox TXT_Fone 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   3960
               MaxLength       =   16
               TabIndex        =   17
               ToolTipText     =   "Telefone da empresa"
               Top             =   2880
               Width           =   1932
            End
            Begin VB.TextBox TXT_InsEst 
               BackColor       =   &H80000016&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   10.5
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   3240
               MaxLength       =   25
               TabIndex        =   10
               ToolTipText     =   "Inscrição Estadual da empresa"
               Top             =   1080
               Width           =   2652
            End
            Begin MSMask.MaskEdBox TXT_CGC 
               Height          =   288
               Left            =   120
               TabIndex        =   9
               ToolTipText     =   "C.N.P.J. (antigo C.G.C.) da empresa."
               Top             =   1080
               Width           =   2892
               _ExtentX        =   5106
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               BackColor       =   -2147483626
               ForeColor       =   -2147483630
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   18
               Mask            =   "99.999.999/9999-99"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TXT_CEP 
               Height          =   288
               Left            =   1680
               TabIndex        =   16
               ToolTipText     =   "CEP da empresa."
               Top             =   2880
               Width           =   1692
               _ExtentX        =   2990
               _ExtentY        =   503
               _Version        =   393216
               ClipMode        =   1
               BackColor       =   -2147483626
               ForeColor       =   -2147483630
               AllowPrompt     =   -1  'True
               AutoTab         =   -1  'True
               MaxLength       =   9
               Mask            =   "99999-999"
               PromptChar      =   "_"
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Praça de Pagamento:"
               Height          =   192
               Left            =   3240
               TabIndex        =   133
               Top             =   1440
               Width           =   1572
            End
            Begin VB.Label LB_Observacao 
               AutoSize        =   -1  'True
               Caption         =   "Observações:"
               Height          =   192
               Left            =   3120
               TabIndex        =   131
               Top             =   3240
               Width           =   1020
            End
            Begin VB.Label LB_Comentarios 
               AutoSize        =   -1  'True
               Caption         =   "Comentários de Nota Fiscal:"
               Height          =   192
               Left            =   120
               TabIndex        =   130
               Top             =   3240
               Width           =   2028
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Empresa:"
               Height          =   192
               Left            =   120
               TabIndex        =   129
               Top             =   240
               Width           =   696
            End
            Begin VB.Label Label15 
               AutoSize        =   -1  'True
               Caption         =   "Endereço:"
               Height          =   192
               Left            =   120
               TabIndex        =   128
               Top             =   1440
               Width           =   744
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Bairro:"
               Height          =   192
               Left            =   120
               TabIndex        =   127
               Top             =   2040
               Width           =   468
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Cidade:"
               Height          =   192
               Left            =   3480
               TabIndex        =   126
               Top             =   2040
               Width           =   564
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Estado:"
               Height          =   192
               Left            =   120
               TabIndex        =   125
               Top             =   2640
               Width           =   552
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "CEP:"
               Height          =   192
               Left            =   1680
               TabIndex        =   124
               Top             =   2640
               Width           =   360
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Fone:"
               Height          =   192
               Left            =   3960
               TabIndex        =   123
               Top             =   2640
               Width           =   408
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Inscrição Estadual:"
               Height          =   192
               Left            =   3240
               TabIndex        =   122
               Top             =   840
               Width           =   1356
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "C.N.P.J.:"
               Height          =   192
               Left            =   120
               TabIndex        =   121
               Top             =   840
               Width           =   600
            End
         End
      End
      Begin VB.Frame FR_1 
         Caption         =   "Operação"
         Height          =   4092
         Left            =   0
         TabIndex        =   113
         ToolTipText     =   "Dados sobre a operação."
         Top             =   0
         Width           =   8412
         Begin VB.ComboBox CB_Natureza 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   6
            ToolTipText     =   "Selecione a natureza da operação"
            Top             =   1440
            Width           =   2892
         End
         Begin VB.ComboBox CB_CFOP 
            Height          =   315
            Left            =   4320
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Selecione a C.F.O.P."
            Top             =   2520
            Width           =   2892
         End
         Begin VB.Frame FR_1_1 
            Caption         =   "Tipo:"
            Height          =   1572
            Left            =   1200
            TabIndex        =   114
            ToolTipText     =   "Tipo de operação"
            Top             =   1200
            Width           =   1692
            Begin VB.OptionButton RD_Saida 
               Caption         =   "Saida"
               Height          =   192
               Left            =   480
               TabIndex        =   4
               ToolTipText     =   "Nota Fiscal de Saída"
               Top             =   480
               Width           =   1092
            End
            Begin VB.OptionButton RD_Entrada 
               Caption         =   "Entrada"
               Height          =   252
               Left            =   480
               TabIndex        =   5
               ToolTipText     =   "Nota Fiscal de Entrada"
               Top             =   1080
               Width           =   1092
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Natureza da Operação:"
            Height          =   192
            Left            =   4332
            TabIndex        =   116
            Top             =   1248
            Width           =   1680
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "CFOP:"
            Height          =   192
            Left            =   4332
            TabIndex        =   115
            Top             =   2328
            Width           =   468
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conesteel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   0
         TabIndex        =   119
         Top             =   4080
         Width           =   1452
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conesteel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   240
         TabIndex        =   118
         Top             =   4440
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Conesteel"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   504
         Left            =   120
         TabIndex        =   117
         Top             =   4200
         Width           =   2292
      End
   End
End
Attribute VB_Name = "Tela_NotaFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Assistente de elaboração da Nota Fiscal"
Dim TRANSPORTADORA As String, tINFPED As T_INFPED, nI As Integer

Private Type T_INFPED
    CON As String
    SPE As String
    CPG As String
    TRA As String
    IVE As String
    ITE As String
    NPE As String
End Type

Private Sub BT_ApagarItem_Click()
    On Error GoTo ERRO_SISCOVAL
    Dim NumLinha, Num
    NumLinha = 0
    For I = 1 To 20
        If FG_1.TextMatrix(I, 1) = "" And _
           FG_1.TextMatrix(I, 2) = "" Then
            NumLinha = I
            If NumLinha = 20 Then
                RespMsg = MsgBox("Não existem ítens para serem removidos.", vbOKOnly, Tela_NotaFiscal.Caption)
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next I
    NumLinha = InputBox("Digite o número da linha/ítem que você deseja remover ( 1 - 20 ):", "Remover ítem")
    If IsNumeric(NumLinha) = False Then
        RespMsg = MsgBox("Você não digitou um número válido.", vbOKOnly, Tela_NotaFiscal.Caption)
        Exit Sub
    ElseIf NumLinha = "" Then
        Exit Sub
    ElseIf NumLinha > 20 Then
        RespMsg = MsgBox("Você deve digitar um número entre 1 e 20.", vbOKOnly, Tela_NotaFiscal.Caption)
        Exit Sub
    ElseIf NumLinha < 1 Then
        RespMsg = MsgBox("Você deve digitar um número entre 1 e 20.", vbOKOnly, Tela_NotaFiscal.Caption)
        Exit Sub
    ElseIf Trim(TXT_Comentario.Text) <> "" And NumLinha = 20 Then
        RespMsg = MsgBox("A última linha não pode ser removida porque se refere aos comentários de nota fiscal cadastrados desta empresa.")
        FG_1.TextMatrix(20, 1) = Trim(TXT_Comentario.Text)
        Exit Sub
    ElseIf NumLinha >= 1 And NumLinha <= 20 Then
        If FG_1.TextMatrix(NumLinha, 1) = "" And _
           FG_1.TextMatrix(NumLinha, 2) = "" Then
                RespMsg = MsgBox("Não existem dados na linha que você deseja remover.", vbOKOnly, Tela_NotaFiscal.Caption)
                Exit Sub
        End If
        TelaNFEmEspera (True)
        'Verifica se o item tem mais de uma linha
        
        'Item Assistente Simples
        If FG_1.TextMatrix(NumLinha, 1) <> "" And _
           FG_1.TextMatrix(NumLinha, 2) <> "" And _
           FG_1.TextMatrix(NumLinha, 3) <> "" And _
           FG_1.TextMatrix(NumLinha, 4) <> "" And _
           FG_2.TextMatrix(NumLinha, 1) <> "" And _
           FG_2.TextMatrix(NumLinha, 2) <> "" Then
            LimpaLinha (NumLinha)
        'Item Assistente Multiplo (Primeira linha)
        ElseIf FG_1.TextMatrix(NumLinha, 1) <> "" And _
           FG_1.TextMatrix(NumLinha, 2) <> "" And _
           FG_1.TextMatrix(NumLinha, 3) = "" And _
           FG_1.TextMatrix(NumLinha, 4) = "" And _
           FG_2.TextMatrix(NumLinha, 1) <> "" And _
           FG_2.TextMatrix(NumLinha, 2) = "" Then
            LimpaLinha (NumLinha)
            For I = (NumLinha + 1) To 20
                If FG_2.TextMatrix(I, 1) = "Idem" Then
                    LimpaLinha (I)
                Else
                    Exit For
                End If
            Next I
        'Item Manual Simples (Primeira linha)
        ElseIf FG_1.TextMatrix(NumLinha, 1) = "" And _
           FG_1.TextMatrix(NumLinha, 2) <> "" And _
           FG_1.TextMatrix(NumLinha, 3) <> "" And _
           FG_1.TextMatrix(NumLinha, 4) <> "" And _
           FG_2.TextMatrix(NumLinha, 1) = "MN" And _
           FG_2.TextMatrix(NumLinha, 2) = "" Then
            LimpaLinha (NumLinha)
        'Item Manual Multiplo (Primeira linha)
        ElseIf FG_1.TextMatrix(NumLinha, 1) = "" And _
           FG_1.TextMatrix(NumLinha, 2) <> "" And _
           FG_1.TextMatrix(NumLinha, 3) = "" And _
           FG_1.TextMatrix(NumLinha, 4) = "" And _
           FG_2.TextMatrix(NumLinha, 1) = "MN" And _
           FG_2.TextMatrix(NumLinha, 2) = "" Then
            LimpaLinha (NumLinha)
            For I = (NumLinha + 1) To 20
                If FG_2.TextMatrix(I, 1) = "Idem" Then
                    LimpaLinha (I)
                Else
                    Exit For
                End If
            Next I
        'Item Multiplo (Qualquer linha)
        ElseIf FG_1.TextMatrix(NumLinha, 1) = "" And _
           FG_1.TextMatrix(NumLinha, 2) <> "" And _
           FG_2.TextMatrix(NumLinha, 1) = "Idem" Then
            'procura 1ª linha do item
            Dim PriLin As Integer
            For I = NumLinha To 1 Step -1
                If FG_2.TextMatrix(I, 1) <> "Idem" Then
                    LimpaLinha (I)
                    PriLin = I
                    Exit For
                End If
            Next I
            'limpa o restante das linhas
            For I = (PriLin + 1) To 20
                If FG_2.TextMatrix(I, 1) = "Idem" Then
                    LimpaLinha (I)
                Else
                    Exit For
                End If
            Next I
        End If
    Else
        Exit Sub
    End If
    
    Dim Y As Integer
    Dim W As Integer
    Y = 1
    'Verifica linhas em branco e puxa linhas com dados para cima
    Do While Y < 20
        If FG_1.TextMatrix(Y, 1) = "" And _
           FG_2.TextMatrix(Y, 1) = "" Then
            W = 1
            For J = (Y + 1) To 19
                If FG_1.TextMatrix(J, 1) = "" And _
                   FG_2.TextMatrix(J, 1) = "" Then
                    W = W + 1
                ElseIf J = 19 Then
                    Exit Do
                Else
                    Exit For
                End If
            Next J
            FG_1.TextMatrix(Y, 1) = FG_1.TextMatrix(Y + W, 1)
            FG_1.TextMatrix(Y, 2) = FG_1.TextMatrix(Y + W, 2)
            FG_1.TextMatrix(Y, 3) = FG_1.TextMatrix(Y + W, 3)
            FG_1.TextMatrix(Y, 4) = FG_1.TextMatrix(Y + W, 4)
            FG_1.TextMatrix(Y, 5) = FG_1.TextMatrix(Y + W, 5)
            FG_1.TextMatrix(Y, 6) = FG_1.TextMatrix(Y + W, 6)
            FG_1.TextMatrix(Y, 7) = FG_1.TextMatrix(Y + W, 7)
            FG_1.TextMatrix(Y, 8) = FG_1.TextMatrix(Y + W, 8)
            FG_1.TextMatrix(Y, 9) = FG_1.TextMatrix(Y + W, 9)
            FG_1.TextMatrix(Y, 10) = FG_1.TextMatrix(Y + W, 10)
            FG_1.TextMatrix(Y, 11) = FG_1.TextMatrix(Y + W, 11)
            FG_2.TextMatrix(Y, 1) = FG_2.TextMatrix(Y + W, 1)
            FG_2.TextMatrix(Y, 2) = FG_2.TextMatrix(Y + W, 2)
            FG_2.TextMatrix(Y, 3) = FG_2.TextMatrix(Y + W, 3)
            FG_2.TextMatrix(Y, 4) = FG_2.TextMatrix(Y + W, 4)
            FG_2.TextMatrix(Y, 5) = FG_2.TextMatrix(Y + W, 5)
            LimpaLinha (Y + W)
        End If
        Y = Y + 1
    Loop
    
    LimpaLinha (20)
    BT_InserirItem.Enabled = True
    
    'Se for incluir declaracoes, verificar as pecas que sobraram
    If BT_DeclaracoesFiscais.Caption = "Remover Dec.Fiscais" Then
        BT_DeclaracoesFiscais.Value = True
        BT_DeclaracoesFiscais.Value = True 'sAO 2
    End If
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_ApagarTudo_Click()
    On Error GoTo ERRO_SISCOVAL
    For I = 1 To 20
        If FG_1.TextMatrix(I, 1) = "" And _
           FG_1.TextMatrix(I, 2) = "" Then
            NumLinha = I
            If NumLinha = 20 Then
                RespMsg = MsgBox("Não existem ítens à serem removidos.", vbOKOnly, Tela_NotaFiscal.Caption)
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next I
    TelaNFEmEspera (True)
    FG_1.Clear
    FG_2.Clear
    CF_I = ""
    CF_J = ""
    BT_DeclaracoesFiscais.Caption = "Incluir Dec.Fiscais"
    BT_DepositoBancario.Caption = "Incluir C/C Banco"
    BT_Comentarios.Caption = "Incluir comentários"
    If Trim(TXT_Comentario.Text) <> "" Then
        FG_1.TextMatrix(20, 1) = Trim(TXT_Comentario.Text)
    End If
    MontaFG
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Avancar_Click()
    On Error GoTo ERRO_SISCOVAL
    If FR_1.Visible = True Then 'Tipo NF
        BT_Concluir.Enabled = False
        BT_Voltar.Enabled = True
        FR_1.Visible = False
        FR_2.Visible = True
        TelaNFEmEspera (True)
        For I = 0 To CB_Exibir.ListCount - 1
            If CB_Exibir.List(I) = "Cliente" Then
                CB_Exibir.ListIndex = I
                Exit For
            End If
        Next I
        TelaNFEmEspera (False)
        LT_Apelido.SetFocus
    ElseIf FR_2.Visible = True Then 'Empresa
        If LT_Apelido.Text = "" Then
            RespMsg = MsgBox("Escolha uma empresa.", vbOKOnly, Tela_NotaFiscal.Caption)
            LT_Apelido.SetFocus
            Exit Sub
        End If
        FR_2.Visible = False
        FR_3.Visible = True
        RD_Inserir.SetFocus
    ElseIf FR_3.Visible = True Then 'Pedidos
        If RD_Pedido.Value = True And FG_P2N.Rows > 1 Then
            'importar pedido para NF
            ImportaPedido
        End If
        FR_3.Visible = False
        FR_4.Visible = True
        TXT_NF.SetFocus
    ElseIf FR_4.Visible = True Then
        If TXT_NF.Text = "" Then
            RespMsg = MsgBox("Confirme o número da nota fiscal.", vbOKOnly, Tela_NotaFiscal.Caption)
            DLL_BD.BDSIS_TBNTF.MoveLast
            TXT_NF.Text = DLL_BD.BDSIS_TBNTF_CPNNF.Value + 1
            TXT_NF.SetFocus
        ElseIf TXT_DataEmissao.Text = "" Then
            RespMsg = MsgBox("Confirme a data de emissão da nota fiscal.", vbOKOnly, Tela_NotaFiscal.Caption)
            TXT_DataEmissao.Text = Date
            TXT_DataEmissao.SetFocus
        ElseIf TXT_DataVenc_A.Text = "" Then
            RespMsg = MsgBox("Confirme a data de vencimento da nota fiscal.", vbOKOnly, Tela_NotaFiscal.Caption)
            RD_28dd.Value = True
            TXT_DataVenc_A.Text = DateAdd("d", 28, Date)
            TXT_DataVenc_A.SetFocus
        Else
            FR_4.Visible = False
            FR_5.Visible = True
        End If
        If Trim(TXT_Comentario.Text) <> "" Then
            FG_1.TextMatrix(20, 1) = Trim(TXT_Comentario.Text)
        End If
        'Limpa Datas
        If TXT_DVB.Text = "" Then
            TXT_DataVenc_B.Text = Space(10)
            TXT_DataVenc_C.Text = Space(10)
            TXT_DataVenc_D.Text = Space(10)
        ElseIf TXT_DVC.Text = "" Then
            TXT_DataVenc_C.Text = Space(10)
            TXT_DataVenc_D.Text = Space(10)
        ElseIf TXT_DVD.Text = "" Then
            TXT_DataVenc_D.Text = Space(10)
        End If
        BT_InserirItem.SetFocus
    ElseIf FR_5.Visible = True Then
        Dim Num
        TXT_BaseICMS.Text = 0
        TXT_ValorICMS.Text = 0
        TXT_ValorTotalIPI.Text = 0
        TXT_ValorTotalProdutos.Text = 0
        TXT_ValorTotalNotaFiscal.Text = 0
        TXT_BaseICMSSub.Text = 0
        TXT_ValorICMSSub.Text = 0
        TXT_ValorFrete.Text = 0
        TXT_ValorSeguro.Text = 0
        TXT_Outras.Text = 0
        Num = 0
        For I = 1 To 20
            If FG_1.TextMatrix(I, 1) = "" And _
               FG_1.TextMatrix(I, 2) = "" And _
               FG_1.TextMatrix(I, 3) = "" Then
               Num = Num + 1
                If Num = 20 Then
                    RespMsg = MsgBox("Não existe nenhum ítem na nota fiscal.", vbOKOnly, Tela_NotaFiscal.Caption)
                    Exit Sub
                End If
            Else
                If FG_2.TextMatrix(I, 3) <> "" Then TXT_BaseICMS.Text = CDbl(TXT_BaseICMS.Text) + CDbl(FG_2.TextMatrix(I, 3))
                If FG_2.TextMatrix(I, 4) <> "" Then TXT_ValorICMS.Text = CDbl(TXT_ValorICMS.Text) + CDbl(FG_2.TextMatrix(I, 4))
                If FG_1.TextMatrix(I, 11) <> "" Then TXT_ValorTotalIPI.Text = CDbl(TXT_ValorTotalIPI.Text) + CDbl(FG_1.TextMatrix(I, 11))
                If FG_1.TextMatrix(I, 8) <> "" Then TXT_ValorTotalProdutos.Text = CDbl(TXT_ValorTotalProdutos.Text) + CDbl(FG_1.TextMatrix(I, 8))
            End If
        Next I
        TelaNFEmEspera (True)
        For I = 1 To 20
            If FG_1.TextMatrix(I, 3) = "I" Then
                If CF_I = "" Then
                    Do While True
                        CF_I = InputBox("Você inseriu um item na nota fiscal que tem classificação fiscal letra I, do qual tem espaço em branco na nota fiscal para entrar com o seu devido código - Por favor digite o código desta classificação fiscal (Ex.: 7307.92.00)", "Código da C.F. letra I", "0000.00.00")
                        If CF_I <> "" Then Exit Do
                    Loop
                End If
            ElseIf FG_1.TextMatrix(I, 3) = "J" Then
                If CF_J = "" Then
                    Do While True
                        CF_J = InputBox("Você inseriu um item na nota fiscal que tem classificação fiscal letra J, do qual tem espaço em branco na nota fiscal para entrar com o seu devido código - Por favor digite o código desta classificação fiscal (Ex.: 7307.92.00)", "Código da C.F. letra J", "0000.00.00")
                        If CF_J <> "" Then Exit Do
                    Loop
                End If
            End If
        Next I
        TXT_ValorTotalNotaFiscal.Text = CDbl(TXT_ValorTotalProdutos.Text) + CDbl(TXT_ValorTotalIPI.Text)
        If TXT_BaseICMSSub.Text = 0 Then TXT_BaseICMSSub.Text = "0,00"
        If TXT_ValorICMSSub.Text = 0 Then TXT_ValorICMSSub.Text = "0,00"
        If TXT_ValorFrete.Text = 0 Then TXT_ValorFrete.Text = "0,00"
        If TXT_ValorSeguro.Text = 0 Then TXT_ValorSeguro.Text = "0,00"
        If TXT_Outras.Text = 0 Then TXT_Outras.Text = "0,00"
        FR_5.Visible = False
        FR_6.Visible = True
        TelaNFEmEspera (False)
        CK_EditarValores.SetFocus
    ElseIf FR_6.Visible = True Then
        FR_6.Visible = False
        FR_7.Visible = True
        RB_CF1.SetFocus
    ElseIf FR_7.Visible = True Then
        'Se for empresa de SP
        TelaNFEmEspera (True)
        CB_Frete.ListIndex = 0
        If TXT_Cidade.Text = "São Paulo" Then
            For I = 0 To LT_NomeTrans.ListCount - 1
                If LT_NomeTrans.List(I) = "NOSSO MOTORISTA" Then
                    LT_NomeTrans.ListIndex = I
                    Exit For
                End If
            Next I
            'Calcula Quantidade
            TXT_QuantVol.Text = 0
            For I = 1 To 20
                If FG_1.TextMatrix(I, 6) <> "" Then
                    TXT_QuantVol.Text = CDbl(TXT_QuantVol.Text) + CDbl(FG_1.TextMatrix(I, 6))
                End If
            Next I
            'Calcula Peso Liquido e Bruto
            TXT_PesoLiquido.Text = 0
            For I = 1 To 20
                If FG_1.TextMatrix(I, 5) <> "" Then
                    TXT_PesoLiquido.Text = CDbl(TXT_PesoLiquido.Text) + CDbl(FG_2.TextMatrix(I, 5))
                End If
            Next I
            TXT_PesoBruto.Text = TXT_PesoLiquido.Text
            CB_Frete.ListIndex = 1
        Else 'Se não for estado de SP
            LT_NomeTrans.ListIndex = -1
            CB_EstVei.ListIndex = -1
            CB_Frete.ListIndex = 0
            CB_Especie.ListIndex = 0
            CB_Placa.ListIndex = -1
            TXT_QuantVol.Text = 0
            TXT_PesoBruto.Text = 0
            TXT_NumVol.Text = ""
            'Calcula Peso Liquido
            TXT_PesoLiquido.Text = 0
            For I = 1 To 20
                If FG_1.TextMatrix(I, 5) <> "" Then
                    TXT_PesoLiquido.Text = CDbl(TXT_PesoLiquido.Text) + CDbl(FG_2.TextMatrix(I, 5))
                End If
            Next I
        End If
        FR_7.Visible = False
        FR_8.Visible = True
        'Se tiver alguma tranp. na ficha
        If TRANSPORTADORA <> "" Then
            For I = 0 To LT_NomeTrans.ListCount - 1
                If Trim(LT_NomeTrans.List(I)) = Trim(TRANSPORTADORA) Then
                    LT_NomeTrans.ListIndex = I
                    Exit For
                End If
            Next I
        End If
        TelaNFEmEspera (False)
        LT_NomeTrans.SetFocus
    ElseIf FR_8.Visible = True Then
        FR_8.Visible = False
        FR_9.Visible = True
        BT_Avancar.Enabled = False
        BT_Concluir.Enabled = True
        BT_Concluir.SetFocus
    End If
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_CadEmp_Click()
    Dim sEmpTemp1 As String, sEmpTemp2 As String, nItem As Integer
    If LT_Apelido.ListIndex > -1 Then
        sEmpTemp1 = LT_Apelido.Text
    Else
        sEmpTemp1 = ""
    End If
    nItem = CB_Exibir.ListIndex
    Me.Hide
    sEmpTemp2 = DLL_CADEMP.CadastroEmpresa(App.ProductName, "Cademp", App.LegalCopyright, sEmpTemp1)
    If sEmpTemp2 <> "" Then
        'reCarregando combo de empresas
        CB_Exibir.ListIndex = nItem
        LT_Apelido.ListIndex = -1
        LT_Apelido.Text = sEmpTemp2
    End If
    Me.Show vbModal
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_NotaFiscal
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Comentarios_Click()
    On Error GoTo ERRO_SISCOVAL
    'Verifica se existem itens na NF
    For I = 1 To 20
        If FG_1.TextMatrix(I, 2) <> "" And _
           FG_2.TextMatrix(I, 1) <> "" Then
            Exit For
        ElseIf I = 20 Then
            RespMsg = MsgBox("Você ainda não inseriu nenhum ítem.", vbOKOnly, Tela_NotaFiscal.Caption)
            Exit Sub
        End If
    Next I
    Tela_NotaFiscal_Dlg_5.Show vbModal
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Concluir_Click()
    On Error GoTo ERRO_SISCOVAL
    'Botoes:
    '1: Cancelado
    '2: Verificando
    '3: OK
    TelaNFEmEspera (True)
    BP.Max = 12
    
    '*********************************************
    'Reseta Operações
    '*********************************************
    BP.Value = 0
    IMG_1.Visible = False
    IMG_2.Visible = False
    IMG_3.Visible = False
    IMG_4.Visible = False
    IMG_5.Visible = False
    IMG_6.Visible = False
    IMG_7.Visible = False
    IMG_8.Visible = False
    IMG_9.Visible = False
    IMG_10.Visible = False
    IMG_11.Visible = False
    IMG_12.Visible = False
    LB_1.FontBold = False
    LB_2.FontBold = False
    LB_3.FontBold = False
    LB_4.FontBold = False
    LB_5.FontBold = False
    LB_6.FontBold = False
    LB_7.FontBold = False
    LB_8.FontBold = False
    LB_9.FontBold = False
    LB_10.FontBold = False
    LB_11.FontBold = False
    LB_12.FontBold = False
    
    '*********************************************
    '1-) Confere número NF
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_1.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_1.Visible = True
    LB_1.FontBold = True
    DLL_BD.BDSIS_TBNTF.MoveLast
    If TXT_NF.Text = DLL_BD.BDSIS_TBNTF_CPNNF.Value + 1 Then
        IMG_1.Picture = LI.ListImages(3).Picture 'OK
        LB_1.FontBold = False
    Else
        IMG_1.Picture = LI.ListImages(1).Picture 'Erro
        RespMsg = MsgBox("Ao conferir o número da nota fiscal, foi constatado que o número " & TXT_NF.Text & " que está no Assistente de Nota Fiscal não corresponde ao número sequencial, que é o " & (DLL_BD.BDSIS_TBNTF_CPNNF.Value + 1) & ". Selecione Sim para alterar o número da nota fiscal ou Não para corrigir o número.", vbInformation + vbYesNo + vbDefaultButton1, "Conflito com o número de Nota Fiscal")
        If RespMsg = vbYes Then
            RespMsg = MsgBox("O número da nota fiscal foi alterado para " & TXT_NF.Text & ".", vbInformation + vbOKOnly, "Alteração do número")
        Else
            TXT_NF.Text = DLL_BD.BDSIS_TBNTF_CPNNF.Value + 1
        End If
        IMG_1.Picture = LI.ListImages(3).Picture 'OK
        LB_1.FontBold = False
    End If
        
    '*********************************************
    '2-) Salva dados da Nota Fiscal
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_2.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_2.Visible = True
    LB_2.FontBold = True
    'salvando tabela Nota Fiscal
    If LB_IMP.Visible = False Then
        DLL_BD.BDSIS_TBNTF.AddNew
    ElseIf LB_IMP.Visible = True Then
        DLL_BD.BDSIS_TBNTF.Seek "=", Val(TXT_NF.Text)
        If DLL_BD.BDSIS_TBNTF.NoMatch Then
            MsgBox "Erro ao procurar a nota fiscal", vbCritical + vbOKOnly, "ERRO"
            Unload Me
        End If
        DLL_BD.BDSIS_TBNTF.Edit
    End If
    DLL_BD.BDSIS_TBNTF_CPNNF.Value = TXT_NF.Text
    If IsDate(TXT_DataEmissao.Text) Then DLL_BD.BDSIS_TBNTF_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
    If RD_Saida.Value = True Then
        DLL_BD.BDSIS_TBNTF_CPTIP.Value = "Saída"
    ElseIf RD_Entrada.Value = True Then
        DLL_BD.BDSIS_TBNTF_CPTIP.Value = "Entrada"
    End If
    DLL_BD.BDSIS_TBNTF_CPNOP.Value = CB_Natureza.Text
    DLL_BD.BDSIS_TBNTF_CPCFO.Value = CB_CFOP.Text
    DLL_BD.BDSIS_TBNTF_CPEMP.Value = Trim(TXT_Apelido.Text)
    If CK_DataSaida.Value = 1 And IsDate(TXT_DataSaida.Text) Then DLL_BD.BDSIS_TBNTF_CPDSA.Value = Format(TXT_DataSaida.Text, "dd/mm/yyyy")
    If TXT_HoraSaida.Text <> "" And Not IsError(Format(TXT_HoraSaida.Text, "hh:mm:ss")) Then DLL_BD.BDSIS_TBNTF_CPHSA.Value = Format(TXT_HoraSaida.Text, "hh:mm:ss")
    DLL_BD.BDSIS_TBNTF_CPHEM.Value = Format(Time, "hh:mm:ss")
    DLL_BD.BDSIS_TBNTF_CPVAL.Value = TXT_ValorTotalNotaFiscal.Text
    If CK_Desdobrar.Value = 0 Then
        DLL_BD.BDSIS_TBNTF_CPNDP.Value = TXT_NF.Text
        DLL_BD.BDSIS_TBNTF_CPVEN.Value = TXT_DataVenc_A.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPNDP.Value = "Vide abaixo"
        DLL_BD.BDSIS_TBNTF_CPVEN.Value = "Vide abaixo"
    End If
    If TXT_PedidoInterno.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPNPI.Value = TXT_PedidoInterno.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPNPI.Value = ""
    End If
    If TXT_SeuPedido.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPNSP.Value = TXT_SeuPedido.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPNSP.Value = ""
    End If
    If TXT_Operacao.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPOPE.Value = TXT_Operacao.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPOPE.Value = ""
    End If
    If BT_DeclaracoesFiscais.Caption = "Remover Dec.Fiscais" Then 'Tem declarações
        DLL_BD.BDSIS_TBNTF_CPDEC.Value = True
    Else
        DLL_BD.BDSIS_TBNTF_CPDEC.Value = False
    End If
    If BT_DepositoBancario.Caption = "Remover C/C Banco" Then 'Tem bancos
        DLL_BD.BDSIS_TBNTF_CPBAN.Value = True
    Else
        DLL_BD.BDSIS_TBNTF_CPBAN.Value = False
    End If
    If BT_Comentarios.Caption = "Remover comentários" Then 'Tem comentários
        DLL_BD.BDSIS_TBNTF_CPCOM.Value = True
    Else
        DLL_BD.BDSIS_TBNTF_CPCOM.Value = False
    End If
    DLL_BD.BDSIS_TBNTF_CPBCI.Value = TXT_BaseICMS.Text
    DLL_BD.BDSIS_TBNTF_CPVIC.Value = TXT_ValorICMS.Text
    DLL_BD.BDSIS_TBNTF_CPBIS.Value = TXT_BaseICMSSub.Text
    DLL_BD.BDSIS_TBNTF_CPVIS.Value = TXT_ValorICMSSub.Text
    DLL_BD.BDSIS_TBNTF_CPVTP.Value = TXT_ValorTotalProdutos.Text
    DLL_BD.BDSIS_TBNTF_CPVFR.Value = TXT_ValorFrete.Text
    DLL_BD.BDSIS_TBNTF_CPVSE.Value = TXT_ValorSeguro.Text
    DLL_BD.BDSIS_TBNTF_CPODA.Value = TXT_Outras.Text
    DLL_BD.BDSIS_TBNTF_CPVIP.Value = TXT_ValorTotalIPI.Text
    DLL_BD.BDSIS_TBNTF_CPVTN.Value = TXT_ValorTotalNotaFiscal.Text
    DLL_BD.BDSIS_TBNTF_CPTRA.Value = LT_NomeTrans.Text
    DLL_BD.BDSIS_TBNTF_CPFCO.Value = CB_Frete.Text
    If CB_Placa.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPPVE.Value = CB_Placa.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPPVE.Value = ""
    End If
    If CB_EstVei.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPEPV.Value = CB_EstVei.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPEPV.Value = ""
    End If
    If TXT_QuantVol.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPQUA.Value = TXT_QuantVol.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPQUA.Value = 0
    End If
    If CB_Especie.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPESP.Value = CB_Especie.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPESP.Value = ""
    End If
    If CB_Marca.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPMAR.Value = CB_Marca.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPMAR.Value = ""
    End If
    If TXT_NumVol.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPNVO.Value = TXT_NumVol.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPNVO.Value = ""
    End If
    If TXT_PesoBruto.Text <> "" Then 'Peso Bruto
        DLL_BD.BDSIS_TBNTF_CPPBR.Value = TXT_PesoBruto.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPPBR.Value = 0
    End If
    If TXT_PesoLiquido.Text <> "" Then 'Peso Liquido
        DLL_BD.BDSIS_TBNTF_CPPLI.Value = TXT_PesoLiquido.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPPLI.Value = 0
    End If
    If TXT_VendInterno.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPVIN.Value = TXT_VendInterno.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPVIN.Value = ""
    End If
    If TXT_VendExterno.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPVEX.Value = TXT_VendExterno.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPVEX.Value = ""
    End If
    If TXT_Setor.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPSET.Value = TXT_Setor.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPSET.Value = ""
    End If
    If TXT_Setor.Text <> "" Then
        DLL_BD.BDSIS_TBNTF_CPSET.Value = TXT_Setor.Text
    Else
        DLL_BD.BDSIS_TBNTF_CPSET.Value = ""
    End If
    If CF_I <> "" Then DLL_BD.BDSIS_TBNTF_CPCFI.Value = CF_I
    If CF_J <> "" Then DLL_BD.BDSIS_TBNTF_CPCFJ.Value = CF_J
    If TXT_DVB.Text = "" Then
        TXT_DataVenc_B.Text = Space(10)
        TXT_DataVenc_C.Text = Space(10)
        TXT_DataVenc_D.Text = Space(10)
    End If
    If TXT_DVC.Text = "" Then
        TXT_DataVenc_C.Text = Space(10)
        TXT_DataVenc_D.Text = Space(10)
    End If
    If TXT_DVD.Text = "" Then
        TXT_DataVenc_D.Text = Space(10)
    End If
    If CK_Desdobrar.Value = 0 Then 'DP NORMAL
        DLL_BD.BDSIS_TBNTF_CPVLA.Value = TXT_ValorTotalNotaFiscal.Text
        DLL_BD.BDSIS_TBNTF_CPVCA.Value = Format(TXT_DataVenc_A.Text, "dd/mm/yyyy")
        DLL_BD.BDSIS_TBNTF_CPVLB.Value = 0
        DLL_BD.BDSIS_TBNTF_CPVCB.Value = 0
        DLL_BD.BDSIS_TBNTF_CPVLC.Value = 0
        DLL_BD.BDSIS_TBNTF_CPVCC.Value = 0
        DLL_BD.BDSIS_TBNTF_CPVLD.Value = 0
        DLL_BD.BDSIS_TBNTF_CPVCD.Value = 0
    ElseIf CK_Desdobrar.Value = 1 Then 'DP DESDOBRADA
        If TXT_DataVenc_A.Text <> Space(10) And _
           TXT_DataVenc_B.Text <> Space(10) And _
           TXT_DataVenc_C.Text = Space(10) And _
           TXT_DataVenc_D.Text = Space(10) Then
            DLL_BD.BDSIS_TBNTF_CPVLA.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 2
            DLL_BD.BDSIS_TBNTF_CPVCA.Value = Format(TXT_DataVenc_A.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLB.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 2
            DLL_BD.BDSIS_TBNTF_CPVCB.Value = Format(TXT_DataVenc_B.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLC.Value = 0
            DLL_BD.BDSIS_TBNTF_CPVCC.Value = 0
            DLL_BD.BDSIS_TBNTF_CPVLD.Value = 0
            DLL_BD.BDSIS_TBNTF_CPVCD.Value = 0
        ElseIf TXT_DataVenc_A.Text <> Space(10) And _
           TXT_DataVenc_B.Text <> Space(10) And _
           TXT_DataVenc_C.Text <> Space(10) And _
           TXT_DataVenc_D.Text = Space(10) Then
            DLL_BD.BDSIS_TBNTF_CPVLA.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 3
            DLL_BD.BDSIS_TBNTF_CPVCA.Value = Format(TXT_DataVenc_A.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLB.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 3
            DLL_BD.BDSIS_TBNTF_CPVCB.Value = Format(TXT_DataVenc_B.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLC.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 3
            DLL_BD.BDSIS_TBNTF_CPVCC.Value = Format(TXT_DataVenc_C.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLD.Value = 0
            DLL_BD.BDSIS_TBNTF_CPVCD.Value = 0
        ElseIf TXT_DataVenc_A.Text <> Space(10) And _
           TXT_DataVenc_B.Text <> Space(10) And _
           TXT_DataVenc_C.Text <> Space(10) And _
           TXT_DataVenc_D.Text <> Space(10) Then
            DLL_BD.BDSIS_TBNTF_CPVLA.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
            DLL_BD.BDSIS_TBNTF_CPVCA.Value = Format(TXT_DataVenc_A.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLB.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
            DLL_BD.BDSIS_TBNTF_CPVCB.Value = Format(TXT_DataVenc_B.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLC.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
            DLL_BD.BDSIS_TBNTF_CPVCC.Value = Format(TXT_DataVenc_C.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBNTF_CPVLD.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
            DLL_BD.BDSIS_TBNTF_CPVCD.Value = Format(TXT_DataVenc_D.Text, "dd/mm/yyyy")
        End If
    End If
    DLL_BD.BDSIS_TBNTF.Update
    'salvando tabela Nota Fiscal - Produtos
    If LB_IMP.Visible = True Then
        If DLL_BD.BDSIS_TBNFP.RecordCount > 0 Then
            DLL_BD.BDSIS_TBNFP.MoveFirst
            Do While Not DLL_BD.BDSIS_TBNFP.EOF
                If DLL_BD.BDSIS_TBNFP_CPNNF.Value = TXT_NF.Text Then
                    DLL_BD.BDSIS_TBNFP.Delete
                End If
                DLL_BD.BDSIS_TBNFP.MoveNext
            Loop
        End If
    End If
    For I = 1 To 20
        If FG_1.TextMatrix(I, 2) <> "" And _
           FG_1.TextMatrix(I, 3) <> "" And _
           FG_1.TextMatrix(I, 4) <> "" And _
           FG_1.TextMatrix(I, 5) <> "" And _
           FG_2.TextMatrix(I, 1) <> "" Then
            DLL_BD.BDSIS_TBNFP.AddNew
            DLL_BD.BDSIS_TBNFP_CPNNF.Value = TXT_NF.Text
            DLL_BD.BDSIS_TBNFP_CPFIG.Value = FG_1.TextMatrix(I, 1)
            DLL_BD.BDSIS_TBNFP_CPBIT.Value = FG_2.TextMatrix(I, 1)
            DLL_BD.BDSIS_TBNFP_CPMAT.Value = FG_2.TextMatrix(I, 2)
            DLL_BD.BDSIS_TBNFP_CPDES.Value = FG_1.TextMatrix(I, 2)
            DLL_BD.BDSIS_TBNFP_CPCCF.Value = FG_1.TextMatrix(I, 3)
            DLL_BD.BDSIS_TBNFP_CPCST.Value = FG_1.TextMatrix(I, 4)
            DLL_BD.BDSIS_TBNFP_CPUNI.Value = FG_1.TextMatrix(I, 5)
            DLL_BD.BDSIS_TBNFP_CPQUA.Value = FG_1.TextMatrix(I, 6)
            DLL_BD.BDSIS_TBNFP_CPPUN.Value = FG_1.TextMatrix(I, 7)
            DLL_BD.BDSIS_TBNFP_CPPTO.Value = FG_1.TextMatrix(I, 8)
            DLL_BD.BDSIS_TBNFP_CPAIC.Value = FG_1.TextMatrix(I, 9)
            DLL_BD.BDSIS_TBNFP_CPAIP.Value = FG_1.TextMatrix(I, 10)
            DLL_BD.BDSIS_TBNFP_CPVIP.Value = FG_1.TextMatrix(I, 11)
            DLL_BD.BDSIS_TBNFP_CPBCI.Value = FG_2.TextMatrix(I, 3)
            DLL_BD.BDSIS_TBNFP_CPVIC.Value = FG_2.TextMatrix(I, 4)
            DLL_BD.BDSIS_TBNFP_CPPPA.Value = FG_2.TextMatrix(I, 5)
            DLL_BD.BDSIS_TBNFP.Update
        ElseIf FG_1.TextMatrix(I, 2) <> "" And _
           FG_1.TextMatrix(I, 3) = "" And _
           FG_1.TextMatrix(I, 4) = "" And _
           FG_1.TextMatrix(I, 5) = "" And _
           FG_2.TextMatrix(I, 1) <> "" Then
            DLL_BD.BDSIS_TBNFP.AddNew
            DLL_BD.BDSIS_TBNFP_CPNNF.Value = TXT_NF.Text
            DLL_BD.BDSIS_TBNFP_CPFIG.Value = FG_1.TextMatrix(I, 1)
            Dim Y As Integer
            Dim CD As String
            CD = FG_1.TextMatrix(I, 2)
            For J = (I + 1) To 20
                If FG_2.TextMatrix(J, 1) = "Idem" And _
                   FG_1.TextMatrix(J, 3) = "" Then
                    CD = CD & " " & FG_1.TextMatrix(J, 2)
                ElseIf FG_2.TextMatrix(J, 1) = "Idem" And _
                   FG_1.TextMatrix(J, 3) <> "" Then
                    CD = CD & " " & FG_1.TextMatrix(J, 2)
                    Y = J
                    I = Y
                    Exit For
                End If
            Next J
            DLL_BD.BDSIS_TBNFP_CPDES.Value = CD
            DLL_BD.BDSIS_TBNFP_CPBIT.Value = FG_2.TextMatrix(Y, 1)
            DLL_BD.BDSIS_TBNFP_CPMAT.Value = FG_2.TextMatrix(Y, 2)
            DLL_BD.BDSIS_TBNFP_CPCCF.Value = FG_1.TextMatrix(Y, 3)
            DLL_BD.BDSIS_TBNFP_CPCST.Value = FG_1.TextMatrix(Y, 4)
            DLL_BD.BDSIS_TBNFP_CPUNI.Value = FG_1.TextMatrix(Y, 5)
            DLL_BD.BDSIS_TBNFP_CPQUA.Value = FG_1.TextMatrix(Y, 6)
            DLL_BD.BDSIS_TBNFP_CPPUN.Value = FG_1.TextMatrix(Y, 7)
            DLL_BD.BDSIS_TBNFP_CPPTO.Value = FG_1.TextMatrix(Y, 8)
            DLL_BD.BDSIS_TBNFP_CPAIC.Value = FG_1.TextMatrix(Y, 9)
            DLL_BD.BDSIS_TBNFP_CPAIP.Value = FG_1.TextMatrix(Y, 10)
            DLL_BD.BDSIS_TBNFP_CPVIP.Value = FG_1.TextMatrix(Y, 11)
            DLL_BD.BDSIS_TBNFP_CPBCI.Value = FG_2.TextMatrix(Y, 3)
            DLL_BD.BDSIS_TBNFP_CPVIC.Value = FG_2.TextMatrix(Y, 4)
            DLL_BD.BDSIS_TBNFP_CPPPA.Value = FG_2.TextMatrix(Y, 5)
            DLL_BD.BDSIS_TBNFP.Update
        End If
    Next I
    'salvando tabela Nota Fiscal - Declarações
    If LB_IMP.Visible = True Then
        If DLL_BD.BDSIS_TBNFD.RecordCount > 0 Then
            DLL_BD.BDSIS_TBNFD.MoveFirst
            Do While Not DLL_BD.BDSIS_TBNFD.EOF
                If DLL_BD.BDSIS_TBNFD_CPNNF.Value = TXT_NF.Text Then
                    DLL_BD.BDSIS_TBNFD.Delete
                End If
                DLL_BD.BDSIS_TBNFD.MoveNext
            Loop
        End If
    End If
    For I = 1 To 20
        If FG_2.TextMatrix(I, 1) = "DF" Then
            DLL_BD.BDSIS_TBNFD.AddNew
            DLL_BD.BDSIS_TBNFD_CPNNF.Value = TXT_NF.Text
            DLL_BD.BDSIS_TBNFD_CPDEC.Value = FG_1.TextMatrix(I, 1)
            DLL_BD.BDSIS_TBNFD_CPLIN.Value = Val(I)
            DLL_BD.BDSIS_TBNFD.Update
        End If
    Next I
    'salvando tabela Nota Fiscal - Bancos
    If LB_IMP.Visible = True Then
        If DLL_BD.BDSIS_TBNFB.RecordCount > 0 Then
            DLL_BD.BDSIS_TBNFB.MoveFirst
            Do While Not DLL_BD.BDSIS_TBNFB.EOF
                If DLL_BD.BDSIS_TBNFB_CPNNF.Value = TXT_NF.Text Then
                    DLL_BD.BDSIS_TBNFB.Delete
                End If
                DLL_BD.BDSIS_TBNFB.MoveNext
            Loop
        End If
    End If
    For I = 1 To 20
        If FG_2.TextMatrix(I, 1) = "DP" Then
            DLL_BD.BDSIS_TBNFB.AddNew
            DLL_BD.BDSIS_TBNFB_CPNNF.Value = TXT_NF.Text
            DLL_BD.BDSIS_TBNFB_CPCON.Value = FG_1.TextMatrix(I, 1)
            DLL_BD.BDSIS_TBNFB_CPLIN.Value = Val(I)
            DLL_BD.BDSIS_TBNFB.Update
        End If
    Next I
    'salvando tabela Nota Fiscal - Comentários
    If LB_IMP.Visible = True Then
        If DLL_BD.BDSIS_TBNFC.RecordCount > 0 Then
            DLL_BD.BDSIS_TBNFC.MoveFirst
            Do While Not DLL_BD.BDSIS_TBNFC.EOF
                If DLL_BD.BDSIS_TBNFC_CPNNF.Value = TXT_NF.Text Then
                    DLL_BD.BDSIS_TBNFC.Delete
                End If
                DLL_BD.BDSIS_TBNFC.MoveNext
            Loop
        End If
    End If
    For I = 1 To 20
        If FG_2.TextMatrix(I, 1) = "CT" Then
            DLL_BD.BDSIS_TBNFC.AddNew
            DLL_BD.BDSIS_TBNFC_CPNNF.Value = TXT_NF.Text
            DLL_BD.BDSIS_TBNFC_CPCOM.Value = FG_1.TextMatrix(I, 1)
            DLL_BD.BDSIS_TBNFC_CPLIN.Value = Val(I)
            DLL_BD.BDSIS_TBNFC.Update
        End If
    Next I
    'Finalizando
    IMG_2.Picture = LI.ListImages(3).Picture 'OK
    LB_2.FontBold = False
        
    '*********************************************
    '3-) Baixa Pedidos
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_3.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_3.Visible = True
    LB_3.FontBold = True
    'Verifica se Foi habilitado lista de pedidos
    If RD_Pedido.Value = True Then
        With DLL_BD
            'liquida os itens
            For J = 1 To (FG_P2N2.Rows - 1)
                .BDSIS_TBPIT.Seek "=", Val(FG_P2N2.TextMatrix(J, 7))
                If Not .BDSIS_TBPIT.NoMatch Then
                    .BDSIS_TBPIT.Edit
                    .BDSIS_TBPIT_CPLIQ.Value = True
                    .BDSIS_TBPIT.Update
                End If
            Next J
            'verifica se todos itens do pedido foram liquidados
            bPedLiq = True
            Dim dUltimoPedido As Double
            dUltimoPedido = 0
            For J = 1 To (FG_P2N2.Rows - 1)
                If dUltimoPedido <> CDbl(FG_P2N2.TextMatrix(J, 0)) Then
                    dUltimoPedido = CDbl(FG_P2N2.TextMatrix(J, 0))
                    bPedLiq = VerificaPedidoLiquidado(dUltimoPedido)
                    'liquida pedido
                    If bPedLiq = True Then
                        .BDSIS_TBPED.Seek "=", dUltimoPedido
                        If Not .BDSIS_TBPIT.NoMatch Then
                            .BDSIS_TBPED.Edit
                            .BDSIS_TBPED_CPLIQ.Value = True
                            .BDSIS_TBPED.Update
                        End If
                    End If
                End If
            Next J
        End With
    ElseIf RD_Inserir.Value = True Then
        IMG_3.Picture = LI.ListImages(1).Picture 'Cancelado
        LB_3.FontBold = False
    End If
    
    '*********************************************
    '4-) Lança NF Mapa FAT
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_4.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_4.Visible = True
    LB_4.FontBold = True
    If LB_IMP.Visible = True Then
        If DLL_BD.BDSIS_TBMFA.RecordCount > 0 Then
            DLL_BD.BDSIS_TBMFA.MoveFirst
            Do While Not DLL_BD.BDSIS_TBMFA.EOF
                If DLL_BD.BDSIS_TBMFA_CPNNF.Value = TXT_NF.Text Then
                    DLL_BD.BDSIS_TBMFA.Delete
                End If
                DLL_BD.BDSIS_TBMFA.MoveNext
            Loop
        End If
    End If
    DLL_BD.BDSIS_TBMFA.AddNew
    DLL_BD.BDSIS_TBMFA_CPNNF.Value = TXT_NF.Text
    DLL_BD.BDSIS_TBMFA_CPDAT.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
    DLL_BD.BDSIS_TBMFA_CPVAL.Value = TXT_ValorTotalNotaFiscal.Text
    DLL_BD.BDSIS_TBMFA_CPEMP.Value = TXT_Apelido.Text
    DLL_BD.BDSIS_TBMFA_CPDEV.Value = False
    DLL_BD.BDSIS_TBMFA.Update
    IMG_4.Picture = LI.ListImages(3).Picture 'OK
    LB_4.FontBold = False
    
    '*********************************************
    '5-) Baixa produtos no banco dados estoque
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_5.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_5.Visible = True
    LB_5.FontBold = True
    If LB_IMP.Visible = False Then
        For I = 1 To 20
            If FG_1.TextMatrix(I, 1) <> "" Then
                If FG_1.TextMatrix(I, 3) <> "" Then 'Uma linha
                    DLL_BD.BDSIS_TBEST.Seek "=", FG_1.TextMatrix(I, 1), FG_2.TextMatrix(I, 1), FG_2.TextMatrix(I, 2)
                    If Not DLL_BD.BDSIS_TBEST.NoMatch Then 'Achou a ficha
                        DLL_BD.BDSIS_TBEST.Edit
                        DLL_BD.BDSIS_TBEST_CPEST.Value = DLL_BD.BDSIS_TBEST_CPEST.Value - CDbl(FG_1.TextMatrix(I, 6))
                        DLL_BD.BDSIS_TBEST.Update
                    End If
                ElseIf FG_1.TextMatrix(I, 3) = "" Then 'Multiplas linhas
                    For J = (I + 1) To 20
                        If FG_2.TextMatrix(J, 1) = "Idem" And _
                           FG_1.TextMatrix(J, 3) <> "" Then
                            Y = J
                            Exit For
                        End If
                    Next J
                    DLL_BD.BDSIS_TBEST.Seek "=", FG_1.TextMatrix(I, 1), FG_2.TextMatrix(I, 1), FG_2.TextMatrix(Y, 2)
                    If Not DLL_BD.BDSIS_TBEST.NoMatch Then 'Achou a ficha
                        DLL_BD.BDSIS_TBEST.Edit
                        DLL_BD.BDSIS_TBEST_CPEST.Value = DLL_BD.BDSIS_TBEST_CPEST.Value - CDbl(FG_1.TextMatrix(Y, 6))
                        DLL_BD.BDSIS_TBEST.Update
                    End If
                End If
            End If
        Next I
    End If
    IMG_5.Picture = LI.ListImages(3).Picture 'OK
    LB_5.FontBold = False
    
    '*********************************************
    '6-) Lança NF Mapa Contas à receber
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_6.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_6.Visible = True
    LB_6.FontBold = True
    If LB_IMP.Visible = True Then
        If DLL_BD.BDSIS_TBMCR.RecordCount > 0 Then
            DLL_BD.BDSIS_TBMCR.MoveFirst
            Do While Not DLL_BD.BDSIS_TBMCR.EOF
                If DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text Then
                    DLL_BD.BDSIS_TBMCR.Delete
                End If
                DLL_BD.BDSIS_TBMCR.MoveNext
            Loop
        End If
        If CK_Desdobrar.Value = 0 Then '1 DP
            DLL_BD.BDSIS_TBMCR.AddNew
            DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
            DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
            DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text
            DLL_BD.BDSIS_TBMCR_CPVAL.Value = TXT_ValorTotalNotaFiscal.Text
            DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_A.Text
            DLL_BD.BDSIS_TBMCR.Update
        Else
            If TXT_DataVenc_A.Text <> Space(10) And _
               TXT_DataVenc_B.Text <> Space(10) And _
               TXT_DataVenc_C.Text = Space(10) And _
               TXT_DataVenc_D.Text = Space(10) Then
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/A"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 2
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_A.Text
                DLL_BD.BDSIS_TBMCR.Update
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/B"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 2
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_B.Text
                DLL_BD.BDSIS_TBMCR.Update
            ElseIf TXT_DataVenc_A.Text <> Space(10) And _
               TXT_DataVenc_B.Text <> Space(10) And _
               TXT_DataVenc_C.Text <> Space(10) And _
               TXT_DataVenc_D.Text = Space(10) Then
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/A"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 3
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_A.Text
                DLL_BD.BDSIS_TBMCR.Update
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/B"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 3
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_B.Text
                DLL_BD.BDSIS_TBMCR.Update
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/C"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 3
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_C.Text
                DLL_BD.BDSIS_TBMCR.Update
            ElseIf TXT_DataVenc_A.Text <> Space(10) And _
               TXT_DataVenc_B.Text <> Space(10) And _
               TXT_DataVenc_C.Text <> Space(10) And _
               TXT_DataVenc_D.Text <> Space(10) Then
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/A"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_A.Text
                DLL_BD.BDSIS_TBMCR.Update
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/B"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_B.Text
                DLL_BD.BDSIS_TBMCR.Update
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/C"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_C.Text
                DLL_BD.BDSIS_TBMCR.Update
                DLL_BD.BDSIS_TBMCR.AddNew
                DLL_BD.BDSIS_TBMCR_CPDEM.Value = Format(TXT_DataEmissao.Text, "dd/mm/yyyy")
                DLL_BD.BDSIS_TBMCR_CPEMP.Value = TXT_Apelido.Text
                DLL_BD.BDSIS_TBMCR_CPNDP.Value = TXT_NF.Text & "/D"
                DLL_BD.BDSIS_TBMCR_CPVAL.Value = CDbl(TXT_ValorTotalNotaFiscal.Text) / 4
                DLL_BD.BDSIS_TBMCR_CPDVE.Value = TXT_DataVenc_D.Text
                DLL_BD.BDSIS_TBMCR.Update
            End If
        End If
    End If
    IMG_6.Picture = LI.ListImages(3).Picture 'OK
    LB_6.FontBold = False
        
    '*********************************************
    '7-) Lança NF Mapa Impostos
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_7.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_7.Visible = True
    LB_7.FontBold = True
    'ICMS
    DLL_BD.BDSIS_TBMIM.Seek "=", "ICMS", Month(Date), Year(Date)
    If DLL_BD.BDSIS_TBMIM.NoMatch Then
        DLL_BD.BDSIS_TBMIM.AddNew
        DLL_BD.BDSIS_TBMIM_CPIMP.Value = "ICMS"
        DLL_BD.BDSIS_TBMIM_CPMES.Value = Month(Date)
        DLL_BD.BDSIS_TBMIM_CPANO.Value = Year(Date)
        DLL_BD.BDSIS_TBMIM_CPVAL.Value = TXT_ValorICMS.Text
    Else
        DLL_BD.BDSIS_TBMIM.Edit
        DLL_BD.BDSIS_TBMIM_CPVAL.Value = CDbl(DLL_BD.BDSIS_TBMIM_CPVAL.Value) + CDbl(TXT_ValorICMS.Text)
    End If
    DLL_BD.BDSIS_TBMIM.Update
    'IPI
    DLL_BD.BDSIS_TBMIM.Seek "=", "IPI", Month(Date), Year(Date)
    If DLL_BD.BDSIS_TBMIM.NoMatch Then
        DLL_BD.BDSIS_TBMIM.AddNew
        DLL_BD.BDSIS_TBMIM_CPIMP.Value = "IPI"
        DLL_BD.BDSIS_TBMIM_CPMES.Value = Month(Date)
        DLL_BD.BDSIS_TBMIM_CPANO.Value = Year(Date)
        DLL_BD.BDSIS_TBMIM_CPVAL.Value = TXT_ValorTotalIPI.Text
    Else
        DLL_BD.BDSIS_TBMIM.Edit
        DLL_BD.BDSIS_TBMIM_CPVAL.Value = DLL_BD.BDSIS_TBMIM_CPVAL.Value + CDbl(TXT_ValorTotalIPI.Text)
    End If
    DLL_BD.BDSIS_TBMIM.Update
    IMG_7.Picture = LI.ListImages(3).Picture 'OK
    LB_7.FontBold = False
    
    '*********************************************
    '8-) Monta Certificado de Qualidade
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_8.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_8.Visible = True
    LB_8.FontBold = True
    If RB_CF2.Value = True Then 'Emite CQ
    
    ElseIf RB_CF1.Value = True Then
        IMG_8.Picture = LI.ListImages(1).Picture 'Cancelado
        LB_8.FontBold = False
    End If
    
    '*********************************************
    '9-) Imprime Certificado de Qualidade
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_9.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_9.Visible = True
    LB_9.FontBold = True
    If RB_CF2.Value = True Then 'Imprime CQ
    
    ElseIf RB_CF1.Value = True Then
        IMG_9.Picture = LI.ListImages(1).Picture 'Cancelado
        LB_9.FontBold = False
    End If
    
    '*********************************************
    '10-) Monta Nota Fiscal
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_10.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_10.Visible = True
    LB_10.FontBold = True
    'Envia dados para NF
    With DLL_IMP
        .NotaFiscal_LimpaItens
        If RD_Saida.Value = True Then
            If .NotaFiscal_MontaNF("TIPO_SAIDA", "X") = False Then Exit Sub
        ElseIf RD_Entrada.Value = True Then
            If .NotaFiscal_MontaNF("TIPO_ENTRADA", "X") = False Then Exit Sub
        End If
        If .NotaFiscal_MontaNF("NO", CB_Natureza.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("CFOP", CB_CFOP.Text) = False Then Exit Sub
        'Destinatário
        If .NotaFiscal_MontaNF("RAZAO", TXT_Empresa.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("CGC", TXT_CGC.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("DATA_EMISSAO", Format(TXT_DataEmissao.Text, "dd/mm/yyyy")) = False Then Exit Sub
        If .NotaFiscal_MontaNF("ENDERECO", TXT_Endereco.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("BAIRRO", TXT_Bairro.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("CEP", TXT_CEP.Text) = False Then Exit Sub
        If CK_DataSaida.Value = 1 Then
            If .NotaFiscal_MontaNF("DATA_SAIDA", Format(TXT_DataSaida.Text, "dd/mm/yyyy")) = False Then Exit Sub
        End If
        If TXT_HoraSaida.Text <> "" Then
            If .NotaFiscal_MontaNF("HORA_SAIDA", Format(TXT_DataSaida.Text, "hh:mm:ss")) = False Then Exit Sub
        End If
        If .NotaFiscal_MontaNF("MUNICIPIO", TXT_Cidade.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("FONE", TXT_Fone.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("ESTADO", CB_Estado.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("INS_EST", TXT_InsEst.Text) = False Then Exit Sub
        'Fatura
        If .NotaFiscal_MontaNF("DATA_EMISSAO_FATURA", Format(TXT_DataEmissao.Text, "dd/mm/yyyy")) = False Then Exit Sub
        If .NotaFiscal_MontaNF("NUM_NOTAFISCAL", TXT_NF.Text) = False Then Exit Sub
        If .NotaFiscal_MontaNF("VALOR_FATURA", Format(TXT_ValorTotalNotaFiscal.Text, "###,###,###,##0.00")) = False Then Exit Sub
        If CK_Desdobrar.Value = 0 Then
            If .NotaFiscal_MontaNF("NUM_DUPLICATA", TXT_NF.Text) = False Then Exit Sub
            If .NotaFiscal_MontaNF("DATA_VENCIMENTO", TXT_DataVenc_A.Text) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("NUM_DUPLICATA", "Vide abaixo") = False Then Exit Sub
            If .NotaFiscal_MontaNF("DATA_VENCIMENTO", "Vide abaixo") = False Then Exit Sub
        End If
        If TXT_PracaPagamento.Text <> "" Then
            If .NotaFiscal_MontaNF("PRACA_PAGAMENTO", TXT_PracaPagamento.Text) = False Then Exit Sub
        Else
            If RD_Saida.Value = True Then
                If .NotaFiscal_MontaNF("PRACA_PAGAMENTO", "O mesmo acima") = False Then Exit Sub
            End If
        End If
        Dim sExtenso As String
        sExtenso = "(" & DLL_FUNCS.ValorExtenso(TXT_ValorTotalNotaFiscal.Text) & ")" & DLL_FUNCS.MultiString("x ", 400)
        If .NotaFiscal_MontaNF("VALOR_EXTENSO_1", Left(Mid(sExtenso, 1, 120), 120)) = False Then Exit Sub
        If .NotaFiscal_MontaNF("VALOR_EXTENSO_2", Left(Mid(sExtenso, 121, 240), 120)) = False Then Exit Sub
        If TXT_PedidoInterno.Text <> "" Then
            If .NotaFiscal_MontaNF("PI", TXT_PedidoInterno.Text) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("PI", "-") = False Then Exit Sub
        End If
        If TXT_SeuPedido.Text <> "" Then
            If .NotaFiscal_MontaNF("SEU_PEDIDO", TXT_SeuPedido.Text) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("SEU_PEDIDO", "-") = False Then Exit Sub
        End If
        If TXT_Operacao.Text <> "" Then
            If .NotaFiscal_MontaNF("OP", TXT_Operacao.Text) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("OP", "-") = False Then Exit Sub
        End If
        'Dados dos Produtos
        Dim cDBTipoNF As String
        Dim cNumDBTipoNF As String
        Dim pCF_I As Boolean, pCF_J As Boolean
        pCF_I = False
        pCF_J = True
        cDBTipoNF = ""
        cNumDBTipoNF = ""
        For I = 1 To 20
            'Verifica o número da linha p/item
            If I < 10 Then
                cNumDBTipoNF = "0" & Trim(Str(I))
            Else
                cNumDBTipoNF = Trim(Str(I))
            End If
            'Aqui e para linhas de produtos
            If FG_2.TextMatrix(I, 1) <> "" And _
               FG_2.TextMatrix(I, 1) <> "CT" And _
               FG_2.TextMatrix(I, 1) <> "DF" And _
               FG_2.TextMatrix(I, 1) <> "DP" Then
                'Figura
                If FG_1.TextMatrix(I, 1) <> "" Then
                    cDBTipoNF = "FIGURA_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 1)) = False Then Exit Sub
                End If
                'Descrição
                If FG_1.TextMatrix(I, 2) <> "" Then
                    cDBTipoNF = "DESCRICAO_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 2)) = False Then Exit Sub
                End If
                'CF
                If FG_1.TextMatrix(I, 3) <> "" Then
                    cDBTipoNF = "CF_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 3)) = False Then Exit Sub
                End If
                'ST
                If FG_1.TextMatrix(I, 4) <> "" Then
                    cDBTipoNF = "ST_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 4)) = False Then Exit Sub
                End If
                'Unidade
                If FG_1.TextMatrix(I, 5) <> "" Then
                    cDBTipoNF = "UNIDADE_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 5)) = False Then Exit Sub
                End If
                'Quantidade
                If FG_1.TextMatrix(I, 6) <> "" Then
                    cDBTipoNF = "QUANTIDADE_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 6)) = False Then Exit Sub
                End If
                'Preço Unitário
                If FG_1.TextMatrix(I, 7) <> "" Then
                    cDBTipoNF = "PRECO_UNITARIO_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 7)) = False Then Exit Sub
                End If
                'Preço Total
                If FG_1.TextMatrix(I, 8) <> "" Then
                    cDBTipoNF = "PRECO_TOTAL_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 8)) = False Then Exit Sub
                End If
                'Aliquota ICMS
                If FG_1.TextMatrix(I, 9) <> "" Then
                    cDBTipoNF = "ALIQUOTA_ICMS_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 9)) = False Then Exit Sub
                End If
                'Aliquota IPI
                If FG_1.TextMatrix(I, 10) <> "" Then
                    cDBTipoNF = "ALIQUOTA_IPI_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 10)) = False Then Exit Sub
                End If
                'Valor IPI
                If FG_1.TextMatrix(I, 11) <> "" Then
                    cDBTipoNF = "VALOR_IPI_" & cNumDBTipoNF
                    If .NotaFiscal_MontaNF(cDBTipoNF, FG_1.TextMatrix(I, 11)) = False Then Exit Sub
                End If
            ElseIf FG_2.TextMatrix(I, 1) = "DF" Then 'Se for Declaracao Fiscal
                If FG_1.TextMatrix(I, 1) <> "" Then
                    If .NotaFiscal_MontaNF("DECLARACAO", FG_1.TextMatrix(I, 1), Val(I)) = False Then Exit Sub
                End If
            ElseIf FG_2.TextMatrix(I, 1) = "CT" Then 'Se for Comentarios
                If FG_1.TextMatrix(I, 1) <> "" Then
                    If .NotaFiscal_MontaNF("COMENTARIO", FG_1.TextMatrix(I, 1), Val(I)) = False Then Exit Sub
                End If
            ElseIf FG_2.TextMatrix(I, 1) = "DP" Then 'Se for Bancos
                If FG_1.TextMatrix(I, 1) <> "" Then
                    If .NotaFiscal_MontaNF("BANCO", FG_1.TextMatrix(I, 1), Val(I)) = False Then Exit Sub
                End If
            End If
            'Verifica se existe CF I e J
            If FG_1.TextMatrix(I, 3) = "I" Then
                pCF_I = True
            ElseIf FG_1.TextMatrix(I, 3) = "J" Then
                pCF_J = True
            End If
        Next I
        'Cálculo do Imposto
        If TXT_BaseICMS.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("BASE_CALCULO_ICMS", Format(TXT_BaseICMS.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("BASE_CALCULO_ICMS", "-.-") = False Then Exit Sub
        End If
        If TXT_ValorICMS.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("VALOR_ICMS", Format(TXT_ValorICMS.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("VALOR_ICMS", "-.-") = False Then Exit Sub
        End If
        If TXT_BaseICMSSub.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("BASE_CALCULO_ICMSSUB", Format(TXT_BaseICMSSub.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("BASE_CALCULO_ICMSSUB", "-.-") = False Then Exit Sub
        End If
        If TXT_ValorICMSSub.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("VALOR_ICMSSUB", Format(TXT_ValorICMSSub.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("VALOR_ICMSSUB", "-.-") = False Then Exit Sub
        End If
        If TXT_ValorTotalProdutos.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("VALOR_TOTAL_PRODUTOS", Format(TXT_ValorTotalProdutos.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("VALOR_TOTAL_PRODUTOS", "-.-") = False Then Exit Sub
        End If
        If TXT_ValorFrete.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("VALOR_FRETE", Format(TXT_ValorFrete.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("VALOR_FRETE", "-.-") = False Then Exit Sub
        End If
        If TXT_ValorSeguro.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("VALOR_SEGURO", Format(TXT_ValorSeguro.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("VALOR_SEGURO", "-.-") = False Then Exit Sub
        End If
        If TXT_Outras.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("OUTRAS_DESPESAS", Format(TXT_Outras.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("OUTRAS_DESPESAS", "-.-") = False Then Exit Sub
        End If
        If TXT_ValorTotalIPI.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("VALOR_TOTAL_IPI", Format(TXT_ValorTotalIPI.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("VALOR_TOTAL_IPI", "-.-") = False Then Exit Sub
        End If
        If TXT_ValorTotalNotaFiscal.Text <> "0,00" Then
            If .NotaFiscal_MontaNF("VALOR_TOTAL_NOTA", Format(TXT_ValorTotalNotaFiscal.Text, "###,###,###,##0.00")) = False Then Exit Sub
        Else
            If .NotaFiscal_MontaNF("VALOR_TOTAL_NOTA", "-.-") = False Then Exit Sub
        End If
        'Transportador
        If .NotaFiscal_MontaNF("TRANSPORTADORA_NOME", TXT_Trans.Text) = False Then Exit Sub
        If CB_Frete.Text = "Remetente" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_FRETE", "1") = False Then Exit Sub
        ElseIf CB_Frete.Text = "Destinatário" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_FRETE", "2") = False Then Exit Sub
        End If
        If CB_Placa.Text <> "" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_PLACA", CB_Placa.Text) = False Then Exit Sub
        End If
        If CB_EstVei.Text <> "" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_UFVEI", CB_EstVei.Text) = False Then Exit Sub
        End If
        If TXT_CGCTrans.Text <> "" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_CGC", TXT_CGCTrans.Text) = False Then Exit Sub
        End If
        If TXT_EndTrans.Text <> "" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_END", TXT_EndTrans.Text) = False Then Exit Sub
        End If
        If TXT_CidTrans.Text <> "" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_MUN", TXT_CidTrans.Text) = False Then Exit Sub
        End If
        If CB_EstTrans.Text <> "" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_UF", CB_EstTrans.Text) = False Then Exit Sub
        End If
        If TXT_IETrans.Text <> "" Then
            If .NotaFiscal_MontaNF("TRANSPORTADORA_IE", TXT_IETrans.Text) = False Then Exit Sub
        End If
        If TXT_QuantVol.Text <> "" Then
            If .NotaFiscal_MontaNF("VOLUMES_QUANTIDADE", CDbl(TXT_QuantVol.Text)) = False Then Exit Sub
        End If
        If CB_Especie.Text <> "" Then
            If .NotaFiscal_MontaNF("VOLUMES_ESPECIE", CB_Especie.Text) = False Then Exit Sub
        End If
        If CB_Marca.Text <> "" Then
            If .NotaFiscal_MontaNF("VOLUMES_MARCA", CB_Marca.Text) = False Then Exit Sub
        End If
        If TXT_NumVol.Text <> "" Then
            If .NotaFiscal_MontaNF("VOLUMES_NUMERO", TXT_NumVol.Text) = False Then Exit Sub
        End If
        If TXT_PesoBruto.Text <> "" Then
            If .NotaFiscal_MontaNF("VOLUMES_PESOBRUTO", Format(CDbl(TXT_PesoBruto.Text), "###,###,###,##0.00")) = False Then Exit Sub
        End If
        If TXT_PesoLiquido.Text <> "" Then
            If .NotaFiscal_MontaNF("VOLUMES_PESOLIQUIDO", Format(CDbl(TXT_PesoLiquido.Text), "###,###,###,##0.00")) = False Then Exit Sub
        End If
        'Dados Adicionais
        If pCF_I = True Then
            If .NotaFiscal_MontaNF("CF_I", CF_I) = False Then Exit Sub
        End If
        If pCF_J = True Then
            If .NotaFiscal_MontaNF("CF_J", CF_J) = False Then Exit Sub
        End If
        If TXT_VendInterno.Text <> "" Then
            If .NotaFiscal_MontaNF("VEND_INT", TXT_VendInterno.Text) = False Then Exit Sub
        End If
        If TXT_VendExterno.Text <> "" Then
            If .NotaFiscal_MontaNF("VEND_EXT", TXT_VendExterno.Text) = False Then Exit Sub
        End If
        If TXT_Setor.Text <> "" Then
            If .NotaFiscal_MontaNF("VEND_SETOR", TXT_Setor.Text) = False Then Exit Sub
        End If
        If CK_Desdobrar.Value = 1 Then 'é desdobrado
            If TXT_DataVenc_A.Text <> Space(10) And _
               TXT_DataVenc_B.Text <> Space(10) And _
               TXT_DataVenc_C.Text = Space(10) And _
               TXT_DataVenc_D.Text = Space(10) Then
                If .NotaFiscal_MontaNF("DESD_DUP_A_VENC", Format(TXT_DataVenc_A.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_B_VENC", Format(TXT_DataVenc_B.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_C_VENC", "-x-") = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_D_VENC", "-x-") = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_A_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 2), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_B_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 2), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_C_VALOR", "-x-") = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_D_VALOR", "-X-") = False Then Exit Sub
            ElseIf TXT_DataVenc_A.Text <> Space(10) And _
               TXT_DataVenc_B.Text <> Space(10) And _
               TXT_DataVenc_C.Text <> Space(10) And _
               TXT_DataVenc_D.Text = Space(10) Then
                If .NotaFiscal_MontaNF("DESD_DUP_A_VENC", Format(TXT_DataVenc_A.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_B_VENC", Format(TXT_DataVenc_B.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_C_VENC", Format(TXT_DataVenc_C.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_D_VENC", "-x-") = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_A_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 3), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_B_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 3), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_C_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 3), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_D_VALOR", "-x-") = False Then Exit Sub
            ElseIf TXT_DataVenc_A.Text <> Space(10) And _
               TXT_DataVenc_B.Text <> Space(10) And _
               TXT_DataVenc_C.Text <> Space(10) And _
               TXT_DataVenc_D.Text <> Space(10) Then
                If .NotaFiscal_MontaNF("DESD_DUP_A_VENC", Format(TXT_DataVenc_A.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_B_VENC", Format(TXT_DataVenc_B.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_C_VENC", Format(TXT_DataVenc_C.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_D_VENC", Format(TXT_DataVenc_D.Text, "dd/mm/yyyy")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_A_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 4), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_B_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 4), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_C_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 4), "###,###,###,##0.00")) = False Then Exit Sub
                If .NotaFiscal_MontaNF("DESD_DUP_D_VALOR", Format((CDbl(TXT_ValorTotalNotaFiscal.Text) / 4), "###,###,###,##0.00")) = False Then Exit Sub
            End If
        Else 'Nao e desdobrado
            If .NotaFiscal_MontaNF("DESD_DUP_A_VENC", TXT_DataVenc_A.Text) = False Then Exit Sub
            If .NotaFiscal_MontaNF("DESD_DUP_B_VENC", "-x-") = False Then Exit Sub
            If .NotaFiscal_MontaNF("DESD_DUP_C_VENC", "-x-") = False Then Exit Sub
            If .NotaFiscal_MontaNF("DESD_DUP_D_VENC", "-x-") = False Then Exit Sub
            If IsDate(TXT_DataVenc_A.Text) Then
                If .NotaFiscal_MontaNF("DESD_DUP_A_VALOR", Format(TXT_ValorTotalNotaFiscal.Text, "###,###,###,##0.00")) = False Then Exit Sub
            Else
                If .NotaFiscal_MontaNF("DESD_DUP_A_VALOR", "-x-") = False Then Exit Sub
            End If
            If .NotaFiscal_MontaNF("DESD_DUP_B_VALOR", "-x-") = False Then Exit Sub
            If .NotaFiscal_MontaNF("DESD_DUP_C_VALOR", "-x-") = False Then Exit Sub
            If .NotaFiscal_MontaNF("DESD_DUP_D_VALOR", "-x-") = False Then Exit Sub
        End If
        'Declaracao do Simples
        'If .NotaFiscal_MontaNF("SIMPLES", "EMITENTE: Empresa Optante Pelo Simples.") = False Then Exit Sub
    End With
    'Finalizando
    IMG_10.Picture = LI.ListImages(3).Picture 'OK
    LB_10.FontBold = False
    
    '*********************************************
    '11-) Imprime Nota Fiscal
    '*********************************************
    BP.Value = BP.Value + 1
    IMG_11.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_11.Visible = True
    LB_11.FontBold = True
    'O Objeto Printer usa medidas em twips
    ' 1 cm = 567 twips
    ' FORMULARIO DA NF: 330 mm alt X 215 mm larg
    DLL_IMP.NotaFiscal_Imprimir (DLL_FUNCS.NomeImpressora("IT_NotaFiscal"))
    IMG_11.Picture = LI.ListImages(3).Picture 'OK
    LB_11.FontBold = False
    
    '*********************************************
    '12-) Finalizando
    '*********************************************
    DLL_FUNCS.RegistraEvento "Imprimir Nota Fiscal", TXT_NF.Text
    BP.Value = BP.Value + 1
    IMG_12.Picture = LI.ListImages(2).Picture 'Verificando
    IMG_12.Visible = True
    LB_12.FontBold = True
    IMG_12.Picture = LI.ListImages(3).Picture 'OK
    LB_12.FontBold = False
    
    TelaNFEmEspera (False)
    Unload Tela_NotaFiscal
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_DeclaracoesFiscais_Click()
    On Error GoTo ERRO_SISCOVAL
    'Verifica se existem itens na NF
    For I = 1 To 20
        If FG_1.TextMatrix(I, 1) = "" And _
           FG_1.TextMatrix(I, 2) = "" Then
            NumLinha = I
            If NumLinha = 20 Then
                RespMsg = MsgBox("Não existem ítens para serem declarados.", vbOKOnly, Tela_NotaFiscal.Caption)
                Exit Sub
            End If
        Else
            Exit For
        End If
    Next I
    
    'Se já existe a declaraçao
    If BT_DeclaracoesFiscais.Caption = "Remover Dec.Fiscais" Then
        For I = 1 To 20
            If FG_2.TextMatrix(I, 1) = "DF" Then
                FG_1.TextMatrix(I, 1) = ""
                FG_1.TextMatrix(I, 2) = ""
                FG_2.TextMatrix(I, 1) = ""
                FG_2.TextMatrix(I, 2) = ""
            End If
        Next I
        BT_DeclaracoesFiscais.Caption = "Incluir Dec.Fiscais"
        Exit Sub
    'Incluir declarações
    ElseIf BT_DeclaracoesFiscais.Caption = "Incluir Dec.Fiscais" Then
        'verfica se existem linhas em branco
        For I = 1 To 20
            If FG_1.TextMatrix(I, 1) = "" And _
               FG_1.TextMatrix(I, 2) = "" Then
                Exit For
            ElseIf I = 20 Then
                RespMsg = MsgBox("Não existem mais linha em branco para inserir declarações fiscais das peças.", vbOKOnly, Tela_NotaFiscal.Caption)
                Exit Sub
            End If
        Next I
        TelaNFEmEspera (True)
        Do
            Dim Decs_NF As String
            Decs_NF = ""
            'Verifica as classificações ficais da nota
            For I = 1 To 20
                    If FG_1.TextMatrix(I, 3) <> "" Then
                        Decs_NF = Decs_NF & Trim(FG_1.TextMatrix(I, 3))
                    End If
            Next I
            'Verifica as declaracoes das CF
            'Funcao inseredesnf grava dados na nota
            Dim NumItem As Integer
            NumItem = 0
            For K = 1 To Len(Decs_NF)
                NumItem = K
                DLL_BD.BDSIS_TBEDC.Seek "=", Mid(Decs_NF, K, 1)
                If DLL_BD.BDSIS_TBEDC.NoMatch Then
                    RespMsg = MsgBox("Ocorreu algum erro durante a procura das declarações de CF no banco de dados...", vbOKOnly, Tela_NotaFiscal.Caption)
                Else
                    If InsereDecsNF(DLL_BD.BDSIS_TBEDC_CPDOU.Value, NumItem) = False Then Exit Do
                    'Verifica se existem declaraçoes para itens qd.cliente for em SP
                    If CB_Estado.Text = "SP" Then
                        If InsereDecsNF(DLL_BD.BDSIS_TBEDC_CPDSP.Value, NumItem) = False Then Exit Do
                    End If
                End If
            Next K
            'Organiza declaracoes
            For I = 1 To 20
                If FG_2.TextMatrix(I, 1) = "DF" Then
                    If Len(FG_1.TextMatrix(I, 1)) < 2 Then
                        FG_1.TextMatrix(I, 1) = "Ítem " & FG_1.TextMatrix(I, 1) & ": " & FG_1.TextMatrix(I, 2)
                    Else
                        FG_1.TextMatrix(I, 1) = "Ítens " & FG_1.TextMatrix(I, 1) & ": " & FG_1.TextMatrix(I, 2)
                    End If
                    FG_1.TextMatrix(I, 2) = ""
                End If
            Next I
            BT_DeclaracoesFiscais.Caption = "Remover Dec.Fiscais"
            Exit Do
        Loop
    End If
    TelaNFEmEspera (False)
    'Verifica se foi inserido declarações na NF
    For I = 1 To 20
        If FG_2.TextMatrix(I, 1) = "DF" And _
           FG_1.TextMatrix(I, 2) <> "" Then
           'Este IF verifica se houve algum EXIT DO (acima)
           FG_1.TextMatrix(I, 1) = ""
           FG_1.TextMatrix(I, 2) = ""
           FG_2.TextMatrix(I, 1) = ""
        ElseIf FG_2.TextMatrix(I, 1) = "DF" Then
            Exit For
        ElseIf I = 20 Then
            RespMsg = MsgBox("De todas as peças que estão digitadas no Assistente de Nota Fiscal, nenhuma tem declarações fiscais. Insira outras peças e tente incluir as declarações novamente.", vbOKOnly, Tela_NotaFiscal.Caption)
            BT_DeclaracoesFiscais.Value = True
            Exit Sub
        End If
    Next I
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_DepositoBancario_Click()
    On Error GoTo ERRO_SISCOVAL
    'Se já existe a C/C
    If BT_DepositoBancario.Caption = "Remover C/C Banco" Then
        For I = 20 To 1 Step -1
            If Tela_NotaFiscal.FG_1.TextMatrix(I, 1) <> "" And _
               Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = "DP" Then
                Tela_NotaFiscal.FG_1.TextMatrix(I, 1) = ""
                Tela_NotaFiscal.FG_2.TextMatrix(I, 1) = ""
            End If
        Next I
        BT_DepositoBancario.Caption = "Incluir C/C Banco"
    
    'Incluir C/C
    ElseIf BT_DepositoBancario.Caption = "Incluir C/C Banco" Then
        'verfica se existem linhas em branco
         For I = 1 To 20
             If FG_1.TextMatrix(I, 1) = "" And _
                FG_1.TextMatrix(I, 2) = "" Then
                 Exit For
             ElseIf I = 20 Then
                 RespMsg = MsgBox("Não existem mais linha em branco para inserir dados para depósito bancário em conta corrente.", vbOKOnly, Tela_NotaFiscal.Caption)
                 Exit Sub
             End If
         Next I
         BT_DepositoBancario.Caption = "Remover C/C Banco"
         Tela_NotaFiscal_Dlg_3.Show vbModal
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_EditarItem_Click()
    On Error GoTo ERRO_SISCOVAL
    'Verifica se existem itens na NF
    For I = 1 To 20
        If FG_1.TextMatrix(I, 2) <> "" And _
           FG_2.TextMatrix(I, 1) <> "" Then
            Exit For
        ElseIf I = 20 Then
            MsgBox "Não existem ítens à serem editados.", vbOKOnly, Tela_NotaFiscal.Caption
            Exit Sub
        End If
    Next I
    Tela_NotaFiscal_Dlg_4.Show vbModal
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_II_Click()
    'importa item do pedido
    If FG_PED.RowSel < 1 Then Exit Sub
    'FG_P2N
    FG_P2N.AddItem (FG_P2N.Rows)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 0) = (FG_P2N.Rows - 1)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 1) = FG_PED.TextMatrix((FG_PED.RowSel), 1)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 2) = FG_PED.TextMatrix((FG_PED.RowSel), 2)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 3) = FG_PED.TextMatrix((FG_PED.RowSel), 3)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 4) = FG_PED.TextMatrix((FG_PED.RowSel), 4)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 5) = FG_PED.TextMatrix((FG_PED.RowSel), 5)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 6) = FG_PED.TextMatrix((FG_PED.RowSel), 6)
    FG_P2N.TextMatrix((FG_P2N.Rows - 1), 7) = FG_PED.TextMatrix((FG_PED.RowSel), 7)
    'FG_P2N2
    FG_P2N2.AddItem (FG_P2N2.Rows)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 0) = FG_PED2.TextMatrix((FG_PED.RowSel), 0)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 1) = FG_PED2.TextMatrix((FG_PED.RowSel), 1)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 2) = FG_PED2.TextMatrix((FG_PED.RowSel), 2)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 3) = FG_PED2.TextMatrix((FG_PED.RowSel), 3)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 4) = FG_PED2.TextMatrix((FG_PED.RowSel), 4)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 5) = FG_PED2.TextMatrix((FG_PED.RowSel), 5)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 6) = FG_PED2.TextMatrix((FG_PED.RowSel), 6)
    FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 7) = FG_PED2.TextMatrix((FG_PED.RowSel), 7)
End Sub
Private Sub BT_InserirItem_Click()
    On Error GoTo ERRO_SISCOVAL
    'verfica se existem linhas em branco
    For I = 1 To 20
        If FG_1.TextMatrix(I, 1) = "" And _
           FG_1.TextMatrix(I, 2) = "" Then
            Exit For
        ElseIf I = 20 Then
            RespMsg = MsgBox("Não existem mais linha em branco para serem preenchidas no quadro dados do produto na nota fiscal.", vbOKOnly, Tela_NotaFiscal.Caption)
            Exit Sub
        End If
    Next I
    Tela_NotaFiscal_Dlg_1.Show vbModal
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_InserirManual_Click()
    On Error GoTo ERRO_SISCOVAL
    'verfica se existem linhas em branco
    For I = 1 To 20
        If FG_1.TextMatrix(I, 1) = "" And _
           FG_1.TextMatrix(I, 2) = "" Then
            Exit For
        ElseIf I = 20 Then
            RespMsg = MsgBox("Não existem mais linha em branco para serem preenchidas no quadro dados do produto na nota fiscal.", vbOKOnly, Tela_NotaFiscal.Caption)
            Exit Sub
        End If
    Next I
    Tela_NotaFiscal_Dlg_2.Show vbModal
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_IT_Click()
    'importa todo pedido
    If FG_PED.Rows < 1 Then Exit Sub
    For I = 1 To (FG_PED.Rows - 1)
        'FG_P2N
        FG_P2N.AddItem (FG_P2N.Rows)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 0) = (FG_P2N.Rows - 1)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 1) = FG_PED.TextMatrix(I, 1)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 2) = FG_PED.TextMatrix(I, 2)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 3) = FG_PED.TextMatrix(I, 3)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 4) = FG_PED.TextMatrix(I, 4)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 5) = FG_PED.TextMatrix(I, 5)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 6) = FG_PED.TextMatrix(I, 6)
        FG_P2N.TextMatrix((FG_P2N.Rows - 1), 7) = FG_PED.TextMatrix(I, 7)
        'FG_P2N2
        FG_P2N2.AddItem (FG_P2N2.Rows)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 0) = FG_PED2.TextMatrix(I, 0)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 1) = FG_PED2.TextMatrix(I, 1)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 2) = FG_PED2.TextMatrix(I, 2)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 3) = FG_PED2.TextMatrix(I, 3)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 4) = FG_PED2.TextMatrix(I, 4)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 5) = FG_PED2.TextMatrix(I, 5)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 6) = FG_PED2.TextMatrix(I, 6)
        FG_P2N2.TextMatrix((FG_P2N2.Rows - 1), 7) = FG_PED2.TextMatrix(I, 7)
    Next I
End Sub
Private Sub BT_LL_Click()
    'limpa tudo
    CB_Pedidos.ListIndex = -1
    MontaFGPedidos
End Sub
Private Sub BT_RI_Click()
    'remove item pedido importado
    If FG_P2N.RowSel < 1 Then Exit Sub
    If (FG_P2N.Rows - 1) = 1 Then
        BT_LL_Click
    Else
        FG_P2N.RemoveItem (FG_P2N.RowSel)
        FG_P2N2.RemoveItem (FG_P2N.RowSel)
    End If
End Sub
Private Sub BT_Voltar_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaNFEmEspera (True)
    If FR_2.Visible = True Then
        BT_Voltar.Enabled = False
        FR_2.Visible = False
        FR_1.Visible = True
    ElseIf FR_3.Visible = True Then
        FR_3.Visible = False
        FR_2.Visible = True
    ElseIf FR_4.Visible = True Then
        FR_4.Visible = False
        FR_3.Visible = True
    ElseIf FR_5.Visible = True Then
        FR_5.Visible = False
        FR_4.Visible = True
    ElseIf FR_6.Visible = True Then
        FR_6.Visible = False
        FR_5.Visible = True
    ElseIf FR_7.Visible = True Then
        FR_7.Visible = False
        FR_6.Visible = True
    ElseIf FR_8.Visible = True Then
        FR_8.Visible = False
        FR_7.Visible = True
    ElseIf FR_9.Visible = True Then
        FR_9.Visible = False
        FR_8.Visible = True
        BT_Avancar.Enabled = True
        BT_Concluir.Enabled = False
    End If
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_CFOP_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Especie_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_QuantVol.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Estado_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And CB_Estado.ListIndex = -1 Then
        TXT_CEP.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_EstTrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Frete.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_EstVei_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Especie.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Exibir_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaNFEmEspera (True)
    LT_Apelido.Clear
    DLL_BD.BDSIS_TBEMP.MoveFirst
    If CB_Exibir.List(CB_Exibir.ListIndex) = "Todos" Then 'Todos
        While Not DLL_BD.BDSIS_TBEMP.EOF
            If DLL_BD.BDSIS_TBEMP_CPAPE.Value <> "" Then
                LT_Apelido.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            DLL_BD.BDSIS_TBEMP.MoveNext
        Wend
    Else
        While Not DLL_BD.BDSIS_TBEMP.EOF
            If DLL_BD.BDSIS_TBEMP_CPTIP = CB_Exibir.List(CB_Exibir.ListIndex) Then
                LT_Apelido.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            DLL_BD.BDSIS_TBEMP.MoveNext
        Wend
    End If
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Exibir_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then LT_Apelido.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Frete_Change()
    If CB_Frete.ListIndex = 0 Then
        CB_Placa.Text = ""
        CB_EstVei.Text = ""
    End If
End Sub
Private Sub CB_Frete_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Placa.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Marca_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_NumVol.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Natureza_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaNFEmEspera (True)
    DLL_BD.BDSIS_TBCDF.MoveFirst
    CB_CFOP.Clear
    Do While Not DLL_BD.BDSIS_TBCDF.EOF
        If DLL_BD.BDSIS_TBCDF_CPNTO.Value = CB_Natureza.Text Then CB_CFOP.AddItem (DLL_BD.BDSIS_TBCDF_CPCFO.Value)
        DLL_BD.BDSIS_TBCDF.MoveNext
    Loop
    If CB_CFOP.ListCount > 0 Then CB_CFOP.ListIndex = 0
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Natureza_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_CFOP.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Pedidos_Change()
    CB_Pedidos_Click
End Sub
Private Sub CB_Pedidos_Click()
    CarregaPedidos
End Sub
Private Sub CB_Placa_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_EstVei.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_TipoTrans_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaNFEmEspera (True)
    LT_NomeTrans.Clear
    DLL_BD.BDSIS_TBEMP.MoveFirst
    If CB_TipoTrans.List(CB_TipoTrans.ListIndex) = "Todos" Then 'Todos
        While Not DLL_BD.BDSIS_TBEMP.EOF
            If DLL_BD.BDSIS_TBEMP_CPAPE.Value <> "" Then
                LT_NomeTrans.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            DLL_BD.BDSIS_TBEMP.MoveNext
        Wend
    Else
        While Not DLL_BD.BDSIS_TBEMP.EOF
            If DLL_BD.BDSIS_TBEMP_CPTIP = CB_TipoTrans.List(CB_TipoTrans.ListIndex) Then
                LT_NomeTrans.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            DLL_BD.BDSIS_TBEMP.MoveNext
        Wend
    End If
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_TipoTrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then LT_NomeTrans.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_DataSaida_Click()
    On Error GoTo ERRO_SISCOVAL
    If CK_DataSaida.Value = 1 Then
        TXT_DataSaida.Enabled = True
        LB_DataSaida.Enabled = True
    Else
        TXT_DataSaida.Enabled = False
        LB_DataSaida.Enabled = False
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_DataSaida_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CK_DataSaida.Value = 0 Then
        RD_28dd.SetFocus
    ElseIf KeyAscii = vbKeyReturn And CK_DataSaida.Value = 1 Then
        TXT_DataSaida.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Desdobrar_Click()
    On Error GoTo ERRO_SISCOVAL
    If CK_Desdobrar.Value = 1 Then
        TXT_DataVenc_B.Enabled = True
        TXT_DataVenc_C.Enabled = True
        TXT_DataVenc_D.Enabled = True
        TXT_DVB.Enabled = True
        TXT_DVC.Enabled = True
        TXT_DVD.Enabled = True
        LB_DVB.Enabled = True
        LB_DVC.Enabled = True
        LB_DVD.Enabled = True
        LB_DataVenc_B.Enabled = True
        LB_DataVenc_C.Enabled = True
        LB_DataVenc_D.Enabled = True
    Else
        TXT_DataVenc_B.Enabled = False
        TXT_DataVenc_C.Enabled = False
        TXT_DataVenc_D.Enabled = False
        TXT_DVB.Enabled = False
        TXT_DVC.Enabled = False
        TXT_DVD.Enabled = False
        LB_DVB.Enabled = False
        LB_DVC.Enabled = False
        LB_DVD.Enabled = False
        LB_DataVenc_B.Enabled = False
        LB_DataVenc_C.Enabled = False
        LB_DataVenc_D.Enabled = False
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_Desdobrar_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And CK_Desdobrar.Value = 0 Then
        BT_Avancar.SetFocus
    ElseIf KeyAscii = 13 And CK_Desdobrar.Value = 1 Then
        TXT_DVB.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EditarTrans_Click()
    On Error GoTo ERRO_SISCOVAL
    If CK_EditarTrans.Value = 0 Then
        Label29.Enabled = False
        TXT_Trans.Enabled = False
        Label39.Enabled = False
        TXT_CGCTrans.Enabled = False
        Label38.Enabled = False
        TXT_IETrans.Enabled = False
        Label34.Enabled = False
        TXT_CidTrans.Enabled = False
        Label31.Enabled = False
        TXT_EndTrans.Enabled = False
        Label35.Enabled = False
        CB_EstTrans.Enabled = False
        TXT_CGCTrans.ForeColor = &H80000011
    ElseIf CK_EditarTrans.Value = 1 Then
        Label29.Enabled = True
        TXT_Trans.Enabled = True
        Label39.Enabled = True
        TXT_CGCTrans.Enabled = True
        Label38.Enabled = True
        TXT_IETrans.Enabled = True
        Label34.Enabled = True
        TXT_CidTrans.Enabled = True
        Label31.Enabled = True
        TXT_EndTrans.Enabled = True
        Label35.Enabled = True
        CB_EstTrans.Enabled = True
        CB_Frete.ListIndex = 0
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EditarTrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CK_EditarTrans.Value = 0 Then
        CB_Frete.SetFocus
    ElseIf KeyAscii = vbKeyReturn And CK_EditarTrans.Value = 1 Then
        TXT_Trans.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EditarValores_Click()
    On Error GoTo ERRO_SISCOVAL
    If CK_EditarValores.Value = 0 Then
        TXT_BaseICMS.Enabled = False
        TXT_ValorICMS.Enabled = False
        TXT_BaseICMSSub.Enabled = False
        TXT_ValorICMSSub.Enabled = False
        TXT_ValorTotalProdutos.Enabled = False
        TXT_ValorFrete.Enabled = False
        TXT_ValorSeguro.Enabled = False
        TXT_Outras.Enabled = False
        TXT_ValorTotalIPI.Enabled = False
        TXT_ValorTotalNotaFiscal.Enabled = False
        TXT_BaseICMS.ForeColor = &H80000011
        TXT_ValorICMS.ForeColor = &H80000011
        TXT_BaseICMSSub.ForeColor = &H80000011
        TXT_ValorICMSSub.ForeColor = &H80000011
        TXT_ValorTotalProdutos.ForeColor = &H80000011
        TXT_ValorFrete.ForeColor = &H80000011
        TXT_ValorSeguro.ForeColor = &H80000011
        TXT_Outras.ForeColor = &H80000011
        TXT_ValorTotalIPI.ForeColor = &H80000011
        TXT_ValorTotalNotaFiscal.ForeColor = &H80000011
    ElseIf CK_EditarValores.Value = 1 Then
        TXT_BaseICMS.Enabled = True
        TXT_ValorICMS.Enabled = True
        TXT_BaseICMSSub.Enabled = True
        TXT_ValorICMSSub.Enabled = True
        TXT_ValorTotalProdutos.Enabled = True
        TXT_ValorFrete.Enabled = True
        TXT_ValorSeguro.Enabled = True
        TXT_Outras.Enabled = True
        TXT_ValorTotalIPI.Enabled = True
        TXT_ValorTotalNotaFiscal.Enabled = True
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CK_EditarValores_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CK_EditarValores.Value = 0 Then
        TXT_HoraSaida.SetFocus
    ElseIf KeyAscii = vbKeyReturn And CK_EditarValores.Value = 1 Then
        TXT_BaseICMS.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    Set DLL_ASFIG = New Assfig.Classe_Assfig
    Set DLL_CADEMP = New Cademp.Classe_Cademp
    Set DLL_IMP = New Impform.Classe_Impform

    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (54)
    DLL_CARGA.ResetaBP
    
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Empresas...")
    If DLL_BD.AbreTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Códigos Fiscais...")
    If DLL_BD.AbreTabela_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal...")
    If DLL_BD.AbreTabela_NotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Produtos...")
    If DLL_BD.AbreTabela_NotaFiscalProdutos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Declarações...")
    If DLL_BD.AbreTabela_NotaFiscalDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Bancos...")
    If DLL_BD.AbreTabela_NotaFiscalBancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela de Nota Fiscal - Comentários...")
    If DLL_BD.AbreTabela_NotaFiscalComentarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Grupos...")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque...")
    If DLL_BD.AbreTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Índice...")
    If DLL_BD.AbreTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Figuras...")
    If DLL_BD.AbreTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Bancos...")
    If DLL_BD.AbreTabela_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - CF e ST...")
    If DLL_BD.AbreTabela_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Alíquotas...")
    If DLL_BD.AbreTabela_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Declarações...")
    If DLL_BD.AbreTabela_EstoqueDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Mapa - Faturamento do Mês...")
    If DLL_BD.AbreTabela_MapaFaturamentoMes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Mapa - Contas à Receber...")
    If DLL_BD.AbreTabela_MapaContasReceber(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Mapa - Impostos...")
    If DLL_BD.AbreTabela_MapaImpostos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Configurações da Nota Fiscal...")
    If DLL_BD.AbreTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Pedidos...")
    If DLL_BD.AbreTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Pedidos - Ítens...")
    If DLL_BD.AbreTabela_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Abre Campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Empresas...")
    If DLL_BD.AbreCampos_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Códigos Fiscais...")
    If DLL_BD.AbreCampos_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal...")
    If DLL_BD.AbreCampos_NotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Produtos...")
    If DLL_BD.AbreCampos_NotaFiscalProdutos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Declarações...")
    If DLL_BD.AbreCampos_NotaFiscalDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Bancos...")
    If DLL_BD.AbreCampos_NotaFiscalBancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Nota Fiscal - Comentários...")
    If DLL_BD.AbreCampos_NotaFiscalComentarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque...")
    If DLL_BD.AbreCampos_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Índice...")
    If DLL_BD.AbreCampos_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Figuras...")
    If DLL_BD.AbreCampos_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Bancos...")
    If DLL_BD.AbreCampos_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - CF e ST...")
    If DLL_BD.AbreCampos_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Alíquotas...")
    If DLL_BD.AbreCampos_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Declarações...")
    If DLL_BD.AbreCampos_EstoqueDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Mapa - Faturamento do Mês...")
    If DLL_BD.AbreCampos_MapaFaturamentoMes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Mapa - Contas à Receber...")
    If DLL_BD.AbreCampos_MapaContasReceber(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Mapa - Impostos...")
    If DLL_BD.AbreCampos_MapaImpostos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela de Configurações da Nota Fiscal")
    If DLL_BD.AbreCampos_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Pedidos...")
    If DLL_BD.AbreCampos_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Pedidos - Ítens...")
    If DLL_BD.AbreCampos_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    'Montando tela
    DLL_CARGA.CarregaTexto ("Organizando tela...")
    FR_2.Left = FR_1.Left
    FR_2.Top = FR_1.Top
    FR_3.Left = FR_1.Left
    FR_3.Top = FR_1.Top
    FR_4.Left = FR_1.Left
    FR_4.Top = FR_1.Top
    FR_5.Left = FR_1.Left
    FR_5.Top = FR_1.Top
    FR_6.Left = FR_1.Left
    FR_6.Top = FR_1.Top
    FR_7.Left = FR_1.Left
    FR_7.Top = FR_1.Top
    FR_8.Left = FR_1.Left
    FR_8.Top = FR_1.Top
    FR_9.Left = FR_1.Left
    FR_9.Top = FR_1.Top

    ' Tela Operacoes
    DLL_CARGA.CarregaTexto ("Carregando Operações...")
    RD_Saida.Value = True
        
    ' Tela Empresas
    DLL_CARGA.CarregaTexto ("Carregando Empresas...")
    CB_Exibir.Clear
    CB_Estado.Clear
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_Exibir, Tela_NotaFiscal.Name)
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_Estado, Tela_NotaFiscal.Name)
    TXT_CGC.ForeColor = &H80000011
    TXT_CEP.ForeColor = &H80000011

    ' Tela Pedidos
    DLL_CARGA.CarregaTexto ("Carregando Pedidos...")
    RD_Inserir.Value = True
    CB_Pedidos.Clear
    MontaFGPedidos
    
    ' Tela Informacoes Gerais
    DLL_CARGA.CarregaTexto ("Carregando Informações sobre a Nota Fiscal...")
    DLL_BD.BDSIS_TBNTF.MoveLast
    TXT_NF.Text = DLL_BD.BDSIS_TBNTF_CPNNF.Value + 1
    TXT_DataEmissao.Text = Format(Date, "dd/mm/yyyy")
    TXT_DataSaida.Text = Format(Date, "dd/mm/yyyy")
    TXT_DataSaida.Enabled = False
    CK_DataSaida.Value = 0
    CK_Desdobrar.Value = 0
    TXT_DataVenc_A.Text = Format(DateAdd("d", 28, Date), "dd/mm/yyyy")
    TXT_DataVenc_B.Text = "__/__/____"
    TXT_DataVenc_C.Text = "__/__/____"
    TXT_DataVenc_D.Text = "__/__/____"
    TXT_DataVenc_B.Enabled = False
    TXT_DataVenc_C.Enabled = False
    TXT_DataVenc_D.Enabled = False
    RD_28dd.Value = True
    LB_IMP.Visible = False
    LB_IMP.Caption = "IMPRIMIR DATA"
    
    ' Tela Produtos
    DLL_CARGA.CarregaTexto ("Carregando Produtos...")
    MontaFG
    CK_EditarValores.Value = 0
    TXT_BaseICMS.Enabled = False
    TXT_ValorICMS.Enabled = False
    TXT_BaseICMSSub.Enabled = False
    TXT_ValorICMSSub.Enabled = False
    TXT_ValorTotalProdutos.Enabled = False
    TXT_ValorFrete.Enabled = False
    TXT_ValorSeguro.Enabled = False
    TXT_Outras.Enabled = False
    TXT_ValorTotalIPI.Enabled = False
    TXT_ValorTotalNotaFiscal.Enabled = False
    
    'Tela Assistente de Ítens
    DLL_CARGA.CarregaTexto ("Carregando Assistente de Ítens...")
    If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEFG.MoveFirst
        While Not DLL_BD.BDSIS_TBEFG.EOF
            If DLL_BD.BDSIS_TBEFG_CPFIG.Value <> "" Then
                Tela_NotaFiscal_Dlg_1.CB_Figura.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value)
            End If
            DLL_BD.BDSIS_TBEFG.MoveNext
        Wend
    End If
            
    'Tela Informações Complementares
    DLL_CARGA.CarregaTexto ("Carregando Informações Complementares...")
    CK_EditarValores.Value = 0
    TXT_BaseICMS.ForeColor = &H80000011
    TXT_ValorICMS.ForeColor = &H80000011
    TXT_BaseICMSSub.ForeColor = &H80000011
    TXT_ValorICMSSub.ForeColor = &H80000011
    TXT_ValorTotalProdutos.ForeColor = &H80000011
    TXT_ValorFrete.ForeColor = &H80000011
    TXT_ValorSeguro.ForeColor = &H80000011
    TXT_Outras.ForeColor = &H80000011
    TXT_ValorTotalIPI.ForeColor = &H80000011
    TXT_ValorTotalNotaFiscal.ForeColor = &H80000011
            
    'Tela Certificados
    DLL_CARGA.CarregaTexto ("Carregando Certificados de Qualidade...")
    RB_CF1.Value = True
    RB_CF2.Enabled = False 'ATIVAR DEPOIS DE FEITO CERTIFICADOS
    
    'Tela Transportador / Volumes
    DLL_CARGA.CarregaTexto ("Carregando Transportador / Volumes...")
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_TipoTrans, Tela_NotaFiscal.Name)
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_Frete, Tela_NotaFiscal.Name)
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_EstVei, Tela_NotaFiscal.Name)
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_EstTrans, Tela_NotaFiscal.Name)
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_Especie, Tela_NotaFiscal.Name)
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_Marca, Tela_NotaFiscal.Name)
    RespMsg = CeI_Combos(Tela_NotaFiscal.CB_Placa, Tela_NotaFiscal.Name)
    CK_EditarTrans.Value = 0
    CF_I = ""
    CF_J = ""
    TXT_CGCTrans.ForeColor = &H80000011
    
    'Finalizando
    InseriuPedido = False
    DLL_FUNCS.RegistraEvento "Abrir Assistente de Notas Fiscais", ""
    DLL_CARGA.CarregaTexto ("Finalizando")
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_NotaFiscal
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Apelido_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_Apelido.ListIndex = -1 Then
        Exit Sub
    End If
    TelaNFEmEspera (True)
    DLL_BD.BDSIS_TBEMP.MoveFirst
    DLL_BD.BDSIS_TBEMP.Seek "=", LT_Apelido.Text
    If DLL_BD.BDSIS_TBEMP.NoMatch Then
        RespMsg = MsgBox("Ocorreu erro durante a procura do nome fantasia da empresa para inserir na lista.", vbOKOnly, Tela_NotaFiscal.Caption)
        Exit Sub
        TelaNFEmEspera (False)
    Else
        TXT_Apelido.Text = DLL_BD.BDSIS_TBEMP_CPAPE
        If DLL_BD.BDSIS_TBEMP_CPEMP <> "" Then
            TXT_Empresa.Text = DLL_BD.BDSIS_TBEMP_CPEMP
        Else
            TXT_Empresa.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCGC <> "" Then
            TXT_CGC.Text = DLL_BD.BDSIS_TBEMP_CPCGC
        Else
            TXT_CGC.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPINE <> "" Then
            TXT_InsEst.Text = DLL_BD.BDSIS_TBEMP_CPINE
        Else
            TXT_InsEst.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPEND <> "" Then
            TXT_Endereco.Text = DLL_BD.BDSIS_TBEMP_CPEND
        Else
            TXT_Endereco.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPBAI <> "" Then
            TXT_Bairro.Text = DLL_BD.BDSIS_TBEMP_CPBAI
        Else
            TXT_Bairro.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPPRA <> "" Then
            TXT_PracaPagamento.Text = DLL_BD.BDSIS_TBEMP_CPPRA
        Else
            TXT_PracaPagamento.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCID <> "" Then
            TXT_Cidade.Text = DLL_BD.BDSIS_TBEMP_CPCID
        Else
            TXT_Cidade.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCEP <> "" Then
            TXT_CEP.Text = DLL_BD.BDSIS_TBEMP_CPCEP
        Else
            TXT_CEP.Text = "_____-___"
        End If
        If DLL_BD.BDSIS_TBEMP_CPFON <> "" Then
            TXT_Fone.Text = DLL_BD.BDSIS_TBEMP_CPFON
        Else
            TXT_Fone.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCOM.Value <> "" Then
            TXT_Comentario.Text = DLL_BD.BDSIS_TBEMP_CPCOM.Value
        Else
            TXT_Comentario.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPOBS <> "" Then
            TXT_Observacao.Text = DLL_BD.BDSIS_TBEMP_CPOBS
        Else
            TXT_Observacao.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPTRA <> "" Then
            TRANSPORTADORA = DLL_BD.BDSIS_TBEMP_CPTRA.Value
        Else
            TRANSPORTADORA = ""
        End If
        CB_Estado.ListIndex = -1
        Do While Not DLL_BD.BDSIS_TBEMP.NoMatch
            For I = 0 To CB_Estado.ListCount - 1
                If DLL_BD.BDSIS_TBEMP_CPEST.Value = CB_Estado.List(I) Then
                    CB_Estado.ListIndex = I
                    Exit For
                End If
            Next
            DLL_BD.BDSIS_TBEMP.MoveNext
            If CB_Estado.ListCount >= 0 Then
                Exit Do
            End If
        Loop
    End If
    RD_Inserir.Value = True
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Terminate()
    On Error GoTo ERRO_SISCOVAL
    Unload Tela_NotaFiscal
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_CodigosFiscais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalProdutos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalBancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_NotaFiscalComentarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Bancos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueDeclaracoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MapaFaturamentoMes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MapaContasReceber(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MapaImpostos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
    Set DLL_FUNCS = Nothing
    Set DLL_ASFIG = Nothing
    Set DLL_CADEMP = Nothing
    Set DLL_IMP = Nothing
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Apelido_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And LT_Apelido.ListIndex <> -1 Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_NomeTrans_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NomeTrans.ListIndex = -1 Then
        TXT_Trans.Text = ""
        TXT_CGCTrans.Text = "__.___.___/____-__"
        TXT_IETrans.Text = ""
        TXT_EndTrans.Text = ""
        TXT_CidTrans.Text = ""
        CB_EstTrans.ListIndex = -1
        Exit Sub
    End If
    DLL_BD.BDSIS_TBEMP.Seek "=", LT_NomeTrans.Text
    If DLL_BD.BDSIS_TBEMP.NoMatch Then
        RespMsg = MsgBox("Ocorreu erro durante a procura do nome fantasia da empresa transportadora para inserir na lista.", vbOKOnly, Tela_NotaFiscal.Caption)
        Exit Sub
    Else
        If DLL_BD.BDSIS_TBEMP_CPEMP <> "" Then
            TXT_Trans.Text = DLL_BD.BDSIS_TBEMP_CPEMP
        Else
            TXT_Trans.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCGC <> "" Then
            TXT_CGCTrans.Text = DLL_BD.BDSIS_TBEMP_CPCGC
        Else
            TXT_CGCTrans.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPINE <> "" Then
            TXT_IETrans.Text = DLL_BD.BDSIS_TBEMP_CPINE
        Else
            TXT_IETrans.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPEND <> "" Then
            TXT_EndTrans.Text = DLL_BD.BDSIS_TBEMP_CPEND
        Else
            TXT_EndTrans.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPCID <> "" Then
            TXT_CidTrans.Text = DLL_BD.BDSIS_TBEMP_CPCID
        Else
            TXT_CidTrans.Text = ""
        End If
        If DLL_BD.BDSIS_TBEMP_CPEST.Value <> "" Then
            For I = 0 To CB_EstTrans.ListCount - 1
                If DLL_BD.BDSIS_TBEMP_CPEST.Value = CB_EstTrans.List(I) Then
                    CB_EstTrans.ListIndex = I
                    Exit For
                End If
            Next
        End If
    End If
    If LT_NomeTrans.Text = "NOSSO MOTORISTA" Then
        CB_EstVei.Text = "SP"
        CB_Frete.ListIndex = 1
        CB_Especie.ListIndex = 1
        CB_Placa.ListIndex = 0
        TXT_NumVol.Text = ""
    ElseIf LT_NomeTrans.Text = "SEU MOTORISTA" Then
        CB_EstVei.ListIndex = -1
        CB_Frete.ListIndex = 0
        CB_Especie.ListIndex = 1
        CB_Placa.Text = ""
        TXT_NumVol.Text = ""
        TXT_Trans.Text = "Seu motorista"
        TXT_CGCTrans.Text = TXT_CGC.Text
        TXT_IETrans.Text = TXT_InsEst.Text
        TXT_CidTrans.Text = TXT_Cidade.Text
        TXT_EndTrans.Text = TXT_Endereco.Text
        CB_EstTrans.ListIndex = CB_Estado.ListIndex
    Else
        CB_EstVei.ListIndex = -1
        CB_Frete.ListIndex = 0
        CB_Especie.ListIndex = 0
        CB_Placa.Text = ""
        TXT_NumVol.Text = ""
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_NomeTrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CK_EditarTrans.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_CF1_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_SV_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = "-x-       "
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RB_SV_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_21dd_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = Format(DateAdd("d", 21, Date), "dd/mm/yyyy")
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_21dd_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_28dd_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = Format(DateAdd("d", 28, Date), "dd/mm/yyyy")
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_28dd_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_30dd_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = Format(DateAdd("d", 30, Date), "dd/mm/yyyy")
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_30dd_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_35dd_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = Format(DateAdd("d", 35, Date), "dd/mm/yyyy")
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_35dd_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_45dd_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = Format(DateAdd("d", 45, Date), "dd/mm/yyyy")
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_45dd_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_AVista_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = "À Vista   "
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_AVista_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_CApres_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_DataVenc_A.Text = "C/Apres.  "
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_CApres_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Entrada_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaNFEmEspera (True)
    CB_Natureza.Clear
    CB_CFOP.Clear
    DLL_BD.BDSIS_TBCDF.MoveFirst
    While Not DLL_BD.BDSIS_TBCDF.EOF
        If DLL_BD.BDSIS_TBCDF_CPNTO.Value <> "" And _
           DLL_BD.BDSIS_TBCDF_CPTIP.Value = "E" Then
            If CB_Natureza.ListCount = 0 Then
                CB_Natureza.AddItem (DLL_BD.BDSIS_TBCDF_CPNTO.Value)
            End If
            If CB_CFOP.ListCount = 0 Then
              CB_CFOP.AddItem (DLL_BD.BDSIS_TBCDF_CPCFO.Value)
            End If
            For I = 0 To CB_Natureza.ListCount - 1
                If I = (CB_Natureza.ListCount - 1) And _
                   CB_Natureza.List(I) <> DLL_BD.BDSIS_TBCDF_CPNTO.Value Then
                    CB_Natureza.AddItem (DLL_BD.BDSIS_TBCDF_CPNTO.Value)
                ElseIf CB_Natureza.List(I) = DLL_BD.BDSIS_TBCDF_CPNTO.Value Then
                    Exit For
                End If
            Next I
            For I = 0 To CB_CFOP.ListCount - 1
                If I = (CB_CFOP.ListCount - 1) And _
                   CB_CFOP.List(I) <> DLL_BD.BDSIS_TBCDF_CPCFO.Value Then
                    CB_CFOP.AddItem (DLL_BD.BDSIS_TBCDF_CPCFO.Value)
                ElseIf CB_CFOP.List(I) = DLL_BD.BDSIS_TBCDF_CPCFO.Value Then
                    Exit For
                End If
            Next I
         End If
        DLL_BD.BDSIS_TBCDF.MoveNext
    Wend
    If CB_Natureza.ListCount > 0 Then
        CB_Natureza.ListIndex = 0
        CB_CFOP.ListIndex = 0
    End If
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Entrada_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then CB_Natureza.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Escdd_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_Escdd.Enabled = True
    TXT_Escdd.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Escdd_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Escdd.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Inserir_Click()
    On Error GoTo ERRO_SISCOVAL
    'Este botao rádio diz que o DLL_FUNCS.PegaUsuario pretende digitar os dados da nota fiscal
    FR_4_2.Enabled = False
    FR_4_2_1.Enabled = False
    FR_4_2_2.Enabled = False
    CB_Pedidos.Enabled = False
    FG_PED.Enabled = False
    FG_P2N.Enabled = False
    BT_IT.Enabled = False
    BT_II.Enabled = False
    BT_RI.Enabled = False
    BT_LL.Enabled = False
    MontaFGPedidos
    CB_Pedidos.Clear
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Inserir_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Pedido_Click()
    On Error GoTo ERRO_SISCOVAL
    'Este botão rádio diz que o usuário pretende verificar todos os pedidos pendentes do cliente e emitir a nota fiscal
    FR_4_2.Enabled = True
    FR_4_2_1.Enabled = True
    FR_4_2_2.Enabled = True
    CB_Pedidos.Enabled = True
    FG_PED.Enabled = True
    FG_P2N.Enabled = True
    BT_IT.Enabled = True
    BT_II.Enabled = True
    BT_RI.Enabled = True
    BT_LL.Enabled = True
    MontaPedidos
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Saida_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaNFEmEspera (True)
    CB_Natureza.Clear
    CB_CFOP.Clear
    DLL_BD.BDSIS_TBCDF.MoveFirst
    While Not DLL_BD.BDSIS_TBCDF.EOF
        If DLL_BD.BDSIS_TBCDF_CPNTO.Value <> "" And _
           DLL_BD.BDSIS_TBCDF_CPTIP.Value = "S" Then
            If CB_Natureza.ListCount = 0 Then
                CB_Natureza.AddItem (DLL_BD.BDSIS_TBCDF_CPNTO.Value)
            End If
            If CB_CFOP.ListCount = 0 Then
              CB_CFOP.AddItem (DLL_BD.BDSIS_TBCDF_CPCFO.Value)
            End If
            For I = 0 To CB_Natureza.ListCount - 1
                If I = (CB_Natureza.ListCount - 1) And _
                   CB_Natureza.List(I) <> DLL_BD.BDSIS_TBCDF_CPNTO.Value Then
                    CB_Natureza.AddItem (DLL_BD.BDSIS_TBCDF_CPNTO.Value)
                ElseIf CB_Natureza.List(I) = DLL_BD.BDSIS_TBCDF_CPNTO.Value Then
                    Exit For
                End If
            Next I
            For I = 0 To CB_CFOP.ListCount - 1
                If I = (CB_CFOP.ListCount - 1) And _
                   CB_CFOP.List(I) <> DLL_BD.BDSIS_TBCDF_CPCFO.Value Then
                    CB_CFOP.AddItem (DLL_BD.BDSIS_TBCDF_CPCFO.Value)
                ElseIf CB_CFOP.List(I) = DLL_BD.BDSIS_TBCDF_CPCFO.Value Then
                    Exit For
                End If
            Next I
         End If
        DLL_BD.BDSIS_TBCDF.MoveNext
    Wend
    CB_Natureza.ListIndex = 0
    CB_CFOP.ListIndex = 0
    TelaNFEmEspera (False)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub RD_Saida_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_Natureza.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Apelido_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Empresa.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Bairro_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Cidade.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_BaseICMS.SelLength = Len(TXT_BaseICMS.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_ValorICMS.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseICMSSub_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_BaseICMSSub.SelLength = Len(TXT_BaseICMSSub.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_BaseICMSSub_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_ValorICMSSub.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CEP_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_CEP.Text <> "" Then
        TXT_Fone.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CGC_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_CGC.Text <> "" Then
        TXT_InsEst.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CGCTrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_IETrans.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Cidade_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_Cidade.Text <> "" Then
        CB_Estado.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_CidTrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_EndTrans.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Comentario_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Observacao.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DataEmissao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 And TXT_DataEmissao.Text <> "" Then
        CK_DataSaida.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DataSaida_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then
        If LB_IMP.Visible = False Then
            LB_IMP.Visible = True
        Else
            LB_IMP.Visible = False
        End If
    End If
End Sub
Private Sub TXT_DataSaida_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DataVenc_A_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DataVenc_B_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DataVenc_C_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DataVenc_D_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DVB_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_DVB.Text = "" Then
        TXT_DataVenc_B.Text = "__/__/____"
    Else
        TXT_DataVenc_B.Text = Format(DateAdd("d", TXT_DVB.Text, Date), "dd/mm/yyyy")
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DVB_GotFocus()
    If TXT_DataVenc_A.Text = "" Or TXT_DataVenc_A.Text = "__/__/____" Then
        MsgBox "Você deve inserir um vencimento /A primeiro.", vbExclamation + vbOKOnly, NOMEAPLIC
        RD_28dd.SetFocus
    End If
End Sub
Private Sub TXT_DVB_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_DVC.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DVC_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_DVC.Text = "" Then
        TXT_DataVenc_C.Text = "__/__/____"
    Else
        TXT_DataVenc_C.Text = Format(DateAdd("d", TXT_DVC.Text, Date), "dd/mm/yyyy")
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DVC_GotFocus()
    If TXT_DVB.Text = "" Then
        MsgBox "Você deve inserir um vencimento /B primeiro.", vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_DVB.SetFocus
    End If
End Sub
Private Sub TXT_DVC_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 And TXT_DataVenc_C.Text <> "" Then
        TXT_DVD.SetFocus
    ElseIf KeyAscii = 13 And TXT_DataVenc_C.Text = "" Then
        BT_Avancar.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DVD_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_DVD.Text = "" Then
        TXT_DataVenc_D.Text = "__/__/____"
    Else
        TXT_DataVenc_D.Text = Format(DateAdd("d", TXT_DVD.Text, Date), "dd/mm/yyyy")
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_DVD_GotFocus()
    If TXT_DVB.Text = "" Then
        MsgBox "Você deve inserir um vencimento /B primeiro.", vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_DVB.SetFocus
    ElseIf TXT_DVB.Text <> "" And TXT_DVC.Text = "" Then
        MsgBox "Você deve inserir um vencimento /C primeiro.", vbExclamation + vbOKOnly, NOMEAPLIC
        TXT_DVC.SetFocus
    End If
End Sub
Private Sub TXT_DVD_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_DVD.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Empresa_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_Empresa.Text <> "" Then
        TXT_CGC.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Endereco_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_Endereco.Text <> "" Then
        TXT_PracaPagamento.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_EndTrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then CB_EstTrans.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Escdd_Change()
    On Error GoTo ERRO_SISCOVAL
    If TXT_Escdd.Text = "" Then
        TXT_DataVenc_A.Text = "__/__/____"
    Else
        TXT_DataVenc_A.Text = Format(DateAdd("d", TXT_Escdd.Text, Date), "dd/mm/yyyy")
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Escdd_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Fone_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then TXT_Comentario.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_HoraSaida_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_HoraSaida.Text = Format(Time, "hh:mm:ss")
    TXT_HoraSaida.SelLength = Len(TXT_HoraSaida.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_HoraSaida_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_PedidoInterno.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_HoraSaida_LostFocus()
    On Error GoTo ERRO_SISCOVAL
    If TXT_HoraSaida.Text <> "" Then
        If IsError(Format(TXT_HoraSaida.Text, "hh:mm:ss")) Then
            MsgBox "A hora digitada é inválida - digite no formato hh:mm:ss", vbInformation + vbOKOnly, NOMEAPLIC
            TXT_HoraSaida.SetFocus
            Exit Sub
        End If
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_IETrans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_CidTrans.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_InsEst_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 And TXT_InsEst.Text <> "" Then
        TXT_Endereco.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NF_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 And TXT_NF.Text <> "" Then
        TXT_DataEmissao.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_NumVol_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_PesoBruto.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Observacao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Operacao_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Operacao.SelLength = Len(TXT_Operacao.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Operacao_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_VendInterno.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Outras_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Outras.SelLength = Len(TXT_Outras.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Outras_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        TXT_ValorTotalIPI.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PedidoInterno_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_PedidoInterno.SelLength = Len(TXT_PedidoInterno.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PedidoInterno_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_SeuPedido.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoBruto_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_PesoLiquido.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PesoLiquido_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then BT_Avancar.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_PracaPagamento_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Bairro.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_QuantVol_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CB_Marca.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Setor_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_Setor.SelLength = Len(TXT_Setor.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Setor_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        BT_Avancar.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_SeuPedido_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_SeuPedido.SelLength = Len(TXT_SeuPedido.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_SeuPedido_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Operacao.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_Trans_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_CGCTrans.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorFrete_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorFrete.SelLength = Len(TXT_ValorFrete.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorFrete_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_ValorSeguro.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorICMS.SelLength = Len(TXT_ValorICMS.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMS_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_BaseICMSSub.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMSSub_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorICMSSub.SelLength = Len(TXT_ValorICMSSub.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorICMSSub_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_ValorTotalProdutos.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorSeguro_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorSeguro.SelLength = Len(TXT_ValorSeguro.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorSeguro_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        TXT_Outras.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotalIPI_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorTotalIPI.SelLength = Len(TXT_ValorTotalIPI.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotalIPI_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        TXT_ValorTotalNotaFiscal.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotalNotaFiscal_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorTotalNotaFiscal.SelLength = Len(TXT_ValorTotalNotaFiscal.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotalNotaFiscal_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = 13 Then
        TXT_PedidoInterno.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotalProdutos_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_ValorTotalProdutos.SelLength = Len(TXT_ValorTotalProdutos.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_ValorTotalProdutos_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_ValorFrete.SetFocus
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_VendExterno_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_VendExterno.SelLength = Len(TXT_VendExterno.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_VendExterno_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_Setor.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_VendInterno_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    TXT_VendInterno.SelLength = Len(TXT_VendInterno.Text)
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_VendInterno_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = 13 Then
        TXT_VendExterno.SetFocus
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Private Static Sub MontaFG()
    On Error GoTo ERRO_SISCOVAL
    FG_1.ColAlignment(0) = flexAlignCenterCenter
    FG_1.ColAlignment(1) = flexAlignLeftCenter
    FG_1.ColAlignment(2) = flexAlignLeftCenter
    FG_1.ColAlignment(3) = flexAlignCenterCenter
    FG_1.ColAlignment(4) = flexAlignCenterCenter
    FG_1.ColAlignment(5) = flexAlignCenterCenter
    FG_1.ColAlignment(9) = flexAlignCenterCenter
    FG_1.ColAlignment(10) = flexAlignCenterCenter
    FG_1.ColWidth(0) = 500
    FG_1.ColWidth(1) = 800
    FG_1.ColWidth(2) = 3300
    FG_1.ColWidth(3) = 400
    FG_1.ColWidth(4) = 400
    FG_1.ColWidth(5) = 700
    FG_1.ColWidth(6) = 900
    FG_1.ColWidth(7) = 1100
    FG_1.ColWidth(8) = 1100
    FG_1.ColWidth(9) = 800
    FG_1.ColWidth(10) = 600
    FG_1.TextArray(0) = "Linha"
    FG_1.TextArray(1) = "Figura"
    FG_1.TextArray(2) = "Descrição dos Produtos"
    FG_1.TextArray(3) = "C.F."
    FG_1.TextArray(4) = "S.T."
    FG_1.TextArray(5) = "Unidade"
    FG_1.TextArray(6) = "Quantidade"
    FG_1.TextArray(7) = "Preço Unitário"
    FG_1.TextArray(8) = "Preço Total"
    FG_1.TextArray(9) = "% I.C.M.S."
    FG_1.TextArray(10) = "% I.P.I."
    FG_1.TextArray(11) = "Valor I.P.I."
    For I = 1 To 20
        FG_1.TextMatrix(I, 0) = I
        FG_2.TextMatrix(I, 0) = I
    Next I
    FG_2.TextArray(0) = "Linha"
    FG_2.TextArray(1) = "Bitola"
    FG_2.TextArray(2) = "Material"
    FG_2.TextArray(3) = "Base Cálculo I.C.M.S."
    FG_2.TextArray(4) = "Valor I.C.M.S."
    FG_2.TextArray(5) = "Peso Unitário"
    
    FG_1.Visible = True
    FG_2.Visible = False
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub LimpaLinha(NumLin As Integer)
    On Error GoTo ERRO_SISCOVAL
    Dim W
    For W = 1 To 11
        FG_1.TextMatrix(NumLin, W) = ""
    Next W
    For W = 1 To 5
        FG_2.TextMatrix(NumLin, W) = ""
    Next W
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function InsereDecsNF(Decl As String, NumItem As Integer) As Boolean
    On Error GoTo ERRO_SISCOVAL
    If Decl = "" Or Decl = "-" Then
        InsereDecsNF = True
        Exit Function
    End If
    
    'Verifica se existem decs. na NF
    For I = 20 To 1 Step -1
        If FG_2.TextMatrix(I, 1) = "DF" Then
            If FG_1.TextMatrix(I, 2) <> Decl Then
                For J = I To 1 Step -1
                    If FG_1.TextMatrix(J, 1) <> "" And _
                       FG_1.TextMatrix(J, 2) = Decl Then
                        'Este IF verifica se nas linhas de cima, existe esta declaraçao
                        If FG_1.TextMatrix(J, 1) = "" Then
                            FG_1.TextMatrix(J, 1) = VBA.Str(NumItem)
                        Else
                            FG_1.TextMatrix(J, 1) = FG_1.TextMatrix(J, 1) & ", " & VBA.Str(NumItem)
                        End If
                        InsereDecsNF = True
                        Exit Function
                    ElseIf FG_1.TextMatrix(J, 1) = "" And _
                       FG_1.TextMatrix(J, 2) = "" Then
                        If FG_1.TextMatrix(J, 1) = "" Then
                            FG_1.TextMatrix(J, 1) = VBA.Str(NumItem)
                        Else
                            FG_1.TextMatrix(J, 1) = FG_1.TextMatrix(J, 1) & ", " & VBA.Str(NumItem)
                        End If
                        FG_1.TextMatrix(J, 2) = Decl
                        FG_2.TextMatrix(J, 1) = "DF"
                        InsereDecsNF = True
                        Exit Function
                    ElseIf J = 1 Then
                        RespMsg = MsgBox("Não existe mais linhas em branco para incluir declarações.", vbOKOnly, Tela_NotaFiscal.Caption)
                        InsereDecsNF = False
                        Exit Function
                    End If
                Next J
            ElseIf FG_1.TextMatrix(I, 2) = Decl Then
                For J = I To 1 Step -1
                    If FG_1.TextMatrix(J, 1) <> "" And _
                       FG_1.TextMatrix(J, 2) = Decl Then
                        If FG_1.TextMatrix(I, 1) = "" Then
                            FG_1.TextMatrix(I, 1) = VBA.Str(NumItem)
                        Else
                            FG_1.TextMatrix(I, 1) = FG_1.TextMatrix(J, 1) & ", " & Str(NumItem)
                        End If
                        InsereDecsNF = True
                        Exit Function
                    End If
                Next J
            End If
        Else
            For J = I To 1 Step -1
                If FG_1.TextMatrix(J, 1) <> "" And _
                   FG_1.TextMatrix(J, 2) = Decl Then
                    'Este IF verifica se nas linhas de cima, existe esta declaraçao
                    If FG_1.TextMatrix(J, 1) = "" Then
                        FG_1.TextMatrix(J, 1) = Str(NumItem)
                    Else
                        FG_1.TextMatrix(J, 1) = FG_1.TextMatrix(J, 1) & ", " & Str(NumItem)
                    End If
                    InsereDecsNF = True
                    Exit Function
                ElseIf FG_1.TextMatrix(J, 1) = "" And _
                   FG_1.TextMatrix(J, 2) = "" Then
                    If FG_1.TextMatrix(J, 1) = "" Then
                        FG_1.TextMatrix(J, 1) = Str(NumItem)
                    Else
                        FG_1.TextMatrix(J, 1) = FG_1.TextMatrix(J, 1) & ", " & Str(NumItem)
                    End If
                    FG_1.TextMatrix(J, 2) = Decl
                    FG_2.TextMatrix(J, 1) = "DF"
                    InsereDecsNF = True
                    Exit Function
                ElseIf J = 1 Then
                    RespMsg = MsgBox("Não existe mais linhas em branco para incluir declarações.", vbOKOnly, Tela_NotaFiscal.Caption)
                    InsereDecsNF = False
                    Exit Function
                End If
            Next J
        End If
    Next I
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Sub TelaNFEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Tela_NotaFiscal.MousePointer = vbHourglass
        Tela_NotaFiscal.Enabled = False
    Else
        Tela_NotaFiscal.MousePointer = vbDefault
        Tela_NotaFiscal.Enabled = True
    End If
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Function CeI_Combos(Combo As ComboBox, Tela As String)
    On Error GoTo ERRO_SISCOVAL
    Set DLL_BD.BDSIS_TBCTG = DLL_BD.BDSIS.OpenRecordset("Configurações Tela-Grupo")
    Set DLL_BD.BDSIS_TBCTG_CPTEL = DLL_BD.BDSIS_TBCTG.Fields("Tela")
    Set DLL_BD.BDSIS_TBCTG_CPCOM = DLL_BD.BDSIS_TBCTG.Fields("Combo")
    Set DLL_BD.BDSIS_TBCTG_CPVAL = DLL_BD.BDSIS_TBCTG.Fields("Valor")
    Set DLL_BD.BDSIS_TBCTG_CPINI = DLL_BD.BDSIS_TBCTG.Fields("Iniciar")
    DLL_BD.BDSIS_TBCTG.Index = "Tela"
    Dim cGru As String
    DLL_BD.BDSIS_TBCTG.Seek "=", Tela, Combo.Name
    If DLL_BD.BDSIS_TBCTG.NoMatch Then
        RespMsg = MsgBox("Erro ao carregar a lista...", vbOKOnly, Tela_NotaFiscal.Caption)
        Exit Function
    Else
        'Carrega combo
        DLL_BD.BDSIS_TBGRU.MoveFirst
        Combo.Clear
        Do While Not DLL_BD.BDSIS_TBGRU.EOF
            If DLL_BD.BDSIS_TBGRU_CPTIP.Value = DLL_BD.BDSIS_TBCTG_CPVAL.Value Then
                Combo.AddItem (DLL_BD.BDSIS_TBGRU_CPVAL.Value)
            End If
            DLL_BD.BDSIS_TBGRU.MoveNext
        Loop
        'Inicia combo
        If DLL_BD.BDSIS_TBCTG_CPINI.Value <> "" Then
            DLL_BD.BDSIS_TBGRU.Seek "=", DLL_BD.BDSIS_TBCTG_CPINI.Value
            If DLL_BD.BDSIS_TBGRU.NoMatch Then
                RespMsg = MsgBox("Erro ao carregar valor da lista...", vbOKOnly, Tela_NotaFiscal.Caption)
            Else
                Combo.Text = DLL_BD.BDSIS_TBGRU_CPVAL.Value
            End If
        End If
    End If
    DLL_BD.BDSIS_TBCTG.Close
ERRO_SISCOVAL:
    If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Function
Private Sub MontaPedidos()
    CB_Pedidos.Clear
    If TXT_Apelido.Text = "" Then GoTo SEM_PEDIDO
    'procura pedidos da empresa
    MontaFGPedidos
    With DLL_BD
        If .BDSIS_TBPED.RecordCount < 1 Then
            GoTo SEM_PEDIDO
        Else
            .BDSIS_TBPED.MoveFirst
            Do While Not .BDSIS_TBPED.EOF
                If Trim(.BDSIS_TBPED_CPEMP.Value) = Trim(TXT_Apelido.Text) And .BDSIS_TBPED_CPLIQ.Value = False Then CB_Pedidos.AddItem .BDSIS_TBPED_CPIND.Value
                .BDSIS_TBPED.MoveNext
            Loop
        End If
    End With
    If CB_Pedidos.ListCount = 0 Then GoTo SEM_PEDIDO
    Exit Sub
SEM_PEDIDO:
    MsgBox "Não existe pedidos da empresa " & Trim(TXT_Apelido.Text) & ".", vbInformation + vbOKOnly, NOMEAPLIC
    RD_Inserir.Value = True
End Sub
Private Sub MontaFGPedidos(Optional nInd As Integer)
    'FG_PED
    With FG_PED
        .Visible = True
        .Enabled = True
        .Clear
        .Cols = 8
        .Rows = 1
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 2000
        .ColWidth(7) = 2000
        .TextArray(0) = "Item"
        .TextArray(1) = "Quantidade"
        .TextArray(2) = "Figura"
        .TextArray(3) = "Bitola"
        .TextArray(4) = "Material"
        .TextArray(5) = "Preço Unitário"
        .TextArray(6) = "Descrição"
        .TextArray(7) = "Complemento"
    End With
    'FG_PED2
    With FG_PED2
        .Visible = False
        .Enabled = False
        .Clear
        .Cols = 8
        .Rows = 1
        .TextArray(0) = "Índice PE"
        .TextArray(1) = "Índice Ficha"
        .TextArray(2) = "Contato"
        .TextArray(3) = "Seu Pedido"
        .TextArray(4) = "Cond.Pagto."
        .TextArray(5) = "Transportadora"
        .TextArray(6) = "Índice Vendedor"
        .TextArray(7) = "Ficha PE"
    End With
    If nInd = 1 Then Exit Sub
    'FG_P2N
    With FG_P2N
        .Visible = True
        .Enabled = True
        .Clear
        .Cols = 8
        .Rows = 1
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
        .ColAlignment(7) = flexAlignLeftCenter
        .ColWidth(0) = 500
        .ColWidth(1) = 1200
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 2000
        .ColWidth(7) = 2000
        .TextArray(0) = "Item"
        .TextArray(1) = "Quantidade"
        .TextArray(2) = "Figura"
        .TextArray(3) = "Bitola"
        .TextArray(4) = "Material"
        .TextArray(5) = "Preço Unitário"
        .TextArray(6) = "Descrição"
        .TextArray(7) = "Complemento"
    End With
    'FG_P2N2
    With FG_P2N2
        .Visible = False
        .Enabled = False
        .Clear
        .Cols = 8
        .Rows = 1
        .TextArray(0) = "Índice PE"
        .TextArray(1) = "Índice Ficha"
        .TextArray(2) = "Contato"
        .TextArray(3) = "Seu Pedido"
        .TextArray(4) = "Cond.Pagto."
        .TextArray(5) = "Transportadora"
        .TextArray(6) = "Índice Vendedor"
        .TextArray(7) = "Ficha PE"
    End With
End Sub
Private Static Sub CarregaPedidos()
    If CB_Pedidos.ListIndex < 0 Then Exit Sub
    MontaFGPedidos 1
    Dim sItens As String
    sItens = ""
    With tINFPED
        .CON = ""
        .CPG = ""
        .ITE = ""
        .IVE = ""
        .SPE = ""
        .TRA = ""
        .NPE = ""
    End With
    With DLL_BD
        If .BDSIS_TBPED.RecordCount > 0 Then
            .BDSIS_TBPED.Seek "=", CB_Pedidos.Text
            If .BDSIS_TBPED.NoMatch = False Then
                If IsNull(.BDSIS_TBPED_CPCON.Value) = False Then tINFPED.CON = .BDSIS_TBPED_CPCON.Value
                If IsNull(.BDSIS_TBPED_CPNSP.Value) = False Then tINFPED.SPE = .BDSIS_TBPED_CPNSP.Value
                If IsNull(.BDSIS_TBPED_CPCPG.Value) = False Then tINFPED.CPG = .BDSIS_TBPED_CPCPG.Value
                If IsNull(.BDSIS_TBPED_CPTRA.Value) = False Then tINFPED.TRA = .BDSIS_TBPED_CPTRA.Value
                If IsNull(.BDSIS_TBPED_CPIVE.Value) = False Then tINFPED.IVE = .BDSIS_TBPED_CPIVE.Value
                If IsNull(.BDSIS_TBPED_CPITE.Value) = False Then tINFPED.ITE = .BDSIS_TBPED_CPITE.Value
                If IsNull(.BDSIS_TBPED_CPIND.Value) = False Then tINFPED.NPE = .BDSIS_TBPED_CPIND.Value
                CarregaItensPedido tINFPED.ITE
            End If
        End If
    End With
End Sub
Private Static Sub CarregaItensPedido(Valor As String)
    Dim sTmp As String
    sTmp = ""
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            CarregaItensPedido_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    CarregaItensPedido_Aux Val(sTmp)
End Sub
Private Static Sub CarregaItensPedido_Aux(Valor As Long)
    Dim sInd As String
    If Valor < 1 Then Exit Sub
    'procura item
    With DLL_BD
        .BDSIS_TBPIT.Seek "=", Valor
        If .BDSIS_TBPIT.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens do Pedido.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        Else
            'FG_PED
            FG_PED.AddItem (FG_PED.Rows)
            FG_PED.TextMatrix((FG_PED.Rows - 1), 0) = (FG_PED.Rows - 1)
            FG_PED.TextMatrix(FG_PED.Rows - 1, 1) = Format(.BDSIS_TBPIT_CPQUA.Value, "###,##0.00")
            If .BDSIS_TBPIT_CPINF.Value > 0 Then
                sInd = .BDSIS_TBEST.Index
                .BDSIS_TBEST.Index = "Índice de Ficha"
                .BDSIS_TBEST.Seek "=", .BDSIS_TBPIT_CPINF.Value
                If .BDSIS_TBEST.NoMatch Then
                    MsgBox "Não foi possível localizar um dos ítens do Pedido.", vbExclamation + vbOKOnly, NOMEAPLIC
                    Exit Sub
                Else
                    If IsNull(.BDSIS_TBEST_CPFIG.Value) = False And IsNull(.BDSIS_TBEST_CPBIT.Value) = False And IsNull(.BDSIS_TBEST_CPMAT.Value) = False Then
                        FG_PED.TextMatrix(FG_PED.Rows - 1, 2) = .BDSIS_TBEST_CPFIG.Value
                        FG_PED.TextMatrix(FG_PED.Rows - 1, 3) = .BDSIS_TBEST_CPBIT.Value
                        FG_PED.TextMatrix(FG_PED.Rows - 1, 4) = .BDSIS_TBEST_CPMAT.Value
                    End If
                    .BDSIS_TBEST.Index = sInd
                    FG_PED.TextMatrix(FG_PED.Rows - 1, 5) = Format(.BDSIS_TBPIT_CPPRE.Value, "###,###,##0.00")
                    If IsNull(.BDSIS_TBPIT_CPDES.Value) = False Then
                        FG_PED.TextMatrix(FG_PED.Rows - 1, 6) = Trim(Trim(.BDSIS_TBPIT_CPDES.Value) & Trim(.BDSIS_TBPIT_CPCOM.Value))
                    End If
                End If
            Else
                'Neste caso é quando o pedido foi tirado manualmente, sendo q o BD de Itens de Pedido NAO guarda o Indice da Ficha,  considera 0.
                FG_PED.TextMatrix(FG_PED.Rows - 1, 6) = .BDSIS_TBPIT_CPDES.Value
                FG_PED.TextMatrix(FG_PED.Rows - 1, 5) = Format(.BDSIS_TBPIT_CPPRE.Value, "###,###,##0.00")
            End If
            If .BDSIS_TBPIT_CPCOM.Value <> "" Then FG_PED.TextMatrix(FG_PED.Rows - 1, 7) = Trim(.BDSIS_TBPIT_CPCOM.Value)
            'FG_PED2
            FG_PED2.AddItem (FG_PED2.Rows)
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 0) = Val(tINFPED.NPE)
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 1) = .BDSIS_TBPIT_CPINF.Value
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 2) = tINFPED.CON
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 3) = tINFPED.SPE
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 4) = tINFPED.CPG
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 5) = tINFPED.TRA
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 6) = tINFPED.IVE
            FG_PED2.TextMatrix(FG_PED2.Rows - 1, 7) = .BDSIS_TBPIT_CPIND.Value
        End If
    End With
End Sub
Private Static Sub ImportaPedido()
    For nI = 1 To (FG_P2N.Rows - 1) 'verifica quantidades ou preços
        If Val(FG_P2N.TextMatrix(nI, 1)) <= 0 Or _
           Val(FG_P2N.TextMatrix(nI, 5)) <= 0 Then
           MsgBox "O ítem " & Str(nI) & " deste pedido está com a quantidade ou o preço irregular. Impossível importar este ítem.", vbInformation + vbOKOnly, "Falta dados"
           Exit Sub
        End If
    Next nI
    Me.Hide
    Dim sTmp As String, sTmp2 As String
    sTmp = ""
    sTmp2 = ""
    'verifica as cond.pagto.
    For I = 1 To (FG_P2N2.Rows - 1)
        If sTmp = "" Then
            sTmp = Trim(FG_P2N2.TextMatrix(I, 4))
        ElseIf sTmp <> "" And sTmp <> FG_P2N2.TextMatrix(I, 4) Then
            If sTmp2 = "" Then
                sTmp2 = Trim(Trim(FG_P2N2.TextMatrix(I, 4)))
            Else
                sTmp2 = Trim(sTmp2 & vbCr & Trim(FG_P2N2.TextMatrix(I, 4)))
            End If
        End If
    Next I
    If sTmp2 <> "" Then
        sTmp = InputBox("Existe mais de uma condição de pagamento nestes ítens de pedido. Digite abaixo qual delas você deseja usar:", "Digite Cond.Pagto.", "28")
    End If
    CarregaCP (sTmp)
    'Insere itens de estoque
    For nI = 1 To (FG_P2N.Rows - 1)
        If Trim(FG_P2N.TextMatrix(nI, 2)) <> "" And Trim(FG_P2N.TextMatrix(nI, 3)) <> "" And Trim(FG_P2N.TextMatrix(nI, 4)) <> "" Then
            'este item é de estoque
            With Tela_NotaFiscal_Dlg_1
                .Hide
                .CB_Figura.Text = FG_P2N.TextMatrix(nI, 2)
                .CB_Bitola.Text = FG_P2N.TextMatrix(nI, 3)
                .CB_Material.Text = FG_P2N.TextMatrix(nI, 4)
                .BT_Procurar.Value = True
                .TXT_Quantidade.Text = FG_P2N.TextMatrix(nI, 1)
                .TXT_ValorUnitario.Text = FG_P2N.TextMatrix(nI, 5)
                .CB_Tratamento.Text = FG_P2N.TextMatrix(nI, 7)
                .BT_Inserir.Value = True
                .BT_Cancelar.Value = True
            End With
        Else
            'este item nao tem em estoque
            With Tela_NotaFiscal_Dlg_2
                .Hide
                If FG_P2N.TextMatrix(nI, 7) = "" Then
                    .TXT_Descricao.Text = FG_P2N.TextMatrix(nI, 6)
                Else
                    .TXT_Descricao.Text = FG_P2N.TextMatrix(nI, 6) + " " + FG_P2N.TextMatrix(nI, 7)
                End If
                .TXT_Quantidade.Text = FG_P2N.TextMatrix(nI, 1)
                .TXT_PrecoUnitario.Text = FG_P2N.TextMatrix(nI, 5)
                .TXT_PesoUnitario.Text = "1"
                MsgBox "É necessário completar os dados do ítem que não é de estoque.", vbInformation + vbOKOnly, "Incluir ítem"
                .Show vbModal
            End With
        End If
    Next nI
    'dados adicionais
    sTmp = ""
    sTmp2 = ""
    For I = 1 To (FG_P2N2.Rows - 1) 'Nosso PE
        If sTmp = "" Then
            sTmp = FG_P2N2.TextMatrix(I, 0)
        ElseIf sTmp <> "" And sTmp <> FG_P2N2.TextMatrix(I, 0) Then
            sTmp2 = Trim(sTmp) & "/" & Trim(FG_P2N2.TextMatrix(I, 0))
        End If
    Next I
    If sTmp2 = "" Then
        TXT_PedidoInterno.Text = sTmp
    Else
        TXT_PedidoInterno.Text = sTmp2
    End If
    sTmp = ""
    sTmp2 = ""
    For I = 1 To (FG_P2N2.Rows - 1) 'Seu Pedido
        If sTmp = "" Then
            sTmp = FG_P2N2.TextMatrix(I, 3)
        ElseIf sTmp <> "" And sTmp <> FG_P2N2.TextMatrix(I, 3) Then
            sTmp2 = Trim(sTmp) & "/" & Trim(FG_P2N2.TextMatrix(I, 3))
        End If
    Next I
    If sTmp2 = "" Then
        TXT_SeuPedido.Text = sTmp
    Else
        TXT_SeuPedido.Text = sTmp2
    End If
    sTmp = ""
    sTmp2 = ""
    For I = 1 To (FG_P2N2.Rows - 1) 'Índice Vendedor
        If sTmp = "" Then
            sTmp = FG_P2N2.TextMatrix(I, 6)
        ElseIf sTmp <> "" And sTmp <> FG_P2N2.TextMatrix(I, 6) Then
            MsgBox "Existe ítens desta N.F. que são de vendedores diferentes.", vbInformation + vbOKOnly, "Vendedores diferentes"
            If sTmp2 = "" Then
                sTmp2 = Trim(sTmp) & "/" & Trim(FG_P2N2.TextMatrix(I, 6))
            Else
                sTmp2 = Trim(sTmp2) & "/" & Trim(FG_P2N2.TextMatrix(I, 6))
            End If
        End If
    Next I
    If sTmp2 = "" Then
        CarregaVendedor sTmp
    Else
        CarregaVendedor sTmp2
    End If
    sTmp = ""
    sTmp2 = ""
    For I = 1 To (FG_P2N2.Rows - 1) 'Transportadora
        If sTmp = "" Then
            sTmp = Trim(FG_P2N2.TextMatrix(I, 5))
        ElseIf sTmp <> "" And sTmp <> FG_P2N2.TextMatrix(I, 5) Then
            If sTmp2 = "" Then
                sTmp2 = Trim(Trim(FG_P2N2.TextMatrix(I, 5)))
            Else
                sTmp2 = Trim(sTmp2 & vbCr & Trim(FG_P2N2.TextMatrix(I, 5)))
            End If
        End If
    Next I
    If sTmp2 <> "" Then
        sTmp = InputBox("Existe mais de uma transportadora nestes ítens de pedido. Digite abaixo qual delas você deseja usar:", "Digite Transportadora")
    End If
    LT_NomeTrans.Text = sTmp
    Me.Show
End Sub
Private Static Sub CarregaCP(Valor As String)
    Dim sTmp1 As String, sTmp2 As String, sTmp3 As String, sTmp4 As String, nInd As Integer, sTmp As String
    nInd = 0
    sTmp = ""
    sTmp1 = ""
    sTmp2 = ""
    sTmp3 = ""
    sTmp4 = ""
    For I = 1 To Len(Valor)
        If Mid(Valor, I, 1) = "/" Then
            If nInd = 0 Then
                sTmp1 = sTmp
            ElseIf nInd = 1 Then
                sTmp2 = sTmp
            ElseIf nInd = 2 Then
                sTmp3 = sTmp
            Else
                sTmp4 = sTmp
            End If
            nInd = nInd + 1
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, I, 1)
        End If
    Next I
    If nInd = 0 Then
        sTmp1 = sTmp
    ElseIf nInd = 1 Then
        sTmp2 = sTmp
    ElseIf nInd = 2 Then
        sTmp3 = sTmp
    Else
        sTmp4 = sTmp
    End If
    If sTmp1 = "21" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RD_21dd.Value = True
    ElseIf sTmp1 = "28" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RD_28dd.Value = True
    ElseIf sTmp1 = "30" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RD_30dd.Value = True
    ElseIf sTmp1 = "35" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RD_35dd.Value = True
    ElseIf sTmp1 = "45" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RD_45dd.Value = True
    ElseIf sTmp1 = "28" And sTmp2 = "30" And sTmp3 = "" And sTmp4 = "" Then
        RD_28dd.Value = True
        CK_Desdobrar.Value = 1
        TXT_DVB.Text = "30"
    ElseIf sTmp1 = "28" And sTmp2 = "35" And sTmp3 = "" And sTmp4 = "" Then
        RD_28dd.Value = True
        CK_Desdobrar.Value = 1
        TXT_DVB.Text = "35"
    ElseIf sTmp1 = "À Vista" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RD_AVista.Value = True
    ElseIf sTmp1 = "C/Apres." And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RD_CApres.Value = True
    ElseIf sTmp1 = "S/Venc." And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        RB_SV.Value = True
    Else
        RD_Escdd.Value = True
        CK_Desdobrar.Value = 1
        TXT_Escdd.Text = sTmp1
        TXT_DVB.Text = sTmp2
        TXT_DVC.Text = sTmp3
        TXT_DVD.Text = sTmp4
    End If
End Sub
Private Static Sub CarregaVendedor(Valor As String)
    TXT_VendInterno.Text = Valor
    'If IsNumeric(Valor) = True Then
    '    With DLL_BD
            '.BDSIS_TBFVE.Seek "=", Val(Valor)
            'If .BDSIS_TBFVE.NoMatch Then
                'TXT_VendInterno.Text = Valor
            'Else
                'If .BDSIS_TBFVE_CPTIP.Value = "I" Then
                    'TXT_VendInterno.Text = Valor
                'ElseIf .BDSIS_TBFVE_CPTIP.Value = "E" Then
                    'TXT_VendExterno.Text = Valor
                'End If
                'If IsNull(.BDSIS_TBFVE_CPSET.Value) = False Then TXT_Setor.Text = Trim(.BDSIS_TBFVE_CPSET.Value)
            'End If
    '    End With
    'Else
        
    'End If
End Sub
Private Static Function VerificaPedidoLiquidado(ByVal NumPed As Double) As Boolean
    VerificaPedidoLiquidado = False
    Dim bItemPedLiq As Boolean
    bItemPedLiq = True
    With DLL_BD
        'procura ficha pedido
        .BDSIS_TBPED.Seek "=", NumPed
        If .BDSIS_TBPED.NoMatch = True Then Exit Function
        Dim sTmp As String
        sTmp = ""
        For K = 1 To Len(.BDSIS_TBPED_CPITE.Value)
            If Mid(.BDSIS_TBPED_CPITE.Value, K, 1) = ";" Then
                If VerificaPedidoLiquidado_Aux(Val(sTmp)) = False Then
                    VerificaPedidoLiquidado = False
                    Exit Function
                End If
                sTmp = ""
            Else
                sTmp = sTmp & Mid(.BDSIS_TBPED_CPITE.Value, K, 1)
            End If
        Next K
        If VerificaPedidoLiquidado_Aux(Val(sTmp)) = False Then
            VerificaPedidoLiquidado = False
            Exit Function
        End If
    End With
    'se o pedido for liq
    VerificaPedidoLiquidado = True
End Function
Private Static Function VerificaPedidoLiquidado_Aux(ItemPed As Double) As Boolean
    VerificaPedidoLiquidado_Aux = False
    With DLL_BD
        .BDSIS_TBPIT.Seek "=", ItemPed
        If Not .BDSIS_TBPIT.NoMatch Then
            If .BDSIS_TBPIT_CPLIQ.Value = True Then VerificaPedidoLiquidado_Aux = True
        End If
    End With
End Function
