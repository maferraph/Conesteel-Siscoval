VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Tela_Pedido 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pedido de Estoque"
   ClientHeight    =   5265
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   8445
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar BP 
      Height          =   255
      Left            =   6000
      TabIndex        =   52
      Top             =   5040
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar BS 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   53
      Top             =   5010
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ST 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "1: Importar"
      TabPicture(0)   =   "Tela_Pedido.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FR_1"
      Tab(0).Control(1)=   "BT_Novo"
      Tab(0).Control(2)=   "BT_Voltar"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "2: Dados Gerais"
      TabPicture(1)   =   "Tela_Pedido.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FR(0)"
      Tab(1).Control(1)=   "BT_Imprimir"
      Tab(1).Control(2)=   "BT_Deletar"
      Tab(1).Control(3)=   "BT_Editar"
      Tab(1).Control(4)=   "FR(3)"
      Tab(1).Control(5)=   "FR(1)"
      Tab(1).Control(6)=   "FR(6)"
      Tab(1).Control(7)=   "FR(4)"
      Tab(1).Control(8)=   "FR(5)"
      Tab(1).Control(9)=   "FR(2)"
      Tab(1).Control(10)=   "TXT_Data"
      Tab(1).Control(11)=   "LB0"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "3: Ítens do Pedido"
      TabPicture(2)   =   "Tela_Pedido.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "BT_AdicionaItem"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "FG"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "BT_DetalhesMP"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "BT_RemoveItem"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "BT_AssitenteFigura"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "BT_AlteraItem"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "FR(7)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "4: Outros"
      TabPicture(3)   =   "Tela_Pedido.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FR(8)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "5: Concluir"
      TabPicture(4)   =   "Tela_Pedido.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FR_Imp"
      Tab(4).Control(1)=   "BT_Apagar"
      Tab(4).Control(2)=   "BT_Cancelar"
      Tab(4).Control(3)=   "BT_Pedido"
      Tab(4).Control(4)=   "FR_9_1"
      Tab(4).Control(5)=   "FR_4"
      Tab(4).Control(6)=   "LI"
      Tab(4).ControlCount=   7
      Begin VB.Frame FR 
         Height          =   1215
         Index           =   7
         Left            =   120
         TabIndex        =   100
         Top             =   420
         Width           =   8175
         Begin VB.ComboBox CB_Prazo 
            Height          =   315
            ItemData        =   "Tela_Pedido.frx":008C
            Left            =   2760
            List            =   "Tela_Pedido.frx":00A5
            TabIndex        =   108
            ToolTipText     =   "Selecione ou digite o prazo de entrega par este ítem"
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox CB_Observacoes 
            Height          =   315
            ItemData        =   "Tela_Pedido.frx":00E6
            Left            =   4080
            List            =   "Tela_Pedido.frx":00E8
            TabIndex        =   107
            ToolTipText     =   "Digite ou selecione o complemento ou observações sobre este ítem"
            Top             =   840
            Width           =   3975
         End
         Begin VB.ComboBox CB_Material 
            Height          =   315
            Left            =   2760
            TabIndex        =   106
            ToolTipText     =   "Selecione um material"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox CB_Bitola 
            Height          =   315
            Left            =   1440
            TabIndex        =   105
            ToolTipText     =   "Sselecione uma bitola"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox CB_Figura 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   104
            ToolTipText     =   "Selecione uma figura"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TXT_Quantidade 
            Height          =   285
            Left            =   120
            TabIndex        =   102
            ToolTipText     =   "Digite aqui a quantidade de peças"
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TXT_Nome 
            Height          =   315
            Left            =   4080
            TabIndex        =   101
            ToolTipText     =   "Se não exisitir figura, digite a descrição da peça à ser cotada neste campo."
            Top             =   240
            Width           =   3975
         End
         Begin MSMask.MaskEdBox TXT_Preco 
            Height          =   285
            Left            =   1440
            TabIndex        =   103
            ToolTipText     =   "Valor unitário desta peça"
            Top             =   840
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   20
            Format          =   "$###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descrição:"
            Height          =   195
            Left            =   4080
            TabIndex        =   97
            Top             =   0
            Width           =   765
         End
         Begin VB.Label LB5 
            AutoSize        =   -1  'True
            Caption         =   "Preço:"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   98
            Top             =   600
            Width           =   465
         End
         Begin VB.Label LB9 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            Height          =   195
            Left            =   4080
            TabIndex        =   114
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label LB8 
            AutoSize        =   -1  'True
            Caption         =   "Material:"
            Height          =   195
            Left            =   2760
            TabIndex        =   113
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LB7 
            AutoSize        =   -1  'True
            Caption         =   "Bitola:"
            Height          =   195
            Left            =   1440
            TabIndex        =   112
            Top             =   0
            Width           =   450
         End
         Begin VB.Label LB6 
            AutoSize        =   -1  'True
            Caption         =   "Figura:"
            Height          =   195
            Left            =   120
            TabIndex        =   111
            Top             =   0
            Width           =   495
         End
         Begin VB.Label LB5 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   110
            Top             =   600
            Width           =   870
         End
         Begin VB.Label LB10 
            AutoSize        =   -1  'True
            Caption         =   "Prazo:"
            Height          =   195
            Left            =   2760
            TabIndex        =   109
            Top             =   600
            Width           =   450
         End
      End
      Begin VB.Frame FR_Imp 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   -71880
         TabIndex        =   94
         Top             =   3600
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton BT_Print 
            Caption         =   "I&mprimir"
            Height          =   855
            Left            =   1200
            Picture         =   "Tela_Pedido.frx":00EA
            Style           =   1  'Graphical
            TabIndex        =   96
            ToolTipText     =   "Imprimir Cotação"
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton BT_Cancel 
            Caption         =   "&Cancelar"
            Height          =   855
            Left            =   3360
            Picture         =   "Tela_Pedido.frx":03F4
            Style           =   1  'Graphical
            TabIndex        =   95
            ToolTipText     =   "Cancela operação"
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.Frame FR 
         Height          =   4215
         Index           =   8
         Left            =   -74160
         TabIndex        =   85
         Top             =   480
         Width           =   6735
         Begin VB.TextBox TXT_SNP 
            Height          =   285
            Left            =   3480
            MaxLength       =   15
            TabIndex        =   43
            ToolTipText     =   "Digite aqui o número do Pedido do Cliente"
            Top             =   3840
            Width           =   3135
         End
         Begin VB.TextBox TXT_OBS 
            Height          =   285
            Left            =   120
            TabIndex        =   42
            ToolTipText     =   "Observações sobre o pedido"
            Top             =   3840
            Width           =   3255
         End
         Begin VB.ComboBox CB_Depto 
            Height          =   315
            ItemData        =   "Tela_Pedido.frx":06FE
            Left            =   120
            List            =   "Tela_Pedido.frx":0708
            Style           =   2  'Dropdown List
            TabIndex        =   35
            ToolTipText     =   "Digite aqui o nome do Departamento"
            Top             =   360
            Width           =   3255
         End
         Begin VB.TextBox TXT_Ramal 
            Height          =   285
            Left            =   3480
            TabIndex        =   36
            ToolTipText     =   "Digite aqui o Ramal do Contato"
            Top             =   360
            Width           =   3135
         End
         Begin VB.ComboBox CB_Frete 
            Height          =   315
            ItemData        =   "Tela_Pedido.frx":071D
            Left            =   3480
            List            =   "Tela_Pedido.frx":0727
            Style           =   2  'Dropdown List
            TabIndex        =   40
            ToolTipText     =   "Digite aqui o tipo do Frete"
            Top             =   1800
            Width           =   3135
         End
         Begin VB.ComboBox CB_Vendedor 
            Height          =   315
            ItemData        =   "Tela_Pedido.frx":0747
            Left            =   120
            List            =   "Tela_Pedido.frx":0749
            Style           =   2  'Dropdown List
            TabIndex        =   37
            ToolTipText     =   "Digite aqui o nome do Vendedor"
            Top             =   1080
            Width           =   3255
         End
         Begin VB.ComboBox CB_Descricao 
            Height          =   315
            ItemData        =   "Tela_Pedido.frx":074B
            Left            =   3480
            List            =   "Tela_Pedido.frx":0758
            Style           =   2  'Dropdown List
            TabIndex        =   38
            ToolTipText     =   "Digite aqui o Departamento ou Cargo do Vendedor"
            Top             =   1080
            Width           =   3135
         End
         Begin VB.TextBox TXT_Dados 
            Height          =   1005
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   41
            ToolTipText     =   "Digite aqui Dados Adicionais sobre o pedido"
            Top             =   2520
            Width           =   6495
         End
         Begin MSMask.MaskEdBox TXT_Outras 
            Height          =   285
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   "Digite aqui o valor de outras despesas"
            Top             =   1800
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   503
            _Version        =   393216
            AllowPrompt     =   -1  'True
            AutoTab         =   -1  'True
            MaxLength       =   20
            Format          =   "$###,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Seu Pedido Nº:"
            Height          =   195
            Left            =   3480
            TabIndex        =   115
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Observações:"
            Height          =   195
            Left            =   120
            TabIndex        =   93
            Top             =   3600
            Width           =   990
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Outras Despesas:"
            Height          =   195
            Left            =   120
            TabIndex        =   92
            Top             =   1560
            Width           =   1260
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ramal do Comprador:"
            Height          =   195
            Left            =   3480
            TabIndex        =   91
            Top             =   120
            Width           =   1530
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Frete:"
            Height          =   195
            Left            =   3480
            TabIndex        =   90
            Top             =   1560
            Width           =   405
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            Height          =   195
            Left            =   120
            TabIndex        =   89
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Dados do Vendedor:"
            Height          =   195
            Left            =   3480
            TabIndex        =   88
            Top             =   840
            Width           =   1470
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Dados Adicionais:"
            Height          =   195
            Left            =   120
            TabIndex        =   87
            Top             =   2280
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Departamento do Comprador:"
            Height          =   195
            Left            =   120
            TabIndex        =   86
            Top             =   120
            Width           =   2085
         End
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "&Apagar"
         Height          =   855
         Left            =   -71160
         Picture         =   "Tela_Pedido.frx":0792
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Apaga campos"
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   855
         Left            =   -69480
         Picture         =   "Tela_Pedido.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Cancela operação"
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton BT_Pedido 
         Caption         =   "Concluir"
         Height          =   855
         Left            =   -72840
         Picture         =   "Tela_Pedido.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Concluir o Pedido"
         Top             =   3960
         Width           =   855
      End
      Begin VB.Frame FR_9_1 
         Caption         =   "Executando"
         Height          =   2295
         Left            =   -74760
         TabIndex        =   70
         Top             =   1560
         Width           =   7935
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "12-) Montando e Imprimindo - Romaneio"
            Height          =   195
            Index           =   12
            Left            =   4200
            TabIndex        =   84
            Top             =   1320
            Width           =   2820
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "11-) Montando e Imprimindo - Pedido Estoque"
            Height          =   195
            Index           =   11
            Left            =   4200
            TabIndex        =   83
            Top             =   1080
            Width           =   3225
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "10-) Montando e Imprimindo - Ordem Expedição"
            Height          =   195
            Index           =   10
            Left            =   4200
            TabIndex        =   82
            Top             =   840
            Width           =   3360
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "9-) Montando e Imprimindo - Ordem Montagem"
            Height          =   195
            Index           =   9
            Left            =   4200
            TabIndex        =   81
            Top             =   600
            Width           =   3270
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   12
            Left            =   3960
            Picture         =   "Tela_Pedido.frx":1320
            Stretch         =   -1  'True
            Top             =   1320
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   11
            Left            =   3960
            Picture         =   "Tela_Pedido.frx":1662
            Stretch         =   -1  'True
            Top             =   1080
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   10
            Left            =   3960
            Picture         =   "Tela_Pedido.frx":19A4
            Stretch         =   -1  'True
            Top             =   840
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   9
            Left            =   3960
            Picture         =   "Tela_Pedido.frx":1CE6
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   8
            Left            =   3960
            Picture         =   "Tela_Pedido.frx":2028
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   7
            Left            =   120
            Picture         =   "Tela_Pedido.frx":236A
            Stretch         =   -1  'True
            Top             =   1800
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "7-) Empenha estoque"
            Height          =   195
            Index           =   7
            Left            =   360
            TabIndex        =   80
            Top             =   1800
            Width           =   1515
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "8-) Montando e Imprimindo - Ordem Fabricação"
            Height          =   195
            Index           =   8
            Left            =   4200
            TabIndex        =   79
            Top             =   360
            Width           =   3315
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "6-) Baixando cotação pendente"
            Height          =   195
            Index           =   6
            Left            =   360
            TabIndex        =   78
            Top             =   1560
            Width           =   2235
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "5-) Lançando pedido no mapa de vendas"
            Height          =   195
            Index           =   5
            Left            =   360
            TabIndex        =   77
            Top             =   1320
            Width           =   2925
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "4-) Lançando peças vendidas"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   76
            Top             =   1080
            Width           =   2115
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "3-) Salvando ítens do pedido"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   75
            Top             =   840
            Width           =   2055
         End
         Begin VB.Image IMG 
            Height          =   252
            Index           =   6
            Left            =   120
            Picture         =   "Tela_Pedido.frx":26AC
            Stretch         =   -1  'True
            Top             =   1560
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Image IMG 
            Height          =   252
            Index           =   5
            Left            =   120
            Picture         =   "Tela_Pedido.frx":29EE
            Stretch         =   -1  'True
            Top             =   1320
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Image IMG 
            Height          =   252
            Index           =   4
            Left            =   120
            Picture         =   "Tela_Pedido.frx":2D30
            Stretch         =   -1  'True
            Top             =   1080
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Image IMG 
            Height          =   252
            Index           =   3
            Left            =   120
            Picture         =   "Tela_Pedido.frx":3072
            Stretch         =   -1  'True
            Top             =   840
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Image IMG 
            Height          =   252
            Index           =   2
            Left            =   120
            Picture         =   "Tela_Pedido.frx":33B4
            Stretch         =   -1  'True
            Top             =   600
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Image IMG 
            Height          =   252
            Index           =   1
            Left            =   120
            Picture         =   "Tela_Pedido.frx":36F6
            Stretch         =   -1  'True
            Top             =   360
            Visible         =   0   'False
            Width           =   252
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "2-) Salvando informações sobre o pedido"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   74
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "1-) Conferindo dados digitados"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   73
            Top             =   360
            Width           =   2145
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   13
            Left            =   3960
            Picture         =   "Tela_Pedido.frx":3A38
            Stretch         =   -1  'True
            Top             =   1560
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Image IMG 
            Height          =   255
            Index           =   14
            Left            =   3960
            Picture         =   "Tela_Pedido.frx":3D7A
            Stretch         =   -1  'True
            Top             =   1800
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "13-) Salvando demais informações"
            Height          =   195
            Index           =   13
            Left            =   4200
            TabIndex        =   72
            Top             =   1560
            Width           =   2430
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "14-) Finalizando variáveis e banco de dados"
            Height          =   195
            Index           =   14
            Left            =   4200
            TabIndex        =   71
            Top             =   1800
            Width           =   3120
         End
      End
      Begin VB.Frame FR_4 
         Caption         =   "Imprimir"
         Height          =   855
         Left            =   -74760
         TabIndex        =   69
         ToolTipText     =   "Imprimir via de Ordem de Expedição"
         Top             =   600
         Width           =   7935
         Begin VB.CheckBox CK_Imp_PE 
            Caption         =   "Pedido Estoque"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            ToolTipText     =   "Imprimir via de Pedido de Estoque"
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox CK_Imp_RO 
            Caption         =   "Romaneio"
            Height          =   255
            Left            =   1680
            TabIndex        =   44
            ToolTipText     =   "Imprimir via de Romaneio"
            Top             =   360
            Width           =   1095
         End
         Begin VB.CheckBox CK_Imp_OE 
            Caption         =   "Ordem Expedição"
            Height          =   255
            Left            =   2880
            TabIndex        =   46
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox CK_Imp_OM 
            Caption         =   "Ordem Montagem"
            Enabled         =   0   'False
            Height          =   255
            Left            =   4560
            TabIndex        =   47
            ToolTipText     =   "Imprimir via(s) de Ordem de Montagem"
            Top             =   360
            Width           =   1695
         End
         Begin VB.CheckBox CK_Imp_OF 
            Caption         =   "Ordem Fabricação"
            Enabled         =   0   'False
            Height          =   255
            Left            =   6240
            TabIndex        =   48
            ToolTipText     =   "Imprimir via(s) de Ordem de Fabricação"
            Top             =   360
            Width           =   1600
         End
      End
      Begin VB.CommandButton BT_AlteraItem 
         Caption         =   "Alterar"
         Height          =   735
         Left            =   4920
         Picture         =   "Tela_Pedido.frx":40BC
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Altera o ítem selecionado na lista abaixo"
         Top             =   1740
         Width           =   1575
      End
      Begin VB.Frame FR 
         Caption         =   "Exibir por:"
         Height          =   1575
         Index           =   0
         Left            =   -74880
         TabIndex        =   57
         Top             =   1200
         Width           =   1455
         Begin VB.OptionButton RB_Pendentes 
            Caption         =   "Pe&ndentes"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            ToolTipText     =   "Exibe Pedidos ainda pendentes"
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton RB_Liquidados 
            Caption         =   "L&iquidados"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Exibe Pedidos liquidados"
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton RB_Incompletos 
            Caption         =   "&Incompleto"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            ToolTipText     =   "Exibe Pedidos em aberto (que não foram concluídos)"
            Top             =   1200
            Width           =   1095
         End
         Begin VB.OptionButton RB_Todos 
            Caption         =   "Todo&s"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            ToolTipText     =   "Exibe todos os Pedidos"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton RB_Empresas 
            Caption         =   "&Empresas"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            ToolTipText     =   "Exibe pedidos pelo nome das empresas"
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.CommandButton BT_Imprimir 
         Caption         =   "I&mprimir"
         Height          =   855
         Left            =   -69360
         Picture         =   "Tela_Pedido.frx":43C6
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Imprimir Pedido"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton BT_Deletar 
         Caption         =   "&Deletar"
         Height          =   855
         Left            =   -71160
         Picture         =   "Tela_Pedido.frx":46D0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Deletar Pedido"
         Top             =   4020
         Width           =   855
      End
      Begin VB.CommandButton BT_Editar 
         Caption         =   "&Editar"
         Height          =   855
         Left            =   -72960
         Picture         =   "Tela_Pedido.frx":4B12
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Editar Pedido"
         Top             =   4020
         Width           =   855
      End
      Begin VB.Frame FR_1 
         Caption         =   "Importar dados da Cotação de Estoque:"
         Height          =   2655
         Left            =   -73200
         TabIndex        =   67
         Top             =   780
         Width           =   4815
         Begin VB.TextBox TXT_NumCot 
            Enabled         =   0   'False
            Height          =   285
            Left            =   3000
            TabIndex        =   2
            Top             =   1320
            Width           =   1455
         End
         Begin VB.CommandButton BT_Importar 
            Caption         =   "Importar Cotação de Estoque"
            Height          =   1815
            Left            =   480
            Picture         =   "Tela_Pedido.frx":4F54
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Abre a tela de Cotação de Estoque para poder importar os dados"
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label LB5 
            AutoSize        =   -1  'True
            Caption         =   "Número da Cotação:"
            Height          =   195
            Index           =   2
            Left            =   3000
            TabIndex        =   68
            Top             =   1080
            Width           =   1470
         End
      End
      Begin VB.Frame FR 
         Height          =   1095
         Index           =   3
         Left            =   -74880
         TabIndex        =   64
         Top             =   2820
         Width           =   1455
         Begin VB.OptionButton RB_Emaberto 
            Caption         =   "Em Aberto"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton RB_Liquidado 
            Caption         =   "Liquidado"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton RB_Pendente 
            Caption         =   "Pendente"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Posição:"
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Número:"
         Height          =   3495
         Index           =   1
         Left            =   -73320
         TabIndex        =   63
         Top             =   420
         Width           =   1455
         Begin VB.ListBox LT_NumPed 
            Height          =   3180
            ItemData        =   "Tela_Pedido.frx":55AE
            Left            =   120
            List            =   "Tela_Pedido.frx":55B0
            TabIndex        =   13
            ToolTipText     =   "Selecione o número do Pedido"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Condições Pgto.:"
         Height          =   3495
         Index           =   6
         Left            =   -68160
         TabIndex        =   58
         Top             =   420
         Width           =   1455
         Begin VB.TextBox TXT_D4 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Digite aqui o número de dias para o 4º vencimento"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.TextBox TXT_D3 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Digite aqui o número de dias para o 3º vencimento"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox TXT_D2 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   23
            ToolTipText     =   "Digite aqui o número de dias para o 2º vencimento"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox TXT_D1 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   "Digite aqui o número de dias para o 1º vencimento"
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox CB_CondPagto 
            Height          =   315
            ItemData        =   "Tela_Pedido.frx":55B2
            Left            =   120
            List            =   "Tela_Pedido.frx":55D7
            Style           =   2  'Dropdown List
            TabIndex        =   21
            ToolTipText     =   "Escolha uma condição de pagamento"
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label LB4 
            AutoSize        =   -1  'True
            Caption         =   "4ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   62
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label LB3 
            AutoSize        =   -1  'True
            Caption         =   "3ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   61
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label LB2 
            AutoSize        =   -1  'True
            Caption         =   "2ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   60
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label LB1 
            AutoSize        =   -1  'True
            Caption         =   "1ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   720
            Width           =   375
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Empresa:"
         Height          =   3495
         Index           =   4
         Left            =   -71760
         TabIndex        =   56
         Top             =   420
         Width           =   1695
         Begin VB.CommandButton BT_CadEmp 
            Caption         =   "Cadastro de Empresas"
            Height          =   495
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   "Abre a tela de cadastro de empresas"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.ListBox LT_Empresa 
            Height          =   2205
            ItemData        =   "Tela_Pedido.frx":5643
            Left            =   120
            List            =   "Tela_Pedido.frx":5645
            TabIndex        =   15
            ToolTipText     =   "Selecione uma empresa"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TXT_Empresa 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   14
            ToolTipText     =   "Digite aqui o nome da empresa caso não exista na lista abaixo"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Transportadora:"
         Height          =   1695
         Index           =   5
         Left            =   -69960
         TabIndex        =   55
         Top             =   2220
         Width           =   1695
         Begin VB.ListBox LT_Transportadora 
            Height          =   1035
            ItemData        =   "Tela_Pedido.frx":5647
            Left            =   120
            List            =   "Tela_Pedido.frx":5649
            TabIndex        =   20
            ToolTipText     =   "Selecione a transportadora deste cliente"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TXT_Transportadora 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   19
            ToolTipText     =   "Digite aqui o nome da tranportadora caso não esteja selecionada na lista abaixo"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Contato:"
         Height          =   1695
         Index           =   2
         Left            =   -69960
         TabIndex        =   54
         Top             =   420
         Width           =   1695
         Begin VB.ListBox LT_Contato 
            Height          =   1035
            ItemData        =   "Tela_Pedido.frx":564B
            Left            =   120
            List            =   "Tela_Pedido.frx":564D
            TabIndex        =   18
            ToolTipText     =   "Selecione o contato da empresa"
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TXT_Contato 
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   17
            ToolTipText     =   "Digite aqui o nome do contato da empresa se não existir na lista abaixo"
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton BT_AssitenteFigura 
         Caption         =   "Assistente Figuras"
         Height          =   735
         Left            =   120
         Picture         =   "Tela_Pedido.frx":564F
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Use o Assistente de Figuras de Estoque caso você não conheça o sistema de figuras"
         Top             =   1740
         Width           =   1575
      End
      Begin VB.CommandButton BT_RemoveItem 
         Caption         =   "Remover"
         Height          =   735
         Left            =   3360
         Picture         =   "Tela_Pedido.frx":5959
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Remove o ítem selecionado na lista abaixo"
         Top             =   1740
         Width           =   1575
      End
      Begin VB.CommandButton BT_DetalhesMP 
         Caption         =   "Detalhes M.P."
         Enabled         =   0   'False
         Height          =   735
         Left            =   6600
         Picture         =   "Tela_Pedido.frx":5C63
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Abre detalhes da matéria-prima do ítem selecionado"
         Top             =   1740
         Width           =   1695
      End
      Begin VB.CommandButton BT_Novo 
         Caption         =   "&Novo"
         Height          =   855
         Left            =   -72600
         Picture         =   "Tela_Pedido.frx":5F6D
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Novo Pedido"
         Top             =   3900
         Width           =   855
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   855
         Left            =   -69720
         Picture         =   "Tela_Pedido.frx":6277
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Volta à Tela Principal"
         Top             =   3900
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   2415
         Left            =   120
         TabIndex        =   34
         ToolTipText     =   "Lista de ítens desta cotação"
         Top             =   2580
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4260
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.CommandButton BT_AdicionaItem 
         Caption         =   "Adicionar"
         Height          =   735
         Left            =   1800
         Picture         =   "Tela_Pedido.frx":66B9
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Adiciona o ítem na lista abaixo"
         Top             =   1740
         Width           =   1575
      End
      Begin MSComctlLib.ImageList LI 
         Left            =   -74880
         Top             =   4200
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
               Picture         =   "Tela_Pedido.frx":69C3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Tela_Pedido.frx":6CDD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Tela_Pedido.frx":6FF7
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSMask.MaskEdBox TXT_Data 
         Height          =   330
         Left            =   -74880
         TabIndex        =   99
         ToolTipText     =   "Data da Cotação"
         Top             =   720
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         BackColor       =   -2147483633
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "99/99/9999"
         PromptChar      =   "_"
      End
      Begin VB.Label LB0 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Da&ta:"
         Height          =   195
         Left            =   -74880
         TabIndex        =   66
         Top             =   540
         Width           =   390
      End
   End
End
Attribute VB_Name = "Tela_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** VARIÁVEIS DLL's ****************
Public DLL_BD As Scvbd.Classe_Scvbd
Public DLL_CARGA As Scvcarr.Classe_Scvcarr
Public DLL_FUNCS As Scvfunc.Classe_Scvfunc
Public DLL_ASFIG As Assfig.Classe_Assfig
Public DLL_COT As Cotest.Classe_Cotest
Public DLL_IMP As Impform.Classe_Impform
Public DLL_CADEMP As Cademp.Classe_Cademp

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Pedido de Estoque"
Dim RespMsg, I As Integer, ESTIND As String, sESTADO As String, J As Integer
Dim FICEST As T_FICEST, nNumPed As Long, sItens As String, bEstadoEdicao As Boolean, sIndBak As String, nNumOE As Long
Public sFax As String, sEmail As String, bCarregaPedido As Boolean
Private nLin As Integer, nInd As Integer, lTeste As Boolean, lFuncTeste As Boolean, sTxtMsg As String, sEmpresa As String, lImprimirPedido As Boolean
Private Type T_FICEST
    FIG As String
    BIT As String
    MAT As String
    NOM As String
    VUN As Currency
    VMI As Currency
    VCU As Currency
    PUN As Long
    EST As Long
    EMI As Long
    VEN As Long
    COT As Long
    INQ As Integer
    INP As Integer
    INN As Integer
    INB As Integer
    INM As Integer
End Type
Private Sub BT_AdicionaItem_Click()
    'On Error GoTo ERRO_SISCOVAL
    'Testes de preenchimento
    ST.Tab = 2
    If CB_Figura.Text = "" And TXT_Nome.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Figura.Text <> "" And CB_Bitola.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Figura.Text <> "" And CB_Material.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Material.SetFocus
        Exit Sub
    ElseIf TXT_Quantidade.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        TXT_Quantidade.SetFocus
        Exit Sub
    ElseIf TXT_Preco.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        TXT_Preco.SetFocus
        Exit Sub
    ElseIf CB_Prazo.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Prazo.SetFocus
        Exit Sub
    End If
    
    TelaEmEspera True
    FG.AddItem (FG.Rows)
    FG.TextMatrix(FG.Rows - 1, 1) = TXT_Quantidade.Text
    FG.TextMatrix(FG.Rows - 1, 2) = CB_Figura.Text
    FG.TextMatrix(FG.Rows - 1, 3) = CB_Bitola.Text
    FG.TextMatrix(FG.Rows - 1, 4) = CB_Material.Text
    FG.TextMatrix(FG.Rows - 1, 5) = TXT_Nome.Text
    FG.TextMatrix(FG.Rows - 1, 6) = CB_Observacoes.Text
    FG.TextMatrix(FG.Rows - 1, 7) = VerificaPrazo("COM")
    FG.TextMatrix(FG.Rows - 1, 8) = VerificaPrazo("PRO")
    FG.TextMatrix(FG.Rows - 1, 9) = VerificaPrazo("MAP")
    FG.TextMatrix(FG.Rows - 1, 10) = Format(TXT_Preco.Text, "###,###,###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 11) = CB_Prazo.Text
    FG.TextMatrix(FG.Rows - 1, 12) = AliquotaImposto("IPI")
    FG.TextMatrix(FG.Rows - 1, 13) = AliquotaImposto("ICMS")
    FG.TextMatrix(FG.Rows - 1, 14) = "Não"
    CB_Figura.Text = ""
    TXT_Nome.Text = ""
    TelaEmEspera False
    CB_Figura.SetFocus
'ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_AlteraItem_Click()
    On Error GoTo ERRO_SISCOVAL
    'Testes de preenchimento
    ST.Tab = 2
    If CB_Figura.Text = "" And TXT_Nome.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Figura.Text <> "" And CB_Bitola.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Figura.Text <> "" And CB_Material.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Material.SetFocus
        Exit Sub
    ElseIf TXT_Quantidade.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        TXT_Quantidade.SetFocus
        Exit Sub
    ElseIf TXT_Preco.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        TXT_Preco.SetFocus
        Exit Sub
    End If
    TelaEmEspera True
    FG.TextMatrix(FG.RowSel, 1) = TXT_Quantidade.Text
    FG.TextMatrix(FG.RowSel, 2) = CB_Figura.Text
    FG.TextMatrix(FG.RowSel, 3) = CB_Bitola.Text
    FG.TextMatrix(FG.RowSel, 4) = CB_Material.Text
    FG.TextMatrix(FG.RowSel, 5) = TXT_Nome.Text
    FG.TextMatrix(FG.RowSel, 6) = CB_Observacoes.Text
    FG.TextMatrix(FG.RowSel, 7) = VerificaPrazo("COM")
    FG.TextMatrix(FG.RowSel, 8) = VerificaPrazo("PRO")
    FG.TextMatrix(FG.RowSel, 9) = VerificaPrazo("MAP")
    FG.TextMatrix(FG.RowSel, 10) = Format(TXT_Preco.Text, "###,###,###,##0.00")
    FG.TextMatrix(FG.RowSel, 11) = CB_Prazo.Text
    FG.TextMatrix(FG.RowSel, 12) = AliquotaImposto("IPI")
    FG.TextMatrix(FG.RowSel, 13) = AliquotaImposto("ICMS")
    FG.TextMatrix(FG.RowSel, 14) = "Não"
    CB_Figura.Text = ""
    TelaEmEspera False
    CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    TXT_NumCot.Text = ""
    TXT_Data.Text = "__/__/____"
    CB_CondPagto.ListIndex = -1
    TXT_D1.Text = ""
    TXT_D2.Text = ""
    TXT_D3.Text = ""
    TXT_D4.Text = ""
    TXT_Empresa.Text = ""
    LT_NumPed.ListIndex = -1
    LT_Empresa.ListIndex = -1
    TXT_Contato.Text = ""
    LT_Contato.ListIndex = -1
    TXT_Transportadora.Text = ""
    LT_Transportadora.ListIndex = -1
    CB_Figura.Text = ""
    CB_Bitola.Text = ""
    CB_Material.Text = ""
    TXT_Quantidade.Text = ""
    TXT_Preco.Text = ""
    TXT_Nome.Text = ""
    CB_Prazo.Text = ""
    CB_Observacoes.Text = ""
    TXT_OBS.Text = ""
    CB_Depto.ListIndex = -1
    TXT_Ramal.Text = ""
    CB_Vendedor.ListIndex = -1
    CB_Descricao.ListIndex = -1
    TXT_Outras.Text = ""
    CB_Frete.ListIndex = -1
    TXT_Dados.Text = ""
    TXT_OBS.Text = ""
    CK_Imp_PE.Value = 0
    CK_Imp_RO.Value = 0
    CK_Imp_OE.Value = 0
    CK_Imp_OM.Value = 0
    CK_Imp_OF.Value = 0
    If bCarregaPedido = False Then
        RB_Todos.Value = False
        RB_Empresas.Value = False
        RB_Pendentes.Value = False
        RB_Liquidados.Value = False
        RB_Incompletos.Value = False
        RB_Pendente.Value = False
        RB_Liquidado.Value = False
        RB_Emaberto.Value = False
    End If
    FG.Clear
    MontaFG
    Limpa_IMG_LB
    ResetaBSEP
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_AssitenteFigura_Click()
    On Error GoTo ERRO_SISCOVAL
    CB_Figura.Text = DLL_ASFIG.AssistenteFigura(App.ProductName, "Assfig", App.LegalCopyright)
    CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_CadEmp_Click()
    Dim sEmpTemp1 As String, sEmpTemp2 As String
    If LT_Empresa.ListIndex > -1 Then
        sEmpTemp1 = LT_Empresa.Text
    Else
        sEmpTemp1 = ""
    End If
    Me.Hide
    sEmpTemp2 = DLL_CADEMP.CadastroEmpresa(App.ProductName, "Cademp", App.LegalCopyright, sEmpTemp1)
    If sEmpTemp2 <> "" Then
        'reCarregando combo de empresas
        If DLL_BD.BDSIS_TBEMP.RecordCount > 0 Then
            LT_Empresa.Clear
            LT_Transportadora.Clear
            ResetaBSEP
            ResetaBP (DLL_BD.BDSIS_TBEMP.RecordCount + 1)
            DLL_BD.BDSIS_TBEMP.MoveFirst
            While Not DLL_BD.BDSIS_TBEMP.EOF
                If DLL_BD.BDSIS_TBEMP_CPAPE.Value <> "" Then
                    LT_Empresa.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
                End If
                If DLL_BD.BDSIS_TBEMP_CPTIP.Value = "Transportadora" Then
                    LT_Transportadora.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
                End If
                CarregaBSEP ("Recarregando lista de empresas...")
                DLL_BD.BDSIS_TBEMP.MoveNext
            Wend
        End If
        ResetaBSEP
        LT_Empresa.Text = sEmpTemp2
        LT_Transportadora.Text = sEmpTemp2
    End If
    Me.Show vbModal
End Sub
Private Sub BT_Cancel_Click()
    ST.TabEnabled(0) = True
    ST.TabEnabled(1) = True
    ST.TabEnabled(2) = True
    ST.TabEnabled(3) = True
    ST.TabEnabled(4) = True
    FR_Imp.Visible = False
    TelaEmEdicao False
    BT_Apagar.Value = True
    ST.Tab = 0
    LT_NumPed.Clear
    BT_Voltar.SetFocus
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEdicao False
    BT_Apagar.Value = True
    ST.Tab = 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Importar_Click()
    'importa dados de Pedidos
    TXT_NumCot.Text = ""
    Me.Hide
    'retorna número do Pedido
    TXT_NumCot.Text = DLL_COT.ImportaCotacao(App.ProductName, "Cotest", App.LegalCopyright)
    If TXT_NumCot.Text <> "" Then
        If IsNumeric(TXT_NumCot.Text) = True Then
            If Val(TXT_NumCot.Text) > 0 Then
                TelaEmEspera True
                ImportaCotacao
                TelaEmEspera False
            End If
        End If
    End If
    Me.Show vbModal
End Sub
Private Sub BT_Pedido_Click()
'    On Error GoTo ERRO_1
    ResetaBP (14)
    lImprimirPedido = False
    'Reseta Operações
    Limpa_IMG_LB
    '***********************************************************
    '1-) verifica se todos campos estao preenchidos
    '***********************************************************
    CarregaBSEP ("Verificando informações preenchidas...")
    Muda_IMG_LB 1, 2
    If LT_Empresa.Text = "" And TXT_Empresa.Text = "" Then
        Concluir_Erro1
        ST.Tab = 1
        LT_Empresa.SetFocus
        Exit Sub
    ElseIf LT_Contato.Text = "" And TXT_Contato.Text = "" Then
        Concluir_Erro1
        ST.Tab = 1
        LT_Contato.SetFocus
        Exit Sub
    ElseIf LT_Transportadora.Text = "" And TXT_Transportadora.Text = "" Then
        Concluir_Erro1
        ST.Tab = 1
        LT_Transportadora.SetFocus
        Exit Sub
    ElseIf TXT_D1.Text = "" And TXT_D2.Text = "" And TXT_D3.Text = "" And TXT_D4.Text = "" Then
        Concluir_Erro1
        ST.Tab = 1
        TXT_D1.SetFocus
        Exit Sub
    ElseIf FG.Rows <= 1 Then
        Concluir_Erro1
        ST.Tab = 2
        FG.SetFocus
        Exit Sub
    ElseIf CB_Vendedor.ListIndex = -1 Then
        Concluir_Erro1
        ST.Tab = 3
        CB_Vendedor.SetFocus
        Exit Sub
    ElseIf TXT_SNP.Text = "" Then
        Concluir_Erro1
        ST.Tab = 3
        TXT_SNP.SetFocus
        Exit Sub
    End If
    Muda_IMG_LB 1, 3
    'começa salvar dados
    TelaEmEspera True
    '***********************************************************
    '2-) Salvando informações sobre o pedido
    '***********************************************************
    CarregaBSEP ("Salvando informações sobre o pedido...")
    Muda_IMG_LB 2, 2
    With DLL_BD
        'procura indice do vendedor
        Dim nVend As Integer
        nVend = 0
        If .BDSIS_TBFUN.RecordCount > 0 Then
            .BDSIS_TBFUN.MoveFirst
            Do While Not .BDSIS_TBFUN.EOF
                If .BDSIS_TBFUN_CPFUN.Value = CB_Vendedor.Text Then
                    nVend = .BDSIS_TBFUN_CPIND.Value
                    Exit Do
                End If
                .BDSIS_TBFUN.MoveNext
            Loop
        End If
    End With
    If bEstadoEdicao = False Then 'Pedido Novo
        With DLL_BD
            'adiciona pedido
            .BDSIS_TBPED.AddNew
            .BDSIS_TBPED_CPDAT.Value = Format(Date, "dd/mm/yyyy")
            .BDSIS_TBPED_CPHOR.Value = Format(Time, "hh:mm:ss")
            .BDSIS_TBPED_CPEMP.Value = Trim(TXT_Empresa.Text)
            .BDSIS_TBPED_CPCON.Value = Trim(TXT_Contato.Text)
            .BDSIS_TBPED_CPCPG.Value = PegaCP(False)
            .BDSIS_TBPED_CPTRA.Value = Trim(TXT_Transportadora.Text)
            .BDSIS_TBPED_CPVAL.Value = CalculaPrecoTotal
            .BDSIS_TBPED_CPIVE.Value = nVend
            .BDSIS_TBPED_CPNSP.Value = Trim(TXT_SNP.Text)
            .BDSIS_TBPED_CPIFR.Value = CB_Frete.ListIndex
            .BDSIS_TBPED_CPOUD.Value = TXT_Outras.Text
            .BDSIS_TBPED_CPDAD.Value = TXT_Dados.Text
            .BDSIS_TBPED_CPOBS.Value = TXT_OBS.Text
            .BDSIS_TBPED_CPLIQ.Value = False
            nNumPed = .BDSIS_TBPED_CPIND.Value
            Muda_IMG_LB 2, 3
            '***********************************************************
            '3-) Salvando ítens do pedido
            '***********************************************************
            CarregaBSEP ("Salvando ítens do pedido...")
            Muda_IMG_LB 3, 2
            sItens = ""
            For I = 1 To FG.Rows - 1
                .BDSIS_TBPIT.AddNew
                If FG.TextMatrix(I, 2) <> "" And FG.TextMatrix(I, 3) <> "" And FG.TextMatrix(I, 4) <> "" Then
                    .BDSIS_TBPIT_CPINF.Value = PegaIndFic(I)
                Else
                    .BDSIS_TBPIT_CPDES.Value = Trim(FG.TextMatrix(I, 5))
                End If
                If FG.TextMatrix(I, 6) <> "" Then .BDSIS_TBPIT_CPCOM.Value = FG.TextMatrix(I, 6)
                .BDSIS_TBPIT_CPQUA.Value = FG.TextMatrix(I, 1)
                .BDSIS_TBPIT_CPNPE.Value = nNumPed
                .BDSIS_TBPIT_CPPRE.Value = FG.TextMatrix(I, 10)
                .BDSIS_TBPIT_CPPRA.Value = FG.TextMatrix(I, 11)
                .BDSIS_TBPIT_CPLIQ.Value = False
                If sItens = "" Then
                    sItens = .BDSIS_TBPIT_CPIND.Value
                Else
                    sItens = sItens & ";" & .BDSIS_TBPIT_CPIND.Value
                End If
                .BDSIS_TBPIT.Update
            Next I
            .BDSIS_TBPED_CPITE.Value = sItens
            .BDSIS_TBPED.Update
            Muda_IMG_LB 3, 3
        End With
        '***********************************************************
        '4-) Lançando peças vendidas
        '***********************************************************
        CarregaBSEP ("Lançando peças vendidas...")
        Muda_IMG_LB 4, 2
        lTeste = True
        For I = 1 To FG.Rows - 1
            lFuncTeste = LancaPedidos(I)
            If lFuncTeste = False Then lTeste = False
        Next I
        Muda_IMG_LB 4, 1
        If lTeste = True Then Muda_IMG_LB 4, 3
        '***********************************************************
        '5-) Lançando pedido no mapa de vendas
        '***********************************************************
        CarregaBSEP ("Lançando pedido no mapa de vendas...")
        Muda_IMG_LB 5, 2
        lTeste = LancaMapaPedido
        Muda_IMG_LB 5, 1
        If lTeste = True Then Muda_IMG_LB 5, 3
    ElseIf bEstadoEdicao = True Then 'Editar Pedido
        With DLL_BD
            Dim cValorVelho As Currency, sMesAno As String
            sMesAno = DLL_FUNCS.NomeMes(Month(.BDSIS_TBPED_CPDAT.Value)) & "/" & Year(.BDSIS_TBPED_CPDAT.Value)
            cValorVelho = .BDSIS_TBPED_CPVAL.Value
            .BDSIS_TBPED.Edit
            .BDSIS_TBPED_CPDAT.Value = Format(Date, "dd/mm/yyyy")
            .BDSIS_TBPED_CPHOR.Value = Format(Time, "hh:mm:ss")
            .BDSIS_TBPED_CPEMP.Value = Trim(TXT_Empresa.Text)
            .BDSIS_TBPED_CPCON.Value = Trim(TXT_Contato.Text)
            .BDSIS_TBPED_CPCPG.Value = PegaCP(False)
            .BDSIS_TBPED_CPTRA.Value = Trim(TXT_Transportadora.Text)
            .BDSIS_TBPED_CPVAL.Value = CalculaPrecoTotal
            .BDSIS_TBPED_CPIVE.Value = nVend
            .BDSIS_TBPED_CPNSP.Value = Trim(TXT_SNP.Text)
            .BDSIS_TBPED_CPABE.Value = False
            nNumPed = .BDSIS_TBPED_CPIND.Value
            sItens = .BDSIS_TBPED_CPITE.Value
            Muda_IMG_LB 2, 3
            '***********************************************************
            '3-) Salvando ítens do pedido
            '***********************************************************
            CarregaBSEP ("Salvando ítens do pedido...")
            Muda_IMG_LB 3, 2
            'retira saldo de Pedidos
            BS.SimpleText = "Retirando saldos de ítens pedidos..."
            RetiraPedidos sItens
            'apaga itens do Pedido velha
            BS.SimpleText = "Retirando ítens antigos do pedido..."
            ApagaItensPedido sItens
            'salva itens do Pedido editado
            BS.SimpleText = "Salvando dados dos ítens do Pedido editado..."
            sItens = ""
            For I = 1 To FG.Rows - 1
                .BDSIS_TBPIT.AddNew
                If FG.TextMatrix(I, 2) <> "" And FG.TextMatrix(I, 3) <> "" And FG.TextMatrix(I, 4) <> "" Then
                    .BDSIS_TBPIT_CPINF.Value = PegaIndFic(I)
                Else
                    .BDSIS_TBPIT_CPDES.Value = Trim(FG.TextMatrix(I, 5))
                End If
                If FG.TextMatrix(I, 6) <> "" Then .BDSIS_TBPIT_CPCOM.Value = FG.TextMatrix(I, 6)
                .BDSIS_TBPIT_CPQUA.Value = FG.TextMatrix(I, 1)
                .BDSIS_TBPIT_CPNPE.Value = nNumPed
                .BDSIS_TBPIT_CPPRE.Value = FG.TextMatrix(I, 10)
                .BDSIS_TBPIT_CPPRA.Value = FG.TextMatrix(I, 11)
                .BDSIS_TBPIT_CPLIQ.Value = False
                If sItens = "" Then
                    sItens = .BDSIS_TBPIT_CPIND.Value
                Else
                    sItens = sItens & ";" & .BDSIS_TBPIT_CPIND.Value
                End If
                .BDSIS_TBPIT.Update
            Next I
            .BDSIS_TBPED_CPITE.Value = sItens
            .BDSIS_TBPED.Update
            Muda_IMG_LB 3, 3
        End With
        '***********************************************************
        '4-) Lançando peças vendidas
        '***********************************************************
        CarregaBSEP ("Lançando peças vendidas...")
        Muda_IMG_LB 4, 2
        lTeste = True
        For I = 1 To FG.Rows - 1
            lFuncTeste = LancaPedidos(I)
            If lFuncTeste = False Then lTeste = False
        Next I
        Muda_IMG_LB 4, 1
        If lTeste = True Then Muda_IMG_LB 4, 3
        '***********************************************************
        '5-) Lançando pedido no mapa de vendas
        '***********************************************************
        CarregaBSEP ("Lançando pedido no mapa de vendas...")
        Muda_IMG_LB 5, 2
        RetiraMapaPedido cValorVelho, sMesAno
        lTeste = LancaMapaPedido
        Muda_IMG_LB 5, 1
        If lTeste = True Then Muda_IMG_LB 5, 3
    End If
    ST.Tab = 4
    '***********************************************************
    '6-) Baixando cotação pendente
    '***********************************************************
    CarregaBSEP ("Baixando cotação pendente...")
    Muda_IMG_LB 6, 2
    lTeste = BaixaCotacao
    Muda_IMG_LB 6, 1
    If lTeste = True Then Muda_IMG_LB 6, 3
    '***********************************************************
    '7-) Empenha estoque
    '***********************************************************
    CarregaBSEP ("Empenha estoque...")
    Muda_IMG_LB 7, 2
    lTeste = EmpenhaEstoque
    Muda_IMG_LB 7, 1
    If lTeste = True Then Muda_IMG_LB 7, 3
    '***********************************************************
    '8-) Montando e Imprimindo - Ordem Fabricação
    '***********************************************************
    CarregaBSEP ("Montando e Imprimindo - Ordem Fabricação...")
    Muda_IMG_LB 8, 2
    lTeste = False
    If CK_Imp_OF.Value = 1 Then lTeste = MI_OF
    Muda_IMG_LB 8, 1
    If lTeste = True Then Muda_IMG_LB 8, 3
    '***********************************************************
    '9-) Montando e Imprimindo - Ordem Montagem
    '***********************************************************
    CarregaBSEP ("Montando e Imprimindo - Ordem Montagem...")
    Muda_IMG_LB 9, 2
    lTeste = False
    If CK_Imp_OM.Value = 1 Then lTeste = MI_OM
    Muda_IMG_LB 9, 1
    If lTeste = True Then Muda_IMG_LB 9, 3
    '***********************************************************
    '10-) Montando e Imprimindo - Ordem Expedição
    '***********************************************************
    CarregaBSEP ("Montando e Imprimindo - Ordem Expedição...")
    Muda_IMG_LB 10, 2
    lTeste = False
    If CK_Imp_OE.Value = 1 Then lTeste = MI_OE
    Muda_IMG_LB 10, 1
    If lTeste = True Then Muda_IMG_LB 10, 3
    '***********************************************************
    '11-) Montando e Imprimindo - Pedido Estoque
    '***********************************************************
    CarregaBSEP ("Montando e Imprimindo - Pedido Estoque...")
    Muda_IMG_LB 11, 2
    lTeste = False
    If CK_Imp_PE.Value = 1 Then lTeste = MI_PE
    Muda_IMG_LB 11, 1
    If lTeste = True Then Muda_IMG_LB 11, 3
    '***********************************************************
    '12-) Montando e Imprimindo - Romaneio
    '***********************************************************
    CarregaBSEP ("Montando e Imprimindo - Romaneio...")
    Muda_IMG_LB 12, 2
    lTeste = False
    If CK_Imp_RO.Value = 1 Then lTeste = MI_RO
    Muda_IMG_LB 12, 1
    If lTeste = True Then Muda_IMG_LB 12, 3
    '***********************************************************
    '13-) Salvando demais informações
    '***********************************************************
    CarregaBSEP ("Salvando demais informações...")
    Muda_IMG_LB 13, 2
    'verifica se a empresa já está cadastrada
    If LT_Empresa.ListIndex = -1 Or LT_Empresa.Text <> TXT_Empresa.Text And TXT_Empresa.Text <> "" Then
        With DLL_BD
            .BDSIS_TBEMP.Seek "=", TXT_Empresa.Text
            If .BDSIS_TBEMP.NoMatch Then
                .BDSIS_TBEMP.AddNew
                .BDSIS_TBEMP_CPAPE.Value = TXT_Empresa.Text
                .BDSIS_TBEMP.Update
                LT_Empresa.AddItem TXT_Empresa.Text
            End If
        End With
    End If
    'verifica o nome do contato
    If LT_Contato.ListIndex = -1 Or LT_Contato.Text <> TXT_Contato.Text And TXT_Contato.Text <> "" Then
        With DLL_BD
            .BDSIS_TBECO.Seek "=", TXT_Empresa.Text, TXT_Contato.Text
            If .BDSIS_TBECO.NoMatch Then
                .BDSIS_TBECO.AddNew
                .BDSIS_TBECO_CPEMP.Value = TXT_Empresa.Text
                .BDSIS_TBECO_CPCON.Value = TXT_Contato.Text
                .BDSIS_TBECO.Update
                LT_Contato.AddItem TXT_Contato.Text
            End If
        End With
    End If
    'verifica se a transportadora já está cadastrada
    If LT_Transportadora.ListIndex = -1 Or LT_Transportadora.Text <> TXT_Transportadora.Text And TXT_Transportadora.Text <> "" Then
        With DLL_BD
            .BDSIS_TBEMP.Seek "=", TXT_Transportadora.Text
            If .BDSIS_TBEMP.NoMatch Then
                .BDSIS_TBEMP.AddNew
                .BDSIS_TBEMP_CPAPE.Value = TXT_Transportadora.Text
                .BDSIS_TBEMP.Update
                LT_Transportadora.AddItem TXT_Transportadora.Text
            End If
        End With
    End If
    Muda_IMG_LB 13, 3
    '***********************************************************
    '14-) Finalizando variáveis e banco de dados
    '***********************************************************
    CarregaBSEP ("Finalizando variáveis e banco de dados...")
    Muda_IMG_LB 14, 3
ERRO_1:
    ResetaBSEP
    BT_Apagar.Value = True
    TelaEmEdicao False
    TelaEmEspera False
    LT_NumPed.Clear
    ST.Tab = 0
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NumPed.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar um pedido na lista de números de Pedidos para poder continuar com esta operação.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    RespMsg = MsgBox("Você tem certeza que deseja apagar o Pedido de nº " & Trim(LT_NumPed.Text) & " do banco de dados ?", vbQuestion + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then
        TelaEmEspera True
        ResetaBSEP
        ResetaBP (3)
        With DLL_BD
            BS.SimpleText = "Procurando Pedido para deletar..."
            'retira mapa de Pedido
            Dim cValorVelho As Currency, sMesAno As String
            sMesAno = DLL_FUNCS.NomeMes(Month(.BDSIS_TBPED_CPDAT.Value)) & "/" & Year(.BDSIS_TBPED_CPDAT.Value)
            cValorVelho = .BDSIS_TBPED_CPVAL.Value
            RetiraMapaPedido cValorVelho, sMesAno
            'retira saldo de Pedido
            RetiraPedidos .BDSIS_TBPED_CPITE.Value
            'apago Pedido
            .BDSIS_TBPED.Delete
            CarregaBSEP ("Procurando ítens do Pedido para deletar...")
            ApagaItensPedido sItens
            CarregaBSEP ("Procurando ítens do Pedido para deletar...")
        End With
        LT_NumPed.RemoveItem LT_NumPed.ListIndex
        BT_Apagar_Click
        LT_Contato.Clear
        LT_Transportadora.ListIndex = -1
        CarregaBSEP ("Limpando campos da tela...")
        TelaEmEspera False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_DetalhesMP_Click()
    ST.Tab = 2
    If CB_Figura.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Bitola.SetFocus
        Exit Sub
    ElseIf CB_Material.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Material.SetFocus
        Exit Sub
    ElseIf TXT_Quantidade.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder incluir um ítem.", vbOKOnly + vbInformation, NOMEAPLIC
        TXT_Quantidade.SetFocus
        Exit Sub
    End If
    Tela_Pedido_MP.DetalhesMP TXT_Quantidade.Text, CB_Figura.Text, CB_Bitola.Text, CB_Material.Text, (Trim(TXT_Nome.Text) & " " & Trim(CB_Observacoes.Text))
    Tela_Pedido_MP.Show vbModal
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NumPed.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar um pedido na lista de números de Pedidos para poder continuar com esta operação.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    TelaEmEdicao True
    FR(1).Enabled = False
    FR(3).Enabled = False
    LT_NumPed.Enabled = False
    bEstadoEdicao = True
    PreValores 2
    LT_Empresa.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Imprimir_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NumPed.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar um pedido na lista de números de Pedidos para poder continuar com esta operação.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    TelaEmEdicao True
    ST.Tab = 4
    ST.TabEnabled(0) = False
    ST.TabEnabled(1) = False
    ST.TabEnabled(2) = False
    ST.TabEnabled(3) = True
    ST.TabEnabled(4) = True
    FR_Imp.Visible = True
    FR_Imp.Enabled = True
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    BT_Apagar.Value = True
    TelaEmEdicao True
    RB_Todos.Value = False
    RB_Empresas.Value = False
    RB_Pendente.Value = True
    FR(1).Enabled = False
    FR(3).Enabled = False
    LT_NumPed.Enabled = False
    bEstadoEdicao = False
    PreValores 1
    nNumPed = 0
    ST.Tab = 1
    LT_Empresa.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Print_Click()
    lTeste = False
    RespMsg = MsgBox("Imprimir agora as vias selecionadas acima ?", vbQuestion + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then
        TelaEmEspera True
        lImprimirPedido = True
        ResetaBSEP
        ResetaBP (12)
        Limpa_IMG_LB
        For I = 1 To 7
            IMG(I).Picture = LI.ListImages(1).Picture
            IMG(I).Visible = True
            CarregaBSEP ("")
        Next I
        '8-) Montando e Imprimindo - Ordem Fabricação
        CarregaBSEP ("")
        Muda_IMG_LB 8, 2
        lTeste = False
        If CK_Imp_OF.Value = 1 Then lTeste = MI_OF
        Muda_IMG_LB 8, 1
        If lTeste = True Then Muda_IMG_LB 8, 3
        '9-) Montando e Imprimindo - Ordem Montagem
        CarregaBSEP ("")
        Muda_IMG_LB 9, 2
        lTeste = False
        If CK_Imp_OM.Value = 1 Then lTeste = MI_OM
        Muda_IMG_LB 9, 1
        If lTeste = True Then Muda_IMG_LB 9, 3
        '10-) Montando e Imprimindo - Ordem Expedição
        CarregaBSEP ("")
        Muda_IMG_LB 10, 2
        lTeste = False
        If CK_Imp_OE.Value = 1 Then lTeste = MI_OE
        Muda_IMG_LB 10, 1
        If lTeste = True Then Muda_IMG_LB 10, 3
        '11-) Montando e Imprimindo - Pedido Estoque
        CarregaBSEP ("")
        Muda_IMG_LB 11, 2
        lTeste = False
        If CK_Imp_PE.Value = 1 Then lTeste = MI_PE
        Muda_IMG_LB 11, 1
        If lTeste = True Then Muda_IMG_LB 11, 3
        '12-) Montando e Imprimindo - Romaneio
        CarregaBSEP ("")
        Muda_IMG_LB 12, 2
        lTeste = False
        If CK_Imp_RO.Value = 1 Then lTeste = MI_RO
        Muda_IMG_LB 12, 1
        If lTeste = True Then Muda_IMG_LB 12, 3
        'finalizando
        Muda_IMG_LB 13, 1
        Muda_IMG_LB 14, 1
        TelaEmEspera False
        BT_Cancel_Click
    End If
End Sub
Private Sub BT_RemoveItem_Click()
    If FG.RowSel > 0 Then
        If FG.Rows = 2 And FG.RowSel = 1 Then
            MontaFG
        Else
            FG.RemoveItem (FG.RowSel)
        End If
    Else
        MsgBox "Não existe linha ou não foi selecionada uma na lista abaixo.", vbInformation + vbOKOnly, NOMEAPLIC
    End If
End Sub
Private Sub BT_Voltar_Click()
    Unload Tela_Pedido
End Sub
Private Sub CB_Bitola_GotFocus()
    If CB_Figura.Text = "" Then
        MsgBox "Selecione primeiro uma figura.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Figura.SetFocus
    End If
    CB_Bitola.SelLength = Len(CB_Bitola.Text)
End Sub
Private Sub CB_Bitola_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And CB_Bitola.Text <> "" Then CB_Material.SetFocus
End Sub
Private Sub CB_Bitola_LostFocus()
    CB_Bitola.Text = UCase(CB_Bitola.Text)
    If CB_Bitola.Text <> "" Then
        For I = 0 To CB_Bitola.ListCount - 1
            If CB_Bitola.Text = CB_Bitola.List(I) Then
                Exit For
            ElseIf CB_Bitola.Text <> CB_Bitola.List(I) And I = CB_Bitola.ListCount - 1 Then
                MsgBox "Essa bitola digitada não existe - consulte esta lista.", vbOKOnly + vbInformation, NOMEAPLIC
                CB_Bitola.SetFocus
                Exit Sub
            End If
        Next I
    End If
    ProcuraFicha
End Sub
Private Sub CB_CondPagto_Click()
    TXT_D1.Text = ""
    TXT_D2.Text = ""
    TXT_D3.Text = ""
    TXT_D4.Text = ""
    TXT_D1.Enabled = False
    TXT_D2.Enabled = False
    TXT_D3.Enabled = False
    TXT_D4.Enabled = False
    If CB_CondPagto.ListIndex = 0 Then
        TXT_D1.Text = "21"
    ElseIf CB_CondPagto.ListIndex = 1 Then
        TXT_D1.Text = "28"
    ElseIf CB_CondPagto.ListIndex = 2 Then
        TXT_D1.Text = "28"
        TXT_D2.Text = "30"
    ElseIf CB_CondPagto.ListIndex = 3 Then
        TXT_D1.Text = "28"
        TXT_D2.Text = "35"
    ElseIf CB_CondPagto.ListIndex = 4 Then
        TXT_D1.Text = "30"
    ElseIf CB_CondPagto.ListIndex = 5 Then
        TXT_D1.Text = "35"
    ElseIf CB_CondPagto.ListIndex = 6 Then
        TXT_D1.Text = "45"
    ElseIf CB_CondPagto.ListIndex = 7 Then
        TXT_D1.Text = "À Vista"
    ElseIf CB_CondPagto.ListIndex = 8 Then
        TXT_D1.Text = "C/Apres."
    ElseIf CB_CondPagto.ListIndex = 9 Then
        TXT_D1.Text = "S/Venc."
    ElseIf CB_CondPagto.ListIndex = 10 Then
        TXT_D1.Enabled = True
        TXT_D2.Enabled = True
        TXT_D3.Enabled = True
        TXT_D4.Enabled = True
    End If
End Sub
Private Sub CB_CondPagto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        ST.Tab = 1
        CB_Figura.SetFocus
    End If
End Sub
Private Sub CB_Depto_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_Ramal.SetFocus
End Sub
Private Sub CB_Descricao_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_Outras.SetFocus
End Sub
Private Sub CB_Figura_Click()
    CarregaFIGBITMAT
    ProcuraFicha
End Sub
Private Sub CB_Figura_GotFocus()
    CB_Figura.SelLength = Len(CB_Figura.Text)
    ZeraCampos 0
End Sub
Private Sub CB_Figura_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And CB_Figura.Text <> "" Then
        CB_Bitola.SetFocus
    ElseIf KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        CB_Figura_Click
    End If
End Sub
Private Sub CB_Figura_LostFocus()
    CB_Figura_Click
End Sub
Private Sub CB_Frete_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_Dados.SetFocus
End Sub
Private Sub CB_Material_Change()
    CB_Material.SelLength = Len(CB_Material.Text)
End Sub
Private Sub CB_Material_GotFocus()
    If CB_Figura.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma figura.", vbOKOnly + vbInformation, NOMEAPLIC)
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma bitola.", vbOKOnly + vbInformation, NOMEAPLIC)
        CB_Bitola.SetFocus
        Exit Sub
    End If
End Sub
Private Sub CB_Material_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn And CB_Material.Text <> "" Then TXT_Quantidade.SetFocus
End Sub
Private Sub CB_Material_LostFocus()
    CB_Material.Text = UCase(CB_Material.Text)
    If CB_Material.Text <> "" Then
        For I = 0 To CB_Material.ListCount - 1
            If CB_Material.Text = CB_Material.List(I) Then
                Exit For
            ElseIf CB_Material.Text <> CB_Material.List(I) And I = CB_Material.ListCount - 1 Then
                MsgBox "Esse material digitado não existe - consulte esta lista.", vbOKOnly + vbInformation, NOMEAPLIC
                CB_Material.SetFocus
                Exit Sub
            End If
        Next I
    End If
    ProcuraFicha
End Sub
Private Sub CB_Observacoes_GotFocus()
    CB_Observacoes.SelLength = Len(CB_Observacoes.Text)
End Sub
Private Sub CB_Observacoes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then BT_AdicionaItem.SetFocus
End Sub
Private Sub CB_Prazo_GotFocus()
    CB_Prazo.SelLength = Len(CB_Prazo.Text)
End Sub
Private Sub CB_Prazo_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_Observacoes.SetFocus
End Sub
Private Sub CB_Vendedor_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_Descricao.SetFocus
End Sub
Private Sub CK_Imp_OE_Click()
    If CK_Imp_OE.Value = 1 Then
        CK_Imp_RO.Value = 1
    Else
        CK_Imp_RO.Value = 0
    End If
End Sub
Private Sub CK_Imp_RO_Click()
    If CK_Imp_RO.Value = 1 Then
        CK_Imp_OE.Value = 1
    Else
        CK_Imp_OE.Value = 0
    End If
End Sub
Private Sub FG_Click()
    If FG.Rows > 1 Then
        TelaEmEspera True
        TXT_Quantidade.Text = FG.TextMatrix(FG.RowSel, 1)
        CB_Figura.Text = FG.TextMatrix(FG.RowSel, 2)
        CB_Bitola.Text = FG.TextMatrix(FG.RowSel, 3)
        CB_Material.Text = FG.TextMatrix(FG.RowSel, 4)
        CarregaFIGBITMAT FG.TextMatrix(FG.RowSel, 3), FG.TextMatrix(FG.RowSel, 4)
        TXT_Nome.Text = FG.TextMatrix(FG.RowSel, 5)
        CB_Observacoes.Text = FG.TextMatrix(FG.RowSel, 6)
        TXT_Preco.Text = FG.TextMatrix(FG.RowSel, 10)
        CB_Prazo.Text = FG.TextMatrix(FG.RowSel, 11)
        TelaEmEspera False
    End If
End Sub
Private Sub FG_SelChange()
    FG_Click
End Sub
Private Sub Form_Load()
    On Error GoTo ERRO_SISCOVAL
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    Set DLL_CARGA = New Scvcarr.Classe_Scvcarr
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    Set DLL_ASFIG = New Assfig.Classe_Assfig
    Set DLL_COT = New Cotest.Classe_Cotest
    Set DLL_IMP = New Impform.Classe_Impform
    Set DLL_CADEMP = New Cademp.Classe_Cademp

    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (55)
    DLL_CARGA.ResetaBP
    
    On Error GoTo ERRO_ACESSO_BANCODADOS
    'Abre bancos de dados
    DLL_CARGA.CarregaTexto ("Abrindo banco de dados...")
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    DLL_CARGA.CarregaTexto ("Abrindo tabela Grupos...")
    If DLL_BD.AbreTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque...")
    If DLL_BD.AbreTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Índice...")
    If DLL_BD.AbreTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Figuras...")
    If DLL_BD.AbreTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Complementos...")
    If DLL_BD.AbreTabela_EstoqueComplementos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - CF e ST...")
    If DLL_BD.AbreTabela_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Estoque - Alíquotas...")
    If DLL_BD.AbreTabela_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Empresas...")
    If DLL_BD.AbreTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Empresas - Contatos...")
    If DLL_BD.AbreTabela_EmpresasContatos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Quantidades...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeQuantidades(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Pecas...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDePecas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Nomes...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeNomes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Bitolas...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeBitolas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Índice de Materiais...")
    If DLL_BD.AbreTabela_MateriaPrimaIndiceDeMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Matéria-Prima - Relação de Materiais...")
    If DLL_BD.AbreTabela_MateriaPrimaRelacaoMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Pedidos...")
    If DLL_BD.AbreTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Pedidos - Ítens...")
    If DLL_BD.AbreTabela_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Mapa - Pedidos...")
    If DLL_BD.AbreTabela_MapaPedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Cotações...")
    If DLL_BD.AbreTabela_Cotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Cotações - Ítens...")
    If DLL_BD.AbreTabela_CotacoesItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Ordem de Fabricação...")
'    If DLL_BD.AbreTabela_OrdemFabricacao(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Ordem de Montagem...")
'    If DLL_BD.AbreTabela_OrdemMontagem(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
'    DLL_CARGA.CarregaTexto ("Abrindo tabela Ordem de Expedição...")
    If DLL_BD.AbreTabela_OrdemExpedicao(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Funcionários...")
    If DLL_BD.AbreTabela_Funcionarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abre Campos
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Grupos...")
    If DLL_BD.AbreCampos_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque...")
    If DLL_BD.AbreCampos_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Índice...")
    If DLL_BD.AbreCampos_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Figuras...")
    If DLL_BD.AbreCampos_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Complementos...")
    If DLL_BD.AbreCampos_EstoqueComplementos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - CF e ST...")
    If DLL_BD.AbreCampos_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Estoque - Alíquotas...")
    If DLL_BD.AbreCampos_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Empresas...")
    If DLL_BD.AbreCampos_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Empresas - Contatos...")
    If DLL_BD.AbreCampos_EmpresasContatos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Quantidades...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeQuantidades(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Pecas...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDePecas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Nomes...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeNomes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Bitolas...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeBitolas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Índice de Materiais...")
    If DLL_BD.AbreCampos_MateriaPrimaIndiceDeMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Matéria-Prima - Relação de Materiais...")
    If DLL_BD.AbreCampos_MateriaPrimaRelacaoMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Pedidos...")
    If DLL_BD.AbreCampos_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Pedidos - Ítens...")
    If DLL_BD.AbreCampos_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Mapa - Pedidos...")
    If DLL_BD.AbreCampos_MapaPedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Cotações...")
    If DLL_BD.AbreCampos_Cotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Cotações - Ítens...")
    If DLL_BD.AbreCampos_CotacoesItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Ordem de Fabricação...")
'    If DLL_BD.AbreCampos_OrdemFabricacao(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Ordem de Montagem...")
'    If DLL_BD.AbreCampos_OrdemMontagem(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Ordem de Expedição...")
    If DLL_BD.AbreCampos_OrdemExpedicao(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Funcionários...")
    If DLL_BD.AbreCampos_Funcionarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS

    'On Error GoTo ERRO_SISCOVAL
    LT_Empresa.Clear
    LT_Contato.Clear
    LT_Transportadora.Clear
    'Carrega lista de figuras
    DLL_CARGA.CarregaTexto ("Carregando lista de figuras...")
    CB_Figura.Clear
    If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEFG.MoveFirst
        While Not DLL_BD.BDSIS_TBEFG.EOF
            If DLL_BD.BDSIS_TBEFG_CPFIG.Value <> "" Then
                CB_Figura.AddItem (DLL_BD.BDSIS_TBEFG_CPFIG.Value)
            End If
            DLL_BD.BDSIS_TBEFG.MoveNext
        Wend
    End If
    'Carregando combo de empresas
    DLL_CARGA.CarregaTexto ("Carregando lista de empresas...")
    If DLL_BD.BDSIS_TBEFG.RecordCount > 0 Then
        DLL_BD.BDSIS_TBEMP.MoveFirst
        While Not DLL_BD.BDSIS_TBEMP.EOF
            If DLL_BD.BDSIS_TBEMP_CPAPE.Value <> "" Then
                LT_Empresa.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            If DLL_BD.BDSIS_TBEMP_CPTIP.Value = "Transportadora" Then
                LT_Transportadora.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
            End If
            DLL_BD.BDSIS_TBEMP.MoveNext
        Wend
    End If
    'Carregando combo de complementos
    DLL_CARGA.CarregaTexto ("Carregando lista de complementos...")
    CB_Observacoes.Clear
    If DLL_BD.BDSIS_TBECM.RecordCount > 0 Then
        DLL_BD.BDSIS_TBECM.MoveFirst
        While Not DLL_BD.BDSIS_TBECM.EOF
            CB_Observacoes.AddItem (DLL_BD.BDSIS_TBECM_CPCOM.Value)
            DLL_BD.BDSIS_TBECM.MoveNext
        Wend
    End If
    'Carregando combo de vendedores
    DLL_CARGA.CarregaTexto ("Carregando lista de vendedores...")
    CB_Vendedor.Clear
    If DLL_BD.BDSIS_TBFUN.RecordCount > 0 Then
        DLL_BD.BDSIS_TBFUN.MoveFirst
        While Not DLL_BD.BDSIS_TBFUN.EOF
            CB_Vendedor.AddItem (DLL_BD.BDSIS_TBFUN_CPFUN.Value)
            DLL_BD.BDSIS_TBFUN.MoveNext
        Wend
    End If
    
    BT_Apagar.Value = True
    TelaEmEdicao False
    Limpa_IMG_LB
    FR_Imp.Visible = False
    FR_Imp.Left = 1680
    FR_Imp.Top = 3960
    ResetaBSEP
    ST.Tab = 0
    bCarregaPedido = False
    TXT_SNP.Text = "verbal"
    NUMPED = ""
    DLL_CARGA.CarregaTexto ("Finalizando")
    DLL_FUNCS.RegistraEvento "Abrir Pedidos de Estoque", ""
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Pedido
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueComplementos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.AbreTabela_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.AbreTabela_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EmpresasContatos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueComplementos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeQuantidades(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDePecas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeNomes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeBitolas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaRelacaoMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MapaPedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Cotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_CotacoesItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_OrdemFabricacao(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_OrdemMontagem(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_OrdemExpedicao(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Funcionarios(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
    Set DLL_FUNCS = Nothing
    Set DLL_ASFIG = Nothing
    Set DLL_COT = Nothing
    Set DLL_IMP = Nothing
    Set DLL_CADEMP = Nothing
End Sub

Private Sub LT_Contato_Click()
    If TXT_Contato.Text <> LT_Contato.Text And LT_Contato.ListIndex >= 0 Then TXT_Contato.Text = LT_Contato.Text
End Sub
Private Sub LT_Contato_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_Transportadora.SetFocus
End Sub
Private Sub LT_Empresa_Click()
    If BT_Novo.Enabled = True And LT_Empresa.Enabled = True Then
        ResetaBSEP
        ResetaBP (3)
        CarregaBSEP ("Limpando campos...")
        LimpaCamposPedido
        CarregaBSEP ("Procurando contatos desta empresa...")
        CarregaContatosEmpresa
        CarregaBSEP ("Procurando Pedidos desta empresa...")
        If LT_NumPed.ListIndex = -1 Then CarregaPedidosPorEmpresa
        ResetaBSEP
    Else
        If TXT_Empresa.Text <> LT_Empresa.Text And LT_Empresa.ListIndex >= 0 Then TXT_Empresa.Text = LT_Empresa.Text
        ProcuraEstado
        ProcuraContato
    End If
End Sub
Private Sub LT_NumPed_Click()
    TelaEmEspera True
    bCarregaPedido = True
    CarregaPedidos
    bCarregaPedido = False
    If LT_NumPed.ListIndex > -1 Then nNumPed = LT_NumPed.Text
    TelaEmEspera False
End Sub
Private Sub LT_Transportadora_Click()
    If TXT_Transportadora.Text <> LT_Transportadora.Text And LT_Transportadora.ListIndex >= 0 Then TXT_Transportadora.Text = LT_Transportadora.Text
End Sub
Private Sub LT_Transportadora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_CondPagto.SetFocus
End Sub
Private Sub RB_Empresas_Click()
    bCarregaPedido = True
    BT_Apagar.Value = True
    bCarregaPedido = False
    LT_NumPed.Clear
    FR(1).Enabled = True
    FR(4).Enabled = True
    LT_NumPed.Enabled = True
    LT_Empresa.Enabled = True
    LT_Empresa.SetFocus
End Sub
Private Sub RB_Empresas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_Empresa.SetFocus
End Sub
Private Sub RB_Incompletos_Click()
    TelaEmEspera True
    bCarregaPedido = True
    BT_Apagar.Value = True
    bCarregaPedido = False
    LT_NumPed.Clear
    'carrego Pedidos em aberto
    With DLL_BD
        If .BDSIS_TBPED.RecordCount > 0 Then
            ResetaBSEP
            ResetaBP (.BDSIS_TBPED.RecordCount)
            .BDSIS_TBPED.MoveFirst
            Do While Not .BDSIS_TBPED.EOF
                If .BDSIS_TBPED_CPABE.Value = True Then LT_NumPed.AddItem .BDSIS_TBPED_CPIND.Value
                .BDSIS_TBPED.MoveNext
                CarregaBSEP ("Carregando Pedidos em aberto (incompletos)...")
            Loop
        End If
    End With
    'habilita listas
    BS.SimpleText = "Finalizando..."
    FR(1).Enabled = True
    FR(4).Enabled = False
    LT_NumPed.Enabled = True
    LT_Empresa.Enabled = False
    ResetaBSEP
    LT_Empresa.ListIndex = -1
    TelaEmEspera False
    LT_NumPed.SetFocus
End Sub
Private Sub RB_Incompletos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_NumPed.SetFocus
End Sub
Private Sub RB_Liquidados_Click()
    TelaEmEspera True
    bCarregaPedido = True
    BT_Apagar.Value = True
    bCarregaPedido = False
    LT_NumPed.Clear
    'carrego Pedidos liquidados
    With DLL_BD
        If .BDSIS_TBPED.RecordCount > 0 Then
            ResetaBSEP
            ResetaBP (.BDSIS_TBPED.RecordCount)
            .BDSIS_TBPED.MoveFirst
            Do While Not .BDSIS_TBPED.EOF
                If .BDSIS_TBPED_CPLIQ.Value = True Then LT_NumPed.AddItem .BDSIS_TBPED_CPIND.Value
                .BDSIS_TBPED.MoveNext
                CarregaBSEP ("Carregando Pedidos liquidados...")
            Loop
        End If
    End With
    'habilita listas
    BS.SimpleText = "Finalizando..."
    FR(1).Enabled = True
    FR(4).Enabled = False
    LT_NumPed.Enabled = True
    LT_Empresa.Enabled = False
    ResetaBSEP
    LT_Empresa.ListIndex = -1
    TelaEmEspera False
    LT_NumPed.SetFocus
End Sub
Private Sub RB_Liquidados_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_NumPed.SetFocus
End Sub
Private Sub RB_Pendentes_Click()
    TelaEmEspera True
    bCarregaPedido = True
    BT_Apagar.Value = True
    bCarregaPedido = False
    LT_NumPed.Clear
    'carrego Pedidos pendentes
    With DLL_BD
        If .BDSIS_TBPED.RecordCount > 0 Then
            ResetaBSEP
            ResetaBP (.BDSIS_TBPED.RecordCount)
            .BDSIS_TBPED.MoveFirst
            Do While Not .BDSIS_TBPED.EOF
                If .BDSIS_TBPED_CPLIQ.Value = False Then LT_NumPed.AddItem .BDSIS_TBPED_CPIND.Value
                .BDSIS_TBPED.MoveNext
                CarregaBSEP ("Carregando Pedidos pendentes...")
            Loop
        End If
    End With
    'habilita listas
    BS.SimpleText = "Finalizando..."
    FR(1).Enabled = True
    FR(4).Enabled = False
    LT_NumPed.Enabled = True
    LT_Empresa.Enabled = False
    ResetaBSEP
    LT_Empresa.ListIndex = -1
    TelaEmEspera False
    LT_NumPed.SetFocus
End Sub
Private Sub RB_Pendentes_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_NumPed.SetFocus
End Sub
Private Sub RB_Todos_Click()
    TelaEmEspera True
    bCarregaPedido = True
    BT_Apagar.Value = True
    bCarregaPedido = False
    LT_NumPed.Clear
    'carrego Pedidos
    With DLL_BD
        If .BDSIS_TBPED.RecordCount > 0 Then
            ResetaBSEP
            ResetaBP (.BDSIS_TBPED.RecordCount)
            .BDSIS_TBPED.MoveFirst
            Do While Not .BDSIS_TBPED.EOF
                LT_NumPed.AddItem .BDSIS_TBPED_CPIND.Value
                .BDSIS_TBPED.MoveNext
                CarregaBSEP ("Carregando Pedidos...")
            Loop
        End If
    End With
    'habilita listas
    BS.SimpleText = "Finalizando..."
    FR(1).Enabled = True
    FR(4).Enabled = False
    LT_NumPed.Enabled = True
    LT_Empresa.Enabled = False
    ResetaBSEP
    LT_Empresa.ListIndex = -1
    TelaEmEspera False
    LT_NumPed.SetFocus
End Sub
Private Sub RB_Todos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_NumPed.SetFocus
End Sub
Private Sub TXT_Contato_Change()
    TXT_Contato.Text = UCase(TXT_Contato.Text)
    If LT_Contato.ListIndex = -1 Then LT_Contato.Text = TXT_Contato.Text
End Sub
Private Sub TXT_Contato_KeyPress(KeyAscii As Integer)
    If LT_Contato.ListIndex >= 0 Then LT_Contato.ListIndex = -1
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then TXT_Contato.SetFocus
End Sub
Private Sub TXT_D1_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
End Sub
Private Sub TXT_D2_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
End Sub
Private Sub TXT_D3_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
End Sub
Private Sub TXT_D4_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
End Sub
Private Sub TXT_Dados_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_OBS.SetFocus
End Sub
Private Sub TXT_Empresa_Change()
    TXT_Empresa.Text = UCase(TXT_Empresa.Text)
    If LT_Empresa.ListIndex = -1 Then LT_Empresa.Text = TXT_Empresa.Text
End Sub
Private Sub TXT_Empresa_KeyPress(KeyAscii As Integer)
    If LT_Empresa.ListIndex >= 0 Then LT_Empresa.ListIndex = -1
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then TXT_Transportadora.SetFocus
End Sub
Private Sub TXT_Nome_GotFocus()
    ZeraCampos 1
    TXT_Nome.SelLength = Len(TXT_Nome.Text)
End Sub
Private Sub TXT_Nome_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn Then TXT_Quantidade.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub TXT_OBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then TXT_SNP.SetFocus
End Sub
Private Sub TXT_Outras_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CB_Frete.SetFocus
End Sub
Private Sub TXT_Preco_GotFocus()
    TXT_Preco.SelLength = Len(TXT_Preco.Text)
End Sub
Private Sub TXT_Preco_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then CB_Prazo.SetFocus
End Sub
Private Sub TXT_Quantidade_GotFocus()
    TXT_Quantidade.SelLength = Len(TXT_Quantidade.Text)
End Sub
Private Sub TXT_Quantidade_KeyPress(KeyAscii As Integer)
    KeyAscii = DLL_FUNCS.ValidaTexto(KeyAscii)
    If KeyAscii = vbKeyReturn Then TXT_Preco.SetFocus
End Sub
Private Sub TXT_Quantidade_LostFocus()
    If TXT_Quantidade.Text <> "" Then If (CDbl(Val(TXT_Quantidade.Text)) < CDbl(Val(FICEST.EST))) Then CB_Prazo.ListIndex = 0
End Sub
Private Sub TXT_Ramal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_Vendedor.SetFocus
End Sub
Private Sub TXT_SNP_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then ST.Tab = 4
End Sub
Private Sub TXT_Transportadora_Change()
    TXT_Transportadora.Text = UCase(TXT_Transportadora.Text)
    If LT_Transportadora.ListIndex = -1 Then LT_Transportadora.Text = TXT_Transportadora.Text
End Sub
Private Sub TXT_Transportadora_KeyPress(KeyAscii As Integer)
    If LT_Transportadora.ListIndex >= 0 Then LT_Transportadora.ListIndex = -1
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************
Private Static Sub TelaEmEspera(Estado As Boolean)
    On Error GoTo ERRO_SISCOVAL
    If Estado = True Then
        Me.MousePointer = vbHourglass
        Me.Enabled = False
    Else
        Me.MousePointer = vbDefault
        Me.Enabled = True
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub MontaFG()
    On Error GoTo ERRO_SISCOVAL
    FG.Cols = 15
    FG.Rows = 1
    
    FG.ColAlignment(0) = flexAlignCenterCenter
    FG.ColAlignment(1) = flexAlignCenterCenter
    FG.ColAlignment(2) = flexAlignLeftCenter
    FG.ColAlignment(3) = flexAlignLeftCenter
    FG.ColAlignment(4) = flexAlignLeftCenter
    FG.ColAlignment(5) = flexAlignLeftCenter
    FG.ColAlignment(6) = flexAlignLeftCenter
    FG.ColAlignment(7) = flexAlignLeftCenter
    FG.ColAlignment(8) = flexAlignLeftCenter
    FG.ColAlignment(9) = flexAlignLeftCenter
    FG.ColAlignment(10) = flexAlignCenterCenter
    FG.ColAlignment(11) = flexAlignLeftCenter
    FG.ColAlignment(12) = flexAlignCenterCenter
    FG.ColAlignment(13) = flexAlignCenterCenter
    FG.ColAlignment(14) = flexAlignCenterCenter
    
    FG.ColWidth(0) = 500
    FG.ColWidth(1) = 1000
    FG.ColWidth(2) = 1200
    FG.ColWidth(3) = 1200
    FG.ColWidth(4) = 1200
    FG.ColWidth(5) = 3500
    FG.ColWidth(6) = 1800
    FG.ColWidth(7) = 1200
    FG.ColWidth(8) = 1200
    FG.ColWidth(9) = 1200
    FG.ColWidth(10) = 1200
    FG.ColWidth(11) = 1200
    FG.ColWidth(12) = 800
    FG.ColWidth(13) = 800
    FG.ColWidth(14) = 800

    FG.TextArray(0) = "Item"
    FG.TextArray(1) = "Quantidade"
    FG.TextArray(2) = "Figura"
    FG.TextArray(3) = "Bitola"
    FG.TextArray(4) = "Material"
    FG.TextArray(5) = "Descrição"
    FG.TextArray(6) = "Complemento"
    FG.TextArray(7) = "Componentes"
    FG.TextArray(8) = "Produção"
    FG.TextArray(9) = "Matéria-Prima"
    FG.TextArray(10) = "Preço Unitário"
    FG.TextArray(11) = "Prazo Entrega"
    FG.TextArray(12) = "I.P.I."
    FG.TextArray(13) = "I.C.M.S."
    FG.TextArray(14) = "Liquidado"
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEdicao(HabilitadoEdicao As Boolean)
    'diversos
    Dim MeuControle As Control
    For Each MeuControle In Tela_Pedido.Controls
        If TypeOf MeuControle Is Label Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is TextBox Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is ComboBox Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is Frame Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is ListBox Then MeuControle.Enabled = HabilitadoEdicao
    Next MeuControle
    'Tab 0
    BT_Novo.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Voltar.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Importar.Enabled = HabilitadoEdicao
    TXT_NumCot.Enabled = False
    'Tab 1
    TXT_Data.Enabled = HabilitadoEdicao
    FR(0).Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Todos.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Empresas.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Pendentes.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Liquidados.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Incompletos.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Pendente.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Liquidado.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Emaberto.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    TXT_D1.Enabled = False
    TXT_D2.Enabled = False
    TXT_D3.Enabled = False
    TXT_D4.Enabled = False
    TXT_Data.Enabled = False
    BT_Editar.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Deletar.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Imprimir.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    'Tab 2
    BT_AssitenteFigura.Enabled = HabilitadoEdicao
    BT_AdicionaItem.Enabled = HabilitadoEdicao
    BT_RemoveItem.Enabled = HabilitadoEdicao
    BT_AlteraItem.Enabled = HabilitadoEdicao
'    BT_DetalhesMP.Enabled = HabilitadoEdicao
    TXT_Preco.Enabled = HabilitadoEdicao
    FG.Enabled = True
    'Tab 3
    TXT_Outras.Enabled = HabilitadoEdicao
    'ja desabilitou tudo
    'Tab 4
    CK_Imp_PE.Enabled = HabilitadoEdicao
    CK_Imp_RO.Enabled = HabilitadoEdicao
    CK_Imp_OE.Enabled = HabilitadoEdicao
'    CK_Imp_OM.Enabled = HabilitadoEdicao
'    CK_Imp_OF.Enabled = HabilitadoEdicao
    BT_Pedido.Enabled = HabilitadoEdicao
    BT_Apagar.Enabled = HabilitadoEdicao
    BT_Cancelar.Enabled = HabilitadoEdicao
End Sub
Private Static Sub CarregaFIGBITMAT(Optional ColocaBIT As String, Optional ColocaMAT As String)
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then Exit Sub
    CB_Figura.Text = UCase(CB_Figura.Text)
    For I = 0 To CB_Figura.ListCount - 1
        If CB_Figura.Text = CB_Figura.List(I) Then
            Exit For
        ElseIf CB_Figura.Text <> CB_Figura.List(I) And I = CB_Figura.ListCount - 1 Then
            MsgBox "Essa figura digitada não existe - consulte esta lista.", vbOKOnly + vbInformation, NOMEAPLIC
            CB_Figura.SetFocus
            Exit Sub
        End If
    Next I
    'procura figura
    DLL_BD.BDSIS_TBEFG.Seek "=", CB_Figura.Text
    've se o indice de figura da nova consulta é igual a figura anteriormente consultada
    If DLL_BD.BDSIS_TBEFG_CPIFG.Value = ESTIND And CB_Bitola.ListCount > 1 Then Exit Sub
    ESTIND = DLL_BD.BDSIS_TBEFG_CPIFG.Value
    DLL_BD.BDSIS_TBEID.Seek "=", DLL_BD.BDSIS_TBEFG_CPIFG.Value
    If DLL_BD.BDSIS_TBEFG.NoMatch And DLL_BD.BDSIS_TBEID.NoMatch Then
        MsgBox "Ocorreu algum erro durante a procura do índice da figura.", vbOKOnly + vbInformation, NOMEAPLIC
        Exit Sub
    End If
    'pega nome peça
    TXT_Nome.Text = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & Trim(DLL_BD.BDSIS_TBEID_CPTNO.Value) & " " & Trim(DLL_BD.BDSIS_TBEFG_CPCOM.Value)
    'Como são tabelas relacionadas, a procura acima ja acha o indice de figura
    Dim cA As String
    'Montando lista de bitolas
    cA = ""
    CB_Bitola.Clear
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGBI.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGBI.Value, I, 1) = ";" Then
            CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Bitola.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    If ColocaBIT <> "" Then CB_Bitola.Text = ColocaBIT
    'Montando lista de materiais
    cA = ""
    CB_Material.Clear
    For I = 1 To Len(DLL_BD.BDSIS_TBEID_CPGMA.Value)
        If Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) <> ";" Then
            cA = cA & Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1)
        ElseIf Mid(DLL_BD.BDSIS_TBEID_CPGMA.Value, I, 1) = ";" Then
            CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
            cA = ""
        End If
    Next I
    CB_Material.AddItem (DLL_FUNCS.ProcuraGrupo(cA))
    'seleciona material A-105
    CB_Material.ListIndex = 0
    If ColocaMAT <> "" Then CB_Material.Text = ColocaMAT
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub ProcuraFicha()
    If CB_Figura.Text = "" Or CB_Bitola.Text = "" Or CB_Material.Text = "" Then Exit Sub
    ZeraFICEST
    Dim gInd, gCla, gExt, gTer As String
    'Procura Figura
    DLL_BD.BDSIS_TBEFG.Seek "=", Trim(CB_Figura.Text)
    If DLL_BD.BDSIS_TBEFG.NoMatch Then
        MsgBox "Ocorreu algum problema na procura da ficha da figura.", vbOKOnly + vbInformation, NOMEAPLIC
        Exit Sub
    Else
        gInd = DLL_BD.BDSIS_TBEFG_CPIFG.Value
        gCla = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGCL.Value))
        gExt = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGEX.Value))
        gTer = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGIN.Value))
    End If
    'Procura Indice de Figura
    DLL_BD.BDSIS_TBEID.Seek "=", gInd
    If DLL_BD.BDSIS_TBEID.NoMatch Then
        MsgBox "Ocorreu algum problema na procura da ficha da figura.", vbOKOnly + vbInformation, NOMEAPLIC
        Exit Sub
    Else
        'Descricoes Normais
        If DLL_BD.BDSIS_TBEID_CPGIN.Value = "" Then 'Nao tem Internos
            If DLL_BD.BDSIS_TBEID_CPTRE = "" Then 'Nao Tem Tipo
                FICEST.NOM = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
            Else 'Tem Tipos
                FICEST.NOM = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTNO.Value) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
            End If
        Else 'Tem Internos
            If DLL_BD.BDSIS_TBEID_CPTRE = "" Then 'Nao Tem Tipo
                FICEST.NOM = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " Int." & _
                            Trim(gTer) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
            Else 'Tem Tipos
                FICEST.NOM = Trim(DLL_BD.BDSIS_TBEID_CPDNO.Value) & " " & _
                            Trim(DLL_BD.BDSIS_TBEID_CPTNO.Value) & " Int." & _
                            Trim(gTer) & " " & _
                            gCla & " " & gExt & " " & _
                            Trim(CB_Material.Text) & " " & _
                            Trim(CB_Bitola.Text)
            End If
        End If
    End If
    DLL_BD.BDSIS_TBEST.Seek "=", CB_Figura.Text, CB_Bitola.Text, CB_Material.Text
    If DLL_BD.BDSIS_TBEST.NoMatch Then
        MsgBox "Ocorreu algum problema na procura da ficha da figura.", vbOKOnly + vbInformation, NOMEAPLIC
        Exit Sub
    Else
        FICEST.FIG = DLL_BD.BDSIS_TBEST_CPFIG.Value
        FICEST.BIT = DLL_BD.BDSIS_TBEST_CPBIT.Value
        FICEST.MAT = DLL_BD.BDSIS_TBEST_CPMAT.Value
        FICEST.VUN = DLL_BD.BDSIS_TBEST_CPVUN.Value
        FICEST.VMI = DLL_BD.BDSIS_TBEST_CPVMI.Value
        FICEST.VCU = DLL_BD.BDSIS_TBEST_CPVCU.Value
        FICEST.PUN = CDbl(Val(DLL_BD.BDSIS_TBEST_CPPUN.Value))
        FICEST.EST = DLL_BD.BDSIS_TBEST_CPEST.Value
        FICEST.EMI = DLL_BD.BDSIS_TBEST_CPETM.Value
        FICEST.VEN = DLL_BD.BDSIS_TBEST_CPVEN.Value
        FICEST.COT = DLL_BD.BDSIS_TBEST_CPCOT.Value
        FICEST.INQ = DLL_BD.BDSIS_TBEST_CPINQ.Value
        FICEST.INP = DLL_BD.BDSIS_TBEST_CPINP.Value
        FICEST.INN = DLL_BD.BDSIS_TBEST_CPINN.Value
        FICEST.INB = DLL_BD.BDSIS_TBEST_CPINB.Value
        FICEST.INM = DLL_BD.BDSIS_TBEST_CPINM.Value
    End If
    TXT_Nome.Text = FICEST.NOM
    If TXT_Quantidade.Text <> "" Then If (CDbl(Val(TXT_Quantidade.Text)) < CDbl(Val(FICEST.EST))) Then CB_Prazo.ListIndex = 0
    TXT_Preco.Text = FICEST.VUN
End Sub
Private Static Sub ZeraFICEST()
    FICEST.FIG = ""
    FICEST.BIT = ""
    FICEST.MAT = ""
    FICEST.NOM = ""
    FICEST.VUN = 0
    FICEST.VMI = 0
    FICEST.VCU = 0
    FICEST.PUN = 0
    FICEST.EST = 0
    FICEST.EMI = 0
    FICEST.VEN = 0
    FICEST.COT = 0
    FICEST.INQ = 0
    FICEST.INP = 0
    FICEST.INN = 0
    FICEST.INB = 0
    FICEST.INM = 0
End Sub
Private Static Sub ZeraCampos(Indice As Integer)
    If Indice = 0 Then
        TXT_Preco.Text = ""
    Else
        CB_Figura.Text = ""
    End If
    CB_Bitola.Text = ""
    CB_Material.Text = ""
    'TXT_Nome.Text = ""
    TXT_Quantidade.Text = ""
    CB_Prazo.Text = ""
    CB_Observacoes.Text = ""
End Sub
Private Static Function AliquotaImposto(TIPO As String) As String
    'If TIPO <> "IPI" Or TIPO <> "ICMS" Then GoTo ERRO_IMPOSTO
    Dim sCF As String, sIMPOSTO As String
    'Procura pela CF da Figura
    DLL_BD.BDSIS_TBCFS.Seek "=", CB_Figura.Text, DLL_FUNCS.ProcuraValorGrupo(CB_Material.Text, "MAT")
    If DLL_BD.BDSIS_TBCFS.NoMatch Then
        sCF = ""
    Else
        sCF = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBCFS_CPGCF.Value))
    End If
    If TIPO = "IPI" Then
        If sCF = "" Then
            sIMPOSTO = InputBox("Não foi encontrado a porcentagem do IPI deste ítem. Digite somente o valor:", "IPI", 0)
            If Not IsNumeric(sIMPOSTO) Then sIMPOSTO = 0
        Else
            'Procura pela alíquota de IPI da Figura
            DLL_BD.BDSIS_TBEAL.Seek "=", "IPI", sESTADO, sCF
            If DLL_BD.BDSIS_TBEAL.NoMatch Then
                GoTo ERRO_IMPOSTO
            Else
                sIMPOSTO = DLL_BD.BDSIS_TBEAL_CPPOR.Value
            End If
        End If
    Else
        If sCF = "" Then
            sIMPOSTO = InputBox("Não foi encontrado a porcentagem do ICMS deste ítem. Digite somente o valor:", "ICMS", 0)
            If Not IsNumeric(sIMPOSTO) Then sIMPOSTO = 0
        Else
            'Procura pela alíquota de ICMS da Figura
            DLL_BD.BDSIS_TBEAL.Seek "=", "ICMS", sESTADO, sCF
            If DLL_BD.BDSIS_TBEAL.NoMatch Then
                GoTo ERRO_IMPOSTO
            Else
                If sESTADO = "SP" Then
                    sIMPOSTO = DLL_BD.BDSIS_TBEAL_CPPOR.Value
                Else
                    sIMPOSTO = DLL_BD.BDSIS_TBEAL_CPPOR.Value
                End If
            End If
        End If
    End If
    sIMPOSTO = sIMPOSTO & "%"
    AliquotaImposto = sIMPOSTO
    Exit Function
ERRO_IMPOSTO:
    AliquotaImposto = "-"
End Function
Private Static Sub ProcuraEstado()
    If LT_Empresa.ListIndex >= 0 Then
        With DLL_BD
            .BDSIS_TBEMP.Seek "=", LT_Empresa.Text
            If .BDSIS_TBEMP.NoMatch Then
                sESTADO = ""
                Exit Sub
            Else
                sESTADO = .BDSIS_TBEMP_CPEST.Value
                If IsNull(.BDSIS_TBEMP_CPFAX.Value) = False Then
                    sFax = .BDSIS_TBEMP_CPFAX.Value
                Else
                    sFax = ""
                End If
                If IsNull(.BDSIS_TBEMP_CPEMA.Value) = False Then
                    sEmail = .BDSIS_TBEMP_CPEMA.Value
                Else
                    sEmail = ""
                End If
                If .BDSIS_TBEMP_CPTRA.Value <> "" Then
                    LT_Transportadora.Text = .BDSIS_TBEMP_CPTRA.Value
                ElseIf sESTADO = "SP" Then
                    LT_Transportadora.Text = "NOSSO MOTORISTA"
                End If
            End If
        End With
    End If
End Sub
Private Static Function PegaCP(ComDD As Boolean) As String
    If ComDD = True Then
        If Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) = "" And Trim(TXT_D3.Text) = "" And Trim(TXT_D4.Text) = "" Then
            PegaCP = Trim(TXT_D1.Text) & "dd"
        ElseIf Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) <> "" And Trim(TXT_D3.Text) = "" And Trim(TXT_D4.Text) = "" Then
            PegaCP = Trim(TXT_D1.Text) & "dd / " & Trim(TXT_D2.Text) & "dd"
        ElseIf Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) <> "" And Trim(TXT_D3.Text) <> "" And Trim(TXT_D4.Text) = "" Then
            PegaCP = Trim(TXT_D1.Text) & "dd / " & Trim(TXT_D2.Text) & "dd / " & Trim(TXT_D3.Text) & "dd"
        ElseIf Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) <> "" And Trim(TXT_D3.Text) <> "" And Trim(TXT_D4.Text) <> "" Then
            PegaCP = Trim(TXT_D1.Text) & "dd / " & Trim(TXT_D2.Text) & "dd / " & Trim(TXT_D3.Text) & "dd / " & Trim(TXT_D4.Text) & "dd"
        End If
    Else
        If Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) = "" And Trim(TXT_D3.Text) = "" And Trim(TXT_D4.Text) = "" Then
            PegaCP = Trim(TXT_D1.Text)
        ElseIf Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) <> "" And Trim(TXT_D3.Text) = "" And Trim(TXT_D4.Text) = "" Then
            PegaCP = Trim(TXT_D1.Text) & "/" & Trim(TXT_D2.Text)
        ElseIf Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) <> "" And Trim(TXT_D3.Text) <> "" And Trim(TXT_D4.Text) = "" Then
            PegaCP = Trim(TXT_D1.Text) & "/" & Trim(TXT_D2.Text) & "/" & Trim(TXT_D3.Text)
        ElseIf Trim(TXT_D1.Text) <> "" And Trim(TXT_D2.Text) <> "" And Trim(TXT_D3.Text) <> "" And Trim(TXT_D4.Text) <> "" Then
            PegaCP = Trim(TXT_D1.Text) & "/" & Trim(TXT_D2.Text) & "/" & Trim(TXT_D3.Text) & "/" & Trim(TXT_D4.Text)
        End If
    End If
End Function
Private Static Function PegaIndFic(IND As Integer) As Long
    With DLL_BD
        .BDSIS_TBEST.Seek "=", FG.TextMatrix(IND, 2), FG.TextMatrix(IND, 3), FG.TextMatrix(IND, 4)
        If .BDSIS_TBEST.NoMatch Then
            PegaIndFic = 0
        Else
            PegaIndFic = .BDSIS_TBEST_CPFIC.Value
        End If
    End With
End Function
Private Static Function LancaPedidos(IND As Integer) As Boolean
    LancaPedidos = False
    Dim nTmp As Long
    With DLL_BD
        .BDSIS_TBEST.Seek "=", FG.TextMatrix(IND, 2), FG.TextMatrix(IND, 3), FG.TextMatrix(IND, 4)
        If Not .BDSIS_TBEST.NoMatch Then
            nTmp = .BDSIS_TBEST_CPVEN.Value
            .BDSIS_TBEST.Edit
            .BDSIS_TBEST_CPVEN.Value = nTmp + CDbl(Val(FG.TextMatrix(IND, 1)))
            .BDSIS_TBEST.Update
            LancaPedidos = True
        End If
    End With
End Function
Private Static Function LancaMapaPedido() As Boolean
    LancaMapaPedido = False
    Dim sTmp As String, nTmp As Long
    sTmp = DLL_FUNCS.NomeMes(Month(Date)) & "/" & Year(Date)
    With DLL_BD
        .BDSIS_TBMPE.Seek "=", sTmp
        If .BDSIS_TBMPE.NoMatch Then
            .BDSIS_TBMPE.AddNew
            .BDSIS_TBMPE_CPMEA.Value = sTmp
            .BDSIS_TBMPE_CPVAL.Value = CalculaPrecoTotal
            .BDSIS_TBMPE.Update
            LancaMapaPedido = True
        Else
            nTmp = .BDSIS_TBMPE_CPVAL.Value
            .BDSIS_TBMPE.Edit
            .BDSIS_TBMPE_CPVAL.Value = (nTmp + CalculaPrecoTotal)
            .BDSIS_TBMPE.Update
            LancaMapaPedido = True
        End If
    End With
End Function
Private Static Function CalculaPrecoTotal() As Currency
    Dim nTmp As Currency
    nTmp = 0
    For I = 1 To FG.Rows - 1
        nTmp = nTmp + (CDbl(FG.TextMatrix(I, 1) * FG.TextMatrix(I, 10)))
    Next I
    CalculaPrecoTotal = nTmp
End Function
Private Static Sub ProcuraContato()
    TXT_Contato.Text = ""
    LT_Contato.Clear
    With DLL_BD
        If .BDSIS_TBECO.RecordCount > 0 Then
            .BDSIS_TBECO.MoveFirst
            Do While Not .BDSIS_TBECO.EOF
                If .BDSIS_TBECO_CPEMP.Value = TXT_Empresa.Text Then
                    LT_Contato.AddItem .BDSIS_TBECO_CPCON.Value
                End If
                .BDSIS_TBECO.MoveNext
            Loop
        End If
    End With
End Sub
Private Static Sub CarregaPedidosPorEmpresa()
    LT_NumPed.Clear
    With DLL_BD
        If .BDSIS_TBPED.RecordCount > 0 Then
            .BDSIS_TBPED.MoveFirst
            Do While Not .BDSIS_TBPED.EOF
                If .BDSIS_TBPED_CPEMP.Value = LT_Empresa.Text Then
                    LT_NumPed.AddItem .BDSIS_TBPED_CPIND.Value
                End If
                .BDSIS_TBPED.MoveNext
            Loop
        End If
    End With
End Sub
Private Static Sub CarregaContatosEmpresa()
    LT_Contato.Clear
    With DLL_BD
        If .BDSIS_TBECO.RecordCount > 0 Then
            .BDSIS_TBECO.MoveFirst
            Do While Not .BDSIS_TBECO.EOF
                If .BDSIS_TBECO_CPEMP.Value = LT_Empresa.Text Then
                    LT_Contato.AddItem .BDSIS_TBECO_CPCON.Value
                End If
                .BDSIS_TBECO.MoveNext
            Loop
        End If
    End With
End Sub
Private Static Sub CarregaPedidos()
    If LT_NumPed.ListIndex < 0 Then Exit Sub
    ApagaDados
    sItens = ""
    With DLL_BD
        ResetaBSEP
        ResetaBSEP
        ResetaBP (4)
        CarregaBSEP ("Carregando Pedidos...")
        If .BDSIS_TBPED.RecordCount > 0 Then
            .BDSIS_TBPED.Seek "=", LT_NumPed.Text
            If .BDSIS_TBPED.NoMatch = False And RB_Empresas.Value = False Then LT_Empresa.Text = .BDSIS_TBPED_CPEMP.Value
            CarregaBSEP ("Procurando Pedidos...")
            .BDSIS_TBPED.Seek "=", LT_NumPed.Text
            If .BDSIS_TBPED.NoMatch = False Then
                CarregaBSEP ("Inserindo dados do Pedido...")
                TXT_Data.Text = Format(.BDSIS_TBPED_CPDAT.Value, "dd/mm/yyyy")
                LT_Contato.Text = .BDSIS_TBPED_CPCON.Value
                CarregaCP .BDSIS_TBPED_CPCPG.Value
                LT_Transportadora.Text = .BDSIS_TBPED_CPTRA.Value
                If IsNull(.BDSIS_TBPED_CPNSP.Value) = False Then TXT_SNP.Text = .BDSIS_TBPED_CPNSP.Value
                TXT_Outras.Text = Format(.BDSIS_TBPED_CPOUD.Value, "###,###,##0.00")
                CB_Frete.ListIndex = .BDSIS_TBPED_CPIFR.Value
                If IsNull(.BDSIS_TBPED_CPDAD.Value) = False Then TXT_Dados.Text = .BDSIS_TBPED_CPDAD.Value
                If IsNull(.BDSIS_TBPED_CPOBS.Value) = False Then TXT_OBS.Text = .BDSIS_TBPED_CPOBS.Value
                sItens = .BDSIS_TBPED_CPITE.Value
                If .BDSIS_TBPED_CPLIQ.Value = True Then
                    RB_Liquidado.Value = True
                ElseIf .BDSIS_TBPED_CPLIQ.Value = False Then
                    RB_Pendente.Value = True
                End If
                If .BDSIS_TBPED_CPABE.Value = True Then
                    RB_Emaberto.Value = True
                End If
                Dim nVend As Integer
                nVend = 0
                If IsNull(.BDSIS_TBPED_CPIVE.Value) = False Then nVend = Val(.BDSIS_TBPED_CPIVE.Value)
                If nVend > 0 Then
                    .BDSIS_TBFUN.Seek "=", nVend
                    If Not .BDSIS_TBFUN.NoMatch Then
                        For I = 0 To (CB_Vendedor.ListCount - 1)
                            If CB_Vendedor.List(I) = .BDSIS_TBFUN_CPFUN.Value Then
                                CB_Vendedor.ListIndex = I
                                CB_Descricao.ListIndex = 0
                                Exit For
                            End If
                        Next I
                    End If
                End If
                CarregaBSEP ("Inserindo ítens do Pedido...")
                CarregaItensPedido sItens
            End If
        End If
    End With
    ResetaBSEP
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
        CB_CondPagto.Text = "21 dd"
    ElseIf sTmp1 = "28" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        CB_CondPagto.Text = "28 dd"
    ElseIf sTmp1 = "30" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        CB_CondPagto.Text = "30 dd"
    ElseIf sTmp1 = "35" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        CB_CondPagto.Text = "35 dd"
    ElseIf sTmp1 = "45" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        CB_CondPagto.Text = "45 dd"
    ElseIf sTmp1 = "28" And sTmp2 = "30" And sTmp3 = "" And sTmp4 = "" Then
        CB_CondPagto.Text = "28 dd / 30 dd"
    ElseIf sTmp1 = "28" And sTmp2 = "35" And sTmp3 = "" And sTmp4 = "" Then
        CB_CondPagto.Text = "28 dd / 35 dd"
    ElseIf sTmp1 = "À Vista" And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        TXT_D1.Text = "À Vista"
        CB_CondPagto.Text = "À Vista"
    ElseIf sTmp1 = "C/Apres." And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        TXT_D1.Text = "C/Apres."
        CB_CondPagto.Text = "C/Apres."
    ElseIf sTmp1 = "S/Venc." And sTmp2 = "" And sTmp3 = "" And sTmp4 = "" Then
        TXT_D1.Text = "S/Venc."
        CB_CondPagto.Text = "S/Venc."
    Else
        CB_CondPagto.Text = "Outros..."
        TXT_D1.Text = sTmp1
        TXT_D2.Text = sTmp2
        TXT_D3.Text = sTmp3
        TXT_D4.Text = sTmp4
    End If
End Sub
Private Static Sub CarregaItensPedido(Valor As String)
    Dim sTmp As String
    sTmp = ""
    FG.Clear
    MontaFG
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            CarregaItensPedido_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    CarregaItensPedido_Aux Val(sTmp)
    TXT_Quantidade.Text = ""
    TXT_Nome.Text = ""
    TXT_Preco.Text = ""
    CB_Figura.Text = ""
    CB_Bitola.Text = ""
    CB_Material.Text = ""
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
            FG.AddItem (FG.Rows)
            FG.TextMatrix(FG.Rows - 1, 1) = Format(.BDSIS_TBPIT_CPQUA.Value, "###,##0.00")
            TXT_Quantidade.Text = .BDSIS_TBPIT_CPQUA.Value
            If .BDSIS_TBPIT_CPINF.Value > 0 Then
                sInd = .BDSIS_TBEST.Index
                .BDSIS_TBEST.Index = "Índice de Ficha"
                .BDSIS_TBEST.Seek "=", .BDSIS_TBPIT_CPINF.Value
                If .BDSIS_TBEST.NoMatch Then
                    MsgBox "Não foi possível localizar um dos ítens do Pedido.", vbExclamation + vbOKOnly, NOMEAPLIC
                    Exit Sub
                Else
                    CB_Figura.Text = .BDSIS_TBEST_CPFIG.Value
                    CB_Bitola.Text = .BDSIS_TBEST_CPBIT.Value
                    CB_Material.Text = .BDSIS_TBEST_CPMAT.Value
                    .BDSIS_TBEST.Index = sInd
                    ProcuraFicha
                    FG.TextMatrix(FG.Rows - 1, 2) = FICEST.FIG
                    FG.TextMatrix(FG.Rows - 1, 3) = FICEST.BIT
                    FG.TextMatrix(FG.Rows - 1, 4) = FICEST.MAT
                    FG.TextMatrix(FG.Rows - 1, 5) = FICEST.NOM
                    FG.TextMatrix(FG.Rows - 1, 7) = VerificaPrazo("COM")
                    FG.TextMatrix(FG.Rows - 1, 8) = VerificaPrazo("PRO")
                    FG.TextMatrix(FG.Rows - 1, 9) = VerificaPrazo("MAP")
                End If
            Else
                FG.TextMatrix(FG.Rows - 1, 5) = .BDSIS_TBPIT_CPDES.Value
            End If
            If .BDSIS_TBPIT_CPCOM.Value <> "" Then FG.TextMatrix(FG.Rows - 1, 6) = .BDSIS_TBPIT_CPCOM.Value
            FG.TextMatrix(FG.Rows - 1, 10) = Format(.BDSIS_TBPIT_CPPRE.Value, "###,###,###,##0.00")
            FG.TextMatrix(FG.Rows - 1, 11) = .BDSIS_TBPIT_CPPRA.Value
            FG.TextMatrix(FG.Rows - 1, 12) = AliquotaImposto("IPI")
            FG.TextMatrix(FG.Rows - 1, 13) = AliquotaImposto("ICMS")
            If .BDSIS_TBPIT_CPLIQ.Value = True Then
                FG.TextMatrix(FG.Rows - 1, 14) = "Sim"
            Else
                FG.TextMatrix(FG.Rows - 1, 14) = "Não"
            End If
        End If
    End With
End Sub
Private Static Sub LimpaCamposPedido()
    FG.Clear
    MontaFG
    TXT_Quantidade.Text = ""
    TXT_Nome.Text = ""
    TXT_Preco.Text = ""
    CB_Figura.Text = ""
    CB_Bitola.Text = ""
    CB_Material.Text = ""
    ZeraFICEST
    LT_Contato.Clear
    LT_NumPed.Clear
    LT_Transportadora.ListIndex = -1
    CB_CondPagto.ListIndex = -1
    TXT_D1.Text = ""
    TXT_D2.Text = ""
    TXT_D3.Text = ""
    TXT_D4.Text = ""
    TXT_Contato.Text = ""
    TXT_Transportadora.Text = ""
End Sub
Private Static Sub ApagaItensPedido(Valor As String)
    Dim sTmp As String
    sTmp = ""
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            ApagaItensPedido_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    ApagaItensPedido_Aux Val(sTmp)
End Sub
Private Static Sub ApagaItensPedido_Aux(Valor As Long)
    With DLL_BD
        .BDSIS_TBPIT.Seek "=", Valor
        If .BDSIS_TBPIT.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens do Pedido para poder deletar.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        Else
            .BDSIS_TBPIT.Delete
        End If
    End With
End Sub
Private Static Sub RetiraPedidos(Valor As String)
    Dim sTmp As String
    sTmp = ""
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            RetiraPedidos_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    RetiraPedidos_Aux Val(sTmp)
End Sub
Private Static Sub RetiraPedidos_Aux(Valor As Long)
    Dim lTmp As Long, lCot As Long
    With DLL_BD
        'procura item do Pedido
        .BDSIS_TBPIT.Seek "=", Valor
        If .BDSIS_TBPIT.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens do Pedido para poder editar.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        lCot = .BDSIS_TBPIT_CPQUA.Value
        'procura ficha de estoque
        Dim sInd As String
        sInd = .BDSIS_TBEST.Index
        .BDSIS_TBEST.Index = "Índice de Ficha"
        .BDSIS_TBEST.Seek "=", .BDSIS_TBPIT_CPINF.Value
        If .BDSIS_TBEST.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens do Pedido para poder editar.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        'retira Pedidos
        lTmp = .BDSIS_TBEST_CPVEN.Value
        .BDSIS_TBEST.Edit
        .BDSIS_TBEST_CPVEN.Value = lTmp - lCot
        .BDSIS_TBEST.Update
        .BDSIS_TBEST.Index = sInd
    End With
End Sub
Private Static Sub RetiraMapaPedido(Valor As Currency, MesAno As String)
    Dim cTmp As Currency
    With DLL_BD
        .BDSIS_TBMPE.Seek "=", MesAno
        If Not .BDSIS_TBMPE.NoMatch Then
            cTmp = .BDSIS_TBMPE_CPVAL.Value
            .BDSIS_TBMPE.Edit
            .BDSIS_TBMPE_CPVAL.Value = (cTmp - Valor)
            .BDSIS_TBMPE.Update
        End If
    End With
End Sub
Private Static Sub GravaItensPedido(Itens As String)
    Dim sTmp As String
    sTmp = ""
    For J = 1 To Len(Itens)
        If Mid(Itens, J, 1) = ";" Then
            GravaItensPedido_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Itens, J, 1)
        End If
    Next J
    GravaItensPedido_Aux Val(sTmp)
End Sub
Private Static Sub GravaItensPedido_Aux(Indice As Long)
    With DLL_BD
        .BDSIS_TBPIT.Seek "=", Indice
        If .BDSIS_TBPIT.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens do Pedido para converter em Pedido.", vbCritical + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        'salva item
        .BDSIS_TBPIT.Edit
        .BDSIS_TBPIT_CPLIQ.Value = True
        .BDSIS_TBPIT.Update
        .BDSIS_TBPIT.AddNew
        .BDSIS_TBPIT_CPQUA.Value = .BDSIS_TBPIT_CPQUA.Value
        .BDSIS_TBPIT_CPINF.Value = .BDSIS_TBPIT_CPINF.Value
        .BDSIS_TBPIT_CPDES.Value = .BDSIS_TBPIT_CPDES.Value
        .BDSIS_TBPIT_CPCOM.Value = .BDSIS_TBPIT_CPCOM.Value
        .BDSIS_TBPIT_CPPRE.Value = .BDSIS_TBPIT_CPPRE.Value
        .BDSIS_TBPIT_CPPRA.Value = .BDSIS_TBPIT_CPPRA.Value
        .BDSIS_TBPIT_CPNPE.Value = nNumPed
        .BDSIS_TBPIT_CPLIQ.Value = False
        If sItens = "" Then
            sItens = .BDSIS_TBPIT_CPIND.Value
        Else
            sItens = sItens & ";" & .BDSIS_TBPIT_CPIND.Value
        End If
        .BDSIS_TBPIT.Update
    End With
End Sub
Private Static Function VerificaPrazo(TIPO As String) As String
    Dim sPrazo As String, nNum As Integer
    nNum = 0
    With Tela_Pedido_MP
        If CB_Figura.Text = "" Then
            sPrazo = "-"
            GoTo SAIDA
        End If
        .DetalhesMP TXT_Quantidade.Text, CB_Figura.Text, CB_Bitola.Text, CB_Material.Text, (Trim(TXT_Nome.Text) & " " & Trim(CB_Observacoes.Text))
        For I = 1 To .FG_MP.Rows - 1
            If TIPO = "COM" Then
                If IsNumeric(.FG_MP.TextMatrix(I, 6)) Then If CDbl(CCur(Val(.FG_MP.TextMatrix(I, 6)))) < CDbl((CCur(Val(.FG_MP.TextMatrix(I, 5))))) Or CDbl((CCur(Val(.FG_MP.TextMatrix(I, 5)))) <= 0) Then nNum = nNum + 1
            ElseIf TIPO = "PRO" Then
                If IsNumeric(.FG_MP.TextMatrix(I, 7)) Then If CDbl(CCur(Val(.FG_MP.TextMatrix(I, 7)))) < CDbl(CCur(Val(.FG_MP.TextMatrix(I, 5)))) Or CDbl(CCur(Val(.FG_MP.TextMatrix(I, 5))) <= 0) Then nNum = nNum + 1
            ElseIf TIPO = "MAP" Then
                If IsNumeric(.FG_MP.TextMatrix(I, 8)) Then If CDbl(CCur(Val(.FG_MP.TextMatrix(I, 8)))) < CDbl(CCur(Val(.FG_MP.TextMatrix(I, 5)))) Or CDbl(CCur(Val(.FG_MP.TextMatrix(I, 5))) <= 0) Then nNum = nNum + 1
            End If
        Next I
        If nNum = 0 Then
            sPrazo = "Imediato"
        ElseIf nNum > 0 And nNum < (.FG_MP.Rows - 1) Then
            sPrazo = "Parcial"
        Else
            sPrazo = "Nenhum"
        End If
    End With
SAIDA:
    VerificaPrazo = sPrazo
End Function
Private Static Sub ImportaCotacao()
    sItens = ""
    With DLL_BD
        ResetaBSEP
        ResetaBP (3)
        CarregaBSEP ("Importando Cotação...")
        If .BDSIS_TBCOT.RecordCount > 0 Then
            .BDSIS_TBCOT.Seek "=", TXT_NumCot.Text
            If .BDSIS_TBCOT.NoMatch Then
                MsgBox "Não foi possível localizar a cotação de preços selecionada.", vbCritical + vbOKOnly, NOMEAPLIC
                Exit Sub
            End If
            CarregaBSEP ("Inserindo dados da Cotação...")
            LT_Empresa.Text = .BDSIS_TBCOT_CPEMP.Value
            TXT_Data.Text = Format(Date, "dd/mm/yyyy")
            LT_Contato.Text = .BDSIS_TBCOT_CPCON.Value
            CarregaCP .BDSIS_TBCOT_CPCPG.Value
            LT_Transportadora.Text = .BDSIS_TBCOT_CPTRA.Value
            sItens = .BDSIS_TBCOT_CPITE.Value
            RB_Pendente.Value = True
            CarregaBSEP ("Inserindo ítens do Pedido...")
            CarregaItensCotacao sItens
            CarregaBSEP ("Inserindo ítens do Pedido...")
        End If
    End With
    ResetaBSEP
End Sub
Private Static Sub CarregaItensCotacao(Valor As String)
    Dim sTmp As String
    sTmp = ""
    FG.Clear
    MontaFG
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            CarregaItensCotacao_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    CarregaItensCotacao_Aux Val(sTmp)
    TXT_Quantidade.Text = ""
    TXT_Nome.Text = ""
    TXT_Preco.Text = ""
    CB_Figura.Text = ""
    CB_Bitola.Text = ""
    CB_Material.Text = ""
End Sub
Private Static Sub CarregaItensCotacao_Aux(Valor As Long)
    Dim sInd As String
    If Valor < 1 Then Exit Sub
    'procura item
    With DLL_BD
        .BDSIS_TBCTI.Seek "=", Valor
        If .BDSIS_TBCTI.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens da Cotação.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        Else
            FG.AddItem (FG.Rows)
            FG.TextMatrix(FG.Rows - 1, 1) = Format(.BDSIS_TBCTI_CPQUA.Value, "###,##0.00")
            TXT_Quantidade.Text = .BDSIS_TBCTI_CPQUA.Value
            If .BDSIS_TBCTI_CPINF.Value > 0 Then
                sInd = .BDSIS_TBEST.Index
                .BDSIS_TBEST.Index = "Índice de Ficha"
                .BDSIS_TBEST.Seek "=", .BDSIS_TBCTI_CPINF.Value
                If .BDSIS_TBEST.NoMatch Then
                    MsgBox "Não foi possível localizar um dos ítens da Cotação.", vbExclamation + vbOKOnly, NOMEAPLIC
                    Exit Sub
                Else
                    CB_Figura.Text = .BDSIS_TBEST_CPFIG.Value
                    CB_Bitola.Text = .BDSIS_TBEST_CPBIT.Value
                    CB_Material.Text = .BDSIS_TBEST_CPMAT.Value
                    .BDSIS_TBEST.Index = sInd
                    ProcuraFicha
                    FG.TextMatrix(FG.Rows - 1, 2) = FICEST.FIG
                    FG.TextMatrix(FG.Rows - 1, 3) = FICEST.BIT
                    FG.TextMatrix(FG.Rows - 1, 4) = FICEST.MAT
                    FG.TextMatrix(FG.Rows - 1, 5) = FICEST.NOM
                    FG.TextMatrix(FG.Rows - 1, 7) = VerificaPrazo("COM")
                    FG.TextMatrix(FG.Rows - 1, 8) = VerificaPrazo("PRO")
                    FG.TextMatrix(FG.Rows - 1, 9) = VerificaPrazo("MAP")
                End If
            Else
                FG.TextMatrix(FG.Rows - 1, 5) = .BDSIS_TBCTI_CPDES.Value
            End If
            If .BDSIS_TBCTI_CPCOM.Value <> "" Then FG.TextMatrix(FG.Rows - 1, 6) = .BDSIS_TBCTI_CPCOM.Value
            FG.TextMatrix(FG.Rows - 1, 10) = Format(.BDSIS_TBCTI_CPPRE.Value, "###,###,###,##0.00")
            FG.TextMatrix(FG.Rows - 1, 11) = .BDSIS_TBCTI_CPPRA.Value
            FG.TextMatrix(FG.Rows - 1, 12) = AliquotaImposto("IPI")
            FG.TextMatrix(FG.Rows - 1, 13) = AliquotaImposto("ICMS")
            FG.TextMatrix(FG.Rows - 1, 14) = "Não"
        End If
    End With
End Sub
Private Static Sub Limpa_IMG_LB()
    ResetaBSEP
    For I = 1 To 14
        IMG(I).Visible = False
    Next I
End Sub
Private Static Sub Muda_IMG_LB(IND As Integer, Num As Integer)
    '1: Erro - 2: Verificando - 3: OK
    If Num = 2 Then
        IMG(IND).Picture = LI.ListImages(2).Picture
        IMG(IND).Visible = True
        LB(IND).FontBold = True
    Else
        If Num = 1 Then
            IMG(IND).Picture = LI.ListImages(1).Picture
        Else
            IMG(IND).Picture = LI.ListImages(3).Picture
        End If
        LB(IND).FontBold = False
    End If
    BS.SimpleText = LB(IND).Caption
End Sub
Private Static Sub Concluir_Erro1()
    Muda_IMG_LB 1, 1
    sTxtMsg = "É necessário preencher todos os campos para poder concluir o Pedido."
    MsgBox sTxtMsg, vbOKOnly + vbInformation, NOMEAPLIC
    Limpa_IMG_LB
End Sub
Private Static Function BaixaCotacao() As Boolean
    BaixaCotacao = False
    If TXT_NumCot.Text = "" Then Exit Function
    If IsNumeric(TXT_NumCot.Text) = False Then Exit Function
    With DLL_BD
        If .BDSIS_TBCOT.RecordCount > 0 Then
            .BDSIS_TBCOT.Seek "=", Val(TXT_NumCot.Text)
            If .BDSIS_TBCOT.NoMatch Then Exit Function
            .BDSIS_TBCOT.Edit
            .BDSIS_TBCOT_CPLIQ.Value = True
            sItens = .BDSIS_TBCOT_CPITE.Value
            .BDSIS_TBCOT.Update
            BaixaItensCotacao sItens
            BaixaCotacao = True
        End If
    End With
End Function
Private Static Sub BaixaItensCotacao(Valor As String)
    Dim sTmp As String
    sTmp = ""
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            BaixaItensCotacao_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    BaixaItensCotacao_Aux Val(sTmp)
End Sub
Private Static Sub BaixaItensCotacao_Aux(Valor As Long)
    With DLL_BD
        .BDSIS_TBCTI.Seek "=", Valor
        If Not .BDSIS_TBCTI.NoMatch Then
            .BDSIS_TBCTI.Edit
            .BDSIS_TBCTI_CPLIQ.Value = True
            .BDSIS_TBCTI.Update
        End If
    End With
End Sub
Private Static Function EmpenhaEstoque() As Boolean
    On Error Resume Next
    EmpenhaEstoque = False
    Dim sTmp As String
    With DLL_BD
        .BDSIS_TBPED.Seek "=", nNumPed
        If .BDSIS_TBPED.NoMatch Then Exit Function
        sTmp = ""
        For J = 1 To Len(.BDSIS_TBPED_CPITE.Value)
            If Mid(.BDSIS_TBPED_CPITE.Value, J, 1) = ";" Then
                EmpenhaEstoque_Aux Val(sTmp)
                sTmp = ""
            Else
                sTmp = sTmp & Mid(.BDSIS_TBPED_CPITE.Value, J, 1)
            End If
        Next J
        EmpenhaEstoque_Aux Val(sTmp)
        EmpenhaEstoque = True
    End With
End Function
Private Static Sub EmpenhaEstoque_Aux(Valor As Long)
    On Error Resume Next
    Dim sInd As String
    With DLL_BD
        .BDSIS_TBPIT.Seek "=", Valor
        If .BDSIS_TBPIT.NoMatch Then Exit Sub
        sInd = .BDSIS_TBEST.Index
        .BDSIS_TBEST.Index = "Índice de Ficha"
        .BDSIS_TBEST.Seek "=", .BDSIS_TBPIT_CPINF.Value
        If Not .BDSIS_TBEST.NoMatch Then
            .BDSIS_TBEST.Edit
            .BDSIS_TBEST_CPVEN.Value = .BDSIS_TBEST_CPVEN.Value + .BDSIS_TBPIT_CPQUA.Value
            .BDSIS_TBEST.Update
        End If
        .BDSIS_TBEST.Index = sInd
    End With
End Sub
Private Static Sub EmpenhaItemEstoque(QUA As Double, FIG As String, BIT As String, MAT As String)
    If FIG = "" Or BIT = "" Or MAT = "" Then Exit Sub
    With DLL_BD
        .BDSIS_TBEST.Seek "=", FIG, BIT, MAT
        If Not .BDSIS_TBEST.NoMatch Then
            .BDSIS_TBEST.Edit
            .BDSIS_TBEST_CPVEN.Value = .BDSIS_TBEST_CPVEN.Value + QUA
            .BDSIS_TBEST.Update
        End If
    End With
End Sub
Private Static Sub DesempenhaItemEstoque(QUA As Double, FIG As String, BIT As String, MAT As String)
    With DLL_BD
        .BDSIS_TBEST.Seek "=", FIG, BIT, MAT
        If Not .BDSIS_TBEST.NoMatch Then
            .BDSIS_TBEST.Edit
            .BDSIS_TBEST_CPVEN.Value = .BDSIS_TBEST_CPVEN.Value - QUA
            .BDSIS_TBEST.Update
        End If
    End With
End Sub
Private Static Function MI_PE() As Boolean
    MI_PE = False
    lTeste = True
    With DLL_BD
        'limpa campos
        lFuncTeste = DLL_IMP.PedidoEstoque_LimpaItens
        If lFuncTeste = False Then lTeste = False
        .BDSIS_TBPED.Seek "=", nNumPed
        If .BDSIS_TBPED.NoMatch Then Exit Function
        'cabecalho
        Dim sNum As String, dDat As Date
        sNum = Str(.BDSIS_TBPED_CPIND.Value)
        dDat = Str(.BDSIS_TBPED_CPDAT.Value)
        .BDSIS_TBEMP.Seek "=", .BDSIS_TBPED_CPEMP.Value 'procura dados sobre a empresa
        If .BDSIS_TBEMP.NoMatch Then Exit Function
        Dim sEmpresa As String, sCGC As String, sINE As String, sFone As String, sFax As String, sEndereco As String, sBairro As String, sCEP As String, sCidade As String, sEsta As String
        sEmpresa = ""
        sCGC = ""
        sINE = ""
        sEndereco = ""
        sBairro = ""
        sCEP = ""
        sCidade = ""
        sEsta = ""
        sFone = ""
        sFax = ""
        If IsNull(.BDSIS_TBEMP_CPEMP.Value) = False Then sEmpresa = .BDSIS_TBEMP_CPEMP.Value
        If IsNull(.BDSIS_TBEMP_CPCGC.Value) = False Then sCGC = .BDSIS_TBEMP_CPCGC.Value
        If IsNull(.BDSIS_TBEMP_CPINE.Value) = False Then sINE = .BDSIS_TBEMP_CPINE.Value
        If IsNull(.BDSIS_TBEMP_CPEND.Value) = False Then sEndereco = .BDSIS_TBEMP_CPEND.Value
        If IsNull(.BDSIS_TBEMP_CPBAI.Value) = False Then sBairro = .BDSIS_TBEMP_CPBAI.Value
        If IsNull(.BDSIS_TBEMP_CPCEP.Value) = False Then sCEP = .BDSIS_TBEMP_CPCEP.Value
        If IsNull(.BDSIS_TBEMP_CPCID.Value) = False Then sCidade = .BDSIS_TBEMP_CPCID.Value
        If IsNull(.BDSIS_TBEMP_CPEST.Value) = False Then sEsta = .BDSIS_TBEMP_CPEST.Value
        If IsNull(.BDSIS_TBEMP_CPFON.Value) = False Then sFone = .BDSIS_TBEMP_CPFON.Value
        If IsNull(.BDSIS_TBEMP_CPFAX.Value) = False Then sFax = .BDSIS_TBEMP_CPFAX.Value
        lFuncTeste = DLL_IMP.PedidoEstoque_Cabecalho(sNum, dDat, sEmpresa, sCGC, sINE, sEndereco, sBairro, sCEP, sCidade, sEsta, sFone, sFax, CB_Depto.Text, TXT_Contato.Text, TXT_Ramal.Text, TXT_SNP.Text)
        If lFuncTeste = False Then lTeste = False
        'rodape
        .BDSIS_TBPED.Seek "=", nNumPed
        If .BDSIS_TBPED.NoMatch Then Exit Function
        lFuncTeste = DLL_IMP.PedidoEstoque_Rodape(Format(.BDSIS_TBPED_CPVAL.Value, "###,###,###,##0.00"), PegaCP(True), .BDSIS_TBPED_CPTRA.Value, CB_Frete.Text, Format(TXT_Outras.Text, "###,###,###,##0.00"), Format(CalculaPrecoIPI, "###,###,###,##0.00"), Format((.BDSIS_TBPED_CPVAL.Value + CDbl(TXT_Outras.Text) + CDbl(CalculaPrecoIPI)), "###,###,###,##0.00"), TXT_Contato.Text, TXT_Empresa.Text, CB_Vendedor.Text, CB_Descricao.Text, TXT_Dados.Text)
        If lFuncTeste = False Then lTeste = False
    End With
    nLin = 0
    nInd = 1
    With FG
        For I = 1 To (.Rows - 1)
            If (Len(Trim(.TextMatrix(I, 5))) + Len(Trim(.TextMatrix(I, 6)))) > 70 Then
                lFuncTeste = MI_PE_DivideDescricao(Trim(Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6))), nInd)
                If lFuncTeste = False Then lTeste = False
            Else
                lFuncTeste = DLL_IMP.PedidoEstoque_Itens(nLin, DLL_FUNCS.PegaNumeroItem(nInd), DLL_FUNCS.PegaUnidade(CDbl(.TextMatrix(I, 1)), 0), .TextMatrix(I, 2), Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6)), .TextMatrix(I, 10), .TextMatrix(I, 13), .TextMatrix(I, 12), .TextMatrix(I, 11))
                If lFuncTeste = False Then lTeste = False
            End If
            nLin = nLin + 1
            nInd = nInd + 1
        Next I
    End With
    'imprimir o pedido
    lFuncTeste = DLL_IMP.PedidoEstoque_Imprimir(DLL_FUNCS.NomeImpressora("IT_PedidoEstoque"))
    If lFuncTeste = False Then lTeste = False
    MI_PE = lTeste
End Function
Private Function MI_PE_DivideDescricao(Texto As String, Indice As Integer) As Boolean
    MI_PE_DivideDescricao = False
    Dim sTmp1 As String, sTmp2 As String, nTmp As Integer, nNumLin As Integer, lT As Boolean
    lTeste = True
    nNumLin = 1
    nTmp = 1
    sTmp1 = Texto
    Do While True
        If Len(sTmp1) > 70 Then
            sTmp2 = Trim(Mid(sTmp1, (((nTmp - 1) * 70) + 1), 70))
            sTmp1 = Trim(Right(sTmp1, Abs(Len(sTmp1) - (nTmp * 70))))
            nTmp = nTmp + 1
            If nNumLin = 1 Then
                lFuncTeste = DLL_IMP.PedidoEstoque_Itens(nLin, DLL_FUNCS.PegaNumeroItem(Indice), DLL_FUNCS.PegaUnidade(CDbl(FG.TextMatrix(I, 1)), 0), FG.TextMatrix(I, 2), sTmp2, "", "", "", "")
                If lFuncTeste = False Then lTeste = False
            Else
                lFuncTeste = DLL_IMP.PedidoEstoque_Itens(nLin, "", "", "", sTmp2, "", "", "", "")
                If lFuncTeste = False Then lTeste = False
            End If
            nNumLin = nNumLin + 1
            nLin = nLin + 1
        Else
            sTmp2 = sTmp1
            lFuncTeste = DLL_IMP.PedidoEstoque_Itens(nLin, "", "", "", sTmp2, FG.TextMatrix(I, 10), FG.TextMatrix(I, 13), FG.TextMatrix(I, 12), FG.TextMatrix(I, 11))
            Exit Do
        End If
    Loop
    MI_PE_DivideDescricao = lT
End Function
Private Static Function CalculaPrecoIPI() As Currency
    CalculaPrecoIPI = 0
    Dim cValor As Currency
    cValor = 0
    For I = 1 To FG.Rows - 1
        If FG.TextMatrix(I, 12) <> "" Or FG.TextMatrix(I, 12) <> "-" Then
            cValor = cValor + DLL_FUNCS.Porcentagem((CDbl(FG.TextMatrix(I, 1)) * CDbl(FG.TextMatrix(I, 10))), CDbl(Left(FG.TextMatrix(I, 12), (Len(FG.TextMatrix(I, 12)) - 1))))
        End If
    Next I
    CalculaPrecoIPI = cValor
End Function
Private Static Function MI_RO() As Boolean
    MI_RO = False
    lTeste = True
    With DLL_BD
        'limpa campos
        lFuncTeste = DLL_IMP.Romaneio_LimpaItens
        If lFuncTeste = False Then lTeste = False
        .BDSIS_TBPED.Seek "=", nNumPed
        If .BDSIS_TBPED.NoMatch Then Exit Function
        .BDSIS_TBEMP.Seek "=", .BDSIS_TBPED_CPEMP.Value 'procura dados sobre a empresa
        If .BDSIS_TBEMP.NoMatch Then Exit Function
        Dim sEmp As String
        sEmp = .BDSIS_TBEMP_CPEMP.Value
        'volta no pedido
        .BDSIS_TBPED.Seek "=", nNumPed
        If .BDSIS_TBPED.NoMatch Then Exit Function
        'cabecalho
        lFuncTeste = DLL_IMP.Romaneio_Cabecalho(Str(nNumOE), sEmp, Format(Date, "dd/mm/yyyy"), Str(nNumPed))
        If lFuncTeste = False Then lTeste = False
    End With
    nLin = 0
    nInd = 1
    With FG
        For I = 1 To (.Rows - 1)
            If (Len(Trim(.TextMatrix(I, 5))) + Len(Trim(.TextMatrix(I, 6)))) > 90 Then
                lFuncTeste = MI_RO_DivideDescricao(Trim(Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6))), nInd)
                If lFuncTeste = False Then lTeste = False
            Else
                lFuncTeste = DLL_IMP.Romaneio_Itens(nLin, DLL_FUNCS.PegaNumeroItem(nInd), DLL_FUNCS.PegaUnidade(CDbl(.TextMatrix(I, 1)), 0), .TextMatrix(I, 2), Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6)), "")
                If lFuncTeste = False Then lTeste = False
            End If
            nLin = nLin + 1
            nInd = nInd + 1
        Next I
    End With
    'imprimir o romaneio
    lFuncTeste = DLL_IMP.Romaneio_Imprimir(DLL_FUNCS.NomeImpressora("IT_Romaneio"))
    If lFuncTeste = False Then lTeste = False
    MI_RO = lTeste
End Function
Private Function MI_RO_DivideDescricao(Texto As String, Indice As Integer) As Boolean
    MI_RO_DivideDescricao = False
    Dim sTmp1 As String, sTmp2 As String, nTmp As Integer, nNumLin As Integer, lT As Boolean
    lTeste = True
    nNumLin = 1
    nTmp = 1
    sTmp1 = Texto
    Do While True
        If Len(sTmp1) > 90 Then
            sTmp2 = Trim(Mid(sTmp1, (((nTmp - 1) * 90) + 1), 90))
            sTmp1 = Trim(Right(sTmp1, (Len(sTmp1) - (nTmp * 90))))
            nTmp = nTmp + 1
            If nNumLin = 1 Then
                lFuncTeste = DLL_IMP.Romaneio_Itens(nLin, DLL_FUNCS.PegaNumeroItem(nInd), DLL_FUNCS.PegaUnidade(CDbl(FG.TextMatrix(I, 1)), 0), FG.TextMatrix(I, 2), sTmp2, "")
                If lFuncTeste = False Then lTeste = False
            Else
                lFuncTeste = DLL_IMP.Romaneio_Itens(nLin, "", "", "", sTmp2, "")
                If lFuncTeste = False Then lTeste = False
            End If
            nNumLin = nNumLin + 1
            nLin = nLin + 1
        Else
            sTmp2 = sTmp1
            lFuncTeste = DLL_IMP.Romaneio_Itens(nLin, "", "", "", sTmp2, "")
            Exit Do
        End If
    Loop
    MI_RO_DivideDescricao = lT
End Function
Private Static Function MI_OF() As Boolean
    MI_OF = False
    lTeste = True
    For I = 1 To (FG.Rows - 1)
        'consulta se existe peca acabada
        If FG.TextMatrix(I, 2) <> "" And FG.TextMatrix(I, 3) <> "" And FG.TextMatrix(I, 4) <> "" Then
            DLL_BD.BDSIS_TBEST.Seek "=", FG.TextMatrix(I, 2), FG.TextMatrix(I, 3), FG.TextMatrix(I, 4)
            If DLL_BD.BDSIS_TBEST.NoMatch Then
                MsgBox "Não foi possível localizar a configuração da matéria-prima do ítem abaixo - será necessário emitir a ordem de fabricação manualmente." & vbCr & vbCr & "Ítem: " & Trim(FG.TextMatrix(I, 0)) & vbCr & " - Figura: " & Trim(FG.TextMatrix(I, 2)) & " - Bitola: " & Trim(FG.TextMatrix(I, 3)) & " - Material: " & Trim(FG.TextMatrix(I, 4)) & vbCr & vbCr & "Peça: " & (Trim(FG.TextMatrix(I, 5)) & " " & Trim(FG.TextMatrix(I, 6))), vbCritical + vbOKOnly, "Falta configuração"
            Else
                'se a quantidade for maior que o estoque, montar OF
                If CDbl(FG.TextMatrix(I, 1)) > (CDbl(DLL_BD.BDSIS_TBEST_CPEST.Value) - CDbl(DLL_BD.BDSIS_TBEST_CPVEN.Value)) Then
                    'verifica cada item da lista de MP
                    'e consulta estoque para ver se precisa emitir OF
                    With Tela_Pedido_MP
                        .DetalhesMP FG.TextMatrix(I, 1), FG.TextMatrix(I, 2), FG.TextMatrix(I, 3), FG.TextMatrix(I, 4), (Trim(FG.TextMatrix(I, 5)) & " " & Trim(FG.TextMatrix(I, 6)))
                        For J = 1 To .FG_MP.Rows - 1
                            If Left(.FG_MP.TextMatrix(J, 0), 2) = "CP" Then 'se for componente
                                'verificar se o componentes e feito fora, ou seja, se existe peças na PA e MP
                                If .FG_MP.TextMatrix(J, 7) <> "-" And .FG_MP.TextMatrix(J, 8) <> "-" Then
                                    'exite peça no BD
                                    If CDbl(.FG_MP.TextMatrix(J, 6)) < CDbl(.FG_MP.TextMatrix(J, 5)) Then
                                        'nao existe peca suficiente, emitir OF
                                        lFuncTeste = MI_OF_Aux(Str(J))
                                        If lFuncTeste = False Then lTeste = False
                                    End If
                                End If
                            ElseIf Left(.FG_MP.TextMatrix(J, 0), 2) = "PA" Then 'se for producao-andamento
                                'verifica se tem peca suficiente
                                If CDbl(.FG_MP.TextMatrix(J, 7)) < CDbl(.FG_MP.TextMatrix(J, 5)) Then
                                    'nao existe peca suficiente, emitir OF
                                    lFuncTeste = MI_OF_Aux(Str(J))
                                    If lFuncTeste = False Then lTeste = False
                                End If
                            End If
                        Next J
                    End With
                End If
            End If
        End If
    Next I
    MI_OF = lTeste
End Function
Private Static Function MI_OF_Aux(Indice As Long) As Boolean
    MI_OF_Aux = True
    lTeste = True
    With Tela_Pedido_MP.FG_MP
        Dim nQuantNec As Long, sFigPA As String
        If Val(.TextMatrix(Indice, 6)) < 1 Then
            nQuantNec = Val(.TextMatrix(Indice, 5))
        Else
            nQuantNec = Val(.TextMatrix(Indice, 5)) - Val(.TextMatrix(Indice, 6))
        End If
        'monta ordem
        lFuncTeste = MI_OF_MontaOrdem(DLL_FUNCS.PegaUnidade(CDbl(nQuantNec), 0), .TextMatrix(Indice, 0), .TextMatrix(Indice, 2), .TextMatrix(Indice, 3), .TextMatrix(Indice, 1))
        If lFuncTeste = False Then lTeste = False
        'empenha estoque
        EmpenhaItemEstoque Str(nQuantNec), .TextMatrix(Indice, 0), .TextMatrix(Indice, 2), .TextMatrix(Indice, 3)
        'procura figura PA
        sFigPA = .TextMatrix(Indice, 0)
        If Left(.TextMatrix(Indice, 0), 2) = "CP" Then sFigPA = "PA" & Trim(Mid(.TextMatrix(Indice, 0), 3, (Len(.TextMatrix(Indice, 0)) - 2)))
        If Left(sFigPA, 2) <> "PA" Then Exit Function
        DLL_BD.BDSIS_TBEST.Seek "=", sFigPA, .TextMatrix(Indice, 2), .TextMatrix(Indice, 3)
        If DLL_BD.BDSIS_TBEST.NoMatch Then Exit Function
        'entra e empenha PA
        DLL_BD.BDSIS_TBEST.Edit
        DLL_BD.BDSIS_TBEST_CPEST.Value = CDbl(DLL_BD.BDSIS_TBEST_CPEST.Value) + CDbl(nQuantNec)
        DLL_BD.BDSIS_TBEST_CPVEN.Value = CDbl(DLL_BD.BDSIS_TBEST_CPVEN.Value) + CDbl(nQuantNec)
        DLL_BD.BDSIS_TBEST.Update
        'procura MP da PA
        Dim fTMP As Tela_Pedido_MP
        Set fTMP = New Tela_Pedido_MP
        fTMP.DetalhesMP Str(nQuantNec), sFigPA, .TextMatrix(Indice, 2), .TextMatrix(Indice, 3), .TextMatrix(Indice, 1)
        If fTMP.FG_MP.Rows < 2 Then Exit Function
        'baixa MP
        DLL_BD.BDSIS_TBEST.Seek "=", fTMP.FG_MP.TextMatrix(1, 0), fTMP.FG_MP.TextMatrix(1, 2), fTMP.FG_MP.TextMatrix(1, 3)
        If DLL_BD.BDSIS_TBEST.NoMatch Then Exit Function
        DLL_BD.BDSIS_TBEST.Edit
        DLL_BD.BDSIS_TBEST_CPEST.Value = CDbl(DLL_BD.BDSIS_TBEST_CPEST.Value) - CDbl(fTMP.FG_MP.TextMatrix(1, 5))
        DLL_BD.BDSIS_TBEST.Update
        Set fTMP = Nothing
    End With
    MI_OF_Aux = lTeste
End Function
Private Static Function MI_OF_MontaOrdem(QUA As String, FIG As String, BIT As String, MAT As String, DES As String) As Boolean
    MI_OF_MontaOrdem = False
    lTeste = True
    Dim nNumOF As Long, nNumFic As Long, nNumQua As Double, nNumDC As Long, sIndBak As String, bAberta As Boolean
    With DLL_BD
        'procura ficha estoque
        .BDSIS_TBEST.Seek "=", FIG, BIT, MAT
        If .BDSIS_TBEST.NoMatch Then Exit Function
        nNumFic = Val(.BDSIS_TBEST_CPFIC.Value)
        nNumQua = (CDbl(DLL_BD.BDSIS_TBEST_CPEST.Value) - CDbl(DLL_BD.BDSIS_TBEST_CPVEN.Value))
        nNumDC = 0
        If IsNull(.BDSIS_TBEST_CPNDC.Value) = False Then nNumDC = Val(.BDSIS_TBEST_CPNDC.Value)
        'verifica se realmente é necessário fazer OF
        If nNumQua > Val(QUA) Then Exit Function
        'verifica se já não existe OF aberta
        sIndBak = .BDSIS_TBODF.Index
        .BDSIS_TBODF.Index = "NumIndFic"
        .BDSIS_TBODF.Seek "=", nNumFic, False
        If Not .BDSIS_TBODF.NoMatch Then 'existe OF aberta
            'altera OF no BD
            .BDSIS_TBODF.Edit
            nNumOF = .BDSIS_TBODF_CPNOF.Value
            .BDSIS_TBODF_CPQES.Value = CDbl(.BDSIS_TBODF_CPQES.Value) + Val(QUA)
            bAberta = True
        Else
            'inclui OF no BD
            .BDSIS_TBODF.AddNew
            nNumOF = .BDSIS_TBODF_CPNOF.Value
            .BDSIS_TBODF_CPDAE.Value = Date + Time()
            .BDSIS_TBODF_CPINF.Value = nNumFic
            .BDSIS_TBODF_CPQES.Value = Val(QUA)
            .BDSIS_TBODF_CPNDC.Value = nNumDC
            .BDSIS_TBODF_CPLIQ.Value = False
            bAberta = False
        End If
        .BDSIS_TBODF.Update
        'monta OF
        lFuncTeste = DLL_IMP.OrdemFabricacao_LimpaItens
        If lFuncTeste = False Then lTeste = False
        lFuncTeste = DLL_IMP.OrdemFabricacao_Cabecalho(Str(nNumOF), Format(Date, "dd/mm/yyyy"), QUA, FIG, BIT, MAT, "", PegaDesc(FIG, DES), Str(nNumDC), "", bAberta)
        If lFuncTeste = False Then lTeste = False
        'imprimi OF
        lFuncTeste = DLL_IMP.OrdemFabricacao_Imprimir(DLL_FUNCS.NomeImpressora("IT_OrdemFabricacao"))
        If lFuncTeste = False Then lTeste = False
        .BDSIS_TBODF.Index = sIndBak
    End With
    MI_OF_MontaOrdem = lTeste
End Function
Private Static Function MI_OM() As Boolean
    MI_OM = False
    lTeste = True
    Dim nQuantNec As Long, nNumOM As Long, nNumFic As Long, nNumQua As Double, nNumDC As Long, sIndBak As String, sDes As String, bAberta As Boolean
    For I = 1 To (FG.Rows - 1)
        'consulta se existe peca acabada
        If FG.TextMatrix(I, 2) <> "" And FG.TextMatrix(I, 3) <> "" And FG.TextMatrix(I, 4) <> "" Then
            DLL_BD.BDSIS_TBEST.Seek "=", FG.TextMatrix(I, 2), FG.TextMatrix(I, 3), FG.TextMatrix(I, 4)
            If DLL_BD.BDSIS_TBEST.NoMatch Then
                MsgBox "Não foi possível localizar a configuração da matéria-prima do ítem abaixo - será necessário emitir a ordem de fabricação manualmente." & vbCr & vbCr & "Ítem: " & Trim(FG.TextMatrix(I, 0)) & vbCr & " - Figura: " & Trim(FG.TextMatrix(I, 2)) & " - Bitola: " & Trim(FG.TextMatrix(I, 3)) & " - Material: " & Trim(FG.TextMatrix(I, 4)) & vbCr & vbCr & "Peça: " & (Trim(FG.TextMatrix(I, 5)) & " " & Trim(FG.TextMatrix(I, 6))), vbCritical + vbOKOnly, "Falta configuração"
            Else
                'se a quantidade for maior que o estoque, montar OM
                If CDbl(FG.TextMatrix(I, 1)) > (CDbl(DLL_BD.BDSIS_TBEST_CPEST.Value) - CDbl(DLL_BD.BDSIS_TBEST_CPVEN.Value)) Then
                    nNumFic = Val(DLL_BD.BDSIS_TBEST_CPFIC.Value)
                    If DLL_BD.BDSIS_TBEST_CPEST.Value < 1 Then
                        nNumQua = CDbl(FG.TextMatrix(I, 1))
                    Else
                        nNumQua = (CDbl(FG.TextMatrix(I, 1)) - (CDbl(DLL_BD.BDSIS_TBEST_CPEST.Value) - CDbl(DLL_BD.BDSIS_TBEST_CPVEN.Value)))
                    End If
                    If IsNull(DLL_BD.BDSIS_TBEST_CPNDC.Value) = False Then nNumDC = Val(DLL_BD.BDSIS_TBEST_CPNDC.Value)
                    sDes = Trim(Trim(FG.TextMatrix(I, 5)) & " " & Trim(FG.TextMatrix(I, 6)))
                    'verifica lista de MP para ver se tem mais de um item, consequentemente é conjunto
                    Tela_Pedido_MP.DetalhesMP FG.TextMatrix(I, 1), FG.TextMatrix(I, 2), FG.TextMatrix(I, 3), FG.TextMatrix(I, 4), (Trim(FG.TextMatrix(I, 5)) & " " & Trim(FG.TextMatrix(I, 6)))
                    If Tela_Pedido_MP.FG_MP.Rows > 1 Then
                        'se tiver mais de 1 item, tem lista de montagem
                        With DLL_BD
                            'verifica se já não existe OF aberta
                            sIndBak = .BDSIS_TBODM.Index
                            .BDSIS_TBODM.Index = "NumIndFic"
                            .BDSIS_TBODM.Seek "=", nNumFic, False
                            If Not .BDSIS_TBODM.NoMatch Then 'existe OM aberta
                                'altera OM no BD
                                .BDSIS_TBODM.Edit
                                nNumOM = .BDSIS_TBODM_CPNOM.Value
                                .BDSIS_TBODM_CPQES.Value = CDbl(.BDSIS_TBODM_CPQES.Value) + CDbl(nNumQua)
                                bAberta = True
                            Else
                                'inclui OM no BD
                                .BDSIS_TBODM.AddNew
                                nNumOM = .BDSIS_TBODM_CPNOM.Value
                                .BDSIS_TBODM_CPDAT.Value = Format(Date, "dd/mm/yyyy")
                                .BDSIS_TBODM_CPHOR.Value = Format(Time, "hh:mm:ss")
                                .BDSIS_TBODM_CPINF.Value = nNumFic
                                .BDSIS_TBODM_CPQES.Value = nNumQua
                                .BDSIS_TBODM_CPNDC.Value = nNumDC
                                .BDSIS_TBODM_CPLIQ.Value = False
                                bAberta = False
                            End If
                            .BDSIS_TBODM.Update
                            .BDSIS_TBODM.Index = sIndBak
                        End With
                        With Tela_Pedido_MP.FG_MP
                            'monta OM
                            lFuncTeste = DLL_IMP.OrdemMontagem_LimpaItens
                            If lFuncTeste = False Then lTeste = False
                            lFuncTeste = DLL_IMP.OrdemMontagem_Cabecalho(Str(nNumOM), Format(Date, "dd/mm/yyyy"), DLL_FUNCS.PegaUnidade(CDbl(nNumQua), 0), FG.TextMatrix(I, 2), FG.TextMatrix(I, 3), FG.TextMatrix(I, 4), Str(nNumDC), "", sDes, "", "", "", "", bAberta)
                            If lFuncTeste = False Then lTeste = False
                            For J = 1 To (.Rows - 1)
                                lFuncTeste = DLL_IMP.OrdemMontagem_Itens((J - 1), DLL_FUNCS.PegaNumeroItem(CInt(J)), DLL_FUNCS.PegaUnidade(CDbl(.TextMatrix(J, 5)), 0), .TextMatrix(J, 1), .TextMatrix(J, 0), .TextMatrix(J, 2), .TextMatrix(J, 3), .TextMatrix(J, 6), .TextMatrix(J, 7), .TextMatrix(J, 8), "", "", "", "", "")
                                If lFuncTeste = False Then lTeste = False
                                'empenha estoque
                                EmpenhaItemEstoque .TextMatrix(J, 5), .TextMatrix(J, 0), .TextMatrix(J, 2), .TextMatrix(J, 3)
                            Next J
                            'imprimi OM
                            lFuncTeste = DLL_IMP.OrdemMontagem_Imprimir(DLL_FUNCS.NomeImpressora("IT_OrdemMontagem"))
                            If lFuncTeste = False Then lTeste = False
                        End With
                    End If
                End If
            End If
        End If
    Next I
    MI_OM = lTeste
End Function
Private Static Function MI_OE() As Boolean
    MI_OE = False
    lTeste = True
    With DLL_BD
        'procura dados sobre a empresa
        .BDSIS_TBEMP.Seek "=", TXT_Empresa.Text
        Dim sCGC As String, sMun As String, sEst As String, sEmpresa As String
        sEmpresa = ""
        sCGC = ""
        sMun = ""
        sEst = ""
        If Not .BDSIS_TBEMP.NoMatch Then
            If IsNull(.BDSIS_TBEMP_CPEMP.Value) = False Then sEmpresa = .BDSIS_TBEMP_CPEMP.Value
            If IsNull(.BDSIS_TBEMP_CPCGC.Value) = False Then sCGC = .BDSIS_TBEMP_CPCGC.Value
            If IsNull(.BDSIS_TBEMP_CPCID.Value) = False Then sMun = .BDSIS_TBEMP_CPCID.Value
            If IsNull(.BDSIS_TBEMP_CPEST.Value) = False Then sEst = .BDSIS_TBEMP_CPEST.Value
        End If
        'procura OE pelo PE e salva
        sIndBak = .BDSIS_TBODE.Index
        .BDSIS_TBODE.Index = "PE"
        .BDSIS_TBODE.Seek "=", nNumPed
        If .BDSIS_TBODE.NoMatch Then
            .BDSIS_TBODE.AddNew
            nNumOE = .BDSIS_TBODE_CPNOE.Value
            .BDSIS_TBODE_CPDAT = Format(Date, "dd/mm/yyyy")
            .BDSIS_TBODE_CPHOR.Value = Format(Time, "hh:mm:ss")
            .BDSIS_TBODE_CPEMP.Value = TXT_Empresa.Text
            .BDSIS_TBODE_CPNPE.Value = nNumPed
            .BDSIS_TBODE_CPTRA.Value = TXT_Transportadora.Text
            .BDSIS_TBODE_CPOBS.Value = TXT_OBS.Text
            .BDSIS_TBODE_CPLIQ.Value = False
            .BDSIS_TBODE.Update
            .BDSIS_TBODE.Index = sIndBak
        Else
            nNumOE = .BDSIS_TBODE_CPNOE.Value
        End If
    End With
    'limpa e monta cabecalho
    lFuncTeste = DLL_IMP.OrdemExpedicao_LimpaItens
    If lFuncTeste = False Then lTeste = False
    lFuncTeste = DLL_IMP.OrdemExpedicao_Cabecalho(Str(nNumOE), Format(Date, "dd/mm/yyyy"), sEmpresa, sCGC, sMun, sEst)
    If lFuncTeste = False Then lTeste = False
    'monta itens
    nLin = 0
    nInd = 1
    With FG
        For I = 1 To (.Rows - 1)
            If (Len(Trim(.TextMatrix(I, 5))) + Len(Trim(.TextMatrix(I, 6)))) > 90 Then
                lFuncTeste = MI_OE_DivideDescricao(Trim(Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6))), nInd)
                If lFuncTeste = False Then lTeste = False
            Else
                lFuncTeste = DLL_IMP.OrdemExpedicao_Itens(nLin, DLL_FUNCS.PegaNumeroItem(nInd), DLL_FUNCS.PegaUnidade(CDbl(.TextMatrix(I, 1)), 0), .TextMatrix(I, 2), Trim(.TextMatrix(I, 5)) & " " & Trim(.TextMatrix(I, 6)), .TextMatrix(I, 11))
                If lFuncTeste = False Then lTeste = False
            End If
            nLin = nLin + 1
            nInd = nInd + 1
        Next I
    End With
     'monta rodape
    lFuncTeste = DLL_IMP.OrdemExpedicao_Rodape(TXT_OBS.Text, TXT_Transportadora.Text, "", Str(nNumPed), "")
    If lFuncTeste = False Then lTeste = False
    'imprimir a OE
    lFuncTeste = DLL_IMP.OrdemExpedicao_Imprimir(DLL_FUNCS.NomeImpressora("IT_OrdemExpedicao"))
    If lFuncTeste = False Then lTeste = False
    MI_OE = lTeste
End Function
Private Function MI_OE_DivideDescricao(Texto As String, Indice As Integer) As Boolean
    MI_OE_DivideDescricao = False
    Dim sTmp1 As String, sTmp2 As String, nTmp As Integer, nNumLin As Integer, lT As Boolean
    lTeste = True
    nNumLin = 1
    nTmp = 1
    sTmp1 = Texto
    Do While True
        If Len(sTmp1) > 90 Then
            sTmp2 = Trim(Mid(sTmp1, (((nTmp - 1) * 90) + 1), 90))
            sTmp1 = Trim(Right(sTmp1, (Len(sTmp1) - (nTmp * 90))))
            nTmp = nTmp + 1
            If nNumLin = 1 Then
                lFuncTeste = DLL_IMP.OrdemExpedicao_Itens(nLin, DLL_FUNCS.PegaNumeroItem(Indice), DLL_FUNCS.PegaUnidade(CDbl(FG.TextMatrix(Indice, 1)), 0), FG.TextMatrix(Indice, 2), sTmp2, "")
                If lFuncTeste = False Then lTeste = False
            Else
                lFuncTeste = DLL_IMP.OrdemExpedicao_Itens(nLin, "", "", "", sTmp2, FG.TextMatrix(Indice, 11))
                If lFuncTeste = False Then lTeste = False
            End If
            nNumLin = nNumLin + 1
            nLin = nLin + 1
        Else
            sTmp2 = sTmp1
            lFuncTeste = DLL_IMP.OrdemExpedicao_Itens(nLin, "", "", "", sTmp2, "")
            Exit Do
        End If
    Loop
    MI_OE_DivideDescricao = lT
End Function
Private Sub PreValores(TIPO As Integer)
    If TIPO = 1 Then 'Novo
        CB_CondPagto.ListIndex = 1
        TXT_Data.Text = Format(Date, "dd/mm/yyyy")
        TXT_Outras.Text = 0
    Else 'Edição
    
    End If
    CB_Descricao.ListIndex = 0
    CK_Imp_RO.Value = 1
    CB_Frete.ListIndex = 0
End Sub
Private Function PegaDesc(PEC As String, DES As String) As String
    If PEC = "" Or DES = "" Then Exit Function
    'procura figura
    DLL_BD.BDSIS_TBEFG.Seek "=", PEC
    If DLL_BD.BDSIS_TBEFG.NoMatch Then
        MsgBox "Ocorreu algum problema na procura da ficha da figura.", vbOKOnly + vbInformation, NOMEAPLIC
        PegaDesc = ""
        Exit Function
    End If
    Dim sCla As String, sExt As String, sCom As String, sFin As String
    sCla = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGCL.Value))
    sExt = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBEFG_CPGEX.Value))
    If IsNull(DLL_BD.BDSIS_TBEFG_CPGIN.Value) = False Then sCom = Trim(DLL_BD.BDSIS_TBEFG_CPGIN.Value)
    If sExt = "" Then
        sFin = Trim(sCom & " " & sCla)
    Else
        sFin = Trim(sCom & " " & sExt & " " & sCla)
    End If
    PegaDesc = PEC & " " & sFin
End Function
Private Sub ApagaDados()
    TXT_Data.Text = "__/__/____"
    CB_CondPagto.ListIndex = -1
    TXT_D1.Text = ""
    TXT_D2.Text = ""
    TXT_D3.Text = ""
    TXT_D4.Text = ""
    TXT_Empresa.Text = ""
    LT_Empresa.ListIndex = -1
    TXT_Contato.Text = ""
    LT_Contato.ListIndex = -1
    TXT_Transportadora.Text = ""
    LT_Transportadora.ListIndex = -1
    CB_Figura.Text = ""
    CB_Bitola.Text = ""
    CB_Material.Text = ""
    TXT_Quantidade.Text = ""
    TXT_Preco.Text = ""
    TXT_Nome.Text = ""
    CB_Prazo.Text = ""
    CB_Observacoes.Text = ""
    TXT_OBS.Text = ""
    CB_Depto.ListIndex = -1
    TXT_Ramal.Text = ""
    CB_Vendedor.ListIndex = -1
    CB_Descricao.ListIndex = -1
    TXT_Outras.Text = ""
    CB_Frete.ListIndex = -1
    TXT_Dados.Text = ""
    TXT_OBS.Text = ""
    TXT_SNP.Text = ""
End Sub
Private Static Sub CarregaBSEP(Texto As String)
    On Error Resume Next
    If BS.SimpleText <> Texto Then BS.SimpleText = Texto
    If BP.Value < BP.Max Then BP.Value = BP.Value + 1
End Sub
Private Static Sub ResetaBP(Max As Integer)
    On Error Resume Next
    BP.Max = Max
    BP.Value = 0
End Sub
Private Static Sub ResetaBSEP()
    On Error Resume Next
    BP.Value = 0
    BS.SimpleText = ""
End Sub
