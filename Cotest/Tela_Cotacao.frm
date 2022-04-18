VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form Tela_Cotacao 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cotação de Estoque"
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
   Begin TabDlg.SSTab ST 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "1: Dados Gerais"
      TabPicture(0)   =   "Tela_Cotacao.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LB0"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FR(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TXT_Data"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "BT_Novo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "BT_Voltar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FR(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "FR(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "FR(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "BT_Pedido"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "BT_Cotacao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FR(4)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FR(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "FR(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "BT_Editar"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "BT_Deletar"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "BT_Imprimir"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "BT_Apagar"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "BT_Cancelar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "BT_Importa"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "BT_CancelaImportacao"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "2: Ítens da Cotação"
      TabPicture(1)   =   "Tela_Cotacao.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "BT_DetalhesMP"
      Tab(1).Control(1)=   "TXT_Nome"
      Tab(1).Control(2)=   "BT_AlteraItem"
      Tab(1).Control(3)=   "BT_RemoveItem"
      Tab(1).Control(4)=   "BT_AdicionaItem"
      Tab(1).Control(5)=   "BT_AssitenteFigura"
      Tab(1).Control(6)=   "FR(7)"
      Tab(1).Control(7)=   "FG"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton BT_CancelaImportacao 
         Caption         =   "&Cancelar Importação"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4560
         Picture         =   "Tela_Cotacao.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Cancela importação"
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton BT_Importa 
         Caption         =   "&Importa a Cotação selecionada"
         Enabled         =   0   'False
         Height          =   855
         Left            =   1440
         Picture         =   "Tela_Cotacao.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Selecione uma Cotação na lista acima para converter em Pedido"
         Top             =   3840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton BT_DetalhesMP 
         Caption         =   "Detalhes M.P."
         Height          =   735
         Left            =   -68400
         Picture         =   "Tela_Cotacao.frx":099C
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Abre detalhes da matéria-prima do ítem selecionado"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox TXT_Nome 
         Height          =   315
         Left            =   -70800
         TabIndex        =   31
         ToolTipText     =   "Se não exisitir figura, digite a descrição da peça à ser cotada neste campo."
         Top             =   600
         Width           =   3855
      End
      Begin VB.CommandButton BT_AlteraItem 
         Caption         =   "Alterar"
         Height          =   735
         Left            =   -70080
         Picture         =   "Tela_Cotacao.frx":0CA6
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Altera o ítem selecionado na lista abaixo"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton BT_RemoveItem 
         Caption         =   "Remover"
         Height          =   735
         Left            =   -71640
         Picture         =   "Tela_Cotacao.frx":0FB0
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Remove o ítem selecionado na lista abaixo"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton BT_AdicionaItem 
         Caption         =   "Adicionar"
         Height          =   735
         Left            =   -73200
         Picture         =   "Tela_Cotacao.frx":12BA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Adiciona o ítem na lista abaixo"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CommandButton BT_AssitenteFigura 
         Caption         =   "Assistente Figuras"
         Height          =   735
         Left            =   -74880
         Picture         =   "Tela_Cotacao.frx":15C4
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Use o Assistente de Figuras de Estoque caso você não conheça o sistema de figuras"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Frame FR 
         Height          =   1215
         Index           =   7
         Left            =   -74880
         TabIndex        =   56
         Top             =   360
         Width           =   8055
         Begin VB.ComboBox CB_Prazo 
            Height          =   315
            ItemData        =   "Tela_Cotacao.frx":18CE
            Left            =   2760
            List            =   "Tela_Cotacao.frx":18E7
            TabIndex        =   34
            ToolTipText     =   "Selecione ou digite o prazo de entrega par este ítem"
            Top             =   840
            Width           =   1215
         End
         Begin VB.ComboBox CB_Observacoes 
            Height          =   315
            ItemData        =   "Tela_Cotacao.frx":1928
            Left            =   4080
            List            =   "Tela_Cotacao.frx":192A
            TabIndex        =   35
            ToolTipText     =   "Digite ou selecione o complemento ou observações sobre este ítem"
            Top             =   840
            Width           =   3855
         End
         Begin VB.ComboBox CB_Material 
            Height          =   315
            Left            =   2760
            TabIndex        =   30
            ToolTipText     =   "Selecione um material"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox CB_Bitola 
            Height          =   315
            Left            =   1440
            TabIndex        =   29
            ToolTipText     =   "Sselecione uma bitola"
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox CB_Figura 
            Height          =   315
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   28
            ToolTipText     =   "Selecione uma figura"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TXT_Quantidade 
            Height          =   285
            Left            =   120
            TabIndex        =   32
            ToolTipText     =   "Digite aqui a quantidade de peças"
            Top             =   840
            Width           =   1215
         End
         Begin MSMask.MaskEdBox TXT_Preco 
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            ToolTipText     =   "Digite aqui o preço unitário para esta cotação"
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
            TabIndex        =   65
            Top             =   0
            Width           =   765
         End
         Begin VB.Label LB5 
            AutoSize        =   -1  'True
            Caption         =   "Preço:"
            Height          =   195
            Index           =   1
            Left            =   1440
            TabIndex        =   64
            Top             =   600
            Width           =   465
         End
         Begin VB.Label LB9 
            AutoSize        =   -1  'True
            Caption         =   "Complemento:"
            Height          =   195
            Left            =   4080
            TabIndex        =   62
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label LB8 
            AutoSize        =   -1  'True
            Caption         =   "Material:"
            Height          =   195
            Left            =   2760
            TabIndex        =   61
            Top             =   0
            Width           =   615
         End
         Begin VB.Label LB7 
            AutoSize        =   -1  'True
            Caption         =   "Bitola:"
            Height          =   195
            Left            =   1440
            TabIndex        =   60
            Top             =   0
            Width           =   450
         End
         Begin VB.Label LB6 
            AutoSize        =   -1  'True
            Caption         =   "Figura:"
            Height          =   195
            Left            =   120
            TabIndex        =   59
            Top             =   0
            Width           =   495
         End
         Begin VB.Label LB5 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   870
         End
         Begin VB.Label LB10 
            AutoSize        =   -1  'True
            Caption         =   "Prazo:"
            Height          =   195
            Left            =   2760
            TabIndex        =   57
            Top             =   600
            Width           =   450
         End
      End
      Begin VB.CommandButton BT_Cancelar 
         Caption         =   "&Cancelar"
         Height          =   855
         Left            =   6240
         Picture         =   "Tela_Cotacao.frx":192C
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cancela operação"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton BT_Apagar 
         Caption         =   "&Apagar"
         Height          =   855
         Left            =   5400
         Picture         =   "Tela_Cotacao.frx":1C36
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Apaga campos"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton BT_Imprimir 
         Caption         =   "I&mprimir"
         Height          =   855
         Left            =   2640
         Picture         =   "Tela_Cotacao.frx":2078
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Imprimir Cotação"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton BT_Deletar 
         Caption         =   "&Deletar"
         Height          =   855
         Left            =   1800
         Picture         =   "Tela_Cotacao.frx":2382
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Deletar Cotação"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton BT_Editar 
         Caption         =   "&Editar"
         Height          =   855
         Left            =   960
         Picture         =   "Tela_Cotacao.frx":27C4
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Editar Cotação"
         Top             =   4080
         Width           =   855
      End
      Begin VB.Frame FR 
         Caption         =   "Contato:"
         Height          =   1695
         Index           =   2
         Left            =   5040
         TabIndex        =   55
         Top             =   360
         Width           =   1695
         Begin VB.TextBox TXT_Contato 
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   19
            ToolTipText     =   "Digite aqui o nome do contato da empresa se não existir na lista abaixo"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox LT_Contato 
            Height          =   1035
            ItemData        =   "Tela_Cotacao.frx":2C06
            Left            =   120
            List            =   "Tela_Cotacao.frx":2C08
            TabIndex        =   20
            ToolTipText     =   "Selecione o contato da empresa"
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Transportadora:"
         Height          =   1695
         Index           =   5
         Left            =   5040
         TabIndex        =   54
         Top             =   2160
         Width           =   1695
         Begin VB.TextBox TXT_Transportadora 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   21
            ToolTipText     =   "Digite aqui o nome da tranportadora caso não esteja selecionada na lista abaixo"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox LT_Transportadora 
            Height          =   1035
            ItemData        =   "Tela_Cotacao.frx":2C0A
            Left            =   120
            List            =   "Tela_Cotacao.frx":2C0C
            TabIndex        =   22
            ToolTipText     =   "Selecione a transportadora deste cliente"
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Empresa:"
         Height          =   3495
         Index           =   4
         Left            =   3240
         TabIndex        =   53
         Top             =   360
         Width           =   1695
         Begin VB.CommandButton BT_CadEmp 
            Caption         =   "Cadastro de Empresas"
            Height          =   495
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   "Abre a tela de cadastro de empresas"
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox TXT_Empresa 
            Height          =   285
            Left            =   120
            MaxLength       =   20
            TabIndex        =   16
            ToolTipText     =   "Digite aqui o nome da empresa caso não exista na lista abaixo"
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox LT_Empresa 
            Height          =   2205
            ItemData        =   "Tela_Cotacao.frx":2C0E
            Left            =   120
            List            =   "Tela_Cotacao.frx":2C10
            TabIndex        =   17
            ToolTipText     =   "Selecione uma empresa"
            Top             =   600
            Width           =   1455
         End
      End
      Begin VB.CommandButton BT_Cotacao 
         Caption         =   "C&oncluir"
         Height          =   855
         Left            =   4560
         Picture         =   "Tela_Cotacao.frx":2C12
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Conclui a Cotação"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton BT_Pedido 
         Caption         =   "&Pedido"
         Height          =   855
         Left            =   3600
         Picture         =   "Tela_Cotacao.frx":3054
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Transformar esta Cotação em Pedido"
         Top             =   4080
         Width           =   855
      End
      Begin VB.Frame FR 
         Caption         =   "Exibir por:"
         Height          =   855
         Index           =   0
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Width           =   1455
         Begin VB.OptionButton RB_Empresas 
            Caption         =   "&Empresas"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   "Exibe cotações pelo nome das empresas"
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton RB_Todos 
            Caption         =   "Todo&s"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Exibe todas as cotações"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Condições Pgto.:"
         Height          =   3495
         Index           =   6
         Left            =   6840
         TabIndex        =   45
         Top             =   360
         Width           =   1455
         Begin VB.ComboBox CB_CondPagto 
            Height          =   315
            ItemData        =   "Tela_Cotacao.frx":3496
            Left            =   120
            List            =   "Tela_Cotacao.frx":34BB
            Style           =   2  'Dropdown List
            TabIndex        =   23
            ToolTipText     =   "Escolha uma condição de pagamento"
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TXT_D1 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   24
            ToolTipText     =   "Digite aqui o número de dias para o 1º vencimento"
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TXT_D2 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   25
            ToolTipText     =   "Digite aqui o número de dias para o 2º vencimento"
            Top             =   1680
            Width           =   1215
         End
         Begin VB.TextBox TXT_D3 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Digite aqui o número de dias para o 3º vencimento"
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox TXT_D4 
            Enabled         =   0   'False
            Height          =   288
            Left            =   120
            TabIndex        =   27
            ToolTipText     =   "Digite aqui o número de dias para o 4º vencimento"
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label LB1 
            AutoSize        =   -1  'True
            Caption         =   "1ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   375
         End
         Begin VB.Label LB2 
            AutoSize        =   -1  'True
            Caption         =   "2ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   48
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label LB3 
            AutoSize        =   -1  'True
            Caption         =   "3ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   47
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label LB4 
            AutoSize        =   -1  'True
            Caption         =   "4ª dd"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   46
            Top             =   2880
            Width           =   375
         End
      End
      Begin VB.Frame FR 
         Caption         =   "Número:"
         Height          =   3495
         Index           =   1
         Left            =   1680
         TabIndex        =   44
         Top             =   360
         Width           =   1455
         Begin VB.ListBox LT_NumCot 
            Height          =   3180
            ItemData        =   "Tela_Cotacao.frx":3527
            Left            =   120
            List            =   "Tela_Cotacao.frx":3529
            TabIndex        =   15
            ToolTipText     =   "Selecione o número da cotação"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton BT_Voltar 
         Caption         =   "&Voltar"
         Height          =   855
         Left            =   7440
         Picture         =   "Tela_Cotacao.frx":352B
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Volta à Tela Principal"
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton BT_Novo 
         Caption         =   "&Novo"
         Height          =   855
         Left            =   120
         Picture         =   "Tela_Cotacao.frx":396D
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Nova Cotação"
         Top             =   4080
         Width           =   855
      End
      Begin MSMask.MaskEdBox TXT_Data 
         Height          =   330
         Left            =   120
         TabIndex        =   10
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
      Begin MSFlexGridLib.MSFlexGrid FG 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   41
         ToolTipText     =   "Lista de ítens desta cotação"
         Top             =   2520
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4260
         _Version        =   393216
         SelectionMode   =   1
         AllowUserResizing=   1
      End
      Begin VB.Frame FR 
         Height          =   855
         Index           =   3
         Left            =   120
         TabIndex        =   52
         Top             =   2280
         Width           =   1455
         Begin VB.OptionButton RB_Pendente 
            Caption         =   "Pe&ndente"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   "Cotação em aberto"
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton RB_Liquidado 
            Caption         =   "L&iquidado"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   "Feito o Pedido"
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Posição:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   0
            Width           =   615
         End
      End
      Begin VB.Label LB0 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Da&ta:"
         Height          =   195
         Left            =   120
         TabIndex        =   50
         Top             =   480
         Width           =   390
      End
   End
   Begin MSComctlLib.ProgressBar BP 
      Height          =   255
      Left            =   6000
      TabIndex        =   42
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
      TabIndex        =   43
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
End
Attribute VB_Name = "Tela_Cotacao"
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
Public DLL_IMP As Impform.Classe_Impform
Public DLL_CADEMP As Cademp.Classe_Cademp

' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Cotação de Estoque"
Dim RespMsg, I As Integer, ESTIND As String, sESTADO As String, J As Integer
Dim FICEST As T_FICEST, nNumCot As Long, sItens As String, bEstadoEdicao As Boolean
Public sFax As String, sEmail As String
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
    ST.Tab = 1
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
    FG.AddItem (FG.Rows)
    FG.TextMatrix(FG.Rows - 1, 1) = TXT_Quantidade.Text
    FG.TextMatrix(FG.Rows - 1, 2) = CB_Figura.Text
    FG.TextMatrix(FG.Rows - 1, 3) = CB_Bitola.Text
    FG.TextMatrix(FG.Rows - 1, 4) = CB_Material.Text
    FG.TextMatrix(FG.Rows - 1, 5) = TXT_Nome.Text
    FG.TextMatrix(FG.Rows - 1, 6) = CB_Observacoes.Text
    FG.TextMatrix(FG.Rows - 1, 7) = Format(DLL_BD.BDSIS_TBEST_CPEST.Value, "###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 8) = Format(DLL_BD.BDSIS_TBEST_CPVEN.Value, "###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 9) = Format(DLL_BD.BDSIS_TBEST_CPCOT.Value, "###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 10) = VerificaPrazo("COM")
    FG.TextMatrix(FG.Rows - 1, 11) = VerificaPrazo("PRO")
    FG.TextMatrix(FG.Rows - 1, 12) = VerificaPrazo("MAP")
    FG.TextMatrix(FG.Rows - 1, 13) = Format(TXT_Preco.Text, "###,###,###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 14) = CB_Prazo.Text
    FG.TextMatrix(FG.Rows - 1, 15) = AliquotaImposto("IPI")
    FG.TextMatrix(FG.Rows - 1, 16) = AliquotaImposto("ICMS")
    FG.TextMatrix(FG.Rows - 1, 17) = Format(DLL_BD.BDSIS_TBEST_CPVUN.Value, "###,###,###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 18) = Format(DLL_BD.BDSIS_TBEST_CPVMI.Value, "###,###,###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 19) = Format(DLL_BD.BDSIS_TBEST_CPVCU.Value, "###,###,###,##0.00")
    FG.TextMatrix(FG.Rows - 1, 20) = "Não"
    CB_Figura.Text = ""
    TelaEmEspera False
    CB_Figura.SetFocus
'ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_AlteraItem_Click()
    On Error GoTo ERRO_SISCOVAL
    'Testes de preenchimento
    ST.Tab = 1
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
    FG.TextMatrix(FG.RowSel, 7) = Format(DLL_BD.BDSIS_TBEST_CPEST.Value, "###,##0.00")
    FG.TextMatrix(FG.RowSel, 8) = Format(DLL_BD.BDSIS_TBEST_CPVEN.Value, "###,##0.00")
    FG.TextMatrix(FG.RowSel, 9) = Format(DLL_BD.BDSIS_TBEST_CPCOT.Value, "###,##0.00")
    FG.TextMatrix(FG.RowSel, 10) = VerificaPrazo("COM")
    FG.TextMatrix(FG.RowSel, 11) = VerificaPrazo("PRO")
    FG.TextMatrix(FG.RowSel, 12) = VerificaPrazo("MAP")
    FG.TextMatrix(FG.RowSel, 13) = Format(TXT_Preco.Text, "###,###,###,##0.00")
    FG.TextMatrix(FG.RowSel, 14) = CB_Prazo.Text
    FG.TextMatrix(FG.RowSel, 15) = AliquotaImposto("IPI")
    FG.TextMatrix(FG.RowSel, 16) = AliquotaImposto("ICMS")
    FG.TextMatrix(FG.RowSel, 17) = Format(DLL_BD.BDSIS_TBEST_CPVUN.Value, "###,###,###,##0.00")
    FG.TextMatrix(FG.RowSel, 18) = Format(DLL_BD.BDSIS_TBEST_CPVMI.Value, "###,###,###,##0.00")
    FG.TextMatrix(FG.RowSel, 19) = Format(DLL_BD.BDSIS_TBEST_CPVCU.Value, "###,###,###,##0.00")
    FG.TextMatrix(FG.RowSel, 20) = "Não"
    CB_Figura.Text = ""
    TelaEmEspera False
    CB_Figura.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Apagar_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEspera True
    TXT_Data.Text = "__/__/____"
    CB_CondPagto.ListIndex = -1
    TXT_D1.Text = ""
    TXT_D2.Text = ""
    TXT_D3.Text = ""
    TXT_D4.Text = ""
    TXT_Empresa.Text = ""
    LT_NumCot.ListIndex = -1
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
    FG.Clear
    MontaFG
    BP.Value = 0
    BS.SimpleText = ""
    TelaEmEspera False
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
        BS.SimpleText = "Recarregando lista de empresas..."
        If DLL_BD.BDSIS_TBEMP.RecordCount > 0 Then
            LT_Empresa.Clear
            LT_Transportadora.Clear
            BP.Max = DLL_BD.BDSIS_TBEMP.RecordCount + 1
            BP.Value = 0
            DLL_BD.BDSIS_TBEMP.MoveFirst
            While Not DLL_BD.BDSIS_TBEMP.EOF
                If DLL_BD.BDSIS_TBEMP_CPAPE.Value <> "" Then
                    LT_Empresa.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
                End If
                If DLL_BD.BDSIS_TBEMP_CPTIP.Value = "Transportadora" Then
                    LT_Transportadora.AddItem (DLL_BD.BDSIS_TBEMP_CPAPE.Value)
                End If
                BP.Value = BP.Value + 1
                DLL_BD.BDSIS_TBEMP.MoveNext
            Wend
        End If
        BS.SimpleText = ""
        BP.Value = 0
        LT_Empresa.Text = sEmpTemp2
        LT_Transportadora.Text = sEmpTemp2
    End If
    Me.Show vbModal
End Sub
Private Sub BT_CancelaImportacao_Click()
    Unload Me
End Sub
Private Sub BT_Cancelar_Click()
    On Error GoTo ERRO_SISCOVAL
    TelaEmEdicao False
    BT_Apagar.Value = True
    ST.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Cotacao_Click()
    On Error GoTo ERRO_SISCOVAL
    'verifica se todos campos estao preenchidos
    If LT_Empresa.Text = "" And TXT_Empresa.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder concluir a cotação.", vbOKOnly + vbInformation, NOMEAPLIC
        ST.Tab = 0
        LT_Empresa.SetFocus
        Exit Sub
    ElseIf LT_Contato.Text = "" And TXT_Contato.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder concluir a cotação.", vbOKOnly + vbInformation, NOMEAPLIC
        ST.Tab = 0
        LT_Contato.SetFocus
        Exit Sub
    ElseIf LT_Transportadora.Text = "" And TXT_Transportadora.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder concluir a cotação.", vbOKOnly + vbInformation, NOMEAPLIC
        ST.Tab = 0
        LT_Transportadora.SetFocus
        Exit Sub
    ElseIf TXT_D1.Text = "" And TXT_D2.Text = "" And TXT_D3.Text = "" And TXT_D4.Text = "" Then
        MsgBox "É necessário preencher todos os campos para poder concluir a cotação.", vbOKOnly + vbInformation, NOMEAPLIC
        ST.Tab = 0
        TXT_D1.SetFocus
        Exit Sub
    ElseIf FG.Rows <= 1 Then
        MsgBox "Não foram incluídos ítens para poder concluir a cotação.", vbOKOnly + vbInformation, NOMEAPLIC
        ST.Tab = 1
        FG.SetFocus
        Exit Sub
    End If
    
    'começa salvar dados
    TelaEmEspera True
    BP.Max = 7
    BP.Value = 0
    BS.SimpleText = ""
    'verifica se a empresa já está cadastrada
    BS.SimpleText = "Verificando cadastro da empresa..."
    BP.Value = BP.Value + 1
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
    BS.SimpleText = "Verificando cadastro de contatos da empresa..."
    BP.Value = BP.Value + 1
    If LT_Contato.ListIndex = -1 Or LT_Contato.Text <> TXT_Contato.Text And TXT_Contato.Text <> "" Then
        With DLL_BD
            .BDSIS_TBECO.Seek "=", TXT_Empresa.Text, TXT_Contato.Text
            If .BDSIS_TBECO.NoMatch Then
                .BDSIS_TBECO.AddNew
                .BDSIS_TBECO_CPEMP.Value = TXT_Empresa.Text
                .BDSIS_TBECO_CPCON.Value = TXT_Contato.Text
                .BDSIS_TBECO.Update
            End If
        End With
    End If
    'verifica se a transportadora já está cadastrada
    BS.SimpleText = "Verificando cadastro da transportadora..."
    BP.Value = BP.Value + 1
    If LT_Transportadora.ListIndex = -1 Or TXT_Transportadora.Text <> TXT_Transportadora.Text And TXT_Transportadora.Text <> "" Then
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
    
    If bEstadoEdicao = False Then 'Cotacao Nova
        'salva dados da cotação
        BS.SimpleText = "Salvando dados sobre a cotação..."
        BP.Value = BP.Value + 1
        With DLL_BD
            .BDSIS_TBCOT.AddNew
            .BDSIS_TBCOT_CPDAT.Value = Format(Date, "dd/mm/yyyy")
            .BDSIS_TBCOT_CPHOR.Value = Format(Time, "hh:mm:ss")
            .BDSIS_TBCOT_CPEMP.Value = Trim(TXT_Empresa.Text)
            .BDSIS_TBCOT_CPCON.Value = Trim(TXT_Contato.Text)
            .BDSIS_TBCOT_CPCPG.Value = PegaCP
            .BDSIS_TBCOT_CPTRA.Value = Trim(TXT_Transportadora.Text)
            .BDSIS_TBCOT_CPVAL.Value = CalculaPrecoTotal
            .BDSIS_TBCOT_CPLIQ.Value = False
            nNumCot = .BDSIS_TBCOT_CPIND.Value
            'salva itens da cotação
            sItens = ""
            BS.SimpleText = "Salvando dados dos ítens da cotação..."
            BP.Value = BP.Value + 1
            For I = 1 To FG.Rows - 1
                .BDSIS_TBCTI.AddNew
                If FG.TextMatrix(I, 2) <> "" And FG.TextMatrix(I, 3) <> "" And FG.TextMatrix(I, 4) <> "" Then
                    .BDSIS_TBCTI_CPINF.Value = PegaIndFic(I)
                Else
                    .BDSIS_TBCTI_CPDES.Value = Trim(FG.TextMatrix(I, 5))
                End If
                If FG.TextMatrix(I, 6) <> "" Then .BDSIS_TBCTI_CPCOM.Value = FG.TextMatrix(I, 6)
                .BDSIS_TBCTI_CPQUA.Value = FG.TextMatrix(I, 1)
                .BDSIS_TBCTI_CPNCO.Value = nNumCot
                .BDSIS_TBCTI_CPPRE.Value = FG.TextMatrix(I, 13)
                .BDSIS_TBCTI_CPPRA.Value = FG.TextMatrix(I, 14)
                .BDSIS_TBCTI_CPLIQ.Value = False
                If sItens = "" Then
                    sItens = .BDSIS_TBCTI_CPIND.Value
                Else
                    sItens = sItens & ";" & .BDSIS_TBCTI_CPIND.Value
                End If
                .BDSIS_TBCTI.Update
            Next I
            .BDSIS_TBCOT_CPITE.Value = sItens
            .BDSIS_TBCOT.Update
        End With
        'lança saldo de cotações
        BS.SimpleText = "Lançando saldos de ítens cotados..."
        BP.Value = BP.Value + 1
        For I = 1 To FG.Rows - 1
            LancaCotadas I
        Next I
        'lança mapa de cotações
        BS.SimpleText = "Lançando mapa de cotação..."
        BP.Value = BP.Value + 1
        LancaMapaCotacao
    ElseIf bEstadoEdicao = True Then 'Editar Cotacao
        'salva dados da cotação
        BS.SimpleText = "Salvando dados sobre a cotação..."
        BP.Value = BP.Value + 1
        With DLL_BD
            Dim cValorVelho As Currency, sMesAno As String
            sMesAno = DLL_FUNCS.NomeMes(Month(.BDSIS_TBCOT_CPDAT.Value)) & "/" & Year(.BDSIS_TBCOT_CPDAT.Value)
            cValorVelho = .BDSIS_TBCOT_CPVAL.Value
            .BDSIS_TBCOT.Edit
            .BDSIS_TBCOT_CPDAT.Value = Format(Date, "dd/mm/yyyy")
            .BDSIS_TBCOT_CPHOR.Value = Format(Time, "hh:mm:ss")
            .BDSIS_TBCOT_CPEMP.Value = Trim(TXT_Empresa.Text)
            .BDSIS_TBCOT_CPCON.Value = Trim(TXT_Contato.Text)
            .BDSIS_TBCOT_CPCPG.Value = PegaCP
            .BDSIS_TBCOT_CPTRA.Value = Trim(TXT_Transportadora.Text)
            .BDSIS_TBCOT_CPVAL.Value = CalculaPrecoTotal
            If RB_Pendente.Value = True Then
                .BDSIS_TBCOT_CPLIQ.Value = False
            ElseIf RB_Liquidado.Value = True Then
                .BDSIS_TBCOT_CPLIQ.Value = True
            End If
            nNumCot = .BDSIS_TBCOT_CPIND.Value
            sItens = .BDSIS_TBCOT_CPITE.Value
            'retira saldo de cotadas
            BS.SimpleText = "Retirando saldos de ítens cotados..."
            RetiraCotadas sItens
            BP.Value = BP.Value + 1
            'apaga itens da cotação velha
            ApagaItensCotacao sItens
            'salva itens da cotação editada
            BS.SimpleText = "Salvando dados dos ítens da cotação editada..."
            sItens = ""
            For I = 1 To FG.Rows - 1
                .BDSIS_TBCTI.AddNew
                If FG.TextMatrix(I, 2) <> "" And FG.TextMatrix(I, 3) <> "" And FG.TextMatrix(I, 4) <> "" Then
                    .BDSIS_TBCTI_CPINF.Value = PegaIndFic(I)
                Else
                    .BDSIS_TBCTI_CPDES.Value = Trim(FG.TextMatrix(I, 5))
                End If
                If FG.TextMatrix(I, 6) <> "" Then .BDSIS_TBCTI_CPCOM.Value = FG.TextMatrix(I, 6)
                .BDSIS_TBCTI_CPQUA.Value = FG.TextMatrix(I, 1)
                .BDSIS_TBCTI_CPNCO.Value = nNumCot
                .BDSIS_TBCTI_CPPRE.Value = FG.TextMatrix(I, 13)
                .BDSIS_TBCTI_CPPRA.Value = FG.TextMatrix(I, 14)
                .BDSIS_TBCTI_CPLIQ.Value = False
                If sItens = "" Then
                    sItens = .BDSIS_TBCTI_CPIND.Value
                Else
                    sItens = sItens & ";" & .BDSIS_TBCTI_CPIND.Value
                End If
                .BDSIS_TBCTI.Update
            Next I
            .BDSIS_TBCOT_CPITE.Value = sItens
            .BDSIS_TBCOT.Update
            BP.Value = BP.Value + 1
        End With
        'lança mapa de cotações
        BS.SimpleText = "Lançando mapa de cotação..."
        BP.Value = BP.Value + 1
        RetiraMapaCotacao cValorVelho, sMesAno
        LancaMapaCotacao
    End If
    'finalizando
    BP.Value = 0
    BS.SimpleText = ""
    BT_Apagar.Value = True
    TelaEmEdicao False
    TelaEmEspera False
    ST.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Deletar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NumCot.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar um pedido na lista de números de cotações para poder continuar com esta operação.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    RespMsg = MsgBox("Você tem certeza que deseja apagar a cotação de nº " & Trim(LT_NumCot.Text) & " do banco de dados ?", vbQuestion + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then
        TelaEmEspera True
        BP.Max = 3
        BP.Value = 0
        With DLL_BD
            BS.SimpleText = "Procurando cotação para deletar..."
            'retira mapa de cotacao
            Dim cValorVelho As Currency, sMesAno As String
            sMesAno = DLL_FUNCS.NomeMes(Month(.BDSIS_TBCOT_CPDAT.Value)) & "/" & Year(.BDSIS_TBCOT_CPDAT.Value)
            cValorVelho = .BDSIS_TBCOT_CPVAL.Value
            RetiraMapaCotacao cValorVelho, sMesAno
            'retira saldo de cotacao
            RetiraCotadas .BDSIS_TBCOT_CPITE.Value
            'apaga cotacao
            .BDSIS_TBCOT.Delete
            BP.Value = BP.Value + 1
            BS.SimpleText = "Procurando ítens da cotação para deletar..."
            ApagaItensCotacao sItens
            BP.Value = BP.Value + 1
        End With
        BS.SimpleText = "Limpando campos da tela..."
        LT_NumCot.RemoveItem LT_NumCot.ListIndex
        BT_Apagar_Click
        LT_Contato.Clear
        LT_Transportadora.ListIndex = -1
        BP.Value = BP.Value + 1
        TelaEmEspera False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_DetalhesMP_Click()
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
    TelaEmEspera True
    Tela_Cotacao_MP.DetalhesMP TXT_Quantidade.Text, CB_Figura.Text, CB_Bitola.Text, CB_Material.Text, TXT_Nome.Text
    TelaEmEspera False
    Tela_Cotacao_MP.Show vbModal
End Sub
Private Sub BT_Editar_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NumCot.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar um pedido na lista de números de cotações para poder continuar com esta operação.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    TelaEmEdicao True
    FR(1).Enabled = False
    FR(3).Enabled = False
    LT_NumCot.Enabled = False
    bEstadoEdicao = True
    LT_Empresa.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Importa_Click()
    If LT_NumCot.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar uma Cotação para poder importar os dados.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    NUMCOT = LT_NumCot.Text
    Unload Me
End Sub
Private Sub BT_Imprimir_Click()
    'On Error GoTo ERRO_SISCOVAL
    If LT_NumCot.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar um pedido na lista de números de cotações para poder continuar com esta operação.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    Tela_Cotacao_Imprimir.Show vbModal
'ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Novo_Click()
    On Error GoTo ERRO_SISCOVAL
    BT_Apagar.Value = True
    TelaEmEdicao True
    CB_CondPagto.ListIndex = 1
    TXT_Data.Text = Format(Date, "dd/mm/yyyy")
    RB_Todos.Value = False
    RB_Empresas.Value = False
    RB_Pendente.Value = True
    FR(1).Enabled = False
    FR(3).Enabled = False
    LT_NumCot.Enabled = False
    bEstadoEdicao = False
    LT_Empresa.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub BT_Pedido_Click()
    On Error GoTo ERRO_SISCOVAL
    If LT_NumCot.ListIndex = -1 Then
        MsgBox "Você deve primeiro selecionar um pedido na lista de números de cotações para poder continuar com esta operação.", vbInformation + vbOKOnly, NOMEAPLIC
        Exit Sub
    End If
    RespMsg = MsgBox("Você tem certeza que deseja salvar esta cotação como Pedido ?", vbQuestion + vbYesNo + vbDefaultButton1, NOMEAPLIC)
    If RespMsg = vbYes Then
        MsgBox "Esta cotação será convertida em PEDIDO em aberto e você deverá posteriormente entrar na tela de Pedidos, selecionar 'Pedidos em Aberto' para poder emitir as Ordens de Fabricação, Montagem e Expedição e converter para 'Pedidos Pendentes'.", vbInformation + vbOKOnly, "Gravando Pedido"
        TelaEmEspera True
        BP.Max = 3
        BP.Value = 0
        'salva pedido
        BS.SimpleText = "Convertendo a cotação em Pedido..."
        With DLL_BD
            .BDSIS_TBPED.AddNew
            .BDSIS_TBPED_CPDAT.Value = Format(Date, "dd/mm/yyyy")
            .BDSIS_TBPED_CPHOR.Value = Format(Time, "hh:mm:ss")
            .BDSIS_TBPED_CPEMP.Value = .BDSIS_TBCOT_CPEMP.Value
            .BDSIS_TBPED_CPCON.Value = .BDSIS_TBCOT_CPCON.Value
            .BDSIS_TBPED_CPCPG.Value = .BDSIS_TBCOT_CPCPG.Value
            .BDSIS_TBPED_CPTRA.Value = .BDSIS_TBCOT_CPTRA.Value
            .BDSIS_TBPED_CPVAL.Value = .BDSIS_TBCOT_CPVAL.Value
            .BDSIS_TBPED_CPLIQ.Value = False
            .BDSIS_TBPED_CPABE.Value = True
            nNumCot = .BDSIS_TBPED_CPIND.Value
            BP.Value = BP.Value + 1
            'salva itens do pedido
            sItens = ""
            BS.SimpleText = "Salvando dados dos ítens do Pedido..."
            GravaItensPedido .BDSIS_TBCOT_CPITE.Value
            .BDSIS_TBPED_CPITE.Value = sItens
            .BDSIS_TBPED.Update
            BP.Value = BP.Value + 1
            'liquida cotacao
            BS.SimpleText = "Liquidando a cotação..."
            .BDSIS_TBCOT.Edit
            .BDSIS_TBCOT_CPLIQ.Value = True
            .BDSIS_TBCOT.Update
            BP.Value = BP.Value + 1
        End With
        BP.Value = 0
        BS.SimpleText = ""
        TelaEmEspera False
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
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
    Unload Tela_Cotacao
End Sub
Private Sub CB_Bitola_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        MsgBox "Selecione primeiro uma figura.", vbOKOnly + vbInformation, NOMEAPLIC
        CB_Figura.SetFocus
    End If
    CB_Bitola.SelLength = Len(CB_Bitola.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CB_Bitola.Text <> "" Then CB_Material.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Bitola_LostFocus()
    On Error GoTo ERRO_SISCOVAL
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
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
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
Private Sub CB_Figura_Click()
    CarregaFIGBITMAT
    ProcuraFicha
End Sub
Private Sub CB_Figura_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    CB_Figura.SelLength = Len(CB_Figura.Text)
    ZeraCampos 0
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CB_Figura.Text <> "" Then
        CB_Bitola.SetFocus
    ElseIf KeyAscii = vbKeyUp Or KeyAscii = vbKeyDown Then
        CB_Figura_Click
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Figura_LostFocus()
    CB_Figura_Click
End Sub
Private Sub CB_Material_Change()
    On Error GoTo ERRO_SISCOVAL
    CB_Material.SelLength = Len(CB_Material.Text)
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_GotFocus()
    On Error GoTo ERRO_SISCOVAL
    If CB_Figura.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma figura.", vbOKOnly + vbInformation, NOMEAPLIC)
        CB_Figura.SetFocus
        Exit Sub
    ElseIf CB_Bitola.Text = "" Then
        RespMsg = MsgBox("Selecione primeiro uma bitola.", vbOKOnly + vbInformation, NOMEAPLIC)
        CB_Bitola.SetFocus
        Exit Sub
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRO_SISCOVAL
    If KeyAscii = vbKeyReturn And CB_Material.Text <> "" Then TXT_Quantidade.SetFocus
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub CB_Material_LostFocus()
    On Error GoTo ERRO_SISCOVAL
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
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
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
Private Sub FG_Click()
    If FG.Rows > 0 Then
        TelaEmEspera True
        TXT_Quantidade.Text = FG.TextMatrix(FG.RowSel, 1)
        CB_Figura.Text = FG.TextMatrix(FG.RowSel, 2)
        CB_Bitola.Text = FG.TextMatrix(FG.RowSel, 3)
        CB_Material.Text = FG.TextMatrix(FG.RowSel, 4)
        CarregaFIGBITMAT FG.TextMatrix(FG.RowSel, 3), FG.TextMatrix(FG.RowSel, 4)
        TXT_Nome.Text = FG.TextMatrix(FG.RowSel, 5)
        CB_Observacoes.Text = FG.TextMatrix(FG.RowSel, 6)
        TXT_Preco.Text = FG.TextMatrix(FG.RowSel, 13)
        CB_Prazo.Text = FG.TextMatrix(FG.RowSel, 14)
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
    Set DLL_IMP = New Impform.Classe_Impform
    Set DLL_CADEMP = New Cademp.Classe_Cademp
    
    DLL_CARGA.Exibe (True)
    DLL_CARGA.Max (46)
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
    DLL_CARGA.CarregaTexto ("Abrindo tabela Cotações...")
    If DLL_BD.AbreTabela_Cotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Cotações - Ítens...")
    If DLL_BD.AbreTabela_CotacoesItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
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
    DLL_CARGA.CarregaTexto ("Abrindo tabela Mapa - Cotações...")
    If DLL_BD.AbreTabela_MapaCotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Pedidos...")
    If DLL_BD.AbreTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo tabela Pedidos - Ítens...")
    If DLL_BD.AbreTabela_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
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
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Cotações...")
    If DLL_BD.AbreCampos_Cotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Cotações - Ítens...")
    If DLL_BD.AbreCampos_CotacoesItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
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
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Mapa - Cotações...")
    If DLL_BD.AbreCampos_MapaCotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Pedidos...")
    If DLL_BD.AbreCampos_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    DLL_CARGA.CarregaTexto ("Abrindo campos da tabela Pedidos - Ítens...")
    If DLL_BD.AbreCampos_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    
    On Error GoTo ERRO_SISCOVAL
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
    If DLL_BD.BDSIS_TBEMP.RecordCount > 0 Then
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
    
    BT_Apagar.Value = True
    TelaEmEdicao False
    BP.Value = 0
    ST.Tab = 0
    NUMCOT = ""
    DLL_CARGA.CarregaTexto ("Finalizando")
    DLL_FUNCS.RegistraEvento "Abrir Cotações de Estoque", ""
    DLL_CARGA.Exibe (False)
    Exit Sub
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    DLL_CARGA.Exibe (False)
    Unload Tela_Cotacao
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ERRO_SISCOVAL
    'Fecha tabelas
    If DLL_BD.FechaTabela_Grupos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Estoque(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueIndice(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueFiguras(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueComplementos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueCFeST(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueAliquotas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Empresas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EmpresasContatos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Cotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_CotacoesItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_EstoqueComplementos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeQuantidades(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDePecas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeNomes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeBitolas(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaIndiceDeMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MateriaPrimaRelacaoMateriais(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_MapaCotacoes(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_Pedidos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaTabela_PedidosItens(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Set DLL_CARGA = Nothing
    Set DLL_FUNCS = Nothing
    Set DLL_ASFIG = Nothing
    Set DLL_IMP = Nothing
    Set DLL_CADEMP = Nothing
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub LT_Contato_Click()
    If TXT_Contato.Text <> LT_Contato.Text And LT_Contato.ListIndex >= 0 Then TXT_Contato.Text = LT_Contato.Text
End Sub
Private Sub LT_Contato_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_Transportadora.SetFocus
End Sub
Private Sub LT_Empresa_Click()
    If BT_Novo.Enabled = True And LT_Empresa.Enabled = True Then
        BP.Max = 3
        BP.Value = 1
        BS.SimpleText = "Limpando campos..."
        LimpaCamposCotacao
        BP.Value = 2
        BS.SimpleText = "Procurando contatos desta empresa..."
        CarregaContatosEmpresa
        BP.Value = 3
        BS.SimpleText = "Procurando cotações desta empresa..."
        If LT_NumCot.ListIndex = -1 Then CarregaCotacoesPorEmpresa
        BP.Value = 0
        BS.SimpleText = ""
    Else
        If TXT_Empresa.Text <> LT_Empresa.Text And LT_Empresa.ListIndex >= 0 Then TXT_Empresa.Text = LT_Empresa.Text
        ProcuraEstado
        ProcuraContato
    End If
End Sub
Private Sub LT_NumCot_Click()
    TelaEmEspera True
    CarregaCotacoes
    TelaEmEspera False
End Sub
Private Sub LT_Transportadora_Click()
    If TXT_Transportadora.Text <> LT_Transportadora.Text And LT_Transportadora.ListIndex >= 0 Then TXT_Transportadora.Text = LT_Transportadora.Text
End Sub
Private Sub LT_Transportadora_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then CB_CondPagto.SetFocus
End Sub
Private Sub RB_Empresas_Click()
    LT_NumCot.Clear
    FR(1).Enabled = True
    FR(4).Enabled = True
    LT_NumCot.Enabled = True
    LT_Empresa.Enabled = True
End Sub
Private Sub RB_Empresas_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_Empresa.SetFocus
End Sub
Private Sub RB_Todos_Click()
    TelaEmEspera True
    LT_NumCot.Clear
    'carrega cotacoes
    With DLL_BD
        If .BDSIS_TBCOT.RecordCount > 0 Then
            BP.Max = .BDSIS_TBCOT.RecordCount + 1
            BP.Value = 0
            BS.SimpleText = "Carregando cotações..."
            .BDSIS_TBCOT.MoveFirst
            Do While Not .BDSIS_TBCOT.EOF
                LT_NumCot.AddItem .BDSIS_TBCOT_CPIND.Value
                .BDSIS_TBCOT.MoveNext
                BP.Value = BP.Value + 1
            Loop
        End If
    End With
    'habilita listas
    BS.SimpleText = "Finalizando..."
    BP.Value = BP.Value + 1
    FR(1).Enabled = True
    FR(4).Enabled = False
    LT_NumCot.Enabled = True
    LT_Empresa.Enabled = False
    BS.SimpleText = ""
    BP.Value = 0
    LT_Empresa.ListIndex = -1
    TelaEmEspera False
    LT_NumCot.SetFocus
End Sub
Private Sub RB_Todos_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then LT_NumCot.SetFocus
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
Private Sub TXT_Empresa_Change()
    TXT_Empresa.Text = UCase(TXT_Empresa.Text)
    If LT_Empresa.ListIndex = -1 Then LT_Empresa.Text = TXT_Empresa.Text
End Sub
Private Sub TXT_Empresa_KeyPress(KeyAscii As Integer)
    If LT_Empresa.ListIndex >= 0 Then LT_Empresa.ListIndex = -1
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = vbKeyReturn Then TXT_Transportadora.SetFocus
End Sub
Private Sub TXT_Nome_Change()
    'ZeraFICEST
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
    FG.Cols = 21
    FG.Rows = 1
    
    FG.ColAlignment(0) = flexAlignCenterCenter
    FG.ColAlignment(1) = flexAlignCenterCenter
    FG.ColAlignment(2) = flexAlignLeftCenter
    FG.ColAlignment(3) = flexAlignLeftCenter
    FG.ColAlignment(4) = flexAlignLeftCenter
    FG.ColAlignment(5) = flexAlignLeftCenter
    FG.ColAlignment(6) = flexAlignLeftCenter
    FG.ColAlignment(7) = flexAlignCenterCenter
    FG.ColAlignment(8) = flexAlignCenterCenter
    FG.ColAlignment(9) = flexAlignCenterCenter
    FG.ColAlignment(10) = flexAlignLeftCenter
    FG.ColAlignment(11) = flexAlignLeftCenter
    FG.ColAlignment(12) = flexAlignLeftCenter
    FG.ColAlignment(13) = flexAlignCenterCenter
    FG.ColAlignment(14) = flexAlignLeftCenter
    FG.ColAlignment(15) = flexAlignCenterCenter
    FG.ColAlignment(16) = flexAlignCenterCenter
    FG.ColAlignment(17) = flexAlignCenterCenter
    FG.ColAlignment(18) = flexAlignCenterCenter
    FG.ColAlignment(19) = flexAlignCenterCenter
    FG.ColAlignment(20) = flexAlignCenterCenter
    
    FG.ColWidth(0) = 500
    FG.ColWidth(1) = 1000
    FG.ColWidth(2) = 1200
    FG.ColWidth(3) = 1200
    FG.ColWidth(4) = 1200
    FG.ColWidth(5) = 3500
    FG.ColWidth(6) = 1800
    FG.ColWidth(7) = 1000
    FG.ColWidth(8) = 1000
    FG.ColWidth(9) = 1000
    FG.ColWidth(10) = 1200
    FG.ColWidth(11) = 1200
    FG.ColWidth(12) = 1200
    FG.ColWidth(13) = 1200
    FG.ColWidth(14) = 1200
    FG.ColWidth(15) = 800
    FG.ColWidth(16) = 800
    FG.ColWidth(17) = 1200
    FG.ColWidth(18) = 1200
    FG.ColWidth(19) = 1200
    FG.ColWidth(20) = 800

    FG.TextArray(0) = "Item"
    FG.TextArray(1) = "Quantidade"
    FG.TextArray(2) = "Figura"
    FG.TextArray(3) = "Bitola"
    FG.TextArray(4) = "Material"
    FG.TextArray(5) = "Descrição"
    FG.TextArray(6) = "Complemento"
    FG.TextArray(7) = "ESTOQUE"
    FG.TextArray(8) = "Empenhadas"
    FG.TextArray(9) = "Cotadas"
    FG.TextArray(10) = "Componentes"
    FG.TextArray(11) = "Produção"
    FG.TextArray(12) = "Matéria-Prima"
    FG.TextArray(13) = "Preço Unitário"
    FG.TextArray(14) = "Prazo Entrega"
    FG.TextArray(15) = "I.P.I."
    FG.TextArray(16) = "I.C.M.S."
    FG.TextArray(17) = "PN"
    FG.TextArray(18) = "PM"
    FG.TextArray(19) = "PC"
    FG.TextArray(20) = "Liquidado"
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Static Sub TelaEmEdicao(HabilitadoEdicao As Boolean)
    BT_Novo.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Editar.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Deletar.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Imprimir.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Pedido.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_Cotacao.Enabled = HabilitadoEdicao
    BT_Apagar.Enabled = HabilitadoEdicao
    BT_Cancelar.Enabled = HabilitadoEdicao
    BT_Voltar.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    BT_AssitenteFigura.Enabled = HabilitadoEdicao
    BT_AdicionaItem.Enabled = HabilitadoEdicao
    BT_RemoveItem.Enabled = HabilitadoEdicao
    BT_AlteraItem.Enabled = HabilitadoEdicao
    BT_DetalhesMP.Enabled = HabilitadoEdicao
    BT_CadEmp.Enabled = HabilitadoEdicao
    FG.Enabled = True
    TXT_Data.Enabled = False
    Dim MeuControle As Control
    For Each MeuControle In Tela_Cotacao.Controls
        If TypeOf MeuControle Is Label Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is TextBox Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is ComboBox Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is Frame Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is ListBox Then MeuControle.Enabled = HabilitadoEdicao
        If TypeOf MeuControle Is OptionButton Then MeuControle.Value = False
    Next MeuControle
    FR(0).Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Todos.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    RB_Empresas.Enabled = DLL_FUNCS.IB(HabilitadoEdicao)
    TXT_D1.Enabled = False
    TXT_D2.Enabled = False
    TXT_D3.Enabled = False
    TXT_D4.Enabled = False
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
    TXT_Nome.Text = ""
    TXT_Quantidade.Text = ""
    CB_Prazo.Text = ""
    CB_Observacoes.Text = ""
End Sub
Private Static Function VerificaPrazo(TIPO As String) As String
    Dim sPrazo As String, nNum As Integer
    nNum = 0
    With Tela_Cotacao_MP
        .DetalhesMP TXT_Quantidade.Text, CB_Figura.Text, CB_Bitola.Text, CB_Material.Text, TXT_Nome.Text
        For I = 1 To .FG_MP.Rows - 1
            If TIPO = "COM" Then
                If IsNumeric(.FG_MP.TextMatrix(I, 6)) Then If (CDbl(Val(.FG_MP.TextMatrix(I, 6))) < CDbl(Val(.FG_MP.TextMatrix(I, 5)))) Or (Val(.FG_MP.TextMatrix(I, 5)) <= 0) Then nNum = nNum + 1
            ElseIf TIPO = "PRO" Then
                If IsNumeric(.FG_MP.TextMatrix(I, 7)) Then If CDbl(Val(.FG_MP.TextMatrix(I, 7))) < CDbl(Val(.FG_MP.TextMatrix(I, 5))) Or (Val(.FG_MP.TextMatrix(I, 5)) <= 0) Then nNum = nNum + 1
            ElseIf TIPO = "MAP" Then
                If IsNumeric(.FG_MP.TextMatrix(I, 8)) Then If CDbl(Val(.FG_MP.TextMatrix(I, 8))) < CDbl(Val(.FG_MP.TextMatrix(I, 5))) Or (Val(.FG_MP.TextMatrix(I, 5)) <= 0) Then nNum = nNum + 1
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
    VerificaPrazo = sPrazo
End Function
Private Static Function AliquotaImposto(TIPO As String) As String
    'If TIPO <> "IPI" Or TIPO <> "ICMS" Then GoTo ERRO_IMPOSTO
    Dim sCF As String, sIMPOSTO As String
    'Procura pela CF da Figura
    DLL_BD.BDSIS_TBCFS.Seek "=", CB_Figura.Text, DLL_FUNCS.ProcuraValorGrupo(CB_Material.Text, "MAT")
    If DLL_BD.BDSIS_TBCFS.NoMatch Then
        GoTo ERRO_IMPOSTO
    Else
        sCF = Trim(DLL_FUNCS.ProcuraGrupo(DLL_BD.BDSIS_TBCFS_CPGCF.Value))
    End If
    If TIPO = "IPI" Then
        'Procura pela alíquota de IPI da Figura
        DLL_BD.BDSIS_TBEAL.Seek "=", "IPI", sESTADO, sCF
        If DLL_BD.BDSIS_TBEAL.NoMatch Then
            GoTo ERRO_IMPOSTO
        Else
            sIMPOSTO = DLL_BD.BDSIS_TBEAL_CPPOR.Value
        End If
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
                If .BDSIS_TBEMP_CPFAX.Value <> Null Then
                    sFax = .BDSIS_TBEMP_CPFAX.Value
                Else
                    sFax = ""
                End If
                If .BDSIS_TBEMP_CPEMA.Value <> Null Then
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
Private Static Function PegaCP() As String
    If TXT_D1.Text <> "" And TXT_D2.Text = "" And TXT_D3.Text = "" And TXT_D4.Text = "" Then
        PegaCP = Trim(TXT_D1.Text)
    ElseIf TXT_D1.Text <> "" And TXT_D2.Text <> "" And TXT_D3.Text = "" And TXT_D4.Text = "" Then
        PegaCP = Trim(TXT_D1.Text) & "/" & Trim(TXT_D2.Text)
    ElseIf TXT_D1.Text <> "" And TXT_D2.Text <> "" And TXT_D3.Text <> "" And TXT_D4.Text = "" Then
        PegaCP = Trim(TXT_D1.Text) & "/" & Trim(TXT_D2.Text) & "/" & Trim(TXT_D3.Text)
    ElseIf TXT_D1.Text <> "" And TXT_D2.Text <> "" And TXT_D3.Text <> "" And TXT_D4.Text <> "" Then
        PegaCP = Trim(TXT_D1.Text) & "/" & Trim(TXT_D2.Text) & "/" & Trim(TXT_D3.Text) & "/" & Trim(TXT_D4.Text)
    End If
End Function
Private Static Function PegaIndFic(Ind As Integer) As Long
    With DLL_BD
        .BDSIS_TBEST.Seek "=", FG.TextMatrix(Ind, 2), FG.TextMatrix(Ind, 3), FG.TextMatrix(Ind, 4)
        If .BDSIS_TBEST.NoMatch Then
            PegaIndFic = 0
        Else
            PegaIndFic = .BDSIS_TBEST_CPFIC.Value
        End If
    End With
End Function
Private Static Sub LancaCotadas(Ind As Integer)
    Dim nTmp As Long
    With DLL_BD
        .BDSIS_TBEST.Seek "=", FG.TextMatrix(Ind, 2), FG.TextMatrix(Ind, 3), FG.TextMatrix(Ind, 4)
        If Not .BDSIS_TBEST.NoMatch Then
            nTmp = .BDSIS_TBEST_CPCOT.Value
            .BDSIS_TBEST.Edit
            .BDSIS_TBEST_CPCOT.Value = nTmp + CDbl(Val(FG.TextMatrix(Ind, 1)))
            .BDSIS_TBEST.Update
        End If
    End With
End Sub
Private Static Sub LancaMapaCotacao()
    Dim sTmp As String, nTmp As Long
    sTmp = DLL_FUNCS.NomeMes(Month(Date)) & "/" & Year(Date)
    With DLL_BD
        .BDSIS_TBMCO.Seek "=", sTmp
        If .BDSIS_TBMCO.NoMatch Then
            .BDSIS_TBMCO.AddNew
            .BDSIS_TBMCO_CPMEA.Value = sTmp
            .BDSIS_TBMCO_CPVAL.Value = CalculaPrecoTotal
            .BDSIS_TBMCO.Update
        Else
            nTmp = .BDSIS_TBMCO_CPVAL.Value
            .BDSIS_TBMCO.Edit
            .BDSIS_TBMCO_CPVAL.Value = (nTmp + CalculaPrecoTotal)
            .BDSIS_TBMCO.Update
        End If
    End With
End Sub
Private Static Function CalculaPrecoTotal() As Currency
    Dim nTmp As Currency
    nTmp = 0
    For I = 1 To FG.Rows - 1
        nTmp = nTmp + (CDbl(FG.TextMatrix(I, 1) * FG.TextMatrix(I, 13)))
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
Private Static Sub CarregaCotacoesPorEmpresa()
    LT_NumCot.Clear
    With DLL_BD
        If .BDSIS_TBCOT.RecordCount > 0 Then
            .BDSIS_TBCOT.MoveFirst
            Do While Not .BDSIS_TBCOT.EOF
                If .BDSIS_TBCOT_CPEMP.Value = LT_Empresa.Text Then
                    LT_NumCot.AddItem .BDSIS_TBCOT_CPIND.Value
                End If
                .BDSIS_TBCOT.MoveNext
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
Private Static Sub CarregaCotacoes()
    If LT_NumCot.ListIndex < 0 Then Exit Sub
    sItens = ""
    With DLL_BD
        BP.Max = 4
        BP.Value = 0
        BS.SimpleText = "Carregando cotações..."
        BP.Value = BP.Value + 1
        If .BDSIS_TBCOT.RecordCount > 0 Then
            .BDSIS_TBCOT.Seek "=", LT_NumCot.Text
            If .BDSIS_TBCOT.NoMatch = False And RB_Empresas.Value = False Then LT_Empresa.Text = .BDSIS_TBCOT_CPEMP.Value
            BS.SimpleText = "Procurando cotação..."
            BP.Value = BP.Value + 1
            .BDSIS_TBCOT.Seek "=", LT_NumCot.Text
            If .BDSIS_TBCOT.NoMatch = False Then
                BS.SimpleText = "Inserindo dados da cotação..."
                BP.Value = BP.Value + 1
                TXT_Data.Text = Format(.BDSIS_TBCOT_CPDAT.Value, "dd/mm/yyyy")
                LT_Contato.Text = .BDSIS_TBCOT_CPCON.Value
                CarregaCP .BDSIS_TBCOT_CPCPG.Value
                LT_Transportadora.Text = .BDSIS_TBCOT_CPTRA.Value
                sItens = .BDSIS_TBCOT_CPITE.Value
                If .BDSIS_TBCOT_CPLIQ.Value = True Then
                    RB_Liquidado.Value = True
                Else
                    RB_Pendente.Value = True
                End If
                BS.SimpleText = "Inserindo ítens da cotação..."
                BP.Value = BP.Value + 1
                CarregaItensCotacao sItens
            End If
        End If
    End With
    BS.SimpleText = ""
    BP.Value = 0
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
            MsgBox "Não foi possível localizar um dos ítens da cotação.", vbExclamation + vbOKOnly, NOMEAPLIC
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
                    MsgBox "Não foi possível localizar um dos ítens da cotação.", vbExclamation + vbOKOnly, NOMEAPLIC
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
                    FG.TextMatrix(FG.Rows - 1, 7) = Format(FICEST.EST, "###,##0.00")
                    FG.TextMatrix(FG.Rows - 1, 8) = Format(FICEST.VEN, "###,##0.00")
                    FG.TextMatrix(FG.Rows - 1, 9) = Format(FICEST.COT, "###,##0.00")
                    FG.TextMatrix(FG.Rows - 1, 10) = VerificaPrazo("COM")
                    FG.TextMatrix(FG.Rows - 1, 11) = VerificaPrazo("PRO")
                    FG.TextMatrix(FG.Rows - 1, 12) = VerificaPrazo("MAP")
                    FG.TextMatrix(FG.Rows - 1, 17) = Format(FICEST.VUN, "###,###,###,##0.00")
                    FG.TextMatrix(FG.Rows - 1, 18) = Format(FICEST.VMI, "###,###,###,##0.00")
                    FG.TextMatrix(FG.Rows - 1, 19) = Format(FICEST.VCU, "###,###,###,##0.00")
                End If
            Else
                FG.TextMatrix(FG.Rows - 1, 5) = .BDSIS_TBCTI_CPDES.Value
            End If
            If .BDSIS_TBCTI_CPCOM.Value <> "" Then FG.TextMatrix(FG.Rows - 1, 6) = .BDSIS_TBCTI_CPCOM.Value
            FG.TextMatrix(FG.Rows - 1, 13) = Format(.BDSIS_TBCTI_CPPRE.Value, "###,###,###,##0.00")
            FG.TextMatrix(FG.Rows - 1, 14) = .BDSIS_TBCTI_CPPRA.Value
            FG.TextMatrix(FG.Rows - 1, 15) = AliquotaImposto("IPI")
            FG.TextMatrix(FG.Rows - 1, 16) = AliquotaImposto("ICMS")
            If .BDSIS_TBCTI_CPLIQ.Value = True Then
                FG.TextMatrix(FG.Rows - 1, 20) = "Sim"
            Else
                FG.TextMatrix(FG.Rows - 1, 20) = "Não"
            End If
        End If
    End With
End Sub
Private Static Sub LimpaCamposCotacao()
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
    LT_NumCot.Clear
    LT_Transportadora.ListIndex = -1
    CB_CondPagto.ListIndex = -1
    TXT_D1.Text = ""
    TXT_D2.Text = ""
    TXT_D3.Text = ""
    TXT_D4.Text = ""
    TXT_Contato.Text = ""
    TXT_Transportadora.Text = ""
End Sub
Private Static Sub ApagaItensCotacao(Valor As String)
    Dim sTmp As String
    sTmp = ""
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            ApagaItensCotacao_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    ApagaItensCotacao_Aux Val(sTmp)
End Sub
Private Static Sub ApagaItensCotacao_Aux(Valor As Long)
    With DLL_BD
        .BDSIS_TBCTI.Seek "=", Valor
        If .BDSIS_TBCTI.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens da cotação para poder deletar.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        Else
            .BDSIS_TBCTI.Delete
        End If
    End With
End Sub
Private Static Sub RetiraCotadas(Valor As String)
    Dim sTmp As String
    sTmp = ""
    For J = 1 To Len(Valor)
        If Mid(Valor, J, 1) = ";" Then
            RetiraCotadas_Aux Val(sTmp)
            sTmp = ""
        Else
            sTmp = sTmp & Mid(Valor, J, 1)
        End If
    Next J
    RetiraCotadas_Aux Val(sTmp)
End Sub
Private Static Sub RetiraCotadas_Aux(Valor As Long)
    Dim lTmp As Long, lCot As Long
    With DLL_BD
        'procura item da cotacao
        .BDSIS_TBCTI.Seek "=", Valor
        If .BDSIS_TBCTI.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens da cotação para poder editar.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        lCot = .BDSIS_TBCTI_CPQUA.Value
        'procura ficha de estoque
        Dim sInd As String
        sInd = .BDSIS_TBEST.Index
        .BDSIS_TBEST.Index = "Índice de Ficha"
        .BDSIS_TBEST.Seek "=", .BDSIS_TBCTI_CPINF.Value
        If .BDSIS_TBEST.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens da cotação para poder editar.", vbExclamation + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        'retira cotadas
        lTmp = .BDSIS_TBEST_CPCOT.Value
        .BDSIS_TBEST.Edit
        .BDSIS_TBEST_CPCOT.Value = lTmp - lCot
        .BDSIS_TBEST.Update
        .BDSIS_TBEST.Index = sInd
    End With
End Sub
Private Static Sub RetiraMapaCotacao(Valor As Currency, MesAno As String)
    Dim cTmp As Currency
    With DLL_BD
        .BDSIS_TBMCO.Seek "=", MesAno
        If Not .BDSIS_TBMCO.NoMatch Then
            cTmp = .BDSIS_TBMCO_CPVAL.Value
            .BDSIS_TBMCO.Edit
            .BDSIS_TBMCO_CPVAL.Value = (cTmp - Valor)
            .BDSIS_TBMCO.Update
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
        .BDSIS_TBCTI.Seek "=", Indice
        If .BDSIS_TBCTI.NoMatch Then
            MsgBox "Não foi possível localizar um dos ítens da Cotação para converter em Pedido.", vbCritical + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
        'salva item
        .BDSIS_TBCTI.Edit
        .BDSIS_TBCTI_CPLIQ.Value = True
        .BDSIS_TBCTI.Update
        .BDSIS_TBPIT.AddNew
        .BDSIS_TBPIT_CPQUA.Value = .BDSIS_TBCTI_CPQUA.Value
        .BDSIS_TBPIT_CPINF.Value = .BDSIS_TBCTI_CPINF.Value
        .BDSIS_TBPIT_CPDES.Value = .BDSIS_TBCTI_CPDES.Value
        .BDSIS_TBPIT_CPCOM.Value = .BDSIS_TBCTI_CPCOM.Value
        .BDSIS_TBPIT_CPPRE.Value = .BDSIS_TBCTI_CPPRE.Value
        .BDSIS_TBPIT_CPPRA.Value = .BDSIS_TBCTI_CPPRA.Value
        .BDSIS_TBPIT_CPNPE.Value = nNumCot
        .BDSIS_TBPIT_CPLIQ.Value = False
        If sItens = "" Then
            sItens = .BDSIS_TBPIT_CPIND.Value
        Else
            sItens = sItens & ";" & .BDSIS_TBPIT_CPIND.Value
        End If
        .BDSIS_TBPIT.Update
    End With
End Sub
