VERSION 5.00
Object = "{95413FF3-4106-4783-8A5F-91F313AFDB3E}#1.0#0"; "Valvulas.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Tela_Fabrica_OM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ordem de Montagem"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   855
      Left            =   6000
      Picture         =   "Tela_Fabrica_OM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Imprimir Pedido"
      Top             =   4200
      Width           =   855
   End
   Begin TabDlg.SSTab ST 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Nova"
      TabPicture(0)   =   "Tela_Fabrica_OM.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Valvula1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Retorno"
      TabPicture(1)   =   "Tela_Fabrica_OM.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FR_Componentes"
      Tab(1).ControlCount=   1
      Begin VB.Frame FR_Componentes 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   2
         Top             =   480
         Width           =   10455
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   7560
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   6480
            TabIndex        =   8
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   5520
            TabIndex        =   7
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text4 
            Height          =   285
            Left            =   4560
            TabIndex        =   6
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3480
            TabIndex        =   5
            Text            =   "Text1"
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   1560
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   240
            Width           =   1815
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   480
            TabIndex        =   3
            Text            =   "Text1"
            Top             =   240
            Width           =   975
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Corrida:"
            Height          =   195
            Index           =   8
            Left            =   7560
            TabIndex        =   18
            Top             =   0
            Width           =   540
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "ORI nº:"
            Height          =   195
            Index           =   7
            Left            =   6480
            TabIndex        =   17
            Top             =   0
            Width           =   525
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "OF nº:"
            Height          =   195
            Index           =   6
            Left            =   5640
            TabIndex        =   16
            Top             =   0
            Width           =   450
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Bitola:"
            Height          =   195
            Index           =   5
            Left            =   4560
            TabIndex        =   15
            Top             =   0
            Width           =   435
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Material:"
            Height          =   195
            Index           =   4
            Left            =   3480
            TabIndex        =   14
            Top             =   0
            Width           =   600
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Componente:"
            Height          =   195
            Index           =   3
            Left            =   1560
            TabIndex        =   13
            Top             =   0
            Width           =   945
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade:"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   12
            Top             =   0
            Width           =   870
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   0
            Width           =   300
         End
         Begin VB.Label LB 
            AutoSize        =   -1  'True
            Caption         =   "01"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   180
         End
      End
      Begin Siscoval_Produtos.Valvula Valvula1 
         Height          =   3615
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   6376
      End
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   10080
      Picture         =   "Tela_Fabrica_OM.frx":0342
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "Tela_Fabrica_OM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Imprimir_Click()
    Tela_Fabrica_OM_Relatorio.PrintForm
End Sub

Private Sub BT_Voltar_Click()
    Unload Me
End Sub

