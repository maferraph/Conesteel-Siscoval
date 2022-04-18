VERSION 5.00
Begin VB.UserControl Bruto 
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   ScaleHeight     =   780
   ScaleWidth      =   7350
   Begin VB.Frame FR_Dados 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.ComboBox CB_Material 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CB_Bitola 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox CB_Peca 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label LB_Peca 
         AutoSize        =   -1  'True
         Caption         =   "Peça em bruto:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1080
      End
      Begin VB.Label LB_Bitola 
         AutoSize        =   -1  'True
         Caption         =   "Bitola:"
         Height          =   195
         Left            =   2880
         TabIndex        =   2
         Top             =   0
         Width           =   435
      End
      Begin VB.Label LB_Material 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   195
         Left            =   4560
         TabIndex        =   1
         Top             =   0
         Width           =   600
      End
   End
End
Attribute VB_Name = "Bruto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub CB_Peca_Change()
    CB_Peca.Clear
    CB_Bitola.Clear
    CB_Material.Clear
    CB_Bitola.Enabled = False
    CB_Material.Enabled = False
    LB_Bitola.Enabled = False
    LB_Material.Enabled = False
    'Carrega lista de peças e bitolas
    Select Case CB_Peca.List(CB_Peca.ListIndex)
        Case "CORPO GAVETA"
            CarregaBitolaPorValvula
            CarregaMaterialForjado
        Case "CASTELO" Or "PREME" Or "TAMPA"
            CarregaBitolaPorValvulaJunta
            CarregaMaterialForjado
        Case "CASTELO"
            CarregaBitolaPorValvulaJunta
            CarregaMaterialForjado
        Case "CUNHA"
            CarregaBitolaPorValvula
            CarregaMaterialFundido
        Case "CONTRA-SEDE"
            With CB_Bitola
                .AddItem "1/2" & Chr(34)
                .AddItem "3/4" & Chr(34)
                .AddItem "1" & Chr(34)
            End With
            CarregaMaterialFundido
        Case "PÊNDULO"
            CB_Bitola.AddItem "1.1/2" & Chr(34) & "e 2" & Chr(34)
            CarregaBitolaPorValvula
            CarregaMaterialFundido
        Case "REDONDO" Or "SEXTAVADO" Or "QUADRADO"
            CarregaBitolaPorBarra
    End Select
    
End Sub
Private Sub CB_Peca_Click()
    CB_Peca_Change
End Sub
Private Sub UserControl_Initialize()
    CB_Peca.Clear
    CB_Bitola.Clear
    CB_Material.Clear
    CB_Bitola.Enabled = False
    CB_Material.Enabled = False
    LB_Bitola.Enabled = False
    LB_Material.Enabled = False
    'Carrega Lista de Peças
    With CB_Peca
        .AddItem "CORPO GAVETA"
        .AddItem "CASTELO"
        .AddItem "PREME"
        .AddItem "TAMPA"
        .AddItem "CUNHA"
        .AddItem "CONTRA-SEDE"
        .AddItem "PÊNDULO"
        .AddItem "REDONDO"
        .AddItem "SEXTAVADO"
        .AddItem "QUADRADO"
        .AddItem "TUBO MECÂNICO"
        .AddItem "TUBO SCH.40"
        .AddItem "TUBO SCH.80"
        .AddItem "TUBO SCH.160"
        .AddItem "TUBO SCH.XXS"
        .AddItem "FLANGE 150"
        .AddItem "FLANGE 300"
        .AddItem "FLANGE 600"
        .AddItem "FLANGE 900"
        .AddItem "FLANGE 1500"
        .AddItem "VOLANTE GAVETA"
        .AddItem "VOLANTE GLOBO"
        .AddItem "SOLDA ADAPTAÇÃO"
        .AddItem "SOLDA REVESTIMENTO"
        .AddItem "COTOVELO 90º"
        .AddItem "COTOVELO 45º"
        .AddItem "PLUG QUADRADO"
        .AddItem "TE"
        .AddItem "PORCA UNIÃO"
        .AddItem "CRUZETA"
        .AddItem "TE TUBULAR"
    End With
End Sub
Private Sub CarregaBitolaPorValvula()
    With CB_Bitola
        .AddItem "1/2" & Chr(34)
        .AddItem "3/4" & Chr(34)
        .AddItem "1" & Chr(34)
        .AddItem "1.1/2" & Chr(34)
        .AddItem "2" & Chr(34)
    End With
End Sub
Private Sub CarregaBitolaPorValvulaJunta()
    With CB_Bitola
        .AddItem "1/2" & Chr(34) & "e 3/4" & Chr(34)
        .AddItem "1" & Chr(34)
        .AddItem "1.1/2" & Chr(34) & "e 2" & Chr(34)
    End With
End Sub
Private Sub CarregaBitolaPorBarra()
    With CB_Bitola
        .AddItem "1/2" & Chr(34)
    End With
End Sub


Private Sub CarregaMaterialForjado()
    With CB_Material
        .AddItem "A105"
        .AddItem "A182 F304"
        .AddItem "A182 F316"
        .AddItem "A182 F304L"
        .AddItem "A182 F316L"
        .AddItem "A182 F5"
        .AddItem "A182 F11"
        .AddItem "A350 LF2"
    End With
End Sub
Private Sub CarregaMaterialFundido()
    With CB_Material
        .AddItem "A216 WCB"
        .AddItem "A217 CA15"
        .AddItem "A351 CF8"
        .AddItem "A351 CF8M"
        .AddItem "A351 CF3"
        .AddItem "A351 CF3M"
        .AddItem "STELLITE"
    End With
End Sub

