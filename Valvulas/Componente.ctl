VERSION 5.00
Begin VB.UserControl Componente 
   ClientHeight    =   765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7365
   ScaleHeight     =   765
   ScaleWidth      =   7365
   Begin VB.Frame FR_Dados 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.ComboBox CB_Componente 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.ComboBox CB_Bitola 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
      Begin VB.ComboBox CB_Material 
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label LB_Material 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   195
         Left            =   4560
         TabIndex        =   6
         Top             =   0
         Width           =   600
      End
      Begin VB.Label LB_Bitola 
         AutoSize        =   -1  'True
         Caption         =   "Bitola:"
         Height          =   195
         Left            =   2880
         TabIndex        =   5
         Top             =   0
         Width           =   435
      End
      Begin VB.Label LB_Peca 
         AutoSize        =   -1  'True
         Caption         =   "Componente:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   945
      End
   End
End
Attribute VB_Name = "Componente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Sub CB_Componente_Change()
    CB_Componente.Clear
    CB_Bitola.Clear
    CB_Material.Clear
    CB_Bitola.Enabled = False
    CB_Material.Enabled = False
    LB_Bitola.Enabled = False
    LB_Material.Enabled = False
    'Carrega lista de peças e bitolas
    Select Case CB_Componente.List(CB_Componente.ListIndex)
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
Private Sub CB_Componente_Click()
    CB_Componente_Change
End Sub
Private Sub UserControl_Initialize()
    CB_Componente.Clear
    CB_Bitola.Clear
    CB_Material.Clear
    CB_Bitola.Enabled = False
    CB_Material.Enabled = False
    LB_Bitola.Enabled = False
    LB_Material.Enabled = False
    'Carrega Lista de Peças
    With CB_Componente
        .AddItem "CORPO GAVETA 800 NPT"
        .AddItem "CORPO GAVETA 800 SW"
        .AddItem "CORPO GAVETA 800 BSP"
        .AddItem "CORPO GAVETA 800 NPT SEDE INTEGRAL"
        .AddItem "CORPO GAVETA 800 SW SEDE INTEGRAL"
        .AddItem "CORPO GAVETA 800 BSP SEDE INTEGRAL"
        .AddItem "CORPO GAVETA 1500 NPT"
        .AddItem "CORPO GAVETA 1500 SW"
        .AddItem "CORPO GAVETA 1500 BSP"
        .AddItem "CORPO GAVETA 1500 NPT SEDE INTEGRAL"
        .AddItem "CORPO GAVETA 1500 SW SEDE INTEGRAL"
        .AddItem "CORPO GAVETA 1500 BSP SEDE INTEGRAL"
        .AddItem "CORPO GAVETA ADAPTADO 150"
        .AddItem "CORPO GAVETA ADAPTADO 300"
        .AddItem "CORPO GAVETA ADAPTADO 600"
        .AddItem "CORPO GAVETA ADAPTADO 900"
        .AddItem "CORPO GAVETA ADAPTADO 1500"
        .AddItem "CORPO GAVETA ADAPTADO 150 SEDE INTEGRAL"
        .AddItem "CORPO GAVETA ADAPTADO 300 SEDE INTEGRAL"
        .AddItem "CORPO GAVETA ADAPTADO 600 SEDE INTEGRAL"
        .AddItem "CORPO GAVETA ADAPTADO 900 SEDE INTEGRAL"
        .AddItem "CORPO GAVETA ADAPTADO 1500 SEDE INTEGRAL"
        .AddItem "CORPO GLOBO 800 NPT"
        .AddItem "CORPO GLOBO 800 SW"
        .AddItem "CORPO GLOBO 800 BSP"
        .AddItem "CORPO GLOBO 800 NPT SEDE INTEGRAL"
        .AddItem "CORPO GLOBO 800 SW SEDE INTEGRAL"
        .AddItem "CORPO GLOBO 800 BSP SEDE INTEGRAL"
        .AddItem "CORPO GLOBO 1500 NPT"
        .AddItem "CORPO GLOBO 1500 SW"
        .AddItem "CORPO GLOBO 1500 BSP"
        .AddItem "CORPO GLOBO 1500 NPT SEDE INTEGRAL"
        .AddItem "CORPO GLOBO 1500 SW SEDE INTEGRAL"
        .AddItem "CORPO GLOBO 1500 BSP SEDE INTEGRAL"
        .AddItem "CORPO GLOBO ADAPTADO 150"
        .AddItem "CORPO GLOBO ADAPTADO 300"
        .AddItem "CORPO GLOBO ADAPTADO 600"
        .AddItem "CORPO GLOBO ADAPTADO 900"
        .AddItem "CORPO GLOBO ADAPTADO 1500"
        .AddItem "CORPO GLOBO ADAPTADO 150 SEDE INTEGRAL"
        .AddItem "CORPO GLOBO ADAPTADO 300 SEDE INTEGRAL"
        .AddItem "CORPO GLOBO ADAPTADO 600 SEDE INTEGRAL"
        .AddItem "CORPO GLOBO ADAPTADO 900 SEDE INTEGRAL"
        .AddItem "CORPO GLOBO ADAPTADO 1500 SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO 800 NPT"
        .AddItem "CORPO PISTÃO 800 SW"
        .AddItem "CORPO PISTÃO 800 BSP"
        .AddItem "CORPO PISTÃO 800 NPT SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO 800 SW SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO 800 BSP SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO 1500 NPT"
        .AddItem "CORPO PISTÃO 1500 SW"
        .AddItem "CORPO PISTÃO 1500 BSP"
        .AddItem "CORPO PISTÃO 1500 NPT SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO 1500 SW SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO 1500 BSP SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO ADAPTADO 150"
        .AddItem "CORPO PISTÃO ADAPTADO 300"
        .AddItem "CORPO PISTÃO ADAPTADO 600"
        .AddItem "CORPO PISTÃO ADAPTADO 900"
        .AddItem "CORPO PISTÃO ADAPTADO 1500"
        .AddItem "CORPO PISTÃO ADAPTADO 150 SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO ADAPTADO 300 SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO ADAPTADO 600 SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO ADAPTADO 900 SEDE INTEGRAL"
        .AddItem "CORPO PISTÃO ADAPTADO 1500 SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA 800 NPT"
        .AddItem "CORPO PORTINHOLA 800 SW"
        .AddItem "CORPO PORTINHOLA 800 BSP"
        .AddItem "CORPO PORTINHOLA 800 NPT SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA 800 SW SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA 800 BSP SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA 1500 NPT"
        .AddItem "CORPO PORTINHOLA 1500 SW"
        .AddItem "CORPO PORTINHOLA 1500 BSP"
        .AddItem "CORPO PORTINHOLA 1500 NPT SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA 1500 SW SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA 1500 BSP SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA ADAPTADO 150"
        .AddItem "CORPO PORTINHOLA ADAPTADO 300"
        .AddItem "CORPO PORTINHOLA ADAPTADO 600"
        .AddItem "CORPO PORTINHOLA ADAPTADO 900"
        .AddItem "CORPO PORTINHOLA ADAPTADO 1500"
        .AddItem "CORPO PORTINHOLA ADAPTADO 150 SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA ADAPTADO 300 SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA ADAPTADO 600 SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA ADAPTADO 900 SEDE INTEGRAL"
        .AddItem "CORPO PORTINHOLA ADAPTADO 1500 SEDE INTEGRAL"
        
        .AddItem "CASTELO APARAFUSADO GAVETA 800"
        .AddItem "CASTELO APARAFUSADO GAVETA 1500"
        .AddItem "CASTELO APARAFUSADO GLOBO 800"
        .AddItem "CASTELO APARAFUSADO GLOBO 1500"
        .AddItem "TAMPA APARAFUSADA PISTÃO 800"
        .AddItem "TAMPA APARAFUSADA PISTÃO 1500"
        
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


