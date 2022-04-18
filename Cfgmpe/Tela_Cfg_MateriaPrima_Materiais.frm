VERSION 5.00
Begin VB.Form Tela_Cfg_MateriaPrima_Materiais 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Selecione um material"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3615
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
   Begin VB.ComboBox CB_MAT 
      Height          =   315
      ItemData        =   "Tela_Cfg_MateriaPrima_Materiais.frx":0000
      Left            =   120
      List            =   "Tela_Cfg_MateriaPrima_Materiais.frx":0002
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Selecione um material nesta lista para a matéria-prima especificada acima"
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Escolha o material da matéria-prima:"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   2535
   End
   Begin VB.Label LB_MP 
      AutoSize        =   -1  'True
      Caption         =   "Matéria-Prima:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   1080
      Width           =   1230
   End
   Begin VB.Label LB_F 
      AutoSize        =   -1  'True
      Caption         =   "Material da Figura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   840
      Width           =   1605
   End
   Begin VB.Label LB_3 
      AutoSize        =   -1  'True
      Caption         =   "Matéria-Prima:"
      Height          =   195
      Left            =   420
      TabIndex        =   4
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label LB_2 
      AutoSize        =   -1  'True
      Caption         =   "Material da Figura:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1305
   End
   Begin VB.Label LB_1 
      Caption         =   "Não foi possível localizar no banco de dados de matéria-prima dados sobre o seguinte material:"
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3480
   End
End
Attribute VB_Name = "Tela_Cfg_MateriaPrima_Materiais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Voltar_Click()
    Tela_Cfg_MateriaPrima.MAT = CB_MAT.Text
    Tela_Cfg_MateriaPrima_Materiais.Hide
End Sub
