VERSION 5.00
Begin VB.Form Tela_Fabrica_CarteiraPedidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carteira de Pedidos"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   11340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TXT_Z 
      Height          =   285
      Index           =   0
      Left            =   8760
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox TXT_O 
      Height          =   285
      Index           =   0
      Left            =   9600
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   10320
      Picture         =   "Tela_Fabrica_CarteiraPedidos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Label LB 
      AutoSize        =   -1  'True
      Caption         =   "Prazo"
      Height          =   195
      Index           =   9
      Left            =   8760
      TabIndex        =   4
      Top             =   120
      Width           =   405
   End
   Begin VB.Label LB 
      AutoSize        =   -1  'True
      Caption         =   "Observações"
      Height          =   195
      Index           =   10
      Left            =   9600
      TabIndex        =   3
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "Tela_Fabrica_CarteiraPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Voltar_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    With Valvula1
        If .Tipo <> "" And .Extremidade <> "" And .Classe <> "" And .Bitola <> "" And .Material <> "" Then
            MsgBox .Tipo & " " & .Extremidade & " " & .Classe & " " & .Bitola & " " & .Material
        End If
    End With
End Sub
