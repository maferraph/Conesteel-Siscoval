VERSION 5.00
Begin VB.Form Tela_Expedicao_EtiquetaSaco_Imprimir 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprimir etiquetas para sacos plásticos"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3720
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TXT_NE 
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Selecione a 1ª etiqueta da folha:"
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3495
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 8"
         Height          =   255
         Index           =   15
         Left            =   1800
         TabIndex        =   18
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 7"
         Height          =   255
         Index           =   14
         Left            =   1800
         TabIndex        =   17
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 6"
         Height          =   255
         Index           =   13
         Left            =   1800
         TabIndex        =   16
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 5"
         Height          =   255
         Index           =   12
         Left            =   1800
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 4"
         Height          =   255
         Index           =   11
         Left            =   1800
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 3"
         Height          =   255
         Index           =   10
         Left            =   1800
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 2"
         Height          =   255
         Index           =   9
         Left            =   1800
         TabIndex        =   12
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 2 Linha 1"
         Height          =   255
         Index           =   8
         Left            =   1800
         TabIndex        =   11
         Top             =   240
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 8"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   10
         Top             =   1920
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 7"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   9
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 6"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 5"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 4"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 3"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 2"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton RB_Etiqueta 
         Caption         =   "Coluna 1 Linha 1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   855
      Left            =   1800
      Picture         =   "Tela_Expedicao_EtiquetaSaco_Imprimir.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimir Pedido"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   2760
      Picture         =   "Tela_Expedicao_EtiquetaSaco_Imprimir.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Volta à Tela Principal"
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Número de Etiquetas:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   2640
      Width           =   1530
   End
End
Attribute VB_Name = "Tela_Expedicao_EtiquetaSaco_Imprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Imprimir_Click()
    Tela_Expedicao_EtiquetaSaco_Relatorio.PrintForm
End Sub
Private Sub BT_Voltar_Click()
    Unload Me
End Sub
