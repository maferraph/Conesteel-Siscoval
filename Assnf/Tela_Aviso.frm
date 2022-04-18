VERSION 5.00
Begin VB.Form Tela_Aviso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Aviso"
   ClientHeight    =   3852
   ClientLeft      =   36
   ClientTop       =   276
   ClientWidth     =   7176
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3852
   ScaleWidth      =   7176
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT 
      Caption         =   "&Fechar"
      Height          =   372
      Left            =   2880
      TabIndex        =   2
      Top             =   3360
      Width           =   1572
   End
   Begin VB.Image IMG 
      Height          =   2376
      Left            =   120
      Picture         =   "Tela_Aviso.frx":0000
      Stretch         =   -1  'True
      Top             =   840
      Width           =   4452
   End
   Begin VB.Label LB_Texto 
      Caption         =   $"Tela_Aviso.frx":185D9
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2472
      Left            =   4800
      TabIndex        =   1
      Top             =   840
      Width           =   2268
   End
   Begin VB.Label LB_Aviso 
      AutoSize        =   -1  'True
      Caption         =   "ATENÇÃO !"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   28.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   804
      Left            =   1800
      TabIndex        =   0
      Top             =   0
      Width           =   3468
   End
End
Attribute VB_Name = "Tela_Aviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BT_Click()
    Unload Tela_Aviso
End Sub
