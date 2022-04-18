VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form Tela_CD 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   864
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1128
   LinkTopic       =   "Form1"
   ScaleHeight     =   864
   ScaleWidth      =   1128
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   677
      _ExtentY        =   677
      _Version        =   327680
   End
   Begin VB.FileListBox LA 
      Height          =   456
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   612
   End
End
Attribute VB_Name = "Tela_CD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

