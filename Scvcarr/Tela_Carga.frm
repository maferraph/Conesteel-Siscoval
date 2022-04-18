VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Carga 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aguarde..."
   ClientHeight    =   372
   ClientLeft      =   36
   ClientTop       =   264
   ClientWidth     =   5628
   ControlBox      =   0   'False
   Enabled         =   0   'False
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   7.8
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   372
   ScaleWidth      =   5628
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar BP 
      Align           =   1  'Align Top
      Height          =   372
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5628
      _ExtentX        =   9927
      _ExtentY        =   656
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "Tela_Carga"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
