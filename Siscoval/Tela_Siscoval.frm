VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Tela_Siscoval 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Siscoval"
   ClientHeight    =   705
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   1  'Dash
   HelpContextID   =   10001
   Icon            =   "Tela_Siscoval.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   705
   ScaleWidth      =   9225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer TEMPO 
      Interval        =   30000
      Left            =   720
      Top             =   840
   End
   Begin MSComctlLib.ImageList LI 
      Left            =   0
      Top             =   720
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   119
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":030A
            Key             =   ""
            Object.Tag             =   "Logon"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":0626
            Key             =   ""
            Object.Tag             =   "Sair"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":0942
            Key             =   ""
            Object.Tag             =   "Siscoval"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":0C5E
            Key             =   ""
            Object.Tag             =   "Nota"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":10B2
            Key             =   ""
            Object.Tag             =   "ConsultaEstoque"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":13CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":16EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1A06
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1D22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":203E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":235A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":2676
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":2992
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":2CAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":2FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":32E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":3602
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":391E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":3C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":3F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":4272
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":458E
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":48AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":4BC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":4EE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":51FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":551A
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":5836
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":5B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":5E6E
            Key             =   ""
            Object.Tag             =   "CQ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":618A
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":64A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":67C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":6ADE
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":6DFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":7116
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":7432
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":774E
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":7A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":7D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":80A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":897E
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":925A
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":9B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":A412
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":ACEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":B5CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":BEA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":C782
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":D05E
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":D93A
            Key             =   ""
            Object.Tag             =   "Pedido"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":E216
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":EAF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":F3CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":F6EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":FFC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":102E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":10736
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":10B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":10FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":112FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1174E
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":11A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":11D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":120A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":124F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":12812
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":12B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":12E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1329E
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":135BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":138D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":13BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":13F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":14362
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":147B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":14AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":14DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1510A
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1555E
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":159B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":15CCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":15FEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1643E
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1675A
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":16A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":16D92
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":170AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":173CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":176E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":17A02
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":17D1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":18172
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1848E
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":187AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":18AC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":18DE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":190FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1941A
            Key             =   ""
            Object.Tag             =   "NF"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1986E
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":19B8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":19EA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1A1C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1A4DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1A7FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1AB16
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1B26A
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1B586
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1B9DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1BCF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1C14A
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1C466
            Key             =   ""
            Object.Tag             =   "Empresas"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1C782
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1CA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1CDBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1D0D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1D3F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1D846
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Tela_Siscoval.frx":1DB62
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BF 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "LI"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Sair"
            Object.ToolTipText     =   "Encerrar o Siscoval"
            Object.Tag             =   "Sair"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Minimizar"
            Object.ToolTipText     =   "Minimiza o Siscoval"
            Object.Tag             =   "Minimizar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Consulta"
            Object.ToolTipText     =   "Consulta Rápida de Estoque"
            Object.Tag             =   "Consulta"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cotacao"
            Object.ToolTipText     =   "Cotação de Preços"
            Object.Tag             =   "Cotacao"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pedido"
            Object.ToolTipText     =   "Pedidos de Clientes"
            Object.Tag             =   "Pedido"
            ImageIndex      =   51
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NF"
            Object.ToolTipText     =   "Assistente de Nota Fiscal"
            Object.Tag             =   "NF"
            ImageIndex      =   99
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CQ"
            Object.ToolTipText     =   "Certificados de Qualidade"
            Object.Tag             =   "CQ"
            ImageIndex      =   30
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Empresas"
            Object.ToolTipText     =   "Cadastro de Empresas"
            Object.Tag             =   "Empresas"
            ImageIndex      =   45
         EndProperty
      EndProperty
   End
   Begin VB.Menu Menu_Principal 
      Caption         =   "&Principal"
      Begin VB.Menu Menu_Principal_MudarSenha 
         Caption         =   "Mudar a Senha"
      End
      Begin VB.Menu Menu_Principal_ModoEspera 
         Caption         =   "&Modo de Espera"
      End
      Begin VB.Menu Menu_Principal_Minimizar 
         Caption         =   "Minimizar"
      End
      Begin VB.Menu Menu_Principal_Lixo_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Principal_Sair 
         Caption         =   "&Sair do Sistema"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Menu_Estoque 
      Caption         =   "&Estoque"
      Begin VB.Menu Menu_Estoque_ConsultaRápida 
         Caption         =   "Consulta Rápida"
      End
      Begin VB.Menu Menu_Estoque_FichaEstoque 
         Caption         =   "Ficha de Estoque"
      End
      Begin VB.Menu Menu_Estoque_Lixo1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Estoque_EntradaPecas 
         Caption         =   "Entrada de Peças"
      End
      Begin VB.Menu Menu_Estoque_RequisicaoPecas 
         Caption         =   "Requisição de Peças"
      End
      Begin VB.Menu Menu_Estoque_Balanco 
         Caption         =   "Balanço"
      End
   End
   Begin VB.Menu Menu_Escritorio 
      Caption         =   "E&scritório"
      Begin VB.Menu Menu_Escritorio_NotaFiscal 
         Caption         =   "Nota Fiscal"
         Begin VB.Menu Menu_Escritorio_NotaFiscal_Assistente 
            Caption         =   "Assistente de Nota Fiscal"
            Shortcut        =   ^N
         End
         Begin VB.Menu Menu_Escritorio_NotaFiscal_Emitidas 
            Caption         =   "Notas Fiscais Emitidas"
         End
      End
      Begin VB.Menu Menu_Escritorio_Certificado 
         Caption         =   "Certificado"
      End
      Begin VB.Menu Menu_Escritorio_Lixo_1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Escritorio_PV 
         Caption         =   "P.V. / A.C.C."
      End
      Begin VB.Menu Menu_Escritorio_Cotacoes 
         Caption         =   "Cotações"
      End
      Begin VB.Menu Menu_Escritorio_Pedidos 
         Caption         =   "Pedidos"
      End
      Begin VB.Menu Menu_Escritorio_Lixo_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Escritorio_CadastrosEmpresas 
         Caption         =   "Cadastros de Empresas"
      End
      Begin VB.Menu Menu_Escritorio_CadastrosBancos 
         Caption         =   "Cadastros de Bancos"
      End
   End
   Begin VB.Menu Menu_Fabrica 
      Caption         =   "Fá&brica"
      Begin VB.Menu Menu_Fabrica_CP 
         Caption         =   "Carteira de Pedidos"
      End
      Begin VB.Menu Menu_Fabrica_OF 
         Caption         =   "Ordem de Fabricação"
      End
      Begin VB.Menu Menu_Fabrica_OM 
         Caption         =   "Ordem de Montagem"
      End
   End
   Begin VB.Menu Menu_Expedicao 
      Caption         =   "E&xpedição"
      Begin VB.Menu Menu_Expedicao_Etiquetas 
         Caption         =   "Etiquetas"
         Begin VB.Menu Menu_Expedicao_Etiquetas_Pacotes 
            Caption         =   "Pacotes"
         End
         Begin VB.Menu Menu_Expedicao_Etiquetas_Sacos 
            Caption         =   "Sacos"
         End
      End
      Begin VB.Menu Menu_Expedicao_MI 
         Caption         =   "Manual de Instruções"
      End
      Begin VB.Menu Menu_Expedicao_RLP 
         Caption         =   "Relatório de Líquido Penetrante"
      End
   End
   Begin VB.Menu Menu_Ferramentas 
      Caption         =   "&Ferramentas"
      Begin VB.Menu Menu_Ferramentas_ConvUnit 
         Caption         =   "Conversor de Unidades"
      End
   End
   Begin VB.Menu Menu_Configuracoes 
      Caption         =   "&Configurações"
      Begin VB.Menu Menu_Configuracoes_NotaFiscal 
         Caption         =   "Nota Fiscal"
         Begin VB.Menu Menu_Configuracoes_NotaFiscal_CodigosFiscais 
            Caption         =   "Códigos Fiscais"
         End
         Begin VB.Menu Menu_Configuracoes_NotaFiscal_Declaracoes 
            Caption         =   "Declarações"
         End
         Begin VB.Menu Menu_Configuracoes_NotaFiscal_ConfiguracoesImpressao 
            Caption         =   "Configurações de Impressão"
         End
      End
      Begin VB.Menu Menu_Configuracoes_Estoque 
         Caption         =   "Estoque"
         Begin VB.Menu Menu_Configuracoes_Estoque_Assistente 
            Caption         =   "Assistente de Estoque"
         End
         Begin VB.Menu Menu_Configuracoes_Estoque_MateriaPrima 
            Caption         =   "Matéria-Prima"
         End
         Begin VB.Menu Menu_Configuracoes_Estoque_Lixo1 
            Caption         =   "-"
         End
         Begin VB.Menu Menu_Configuracoes_Estoque_PrecoPeso 
            Caption         =   "Preços e Pesos"
         End
         Begin VB.Menu Menu_Configuracoes_Estoque_CFeST 
            Caption         =   "C.F. e S.T."
         End
         Begin VB.Menu Menu_Configuracoes_Estoque_Aliquotas 
            Caption         =   "Alíquotas"
         End
      End
      Begin VB.Menu Menu_Configuracoes_Fabrica 
         Caption         =   "Fábrica"
         Begin VB.Menu Menu_Configuracoes_Fabrica_Maquinas 
            Caption         =   "Equipamentos"
         End
         Begin VB.Menu Menu_Configuracoes_Fabrica_Processos 
            Caption         =   "Etapas dos Processos"
         End
      End
      Begin VB.Menu Menu_Configuracoes_lixo1 
         Caption         =   "-"
      End
      Begin VB.Menu Menu_Configuracoes_Imagens 
         Caption         =   "Imagens"
      End
      Begin VB.Menu Menu_Configuracoes_Grupos 
         Caption         =   "Grupos"
      End
      Begin VB.Menu Menu_Configuracoes_Usuarios 
         Caption         =   "Usuários"
      End
   End
End
Attribute VB_Name = "Tela_Siscoval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const NOMEAPLIC As String = "Sistema Siscoval"
Private Sub BF_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ERRO_SISCOVAL
    If Button.Key = "Sair" Then 'Encerra o Sistema
        Menu_Principal_Sair_Click
    ElseIf Button.Key = "Minimizar" Then 'Minimiza a Barra
        Menu_Principal_Minimizar_Click
    ElseIf Button.Key = "NF" Then 'Assistente da Nota Fiscal
        Menu_Escritorio_NotaFiscal_Assistente_Click
    ElseIf Button.Key = "Cotacao" Then 'Cotação de Preços
        Menu_Escritorio_Cotacoes_Click
    ElseIf Button.Key = "Pedido" Then 'Pedido de Estoque
        Menu_Escritorio_Pedidos_Click
    ElseIf Button.Key = "Consulta" Then 'Consulta Rápida ao Estoque
        Menu_Estoque_ConsultaRápida_Click
    ElseIf Button.Key = "CQ" Then 'Certificados de Qualidade
'        Menu_Escritorio_Certificado_Click
    ElseIf Button.Key = "Empresas" Then 'Cadastro de Empresas
        Menu_Escritorio_CadastrosEmpresas_Click
    End If
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Estoque_Aliquotas_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgali")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Estoque_Assistente_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgest")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Estoque_CFeST_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgecf")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Estoque_MateriaPrima_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgmpe")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Estoque_PrecoPeso_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgepp")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Fabrica_Maquinas_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgfamaq")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Fabrica_Processos_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgfapro")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Grupos_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfggru")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Imagens_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgimg")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_NotaFiscal_CodigosFiscais_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgcfi")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_NotaFiscal_ConfiguracoesImpressao_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfginf")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_NotaFiscal_Declaracoes_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgedc")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Configuracoes_Usuarios_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgusu")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Escritorio_CadastrosBancos_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cfgbco")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Escritorio_CadastrosEmpresas_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cademp")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Escritorio_Cotacoes_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Cotest")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Escritorio_Pedidos_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Pedest")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Escritorio_PV_Click()
    Tela_Escritorio_PropostaVendas.Show vbModal
End Sub
Private Sub Menu_Estoque_ConsultaRápida_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Estcrp")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Estoque_FichaEstoque_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Estcon")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Expedicao_Etiquetas_Sacos_Click()
    Tela_Expedicao_EtiquetaSaco.Show vbModal
End Sub

Private Sub Menu_Expedicao_MI_Click()
    Tela_ManualInstrucao.Show vbModal
End Sub

Private Sub Menu_Fabrica_CP_Click()
    Tela_Fabrica_CarteiraPedidos.Show vbModal
End Sub
Private Sub Menu_Fabrica_OM_Click()
    Tela_Fabrica_OM.Show vbModal
End Sub
Private Sub Menu_Ferramentas_ConvUnit_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Convunit")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Escritorio_NotaFiscal_Assistente_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Assnf")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Escritorio_NotaFiscal_Emitidas_Click()
    On Error GoTo ERRO_SISCOVAL
    ChamaDLL ("Nfemit")
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Principal_Minimizar_Click()
    On Error GoTo ERRO_SISCOVAL
    MinimizaSiscoval
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Principal_ModoEspera_Click()
    On Error GoTo ERRO_SISCOVAL
    'O modo de espera coloca o Sistema como icone na barra de tarefas
    'sem estar conectado usuario; ao clicar no icone, abrira tela de
    'logon de usuario
    Tela_Siscoval.Visible = False
    Tela_Principal.Visible = False
    With IconeTela
        .cbSize = Len(IconeTela)
        .hwnd = Tela_Principal.ICONE_INICIAR.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Tela_Principal.ICONE_INICIAR.Picture
        .szTip = Tela_Principal.ICONE_INICIAR.ToolTipText & vbNullChar
        .Tela = "LOGON"
    End With
    Shell_NotifyIcon NIM_ADD, IconeTela
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Principal_MudarSenha_Click()
    On Error GoTo ERRO_SISCOVAL
    Tela_MudarSenha.Show vbModal
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
Private Sub Menu_Principal_Sair_Click()
    DLL_FUNCS.RegistraEvento "Finalização de Sistema", ""
    End
End Sub
Sub Form_Load()
    Set DLL_BD = New Scvbd.Classe_Scvbd
    'Abre bancos de dados
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Abrindo Tabelas
    If DLL_BD.AbreTabela_Avisos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Abrindo Campos
    If DLL_BD.AbreCampos_Avisos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
'    Exit Sub
'ERRO_ABERTURA:
'    MsgBox "Não foi possível abrir o banco de dados para o Siscoval", vbCritical + vbOKOnly, "ERRO"
'    End
End Sub
Sub Form_Unload(Cancel As Integer)
    'Fecha tabelas
    If DLL_BD.FechaTabela_Avisos(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
End Sub
Private Sub TEMPO_Timer()
    On Error Resume Next
    If lTempo = False Then Exit Sub
    With DLL_BD
        If .BDSIS_TBAVI.RecordCount > 0 Then
            .BDSIS_TBAVI.MoveFirst
            Do While Not .BDSIS_TBAVI.EOF
                If (.BDSIS_TBAVI_CPDES.Value = Usuario And .BDSIS_TBAVI_CPVEN.Value = VBA.Date) Or _
                   (.BDSIS_TBAVI_CPDES.Value = DLL_FUNCS.PegaNomeComputador And .BDSIS_TBAVI_CPVEN.Value = VBA.Date) Then
                    Tela_Avisos.Aviso
                    Exit Do
                End If
                .BDSIS_TBAVI.MoveNext
            Loop
        End If
    End With
End Sub


'**************   FUNCOES   **************

Private Static Sub MinimizaSiscoval()
    On Error GoTo ERRO_SISCOVAL
    'Esconde tela_siscoval e insere icone na barra de tarefas
    Tela_Siscoval.Visible = False
    Tela_Principal.Visible = False
    With IconeTela
        .cbSize = Len(IconeTela)
        .hwnd = Tela_Principal.ICONE_SISCOVAL.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Tela_Principal.ICONE_SISCOVAL.Picture
        .szTip = Tela_Principal.ICONE_SISCOVAL.ToolTipText & vbNullChar
        .Tela = "SISCOVAL"
    End With
    Shell_NotifyIcon NIM_ADD, IconeTela
ERRO_SISCOVAL: If Err Then If DLL_FUNCS.MensagemErro(DLL_FUNCS.PegaUsuario, DLL_FUNCS.PegaNomeComputador, Err.Number, Err.Description, Err.Source, NOMEAPLIC, Err.HelpFile, Err.HelpContext) = True Then Resume Next
End Sub
