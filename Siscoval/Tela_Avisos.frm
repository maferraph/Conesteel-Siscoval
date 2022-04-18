VERSION 5.00
Begin VB.Form Tela_Avisos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AVISO"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Mostrar este aviso novamente:"
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   2880
      Width           =   6735
      Begin VB.ComboBox CB_Aviso 
         Height          =   315
         ItemData        =   "Tela_Avisos.frx":0000
         Left            =   120
         List            =   "Tela_Avisos.frx":0010
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
      Begin VB.CommandButton BT_OK 
         Caption         =   "&OK"
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Mensagem:"
      Height          =   2055
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   6735
      Begin VB.TextBox TXT 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remetente:"
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.Label LB_Data 
         AutoSize        =   -1  'True
         Caption         =   "01/01/2001"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   3
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label LB 
         AutoSize        =   -1  'True
         Caption         =   "Data de Emissão:"
         Height          =   195
         Left            =   5400
         TabIndex        =   2
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label LB_Remetente 
         AutoSize        =   -1  'True
         Caption         =   "MAURÍCIO FERNANDES RAPHAEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3675
      End
   End
End
Attribute VB_Name = "Tela_Avisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Aviso()
    With DLL_BD
        LB_Remetente.Caption = .BDSIS_TBAVI_CPREM.Value
        LB_Data.Caption = .BDSIS_TBAVI_CPEMI.Value
        TXT.Text = .BDSIS_TBAVI_CPAVI.Value
        CB_Aviso.ListIndex = 0
        Me.Show vbModal
    End With
End Sub
Private Sub BT_OK_Click()
    With DLL_BD
        If CB_Aviso.ListIndex = -1 Then
            MsgBox "Escolha uma opção na lista.", vbInformation + vbOKOnly, "Falta dados"
            CB_Aviso.SetFocus
            Exit Sub
        ElseIf CB_Aviso.ListIndex = 0 Then 'Nunca
            .BDSIS_TBAVI.Delete
        ElseIf CB_Aviso.ListIndex = 1 Then 'Amanha
            .BDSIS_TBAVI.Edit
            .BDSIS_TBAVI_CPVEN.Value = DateAdd("d", 1, VBA.Date)
            .BDSIS_TBAVI.Update
        ElseIf CB_Aviso.ListIndex = 2 Then '1 Semana
            .BDSIS_TBAVI.Edit
            .BDSIS_TBAVI_CPVEN.Value = DateAdd("w", 1, VBA.Date)
            .BDSIS_TBAVI.Update
        ElseIf CB_Aviso.ListIndex = 3 Then '1 Mês
            .BDSIS_TBAVI.Edit
            .BDSIS_TBAVI_CPVEN.Value = DateAdd("m", 1, VBA.Date)
            .BDSIS_TBAVI.Update
        End If
    End With
    Unload Me
End Sub
