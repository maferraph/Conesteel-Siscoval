VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Tela_Cotacao_MP 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuração da Matéria-Prima"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BT_Voltar 
      Caption         =   "&Voltar"
      Height          =   855
      Left            =   4200
      Picture         =   "Tela_Cotacao_MP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Volta à Cotação"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton BT_Imprimir 
      Caption         =   "I&mprimir"
      Height          =   855
      Left            =   2400
      Picture         =   "Tela_Cotacao_MP.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Imprimir a dados sobre matéria-prima"
      Top             =   3240
      Width           =   855
   End
   Begin VB.Frame FR 
      Height          =   615
      Index           =   7
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7695
      Begin VB.Label LB_Quantidade 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quantidade:"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   0
         Width           =   870
      End
      Begin VB.Label LB_Material 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5520
         TabIndex        =   7
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label LB_Bitola 
         AutoSize        =   -1  'True
         Caption         =   "Bitola:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   3960
         TabIndex        =   6
         Top             =   240
         Width           =   780
      End
      Begin VB.Label LB_Figura 
         AutoSize        =   -1  'True
         Caption         =   "Figura:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2400
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.Label LB6 
         AutoSize        =   -1  'True
         Caption         =   "Figura:"
         Height          =   195
         Left            =   2400
         TabIndex        =   4
         Top             =   0
         Width           =   495
      End
      Begin VB.Label LB7 
         AutoSize        =   -1  'True
         Caption         =   "Bitola:"
         Height          =   195
         Left            =   3960
         TabIndex        =   3
         Top             =   0
         Width           =   450
      End
      Begin VB.Label LB8 
         AutoSize        =   -1  'True
         Caption         =   "Material:"
         Height          =   195
         Left            =   5520
         TabIndex        =   2
         Top             =   0
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid FG_MP 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Lista de Matéria-Prima"
      Top             =   720
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   4260
      _Version        =   393216
      SelectionMode   =   1
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "Tela_Cotacao_MP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Cotação de Estoque - MP"
Dim RespMsg, I As Integer, SeekErro As Boolean, sDescricao As String, J As Integer
Dim MATPRI As MP
Private Type MP
    QUA As String
    PEC As String
    NOM As String
    BIT As String
    MAT As String
End Type
Private Sub BT_Voltar_Click()
    Me.Hide
End Sub

'***************************************************
'                  FUNÇÕES E ROTINAS
'***************************************************

Public Static Sub DetalhesMP(QUA As String, FIG As String, BIT As String, MAT As String, DES As String)
    If FIG = "" Or BIT = "" Or MAT = "" Then Exit Sub
    TelaEmEspera True
    Tela_Cotacao.BP.Max = 5
    Tela_Cotacao.BP.Value = 0
    Tela_Cotacao.BP.Value = Tela_Cotacao.BP.Value + 1
    LB_Quantidade.Caption = QUA
    LB_Figura.Caption = FIG
    LB_Bitola.Caption = BIT
    LB_Material.Caption = MAT
    sDescricao = DES
    Tela_Cotacao.BP.Value = Tela_Cotacao.BP.Value + 1
    MontaFG_MP
    Tela_Cotacao.BP.Value = Tela_Cotacao.BP.Value + 1
    LeMP
    Tela_Cotacao.BP.Value = Tela_Cotacao.BP.Value + 1
    ConsultaSaldo
    Tela_Cotacao.BP.Value = Tela_Cotacao.BP.Value + 1
    TelaEmEspera False
    Tela_Cotacao.BP.Value = 0
End Sub
Private Static Sub TelaEmEspera(Estado As Boolean)
    If Estado = True Then
        Me.MousePointer = vbHourglass
        Me.Enabled = False
    Else
        Me.MousePointer = vbDefault
        Me.Enabled = True
    End If
End Sub
Private Static Sub MontaFG_MP()
    With FG_MP
        .Clear
        .FixedCols = 0
        .Cols = 9
        .Rows = 1
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignCenterCenter
        .ColAlignment(8) = flexAlignCenterCenter
        .ColWidth(0) = 1800
        .ColWidth(1) = 1600
        .ColWidth(2) = 1600
        .ColWidth(3) = 1600
        .ColWidth(4) = 1800
        .ColWidth(5) = 1800
        .ColWidth(6) = 2200
        .ColWidth(7) = 2200
        .ColWidth(8) = 2200
        .TextArray(0) = "Componente"
        .TextArray(1) = "Nome"
        .TextArray(2) = "Bitola"
        .TextArray(3) = "Material"
        .TextArray(4) = "Quant.Comp.Unitário"
        .TextArray(5) = "Quant.Comp.Necessário"
        .TextArray(6) = "ESTOQUE - Componentes"
        .TextArray(7) = "ESTOQUE - Prod.Andamento"
        .TextArray(8) = "ESTOQUE - Matéria-Prima"
    End With
End Sub
Private Static Sub ProcuraMP(QUA As String, PEC As String, NOM As String, BIT As String, MAT As String)
    MATPRI.QUA = ""
    MATPRI.PEC = ""
    MATPRI.NOM = ""
    MATPRI.BIT = ""
    MATPRI.MAT = ""
    With Tela_Cotacao.DLL_BD
        'altera para novos indices
        .BDSIS_TBMPQ.Index = "Índice de Quantidades"
        .BDSIS_TBMPP.Index = "Índice de Peças"
        .BDSIS_TBMPN.Index = "Índice de Nomes"
        .BDSIS_TBMPB.Index = "Índice de Bitolas"
        .BDSIS_TBMPM.Index = "Índice de Materiais"
        'procura pecas
        .BDSIS_TBMPQ.Seek "=", QUA
        If Not .BDSIS_TBMPQ.NoMatch Then MATPRI.QUA = .BDSIS_TBMPQ_CPQUA.Value
        .BDSIS_TBMPP.Seek "=", PEC
        If Not .BDSIS_TBMPP.NoMatch Then MATPRI.PEC = .BDSIS_TBMPP_CPPEC.Value
        .BDSIS_TBMPN.Seek "=", NOM
        If Not .BDSIS_TBMPN.NoMatch Then MATPRI.NOM = .BDSIS_TBMPN_CPNOM.Value
        .BDSIS_TBMPB.Seek "=", BIT
        If Not .BDSIS_TBMPB.NoMatch Then MATPRI.BIT = .BDSIS_TBMPB_CPBIT.Value
        .BDSIS_TBMPM.Seek "=", MAT
        If Not .BDSIS_TBMPM.NoMatch Then MATPRI.MAT = .BDSIS_TBMPM_CPMAT.Value
        'altera para velhos indices
        .BDSIS_TBMPQ.Index = "Quantidades"
        .BDSIS_TBMPP.Index = "Peças"
        'DLL_BD.BDSIS_TBMPN.Index = "Nomes"
        .BDSIS_TBMPB.Index = "Bitolas"
        .BDSIS_TBMPM.Index = "Materiais"
    End With
End Sub
Private Static Sub DivideMP(TIPO As String, Valor As String, LinIni As Integer)
    If TIPO = "" Or Valor = "" Then Exit Sub
    'comeca dividir
    Dim cA As String, nA As Integer
    cA = ""
    nA = LinIni
    For I = 1 To Len(Valor)
        If Mid(Valor, I, 1) <> ";" Then
            cA = cA & Mid(Valor, I, 1)
        ElseIf Mid(Valor, I, 1) = ";" Then
            'insere dados
            If TIPO = "QUA" Then 'quantidade
                FG_MP.AddItem (1)
                FG_MP.TextMatrix(FG_MP.Rows - 1, 4) = Format(cA, "###,##0.00")
                FG_MP.TextMatrix(FG_MP.Rows - 1, 5) = Format((Val(LB_Quantidade.Caption) * Val(cA)), "###,##0.00")
            ElseIf TIPO = "PEC" Then 'peca
                FG_MP.TextMatrix(nA, 0) = cA
            ElseIf TIPO = "NOM" Then 'nome
                FG_MP.TextMatrix(nA, 1) = cA
            ElseIf TIPO = "BIT" Then 'bitola
                FG_MP.TextMatrix(nA, 2) = cA
            ElseIf TIPO = "MAT" Then 'material
                FG_MP.TextMatrix(nA, 3) = cA
            End If
            cA = ""
            If TIPO <> "QUA" Then nA = nA + 1
        End If
    Next I
    'insere o primeiro ou o ultimo
    If TIPO = "QUA" Then 'quantidade
        FG_MP.AddItem (1)
        FG_MP.TextMatrix(FG_MP.Rows - 1, 4) = Format(cA, "###,##0.00")
        FG_MP.TextMatrix(FG_MP.Rows - 1, 5) = Format((Val(LB_Quantidade.Caption) * Val(cA)), "###,##0.00")
    ElseIf TIPO = "PEC" Then 'peca
        FG_MP.TextMatrix(FG_MP.Rows - 1, 0) = cA
    ElseIf TIPO = "NOM" Then 'nome
        FG_MP.TextMatrix(FG_MP.Rows - 1, 1) = cA
    ElseIf TIPO = "BIT" Then 'bitola
        FG_MP.TextMatrix(FG_MP.Rows - 1, 2) = cA
    ElseIf TIPO = "MAT" Then 'material
        FG_MP.TextMatrix(FG_MP.Rows - 1, 3) = cA
    End If
End Sub
Private Static Sub LeMP()
    'procura configuração desta peça
    With Tela_Cotacao.DLL_BD
        .BDSIS_TBEST.Seek "=", LB_Figura.Caption, LB_Bitola.Caption, LB_Material.Caption
        If Not .BDSIS_TBEST.NoMatch Then
            If IsEmpty(.BDSIS_TBEST_CPINQ.Value) = True Or _
               IsEmpty(.BDSIS_TBEST_CPINP.Value) = True Or _
               IsEmpty(.BDSIS_TBEST_CPINN.Value) = True Or _
               IsEmpty(.BDSIS_TBEST_CPINB.Value) = True Or _
               IsEmpty(.BDSIS_TBEST_CPINM.Value) Then
                MsgBox "Um ou mais dados sobre a matéria-prima deste ítem podem estar faltando, configure-a antes de consultar.", vbInformation + vbOKOnly, NOMEAPLIC
                Exit Sub
            End If
            ProcuraMP .BDSIS_TBEST_CPINQ.Value, .BDSIS_TBEST_CPINP.Value, .BDSIS_TBEST_CPINN.Value, .BDSIS_TBEST_CPINB.Value, .BDSIS_TBEST_CPINM.Value
            DivideMP "QUA", MATPRI.QUA, 1
            DivideMP "PEC", MATPRI.PEC, 1
            DivideMP "NOM", MATPRI.NOM, 1
            DivideMP "BIT", MATPRI.BIT, 1
            DivideMP "MAT", MATPRI.MAT, 1
        Else
            MsgBox "Não foi possível localizar a ficha de estoque - verifique os dados digitados.", vbInformation + vbOKOnly, NOMEAPLIC
            Exit Sub
        End If
    End With
End Sub
Private Static Sub ConsultaSaldo()
    Dim sTipoPeca As String
    'pega MP dos componentes carregados
    With Tela_Cotacao.DLL_BD
        For I = 1 To FG_MP.Rows - 1
            SeekErro = False
            If FG_MP.TextMatrix(I, 0) <> "" And FG_MP.TextMatrix(I, 2) <> "" And FG_MP.TextMatrix(I, 3) <> "" Then
                .BDSIS_TBEST.Seek "=", FG_MP.TextMatrix(I, 0), FG_MP.TextMatrix(I, 2), FG_MP.TextMatrix(I, 3)
                If Not .BDSIS_TBEST.NoMatch Then
                    'verifica o tipo de peca
                    If Left(FG_MP.TextMatrix(I, 0), 2) = "CP" Then 'se for COMPONENTE
                        sTipoPeca = "COM"
                        FG_MP.TextMatrix(I, 6) = Format((CDbl(.BDSIS_TBEST_CPEST.Value) - CDbl(.BDSIS_TBEST_CPVEN.Value)), "###,##0.00")
                    ElseIf Left(FG_MP.TextMatrix(I, 0), 2) = "PA" Then 'se for PRODUÇAO-ANDAMENTO
                        sTipoPeca = "PA"
                        FG_MP.TextMatrix(I, 6) = "Não exite"
                        FG_MP.TextMatrix(I, 7) = Format((CDbl(.BDSIS_TBEST_CPEST.Value) - CDbl(.BDSIS_TBEST_CPVEN.Value)), "###,##0.00")
                    ElseIf Left(FG_MP.TextMatrix(I, 0), 2) = "MP" Then 'se for MATERIA-PRIMA
                        sTipoPeca = "MP"
                        FG_MP.TextMatrix(I, 6) = "Não exite"
                        FG_MP.TextMatrix(I, 7) = "Não exite"
                        FG_MP.TextMatrix(I, 8) = Format((CDbl(.BDSIS_TBEST_CPEST.Value) - CDbl(.BDSIS_TBEST_CPVEN.Value)), "###,##0.00")
                    End If
                Else
                    SeekErro = True
                End If
            Else
                SeekErro = True
            End If
            If SeekErro = True Then
                FG_MP.TextMatrix(I, 6) = "-"
                FG_MP.TextMatrix(I, 7) = "-"
                FG_MP.TextMatrix(I, 8) = "-"
                GoTo PROXIMO_I
            End If
            SeekErro = False
            'procura se for componentes
            If sTipoPeca = "COM" Then
                'componente tem so um item por peca
                ProcuraMP .BDSIS_TBEST_CPINQ.Value, .BDSIS_TBEST_CPINP.Value, .BDSIS_TBEST_CPINN.Value, .BDSIS_TBEST_CPINB.Value, .BDSIS_TBEST_CPINM.Value
                If MATPRI.PEC <> "" And MATPRI.BIT <> "" And MATPRI.MAT <> "" Then
                    .BDSIS_TBEST.Seek "=", MATPRI.PEC, MATPRI.BIT, MATPRI.MAT
                    If Not .BDSIS_TBEST.NoMatch Then
                        FG_MP.TextMatrix(I, 7) = Format((CDbl(.BDSIS_TBEST_CPEST.Value) - CDbl(.BDSIS_TBEST_CPVEN.Value)), "###,##0.00")
                    Else
                        SeekErro = True
                    End If
                    If SeekErro = True Then
                        FG_MP.TextMatrix(I, 7) = "-"
                        SeekErro = False
                    End If
                    'matéria-prima tem so um item por peca
                    ProcuraMP .BDSIS_TBEST_CPINQ.Value, .BDSIS_TBEST_CPINP.Value, .BDSIS_TBEST_CPINN.Value, .BDSIS_TBEST_CPINB.Value, .BDSIS_TBEST_CPINM.Value
                    If MATPRI.PEC <> "" And MATPRI.BIT <> "" And MATPRI.MAT <> "" Then
                        .BDSIS_TBEST.Seek "=", MATPRI.PEC, MATPRI.BIT, MATPRI.MAT
                        If Not .BDSIS_TBEST.NoMatch Then
                            FG_MP.TextMatrix(I, 8) = Format((CDbl(.BDSIS_TBEST_CPEST.Value) - CDbl(.BDSIS_TBEST_CPVEN.Value)), "###,##0.00")
                        Else
                            SeekErro = True
                        End If
                    Else
                        SeekErro = True
                    End If
                    If SeekErro = True Then FG_MP.TextMatrix(I, 8) = "-"
                Else
                    SeekErro = True
                End If
            End If
            'procura por matéria-prima
            If sTipoPeca = "PA" Then
                'matéria-prima tem so um item por peca
                ProcuraMP .BDSIS_TBEST_CPINQ.Value, .BDSIS_TBEST_CPINP.Value, .BDSIS_TBEST_CPINN.Value, .BDSIS_TBEST_CPINB.Value, .BDSIS_TBEST_CPINM.Value
                If MATPRI.PEC <> "" And MATPRI.BIT <> "" And MATPRI.MAT <> "" Then
                    .BDSIS_TBEST.Seek "=", MATPRI.PEC, MATPRI.BIT, MATPRI.MAT
                    If Not .BDSIS_TBEST.NoMatch Then
                        FG_MP.TextMatrix(I, 8) = Format((CDbl(.BDSIS_TBEST_CPEST.Value) - CDbl(.BDSIS_TBEST_CPVEN.Value)), "###,##0.00")
                    Else
                        SeekErro = True
                    End If
                Else
                    SeekErro = True
                End If
                If SeekErro = True Then FG_MP.TextMatrix(I, 8) = "-"
            End If
PROXIMO_I:
        Next I
    End With
End Sub
