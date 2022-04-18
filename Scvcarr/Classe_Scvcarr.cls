VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classe_Scvcarr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub CarregaTexto(Texto As String)
    Tela_Carga.Caption = Texto
    Tela_Carga.BP.Value = Tela_Carga.BP.Value + 1
End Sub
Public Sub Max(Valor As Integer)
    Tela_Carga.BP.Max = Valor
    ResetaBP
End Sub
Public Sub ResetaBP()
    Tela_Carga.BP.Value = 0
    Tela_Carga.Caption = ""
End Sub
Public Sub Exibe(Exibir As Boolean)
    If Exibir = True Then
        ResetaBP
        Tela_Carga.Visible = True
    Else
        Tela_Carga.Visible = False
    End If
End Sub
