VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Classe_Assfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function AssistenteFigura(X, Y, Z) As String
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then
        AssistenteFigura = ""
    Else
        Tela_AssistenteFigura.Show vbModal
        AssistenteFigura = FIG
    End If
End Function
