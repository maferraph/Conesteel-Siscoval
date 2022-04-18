VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classe_Nfemit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function Siscoval(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then
        Siscoval = False
    Else
        Siscoval = True
        Tela_NFEmitidas.Show vbModal
    End If
End Function

