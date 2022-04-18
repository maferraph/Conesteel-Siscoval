Attribute VB_Name = "Modulo_Cotacao"
Option Explicit

Global NUMCOT As String

Public Sub HabilitaFrameImportacao()
    NUMCOT = ""
    With Tela_Cotacao
        'habilita
        .BT_Novo.Enabled = False
        .BT_Editar.Enabled = False
        .BT_Deletar.Enabled = False
        .BT_Imprimir.Enabled = False
        .BT_Pedido.Enabled = False
        .BT_Cotacao.Enabled = False
        .BT_Apagar.Enabled = False
        .BT_Cancelar.Enabled = False
        .BT_Voltar.Enabled = False
        .BT_CancelaImportacao.Enabled = True
        .BT_Importa.Enabled = True
        'Top
        .BT_CancelaImportacao.Top = .BT_Novo.Top
        .BT_Importa.Top = .BT_Novo.Top
        'exibe
        .BT_Novo.Visible = False
        .BT_Editar.Visible = False
        .BT_Deletar.Visible = False
        .BT_Imprimir.Visible = False
        .BT_Pedido.Visible = False
        .BT_Cotacao.Visible = False
        .BT_Apagar.Visible = False
        .BT_Cancelar.Visible = False
        .BT_Voltar.Visible = False
        .BT_CancelaImportacao.Visible = True
        .BT_Importa.Visible = True
    End With
End Sub


