VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classe_Impform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim I As Integer, NumLB As Integer
Private DLL_FUNCS As Scvfunc.Classe_Scvfunc
Private DLL_BD As Scvbd.Classe_Scvbd
Dim MeuControle As Control


'**************************************************************
'                         FUNÇÕES
'**************************************************************
Private Sub SelecionaImpressora(Optional Nome As String, Optional FolhaDeitada As Integer)
    'If Nome = "Epson LX-300" Then Printer.PaperSize = vbPRPSLetter
    For I = 0 To Printers.Count - 1
        If Nome <> "" Then
            If Printers(I).DeviceName = Nome Then
                Set Printer = Printers(I)
                Exit For
            End If
        Else
            If Printers(I).DeviceName = "PADRÃO" Then
                Set Printer = Printers(I)
                Exit For
            End If
        End If
    Next I
    If FolhaDeitada = 1 Then
        Printer.Orientation = vbPRORLandscape
    End If
End Sub
Private Function DataExtenso(DataDia As Date) As String
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DataExtenso = "São Paulo, " & Trim(Day(DataDia)) & " de " & Trim(DLL_FUNCS.NomeMes(Month(DataDia))) & " de " & Trim(Year(DataDia))
    Set DLL_FUNCS = Nothing
End Function
Private Function ContaControles(Tela As Form) As Integer
    ContaControles = Tela.Controls.Count
End Function
Private Static Sub ImprimeTela(Tela As Form)
    On Error Resume Next
    For Each MeuControle In Tela.Controls
        If TypeOf MeuControle Is Label Then
            If MeuControle.Caption <> "" Then
                Printer.Font = MeuControle.Font
                Printer.FontSize = MeuControle.FontSize
                Printer.FontBold = MeuControle.FontBold
                Printer.FontItalic = MeuControle.FontItalic
                Printer.ForeColor = MeuControle.FontUnderline
                Printer.CurrentX = MeuControle.Left
                Printer.CurrentY = MeuControle.Top
                Printer.Print (MeuControle.Caption)
            End If
        ElseIf TypeOf MeuControle Is TextBox Then
            If MeuControle.Text <> "" Then
                Printer.Font = MeuControle.Font
                Printer.FontSize = MeuControle.FontSize
                Printer.FontBold = MeuControle.FontBold
                Printer.FontItalic = MeuControle.FontItalic
                Printer.ForeColor = MeuControle.FontUnderline
                Printer.CurrentX = MeuControle.Left
                Printer.CurrentY = MeuControle.Top
                Printer.Print (MeuControle.Text)
            End If
        ElseIf TypeOf MeuControle Is Line Then
            Printer.FillColor = MeuControle.FillColor
            Printer.FillStyle = MeuControle.FillStyle
            Printer.FillStyle = MeuControle.FillStyle
            Printer.CurrentX = MeuControle.Left
            Printer.CurrentY = MeuControle.Top
            Printer.Line (MeuControle.X1, MeuControle.Y1)-(MeuControle.X2, MeuControle.Y2)
        ElseIf TypeOf MeuControle Is PictureBox Then
            Printer.CurrentX = MeuControle.Left
            Printer.CurrentY = MeuControle.Top
            Printer.PaintPicture MeuControle, MeuControle.Left, MeuControle.Top, MeuControle.Width, MeuControle.Height
        End If
    Next MeuControle
    Printer.EndDoc
End Sub


'**************************************************************
'                        NOTA FISCAL
'**************************************************************
Public Function NotaFiscal_LimpaItens() As Integer
    NotaFiscal_LimpaItens = 0
    With IT_NotaFiscal
        For I = 0 To .Controls.Count - 1
            .Controls(I).Visible = False
            .Controls(I).Caption = ""
        Next I
    End With
    NumLB = 0
    NotaFiscal_LimpaItens = ContaControles(IT_NotaFiscal)
End Function
Public Function NotaFiscal_MontaNF(TipoBD As String, Texto As String, Optional NumI As Integer) As Boolean
    NotaFiscal_MontaNF = False
    'abre bd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    If DLL_BD.AbreCampos_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Monta Item
    If TipoBD = "COMENTARIO" Or TipoBD = "DECLARACAO" Or TipoBD = "BANCO" Then
        DLL_BD.BDSIS_TBCNF.Seek "=", TipoBD
        If DLL_BD.BDSIS_TBCNF.NoMatch Then
            GoTo ERRO_ACESSO_BANCODADOS
        Else
            NumLB = NumLB + 1
            IT_NotaFiscal.LB_NF(NumLB).Visible = True
            IT_NotaFiscal.LB_NF(NumLB).Caption = Texto
            IT_NotaFiscal.LB_NF(NumLB).FontName = DLL_BD.BDSIS_TBCNF_CPFON.Value
            IT_NotaFiscal.LB_NF(NumLB).FontSize = DLL_BD.BDSIS_TBCNF_CPTAM.Value
            IT_NotaFiscal.LB_NF(NumLB).FontBold = DLL_BD.BDSIS_TBCNF_CPNEG.Value
            IT_NotaFiscal.LB_NF(NumLB).FontItalic = DLL_BD.BDSIS_TBCNF_CPITA.Value
            IT_NotaFiscal.LB_NF(NumLB).FontUnderline = DLL_BD.BDSIS_TBCNF_CPSUB.Value
        End If
        Dim sTipo As String
        sTipo = ""
        If Val(NumI) < 10 And Val(NumI) > 0 Then
            sTipo = "FIGURA_0" & CStr(NumI)
        ElseIf Val(NumI) > 9 Then
            sTipo = "FIGURA_" & CStr(NumI)
        End If
        If sTipo = "" Then Exit Function
        DLL_BD.BDSIS_TBCNF.Seek "=", sTipo
        If DLL_BD.BDSIS_TBCNF.NoMatch Then
            GoTo ERRO_ACESSO_BANCODADOS
        Else
            IT_NotaFiscal.LB_NF(NumLB).Left = DLL_BD.BDSIS_TBCNF_CPCOL.Value
            IT_NotaFiscal.LB_NF(NumLB).Top = DLL_BD.BDSIS_TBCNF_CPLIN.Value
        End If
    Else
        DLL_BD.BDSIS_TBCNF.Seek "=", TipoBD
        If DLL_BD.BDSIS_TBCNF.NoMatch Then
            GoTo ERRO_ACESSO_BANCODADOS
        Else
            NumLB = NumLB + 1
            IT_NotaFiscal.LB_NF(NumLB).Visible = True
            If TipoBD = "VALOR_EXTENSO_1" Or TipoBD = "VALOR_EXTENSO_2" Then
                IT_NotaFiscal.LB_NF(NumLB).Caption = Left(Texto, 120)
            Else
                IT_NotaFiscal.LB_NF(NumLB).Caption = Texto
            End If
            IT_NotaFiscal.LB_NF(NumLB).Left = DLL_BD.BDSIS_TBCNF_CPCOL.Value
            IT_NotaFiscal.LB_NF(NumLB).Top = DLL_BD.BDSIS_TBCNF_CPLIN.Value
            IT_NotaFiscal.LB_NF(NumLB).FontName = DLL_BD.BDSIS_TBCNF_CPFON.Value
            IT_NotaFiscal.LB_NF(NumLB).FontSize = DLL_BD.BDSIS_TBCNF_CPTAM.Value
            IT_NotaFiscal.LB_NF(NumLB).FontBold = DLL_BD.BDSIS_TBCNF_CPNEG.Value
            IT_NotaFiscal.LB_NF(NumLB).FontItalic = DLL_BD.BDSIS_TBCNF_CPITA.Value
            IT_NotaFiscal.LB_NF(NumLB).FontUnderline = DLL_BD.BDSIS_TBCNF_CPSUB.Value
        End If
    End If
    'fecha bd
    If DLL_BD.FechaTabela_ConfiguracoesNotaFiscal(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    Set DLL_BD = Nothing
    NotaFiscal_MontaNF = True
    Exit Function
ERRO_ACESSO_BANCODADOS:
    MsgBox "Ocorreu algum erro quando o sistema estava procurando informações no banco de dados para montar a nota fiscal. Erro: " & TipoBD, vbCritical + vbOKOnly, "Erro"
    Set DLL_BD = Nothing
End Function
Public Function NotaFiscal_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    NotaFiscal_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Nota Fiscal", ""
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_NotaFiscal
    SelecionaImpressora 'Seleciona impressora padrão
    NotaFiscal_Imprimir = True
End Function


'**************************************************************
'                  ORDEM DE FABRICAÇÃO
'**************************************************************
Public Function OrdemFabricacao_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    OrdemFabricacao_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Ordem de Fabricação", IT_OrdemFabricacao.LB_Num.Caption
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_OrdemFabricacao
    SelecionaImpressora 'Seleciona impressora padrão
    Set DLL_FUNCS = Nothing
    OrdemFabricacao_Imprimir = True
End Function
Public Function OrdemFabricacao_LimpaItens() As Boolean
    OrdemFabricacao_LimpaItens = False
    With IT_OrdemFabricacao
        .LB_Num.Caption = ""
        .LB_OFA.Visible = False
        .LB_Data.Caption = ""
        .LB_QuantEst.Caption = ""
        .LB_Fig.Caption = ""
        .LB_Bit.Caption = ""
        .LB_Mat.Caption = ""
        .LB_QuantPro.Caption = ""
        .LB_Des.Caption = ""
        .LB_DC.Caption = ""
        .LB_Corrida.Caption = ""
        For I = 0 To 9
            .LB_E(I).Caption = ""
            .LB_M(I).Caption = ""
            .LB_DI(I).Caption = "____/____/____"
            .LB_HI(I).Caption = "_____:_____"
            .LB_DF(I).Caption = "____/____/____"
            .LB_HF(I).Caption = "_____:_____"
            .LB_HE(I).Caption = "_______________"
        Next I
        OrdemFabricacao_LimpaItens = True
    End With
End Function
Public Function OrdemFabricacao_Cabecalho(Num As String, Dat As String, Qes As String, Fig As String, Bit As String, Mat As String, Qpr As String, Des As String, Ndc As String, Cor As String, lAberta As Boolean) As Boolean
    OrdemFabricacao_Cabecalho = False
    With IT_OrdemFabricacao
        .LB_Num.Caption = Num
        If lAberta = True Then .LB_OFA.Visible = True
        .LB_Data.Caption = Format(Dat, "dd/mm/yyyy")
        .LB_QuantEst.Caption = Qes
        .LB_Fig.Caption = Fig
        .LB_Bit.Caption = Bit
        .LB_Mat.Caption = Mat
        .LB_QuantPro.Caption = Qpr
        .LB_Des.Caption = Des
        .LB_DC.Caption = Ndc
        .LB_Corrida.Caption = Cor
    End With
    OrdemFabricacao_Cabecalho = True
End Function
Public Function OrdemFabricacao_Itens(Indice As Integer, Net As String, Nma As String, Dai As String, Hoi As String, Daf As String, Hof As String, Qhe As String) As Boolean
    OrdemFabricacao_Itens = False
    With IT_OrdemFabricacao
        .LB_E(Indice).Caption = Net
        .LB_M(Indice).Caption = Nma
        .LB_DI(Indice).Caption = Dai
        .LB_HI(Indice).Caption = Hoi
        .LB_DF(Indice).Caption = Daf
        .LB_HF(Indice).Caption = Hof
        .LB_HE(Indice).Caption = Qhe
    End With
    OrdemFabricacao_Itens = True
End Function


'**************************************************************
'                        ORDEM DE MONTAGEM
'**************************************************************
Public Function OrdemMontagem_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    OrdemMontagem_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Ordem de Montagem", IT_OrdemMontagem.LB_Num.Caption
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_OrdemMontagem
    SelecionaImpressora 'Seleciona impressora padrão
    Set DLL_FUNCS = Nothing
    OrdemMontagem_Imprimir = True
End Function
Public Function OrdemMontagem_LimpaItens() As Boolean
    OrdemMontagem_LimpaItens = False
    With IT_OrdemMontagem
        .LB_Num.Caption = ""
        .LB_OMA.Visible = False
        .LB_Data.Caption = ""
        .LB_QuantEst.Caption = ""
        .LB_Figura.Caption = ""
        .LB_Bitola.Caption = ""
        .LB_Material.Caption = ""
        .LB_DC.Caption = ""
        .LB_QuantMon.Caption = ""
        .LB_Descricao.Caption = ""
        .LB_DI.Caption = "____/____/____"
        .LB_HI.Caption = "_____:_____"
        .LB_DF.Caption = "____/____/____"
        .LB_HF.Caption = "_____:_____"
        For I = 0 To 18
            .LB_Ite(I).Caption = ""
            .LB_Qua(I).Caption = ""
            .LB_Des(I).Caption = ""
            .LB_Pec(I).Caption = ""
            .LB_Bit(I).Caption = ""
            .LB_Mat(I).Caption = ""
            .LB_Ece(I).Caption = ""
            .LB_Epe(I).Caption = ""
            .LB_Eme(I).Caption = ""
            .LB_CA(I).Caption = ""
            .LB_CB(I).Caption = ""
            .LB_CC(I).Caption = ""
            .LB_CD(I).Caption = ""
            .LB_CE(I).Caption = ""
        Next I
    End With
    OrdemMontagem_LimpaItens = True
End Function
Public Function OrdemMontagem_Cabecalho(Num As String, Dat As String, Qes As String, Fig As String, Bit As String, Mat As String, Ndc As String, Qmo As String, Des As String, Din As String, Hin As String, Dfi As String, Hfi As String, lAberta As Boolean) As Boolean
    OrdemMontagem_Cabecalho = False
    With IT_OrdemMontagem
        .LB_Num.Caption = Num
        If lAberta = True Then
            .LB_OMA.Visible = True
        Else
            .LB_OMA.Visible = False
        End If
        .LB_Data.Caption = Format(Dat, "dd/mm/yyyy")
        .LB_QuantEst.Caption = Qes
        .LB_Figura.Caption = Fig
        .LB_Bitola.Caption = Bit
        .LB_Material.Caption = Mat
        .LB_DC.Caption = Ndc
        .LB_QuantMon.Caption = Qmo
        .LB_Descricao.Caption = Des
        If Din = "" Then
            .LB_DI.Caption = "____/____/____"
        Else
            .LB_DI.Caption = Din
        End If
        If Hin = "" Then
            .LB_HI.Caption = "_____:_____"
        Else
            .LB_HI.Caption = Hin
        End If
        If Dfi = "" Then
            .LB_DF.Caption = "____/____/____"
        Else
            .LB_DF.Caption = Dfi
        End If
        If Hfi = "" Then
            .LB_HF.Caption = "_____:_____"
        Else
            .LB_HF.Caption = Hfi
        End If
    End With
    OrdemMontagem_Cabecalho = True
End Function
Public Function OrdemMontagem_Itens(Indice As Integer, Nit As String, Qua As String, Des As String, Pec As String, Bit As String, Mat As String, Eco As String, Epa As String, Emp As String, NCA As String, NCB As String, NCC As String, NCD As String, NCE As String) As Boolean
    OrdemMontagem_Itens = False
    With IT_OrdemMontagem
        .LB_Ite(Indice).Caption = Nit
        .LB_Qua(Indice).Caption = Qua
        .LB_Des(Indice).Caption = Des
        .LB_Pec(Indice).Caption = Pec
        .LB_Bit(Indice).Caption = Bit
        .LB_Mat(Indice).Caption = Mat
        .LB_Ece(Indice).Caption = Eco
        .LB_Epe(Indice).Caption = Epa
        .LB_Eme(Indice).Caption = Emp
        .LB_CA(Indice).Caption = NCA
        .LB_CB(Indice).Caption = NCB
        .LB_CC(Indice).Caption = NCC
        .LB_CD(Indice).Caption = NCD
        .LB_CE(Indice).Caption = NCE
    End With
    OrdemMontagem_Itens = True
End Function


'**************************************************************
'                       PEDIDO DE ESTOQUE
'**************************************************************
Public Function PedidoEstoque_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    PedidoEstoque_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Pedido de Estoque", IT_PedidoEstoque.LB_Num.Caption
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_PedidoEstoque
    'IT_PedidoEstoque.PrintForm
    'Printer.EndDoc
    SelecionaImpressora 'Seleciona impressora padrão
    Set DLL_FUNCS = Nothing
    PedidoEstoque_Imprimir = True
End Function
Public Function PedidoEstoque_Cabecalho(Num As String, Data As Date, Empresa As String, CNPJ As String, IE As String, Endereco As String, Bairro As String, CEP As String, Cidade As String, Estado As String, Fone As String, Fax As String, Depto As String, Contato As String, Ramal As String, NumSeuPed As String) As Boolean
    PedidoEstoque_Cabecalho = False
    With IT_PedidoEstoque
        .LB_Num.Caption = Num
        .LB_Data.Caption = DataExtenso(Data)
        .LB_Empresa.Caption = Empresa
        .LB_CNPJ.Caption = CNPJ
        .LB_IE.Caption = IE
        .LB_Endereco.Caption = Endereco
        .LB_Bairro.Caption = Bairro
        .LB_CEP.Caption = CEP
        .LB_Cidade.Caption = Cidade
        .LB_Estado.Caption = Estado
        .LB_Fone.Caption = Fone
        .LB_Fax.Caption = Fax
        .LB_Depto.Caption = Depto
        .LB_Contato.Caption = Contato
        .LB_Ramal.Caption = Ramal
        .LB_NSP.Caption = NumSeuPed
    End With
    PedidoEstoque_Cabecalho = True
End Function
Public Function PedidoEstoque_Rodape(ValPed As String, CondPagto As String, Trans As String, Frete As String, OutDes As String, ValIpi As String, ValTotPed As String, NomCon As String, NomEmp As String, NomVen As String, NomDep As String, Cond As String) As Boolean
    PedidoEstoque_Rodape = False
    With IT_PedidoEstoque
        .LB_ValPed.Caption = ValPed
        .LB_CondPagto.Caption = CondPagto
        .LB_Trans.Caption = Trans
        .LB_Frete.Caption = Frete
        .LB_OutDes.Caption = OutDes
        .LB_ValIpi.Caption = ValIpi
        .LB_ValTotPed.Caption = ValTotPed
        .LB_NomCon.Caption = NomCon
        .LB_NomEmp.Caption = NomEmp
        .LB_NomVen.Caption = NomVen
        .LB_NomDep.Caption = NomDep
        .TXT_Cond.Text = Cond
    End With
    PedidoEstoque_Rodape = True
End Function
Public Function PedidoEstoque_LimpaItens() As Boolean
    PedidoEstoque_LimpaItens = False
    With IT_PedidoEstoque
        .LB_Num.Caption = ""
        .LB_Data.Caption = ""
        .LB_Empresa.Caption = ""
        .LB_CNPJ.Caption = ""
        .LB_IE.Caption = ""
        .LB_Endereco.Caption = ""
        .LB_Bairro.Caption = ""
        .LB_CEP.Caption = ""
        .LB_Cidade.Caption = ""
        .LB_Estado.Caption = ""
        .LB_Fone.Caption = ""
        .LB_Fax.Caption = ""
        .LB_Depto.Caption = ""
        .LB_Contato.Caption = ""
        .LB_Ramal.Caption = ""
        .LB_NSP.Caption = ""
        For I = 0 To 23
            .LB_Item(I).Caption = ""
            .LB_Quant(I).Caption = ""
            .LB_Figura(I).Caption = ""
            .LB_Descricao(I).Caption = ""
            .LB_Preco(I).Caption = ""
            .LB_ICMS(I).Caption = ""
            .LB_IPI(I).Caption = ""
            .LB_Prazo(I).Caption = ""
        Next I
        .LB_ValPed.Caption = ""
        .LB_CondPagto.Caption = ""
        .LB_Trans.Caption = ""
        .LB_Frete.Caption = ""
        .LB_OutDes.Caption = ""
        .LB_ValIpi.Caption = ""
        .LB_ValTotPed.Caption = ""
        .LB_NomCon.Caption = ""
        .LB_NomEmp.Caption = ""
        .LB_NomVen.Caption = ""
        .LB_NomDep.Caption = ""
        .TXT_Cond.Text = ""
    End With
    PedidoEstoque_LimpaItens = True
End Function
Public Function PedidoEstoque_Itens(Indice As Integer, Item As String, Quant As String, Figura As String, Descricao As String, Preco As String, ICMS As String, IPI As String, Prazo As String) As Boolean
    PedidoEstoque_Itens = False
    With IT_PedidoEstoque
        .LB_Item(Indice).Caption = Item
        .LB_Quant(Indice).Caption = Quant
        .LB_Figura(Indice).Caption = Figura
        .LB_Descricao(Indice).Caption = Descricao
        .LB_Preco(Indice).Caption = Preco
        .LB_ICMS(Indice).Caption = ICMS
        .LB_IPI(Indice).Caption = IPI
        .LB_Prazo(Indice).Caption = Prazo
    End With
    PedidoEstoque_Itens = True
End Function


'**************************************************************
'                        ORDEM DE EXPEDIÇÃO
'**************************************************************
Public Function OrdemExpedicao_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    OrdemExpedicao_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Ordem de Expedição", IT_OrdemExpedicao.LB_Num.Caption
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_OrdemExpedicao
    SelecionaImpressora 'Seleciona impressora padrão
    OrdemExpedicao_Imprimir = True
End Function
Public Function OrdemExpedicao_LimpaItens() As Boolean
    OrdemExpedicao_LimpaItens = False
    With IT_OrdemExpedicao
        .LB_Num.Caption = ""
        .LB_Data.Caption = ""
        .LB_Empresa.Caption = ""
        .LB_CNPJ.Caption = ""
        .LB_Cidade.Caption = ""
        .LB_Estado.Caption = ""
        For I = 0 To 23
            .LB_Ite(I).Caption = ""
            .LB_Qua(I).Caption = ""
            .LB_Fig(I).Caption = ""
            .LB_Des(I).Caption = ""
            .LB_Con(I).Caption = ""
        Next I
        .LB_Obs.Caption = ""
        .LB_TRA.Caption = ""
        .LB_TIP.Caption = ""
        .LB_PE.Caption = ""
        .LB_PL.Caption = ""
    End With
    OrdemExpedicao_LimpaItens = True
End Function
Public Function OrdemExpedicao_Cabecalho(Num As String, Data As Date, Empresa As String, CNPJ As String, Cidade As String, Estado As String) As Boolean
    OrdemExpedicao_Cabecalho = False
    With IT_OrdemExpedicao
        .LB_Num.Caption = Num
        .LB_Data.Caption = Format(Data, "dd/mm/yyyy")
        .LB_Empresa.Caption = Empresa
        .LB_CNPJ.Caption = CNPJ
        .LB_Cidade.Caption = Cidade
        .LB_Estado.Caption = Estado
    End With
    OrdemExpedicao_Cabecalho = True
End Function
Public Function OrdemExpedicao_Rodape(OBS As String, TRA As String, TIP As String, PE As String, PL As String) As Boolean
    OrdemExpedicao_Rodape = False
    With IT_OrdemExpedicao
        .LB_Obs.Caption = OBS
        .LB_TRA.Caption = TRA
        .LB_TIP.Caption = TIP
        .LB_PE.Caption = PE
        .LB_PL.Caption = PL
    End With
    OrdemExpedicao_Rodape = True
End Function
Public Function OrdemExpedicao_Itens(Indice As Integer, Ite As String, Qua As String, Fig As String, Des As String, Con As String) As Boolean
    OrdemExpedicao_Itens = False
    With IT_OrdemExpedicao
        .LB_Ite(Indice).Caption = Ite
        .LB_Qua(Indice).Caption = Qua
        .LB_Fig(Indice).Caption = Fig
        .LB_Des(Indice).Caption = Des
        .LB_Con(Indice).Caption = Con
    End With
    OrdemExpedicao_Itens = True
End Function


'**************************************************************
'                            ROMANEIO
'**************************************************************
Public Function Romaneio_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    Romaneio_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Romaneio", IT_Romaneio.LB_Num.Caption
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_Romaneio
    SelecionaImpressora 'Seleciona impressora padrão
    Set DLL_FUNCS = Nothing
    Romaneio_Imprimir = True
End Function
Public Function Romaneio_Cabecalho(Num As String, Empresa As String, DataDia As Date, NumPed As String) As Boolean
    Romaneio_Cabecalho = False
    With IT_Romaneio
        .LB_Num.Caption = Num
        .LB_Empresa.Caption = Empresa
        .LB_Data.Caption = Format(DataDia, "dd/mm/yyyy")
        .LB_NumPed.Caption = NumPed
    End With
    Romaneio_Cabecalho = True
End Function
Public Function Romaneio_LimpaItens() As Boolean
    Romaneio_LimpaItens = False
    With IT_Romaneio
        .LB_Num.Caption = ""
        .LB_Empresa.Caption = ""
        .LB_Data.Caption = ""
        .LB_NumPed.Caption = ""
        For I = 0 To 23
            .LB_Ite(I).Caption = ""
            .LB_Qua(I).Caption = ""
            .LB_Pec(I).Caption = ""
            .LB_Descricao(I).Caption = ""
            .LB_Corrida(I).Caption = ""
        Next I
        .LB_Obs.Caption = ""
        .LB_Separado.Caption = ""
        .LB_Conferido.Caption = ""
        .LB_MarcP.Caption = ""
        .LB_MarcS.Caption = ""
        .LB_MarcO.Caption = ""
        .LB_Outros.Caption = "___________"
        .LB_Marca.Caption = ""
        .LB_Numeros.Caption = ""
        .LB_PesoBru.Caption = ""
        .LB_PesoLiq.Caption = ""
    End With
    Romaneio_LimpaItens = True
End Function
Public Function Romaneio_Itens(Indice As Integer, Ite As String, Qua As String, Pec As String, Descricao As String, Corrida As String) As Boolean
    Romaneio_Itens = False
    With IT_Romaneio
        .LB_Ite(Indice).Caption = Ite
        .LB_Qua(Indice).Caption = Qua
        .LB_Pec(Indice).Caption = Pec
        .LB_Descricao(Indice).Caption = Descricao
        .LB_Corrida(Indice).Caption = Corrida
    End With
    Romaneio_Itens = True
End Function
Public Function Romaneio_Rodape(OBS As String, Separado As String, Conferido As String, MarcP As String, MarcS As String, MarcO As String, Outros As String, Marca As String, Numeros As String, PesoBru As String, PesoLiq As String) As Boolean
    Romaneio_Rodape = False
    With IT_Romaneio
        .LB_Obs.Caption = OBS
        .LB_Separado.Caption = Separado
        .LB_Conferido.Caption = Conferido
        .LB_MarcP.Caption = MarcP
        .LB_MarcS.Caption = MarcS
        .LB_MarcO.Caption = MarcO
        .LB_Outros.Caption = Outros
        .LB_Marca.Caption = Marca
        .LB_Numeros.Caption = Numeros
        .LB_PesoBru.Caption = PesoBru
        .LB_PesoLiq.Caption = PesoLiq
    End With
    Romaneio_Rodape = True
End Function


'**************************************************************
'                       COTAÇÃO DE ESTOQUE
'**************************************************************
Public Function CotacaoEstoque_LimpaItens() As Boolean
    CotacaoEstoque_LimpaItens = False
    With IT_CotacaoEstoque
        .LB_Num.Caption = ""
        .LB_Data.Caption = ""
        .LB_Empresa.Caption = ""
        .LB_CNPJ.Caption = ""
        .LB_IE.Caption = ""
        .LB_Endereco.Caption = ""
        .LB_Bairro.Caption = ""
        .LB_CEP.Caption = ""
        .LB_Cidade.Caption = ""
        .LB_Estado.Caption = ""
        .LB_Fone.Caption = ""
        .LB_Fax.Caption = ""
        .LB_Depto.Caption = ""
        .LB_Contato.Caption = ""
        .LB_Ramal.Caption = ""
        For I = 0 To 27
            .LB_Item(I).Caption = ""
            .LB_Figura(I).Caption = ""
            .LB_Quant(I).Caption = ""
            .LB_Descricao(I).Caption = ""
            .LB_Preco(I).Caption = ""
            .LB_ICMS(I).Caption = ""
            .LB_IPI(I).Caption = ""
            .LB_Prazo(I).Caption = ""
        Next I
        .LB_CondPagto.Caption = ""
        .LB_Trans.Caption = ""
        .LB_Frete.Caption = ""
        .TXT_Cond.Text = ""
        .LB_Nome.Caption = ""
        .LB_Vend.Caption = ""
    End With
    CotacaoEstoque_LimpaItens = True
End Function
Public Function CotacaoEstoque_Cabecalho(Num As String, Data As Date, Empresa As String, CNPJ As String, IE As String, Endereco As String, Bairro As String, CEP As String, Cidade As String, Estado As String, Fone As String, Fax As String, Depto As String, Contato As String, Ramal As String) As Boolean
    CotacaoEstoque_Cabecalho = False
    With IT_CotacaoEstoque
        .LB_Num.Caption = Num
        .LB_Data.Caption = DataExtenso(Data)
        .LB_Empresa.Caption = Empresa
        .LB_CNPJ.Caption = CNPJ
        .LB_IE.Caption = IE
        .LB_Endereco.Caption = Endereco
        .LB_Bairro.Caption = Bairro
        .LB_CEP.Caption = CEP
        .LB_Cidade.Caption = Cidade
        .LB_Estado.Caption = Estado
        .LB_Fone.Caption = Fone
        .LB_Fax.Caption = Fax
        .LB_Depto.Caption = Depto
        .LB_Contato.Caption = Contato
        .LB_Ramal.Caption = Ramal
    End With
    CotacaoEstoque_Cabecalho = True
End Function
Public Function CotacaoEstoque_Itens(Indice As Integer, Item As String, Quant As String, Figura As String, Descricao As String, Preco As String, ICMS As String, IPI As String, Prazo As String) As Boolean
    CotacaoEstoque_Itens = False
    With IT_CotacaoEstoque
        .LB_Item(Indice).Caption = Item
        .LB_Quant(Indice).Caption = Quant
        .LB_Figura(Indice).Caption = Figura
        .LB_Descricao(Indice).Caption = Descricao
        .LB_Preco(Indice).Caption = Preco
        .LB_ICMS(Indice).Caption = ICMS
        .LB_IPI(Indice).Caption = IPI
        .LB_Prazo(Indice).Caption = Prazo
    End With
    CotacaoEstoque_Itens = True
End Function
Public Function CotacaoEstoque_Rodape(CondPagto As String, Trans As String, Frete As String, Nome As String, Vend As String, Cond As String) As Boolean
    CotacaoEstoque_Rodape = False
    With IT_CotacaoEstoque
        .LB_CondPagto.Caption = CondPagto
        .LB_Trans.Caption = Trans
        .LB_Frete.Caption = Frete
        .LB_Nome.Caption = Nome
        .LB_Vend.Caption = Vend
        .TXT_Cond.Text = Cond
    End With
    CotacaoEstoque_Rodape = True
End Function
Public Function CotacaoEstoque_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    CotacaoEstoque_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Cotação de Estoque", IT_CotacaoEstoque.LB_Num.Caption
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
'    ImprimeTela IT_CotacaoEstoque
     IT_CotacaoEstoque.PrintForm
    SelecionaImpressora 'Seleciona impressora padrão
    Set DLL_FUNCS = Nothing
    CotacaoEstoque_Imprimir = True
End Function


'**************************************************************
'                CONFIGURAÇÃO DE MATÉRIA-PRIMA
'**************************************************************
Public Function ConfigMP_LimpaItens() As Boolean
    ConfigMP_LimpaItens = False
    With IT_ConfigMP
        .LB_Data.Caption = ""
        .LB_Quan.Caption = ""
        .LB_Figu.Caption = ""
        .LB_Bito.Caption = ""
        .LB_Mate.Caption = ""
        .LB_Desc.Caption = ""
        For I = 0 To 43
            .LB_Ite(I).Caption = ""
            .LB_Pec(I).Caption = ""
            .LB_Des(I).Caption = ""
            .LB_Bit(I).Caption = ""
            .LB_Mat(I).Caption = ""
            .LB_Qun(I).Caption = ""
            .LB_Qne(I).Caption = ""
            .LB_Eco(I).Caption = ""
            .LB_Epr(I).Caption = ""
            .LB_Ema(I).Caption = ""
        Next I
    End With
    ConfigMP_LimpaItens = True
End Function
Public Function ConfigMP_Cabecalho(Data As Date, Quan As String, Figu As String, Bito As String, Mate As String, Desc As String) As Boolean
    ConfigMP_Cabecalho = False
    With IT_ConfigMP
        .LB_Data.Caption = Format(Data, "dd/mm/yyyy")
        .LB_Quan.Caption = Quan
        .LB_Figu.Caption = Figu
        .LB_Bito.Caption = Bito
        .LB_Mate.Caption = Mate
        .LB_Desc.Caption = Desc
    End With
    ConfigMP_Cabecalho = True
End Function
Public Function ConfigMP_Itens(Indice As Integer, Ite As String, Pec As String, Des As String, Bit As String, Mat As String, Qun As String, Qne As String, Eco As String, Epr As String, Ema As String) As Boolean
    ConfigMP_Itens = False
    With IT_ConfigMP
        .LB_Ite(Indice).Caption = Ite
        .LB_Pec(Indice).Caption = Pec
        .LB_Des(Indice).Caption = Des
        .LB_Bit(Indice).Caption = Bit
        .LB_Mat(Indice).Caption = Mat
        .LB_Qun(Indice).Caption = Qun
        .LB_Qne(Indice).Caption = Qne
        .LB_Eco(Indice).Caption = Eco
        .LB_Epr(Indice).Caption = Epr
        .LB_Ema(Indice).Caption = Ema
    End With
    ConfigMP_Itens = True
End Function
Public Function ConfigMP_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    ConfigMP_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Configuração de MP", ""
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_ConfigMP
    SelecionaImpressora 'Seleciona impressora padrão
    Set DLL_FUNCS = Nothing
    ConfigMP_Imprimir = True
End Function


'**************************************************************
'                       RELATORIO DE CONTAS
'**************************************************************
Public Function RelatorioContas_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    RelatorioContas_Imprimir = False
    SelecionaImpressora NomeImpressora 'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Relatorio Contas", " "
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_Contas_Relatorio
    SelecionaImpressora 'Seleciona impressora padrão
    Set DLL_FUNCS = Nothing
    RelatorioContas_Imprimir = True
End Function


Public Function RelatorioContas_LimpaItens() As Boolean
    RelatorioContas_LimpaItens = False
    With IT_Contas_Relatorio
        
        RelatorioContas_LimpaItens = True
    End With
End Function
Public Function RelatorioContas_Cabecalho(Num As String, Dat As String, Qes As String, Fig As String, Bit As String, Mat As String, Qpr As String, Des As String, Ndc As String, Cor As String, lAberta As Boolean) As Boolean
    RelatorioContas_Cabecalho = False
    With IT_Contas_Relatorio
    
    End With
    RelatorioContas_Cabecalho = True
End Function
Public Function RelatorioContas_Itens(Indice As Integer, Net As String, Nma As String, Dai As String, Hoi As String, Daf As String, Hof As String, Qhe As String) As Boolean
    RelatorioContas_Itens = False
    With IT_Contas_Relatorio
    
    End With
    RelatorioContas_Itens = True
End Function




'**************************************************************
'                       RELATORIO DE EQUIPAMENTOS
'**************************************************************
Public Function Equipamentos_LimpaItens() As Boolean
    Equipamentos_LimpaItens = False
    For Each MeuControle In IT_Equipamentos.Controls
        If (TypeOf MeuControle Is Label) And (MeuControle.Name <> "LB_Fixo") Then
            MeuControle.Caption = ""
        End If
    Next MeuControle
    Equipamentos_LimpaItens = True
End Function
Public Function Equipamentos_Cabecalho(Data As Date, NE As String, Nome As String, Tipo As String, Modelo As String, Fabricante As String, Valor As String, Foto As String, Custo As String, Processo As String, Funcionario As String) As Boolean
    Equipamentos_Cabecalho = False
    With IT_Equipamentos
        .LB_Data.Caption = Format(Data, "dd/mm/yyyy")
        .LB_NE.Caption = NE
        .LB_Nome.Caption = Nome
        .LB_Tipo.Caption = Tipo
        .LB_Modelo.Caption = Modelo
        .LB_Fabricante.Caption = Fabricante
        .LB_Valor.Caption = Format(Valor, "###,###,###,##0.00")
        .LB_Foto.Caption = Foto
        .LB_Custo.Caption = Format(Custo, "###,###,###,##0.00")
        .LB_Processo.Caption = Processo
        .LB_Funcionario.Caption = Funcionario
    End With
    Equipamentos_Cabecalho = True
End Function
Public Function Equipamentos_Itens(Item As Integer, ItemMan As String, DataUltima As String, DataProxima As String, ValorMan As String, EmpMan As String, ConMan As String, TipMan As String, Aviso As String, Usuario As String) As Boolean
    Equipamentos_Itens = False
    With IT_Equipamentos
        .LB_Item(Item).Caption = ItemMan
        .LB_DataUltima(Item).Caption = Format(DataUltima, "dd/mm/yyyy")
        .LB_DataProxima(Item).Caption = Format(DataProxima, "dd/mm/yyyy")
        .LB_ValorMan(Item).Caption = Format(ValorMan, "###,###,###,##0.00")
        .LB_EmpMan(Item).Caption = EmpMan
        .LB_ConMan(Item).Caption = ConMan
        .LB_TipMan(Item).Caption = TipMan
        .LB_Aviso(Item).Caption = Aviso
        .LB_Usuario(Item).Caption = Usuario
    End With
    Equipamentos_Itens = True
End Function
Public Function Equipamentos_Imprimir(NomeImpressora As String, Optional NumVias As Integer) As Boolean
    Equipamentos_Imprimir = False
    SelecionaImpressora NomeImpressora, 1  'Seleciona impressora deste formulario
    Set DLL_FUNCS = New Scvfunc.Classe_Scvfunc
    DLL_FUNCS.RegistraEvento "Imprimir Configurações de Equipamentos", ""
    'Imprimi o formulário
    If NumVias > 1 Then Printer.Copies = NumVias
    ImprimeTela IT_Equipamentos
    SelecionaImpressora 'Seleciona impressora padrão
    Equipamentos_Imprimir = True
End Function
