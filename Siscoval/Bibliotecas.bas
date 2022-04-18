Attribute VB_Name = "Bibliotecas"
Option Explicit
' ****************** DECLARAÇÕES ****************
Dim DLL_BD As Scvbd.Classe_Scvbd
'Biblioteca Scvcarr não aplicada
Dim DLL_FUNCS As Scvfunc.Classe_Scvfunc
'Biblioteca Assfig não aplicada
Dim DLL_CFGGRU As Cfggru.Classe_Cfggru
Dim DLL_CFGUSU As Cfgusu.Classe_Cfgusu
Dim DLL_CADEMP As Cademp.Classe_Cademp
Dim DLL_CFGALI As Cfgali.Classe_Cfgali
Dim DLL_CFGBCO As Cfgbco.Classe_Cfgbco
Dim DLL_CFGCFI As Cfgcfi.Classe_Cfgcfi
Dim DLL_CFGECF As Cfgecf.Classe_Cfgecf
Dim DLL_CFGEDF As Cfgedc.Classe_Cfgedc
Dim DLL_CFGEPP As Cfgepp.Classe_Cfgepp
Dim DLL_CFGEST As Cfgest.Classe_Cfgest
Dim DLL_CFGINF As Cfginf.Classe_Cfginf
Dim DLL_ASSNF As Assnf.Classe_Assnf
Dim DLL_ESTCRP As Estcrp.Classe_Estcrp
Dim DLL_ESTCON As Estcon.Classe_Estcon
Dim DLL_COTEST As Cotest.Classe_Cotest
Dim DLL_NFEMIT As Nfemit.Classe_Nfemit
Dim DLL_CONUNI As Convunit.Classe_Convunit
Dim DLL_CFGMPE As Cfgmpe.Classe_Cfgmpe
Dim DLL_PEDEST As Pedest.Classe_Pedest
Dim DLL_ESTBAL As Estbal.Classe_Estbal
Dim DLL_CFGMAQ As Cfgfamaq.Classe_Cfgfamaq
Dim DLL_CFGIMG As Cfgimg.Classe_Cfgimg

Public Static Sub ChamaDLL(DLL As String)
    'Essa função irá carregar a DLL chamada,
    'irá executá-la e depois encerrá-la.
    If DLL = "Cfgali" Then 'Configurações de Alíquotas
        Set DLL_CFGALI = New Cfgali.Classe_Cfgali
        If DLL_CFGALI.Siscoval(App.ProductName, "Cfgali", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgali")
        Set DLL_CFGALI = Nothing
    ElseIf DLL = "Cfgbco" Then 'Configurações de bancos
        Set DLL_CFGBCO = New Cfgbco.Classe_Cfgbco
        If DLL_CFGBCO.Siscoval(App.ProductName, "Cfgbco", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgbco")
        Set DLL_CFGBCO = Nothing
    ElseIf DLL = "Cfgcfi" Then 'Configuração de códigos fiscais
        Set DLL_CFGCFI = New Cfgcfi.Classe_Cfgcfi
        If DLL_CFGCFI.Siscoval(App.ProductName, "Cfgcfi", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgcfi")
        Set DLL_CFGCFI = Nothing
    ElseIf DLL = "Cfgest" Then 'Configuração de estoque
        Set DLL_CFGEST = New Cfgest.Classe_Cfgest
        If DLL_CFGEST.Siscoval(App.ProductName, "Cfgest", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgest")
        Set DLL_CFGEST = Nothing
    ElseIf DLL = "Cfgecf" Then 'Configuração de estoque CF E ST
        Set DLL_CFGECF = New Cfgecf.Classe_Cfgecf
        If DLL_CFGECF.Siscoval(App.ProductName, "Cfgecf", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgecf")
        Set DLL_CFGECF = Nothing
    ElseIf DLL = "Cfgedc" Then 'Configuração declarações fiscais
        Set DLL_CFGEDF = New Cfgedc.Classe_Cfgedc
        If DLL_CFGEDF.Siscoval(App.ProductName, "Cfgedc", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgedc")
        Set DLL_CFGEDF = Nothing
    ElseIf DLL = "Cfgepp" Then 'Configuração Preço e Peso Estoque
        Set DLL_CFGEPP = New Cfgepp.Classe_Cfgepp
        If DLL_CFGEPP.Siscoval(App.ProductName, "Cfgepp", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgepp")
        Set DLL_CFGEPP = Nothing
    ElseIf DLL = "Cfggru" Then  'Configuração de Grupos
        Set DLL_CFGGRU = New Cfggru.Classe_Cfggru
        If DLL_CFGGRU.Siscoval(App.ProductName, "Cfggru", App.LegalCopyright) = False Then ErroLogonDLL ("Cfggru")
        Set DLL_CFGGRU = Nothing
    ElseIf DLL = "Cfgusu" Then
        Set DLL_CFGUSU = New Cfgusu.Classe_Cfgusu
        If Aux_Cfgusu = True Then
            If DLL_CFGUSU.Siscoval(App.ProductName, "Cfgusu", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgusu")
            Set DLL_CFGUSU = Nothing
        End If
    ElseIf DLL = "Cfginf" Then 'Configurações Nota Fiscal
        Set DLL_CFGINF = New Cfginf.Classe_Cfginf
        If DLL_CFGINF.Siscoval(App.ProductName, "Cfginf", App.LegalCopyright) = False Then ErroLogonDLL ("Cfginf")
        Set DLL_CFGINF = Nothing
    ElseIf DLL = "Cademp" Then 'Cadastro de Empresas
        Set DLL_CADEMP = New Cademp.Classe_Cademp
        If DLL_CADEMP.Siscoval(App.ProductName, "Cademp", App.LegalCopyright) = False Then ErroLogonDLL ("Cademp")
        Set DLL_CADEMP = Nothing
    ElseIf DLL = "Assnf" Then 'Assistente Nota Fiscal
        Set DLL_ASSNF = New Assnf.Classe_Assnf
        If DLL_ASSNF.Siscoval(App.ProductName, "Assnf", App.LegalCopyright) = False Then ErroLogonDLL ("Assnf")
        Set DLL_ASSNF = Nothing
    ElseIf DLL = "Estcrp" Then 'Estoque - Consulta Rápida
        Set DLL_ESTCRP = New Estcrp.Classe_Estcrp
        If DLL_ESTCRP.Siscoval(App.ProductName, "Estcrp", App.LegalCopyright) = False Then ErroLogonDLL ("Estcrp")
        Set DLL_ESTCRP = Nothing
    ElseIf DLL = "Estcon" Then 'Estoque - Fichas de Estoque
        Set DLL_ESTCON = New Estcon.Classe_Estcon
        If DLL_ESTCON.Siscoval(App.ProductName, "Estcon", App.LegalCopyright) = False Then ErroLogonDLL ("Estcon")
        Set DLL_ESTCON = Nothing
    ElseIf DLL = "Cotest" Then 'Cotação de Estoque
        Set DLL_COTEST = New Cotest.Classe_Cotest
        If DLL_COTEST.Siscoval(App.ProductName, "Cotest", App.LegalCopyright) = False Then ErroLogonDLL ("Cotest")
        Set DLL_COTEST = Nothing
    ElseIf DLL = "Nfemit" Then 'Notas Fiscais Emitidas
        Set DLL_NFEMIT = New Nfemit.Classe_Nfemit
        If DLL_NFEMIT.Siscoval(App.ProductName, "Nfemit", App.LegalCopyright) = False Then ErroLogonDLL ("Nfemit")
        Set DLL_NFEMIT = Nothing
    ElseIf DLL = "Convunit" Then 'Conversor de Unidades
        Set DLL_CONUNI = New Convunit.Classe_Convunit
        If DLL_CONUNI.Siscoval(App.ProductName, "Convunit", App.LegalCopyright) = False Then ErroLogonDLL ("Convunit")
        Set DLL_CONUNI = Nothing
    ElseIf DLL = "Cfgmpe" Then 'Configuração Matéria-Prima
        Set DLL_CFGMPE = New Cfgmpe.Classe_Cfgmpe
        If DLL_CFGMPE.Siscoval(App.ProductName, "Cfgmpe", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgmpe")
        Set DLL_CFGMPE = Nothing
    ElseIf DLL = "Pedest" Then 'Pedido
        Set DLL_PEDEST = New Pedest.Classe_Pedest
        If DLL_PEDEST.Siscoval(App.ProductName, "Pedest", App.LegalCopyright) = False Then ErroLogonDLL ("Pedest")
        Set DLL_PEDEST = Nothing
    ElseIf DLL = "Cfgfamaq" Then 'Configuração de Máquinas
        Set DLL_CFGMAQ = New Cfgfamaq.Classe_Cfgfamaq
        If DLL_CFGMAQ.Siscoval(App.ProductName, "Cfgfamaq", App.LegalCopyright, Usuario) = False Then ErroLogonDLL ("Cfgfamaq")
        Set DLL_CFGMAQ = Nothing
    ElseIf DLL = "Cfgimg" Then 'Configuração de Imagens
        Set DLL_CFGIMG = New Cfgimg.Classe_Cfgimg
        If DLL_CFGIMG.Siscoval(App.ProductName, "Cfgimg", App.LegalCopyright) = False Then ErroLogonDLL ("Cfgimg")
        Set DLL_CFGIMG = Nothing
    End If
    Screen.MousePointer = vbNormal
End Sub
Private Static Sub ErroLogonDLL(Biblioteca As String)
    RespMsg = MsgBox("Não foi possível acessar a biblioteca " & VBA.Trim(Biblioteca) & ". Tente mais tarde.", vbCritical + vbOKOnly, "Erro de Acesso")
End Sub
Private Static Function Aux_Cfgusu() As Boolean 'Esta e uma funcao da dll Cfgusu
    Screen.MousePointer = vbHourglass
    'Abre Classes DLL's
    Set DLL_BD = New Scvbd.Classe_Scvbd
    'Abre bancos de dados
    If DLL_BD.AbreBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo Tabelas
    If DLL_BD.AbreTabela_UsuariosMenus(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Abrindo campos
    If DLL_BD.AbreCampos_UsuariosMenus(App.ProductName, "Scvbd", App.LegalCopyright) = False Then GoTo ERRO_ACESSO_BANCODADOS
    'Verifica os menus
    Dim MenuTP As Control
    For Each MenuTP In Tela_Siscoval
        If TypeOf MenuTP Is Menu Then
            DLL_BD.BDSIS_TBUME.Seek "=", MenuTP.Name
            If DLL_BD.BDSIS_TBUME.NoMatch Then
                DLL_BD.BDSIS_TBUME.AddNew
                DLL_BD.BDSIS_TBUME_CPMEN.Value = MenuTP.Name
                DLL_BD.BDSIS_TBUME.Update
            End If
        End If
    Next MenuTP
    'Fecha tabelas
    If DLL_BD.FechaTabela_UsuariosMenus(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha banco de dados
    If DLL_BD.FechaBD(App.ProductName, "Scvbd", App.LegalCopyright) = False Then Beep
    'Fecha classes de DLL's
    Set DLL_BD = Nothing
    Aux_Cfgusu = True
    Screen.MousePointer = vbNormal
    Exit Function
ERRO_ACESSO_BANCODADOS:
    RespMsg = MsgBox("Ocorreu algum erro durante o acesso ao banco de dados.", vbCritical + vbOKOnly, "Erro de abertura")
    Aux_Cfgusu = False
End Function
