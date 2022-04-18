VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classe_Scvbd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Esta classe sera usada para iniciar todas tabelas
'do banco de dados do Siscoval.
Dim Dir_EXE, Dir_BD As String

'*****************************************************
'                   W O R K S P A C E
'*****************************************************
Public WKSP As Workspace 'Abre uma área de trabalho

'*****************************************************
'    B A N C O   D E   D A D O S   S I S C O V A L
'*****************************************************
Public BDSIS As Database 'Siscoval

' ******************   TABELAS    ******************
Public BDSIS_TBEMP As Recordset 'Empresas
Public BDSIS_TBCDF As Recordset 'Códigos Fiscais
Public BDSIS_TBNTF As Recordset 'Nota Fiscal
Public BDSIS_TBNFP As Recordset 'Nota Fiscal - Produtos
Public BDSIS_TBNFD As Recordset 'Nota Fiscal - Declarações
Public BDSIS_TBNFB As Recordset 'Nota Fiscal - Bancos
Public BDSIS_TBNFC As Recordset 'Nota Fiscal - Comentários
Public BDSIS_TBGRU As Recordset 'Grupos
Public BDSIS_TBEST As Recordset 'Estoque
Public BDSIS_TBEID As Recordset 'Estoque - Índice
Public BDSIS_TBEFG As Recordset 'Estoque - Figuras
Public BDSIS_TBECM As Recordset 'Estoque - Complementos
Public BDSIS_TBBAN As Recordset 'Bancos
Public BDSIS_TBCFS As Recordset 'Estoque - CF e ST
Public BDSIS_TBEAL As Recordset 'Estoque - Alíquotas
Public BDSIS_TBEDC As Recordset 'Estoque - Declarações
Public BDSIS_TBCTG As Recordset 'Configurações Tela-Grupo
Public BDSIS_TBMFA As Recordset 'Mapa - Faturamento do Mês
Public BDSIS_TBMCR As Recordset 'Mapa - Contas à Receber
Public BDSIS_TBMIM As Recordset 'Mapa - Impostos
Public BDSIS_TBMCO As Recordset 'Mapa - Cotações
Public BDSIS_TBCNF As Recordset 'Configurações da Nota Fiscal
Public BDSIS_TBUSU As Recordset 'Usuários
Public BDSIS_TBUCO As Recordset 'Usuários - Computadores
Public BDSIS_TBUME As Recordset 'Usuários - Menus
Public BDSIS_TBINF As Recordset 'Configurações da Nota Fiscal - Imagem
Public BDSIS_TBCOT As Recordset 'Cotações
Public BDSIS_TBCTI As Recordset 'Cotações - Ítens
Public BDSIS_TBECO As Recordset 'Empresas - Contatos
Public BDSIS_TBTIM As Recordset 'Configurações de Telas para Impressoras
Public BDSIS_TBIMP As Recordset 'Impressoras
Public BDSIS_TBGEV As Recordset 'Log de Eventos
Public BDSIS_TBGRR As Recordset 'Log de Erros
Public BDSIS_TBMPP As Recordset 'Matéria-Prima - Índice de Peças
Public BDSIS_TBMPB As Recordset 'Matéria-Prima - Índice de Bitolas
Public BDSIS_TBMPM As Recordset 'Matéria-Prima - Índice de Materiais
Public BDSIS_TBMPQ As Recordset 'Matéria-Prima - Índice de Quantidades
Public BDSIS_TBMPN As Recordset 'Matéria-Prima - Índice de Nomes
Public BDSIS_TBMPR As Recordset 'Matéria-Prima - Relação de Materiais
Public BDSIS_TBEPE As Recordset 'Entrada de Peças
Public BDSIS_TBEPI As Recordset 'Entrada de Peças - Índice
Public BDSIS_TBPED As Recordset 'Pedidos
Public BDSIS_TBPIT As Recordset 'Pedidos - Ítens
Public BDSIS_TBMPE As Recordset 'Mapa - Pedidos
Public BDSIS_TBODF As Recordset 'Ordem de Fabricação
Public BDSIS_TBOFP As Recordset 'Ordem de Fabricação - Processos
Public BDSIS_TBODM As Recordset 'Ordem de Montagem
Public BDSIS_TBODE As Recordset 'Ordem de Expedição
Public BDSIS_TBFUN As Recordset 'Funcionários
Public BDSIS_TBCPR As Recordset 'Contas à Pagar e Receber
Public BDSIS_TBIMA As Recordset 'Imagens
Public BDSIS_TBEQU As Recordset 'Equipamentos
Public BDSIS_TBEMA As Recordset 'Equipamentos - Manutenção
Public BDSIS_TBAVI As Recordset 'Avisos
Public BDSIS_TBCFU As Recordset 'CNC - Funções
Public BDSIS_TBCMA As Recordset 'CNC - Máquinas
Public BDSIS_TBCNC As Recordset 'CNC - Programas
Public BDSIS_TBLAY As Recordset 'Layouts
Public BDSIS_TBLES As Recordset 'Layouts - Esquemas
Public BDSIS_TBLDI As Recordset 'Layouts - Dimensões


' ******************   CAMPOS    *******************
' Campos da Tabela Empresas
Public BDSIS_TBEMP_CPAPE As Field
Public BDSIS_TBEMP_CPEMP As Field
Public BDSIS_TBEMP_CPINE As Field
Public BDSIS_TBEMP_CPCGC As Field
Public BDSIS_TBEMP_CPEND As Field
Public BDSIS_TBEMP_CPBAI As Field
Public BDSIS_TBEMP_CPCID As Field
Public BDSIS_TBEMP_CPEST As Field
Public BDSIS_TBEMP_CPCEP As Field
Public BDSIS_TBEMP_CPFON As Field
Public BDSIS_TBEMP_CPFAX As Field
Public BDSIS_TBEMP_CPPRA As Field
Public BDSIS_TBEMP_CPTIP As Field
Public BDSIS_TBEMP_CPCOM As Field
Public BDSIS_TBEMP_CPOBS As Field
Public BDSIS_TBEMP_CPTRA As Field 'Transportadora
Public BDSIS_TBEMP_CPEMA As Field 'E-Mail
Public BDSIS_TBEMP_CPSIT As Field 'Site
'Campos da Tabela Códigos Fiscais
Public BDSIS_TBCDF_CPNRG As Field 'Número do Registro
Public BDSIS_TBCDF_CPNTO As Field 'Natureza da Operação
Public BDSIS_TBCDF_CPCFO As Field 'CFOP
Public BDSIS_TBCDF_CPTIP As Field 'Tipo
'Campos da Tabela Nota Fiscal
Public BDSIS_TBNTF_CPNNF As Field 'Número da Nota Fiscal
Public BDSIS_TBNTF_CPDEM As Field 'Data Emissao
Public BDSIS_TBNTF_CPHEM As Field 'Hora Emissao
Public BDSIS_TBNTF_CPTIP As Field 'Tipo
Public BDSIS_TBNTF_CPNOP As Field 'Natureza Operação
Public BDSIS_TBNTF_CPCFO As Field 'CFOP
Public BDSIS_TBNTF_CPEMP As Field 'Empresa
Public BDSIS_TBNTF_CPDSA As Field 'Data Saída
Public BDSIS_TBNTF_CPHSA As Field 'Hora Saída
Public BDSIS_TBNTF_CPVAL As Field 'Valor
Public BDSIS_TBNTF_CPNDP As Field 'Número Duplicata
Public BDSIS_TBNTF_CPVEN As Field 'Vencimento
Public BDSIS_TBNTF_CPNPI As Field 'Número Pedido Interno
Public BDSIS_TBNTF_CPNSP As Field 'Número Seu Pedido
Public BDSIS_TBNTF_CPOPE As Field 'Operação
Public BDSIS_TBNTF_CPDEC As Field 'Declarações
Public BDSIS_TBNTF_CPBAN As Field 'Bancos
Public BDSIS_TBNTF_CPCOM As Field 'Comentários
Public BDSIS_TBNTF_CPBCI As Field 'Base Cálculo ICMS
Public BDSIS_TBNTF_CPVIC As Field 'Valor ICMS
Public BDSIS_TBNTF_CPBIS As Field 'Base Cálculo ICMS Substituição
Public BDSIS_TBNTF_CPVIS As Field 'Valor ICMS Substituição
Public BDSIS_TBNTF_CPVTP As Field 'Valor Total Produtos
Public BDSIS_TBNTF_CPVFR As Field 'Valor Frete
Public BDSIS_TBNTF_CPVSE As Field 'Valor Seguro
Public BDSIS_TBNTF_CPODA As Field 'Outras Despesas Acessórias
Public BDSIS_TBNTF_CPVIP As Field 'Valor Total IPI
Public BDSIS_TBNTF_CPVTN As Field 'Valor Total Nota
Public BDSIS_TBNTF_CPTRA As Field 'Transportadora
Public BDSIS_TBNTF_CPFCO As Field 'Frete por Conta
Public BDSIS_TBNTF_CPPVE As Field 'Placa Veículo
Public BDSIS_TBNTF_CPEPV As Field 'Estado Placa Veículo
Public BDSIS_TBNTF_CPQUA As Field 'Quantidade
Public BDSIS_TBNTF_CPESP As Field 'Espécie
Public BDSIS_TBNTF_CPMAR As Field 'Marca
Public BDSIS_TBNTF_CPNVO As Field 'Número
Public BDSIS_TBNTF_CPPBR As Field 'Peso Bruto
Public BDSIS_TBNTF_CPPLI As Field 'Peso Líquido
Public BDSIS_TBNTF_CPVIN As Field 'Vendedor Interno
Public BDSIS_TBNTF_CPVEX As Field 'Vendedor Externo
Public BDSIS_TBNTF_CPSET As Field 'Setor
Public BDSIS_TBNTF_CPCFI As Field 'CF I
Public BDSIS_TBNTF_CPCFJ As Field 'CF J
Public BDSIS_TBNTF_CPVLA As Field 'Valor A
Public BDSIS_TBNTF_CPVCA As Field 'Vencimento A
Public BDSIS_TBNTF_CPVLB As Field 'Valor B
Public BDSIS_TBNTF_CPVCB As Field 'Vencimento B
Public BDSIS_TBNTF_CPVLC As Field 'Valor C
Public BDSIS_TBNTF_CPVCC As Field 'Vencimento C
Public BDSIS_TBNTF_CPVLD As Field 'Valor D
Public BDSIS_TBNTF_CPVCD As Field 'Vencimento D
'Campos da Tabela Nota Fiscal - Produtos
Public BDSIS_TBNFP_CPNNF As Field 'Número da Nota Fiscal
Public BDSIS_TBNFP_CPFIG As Field 'Figura
Public BDSIS_TBNFP_CPBIT As Field 'Bitola
Public BDSIS_TBNFP_CPMAT As Field 'Material
Public BDSIS_TBNFP_CPDES As Field 'Descrição
Public BDSIS_TBNFP_CPCCF As Field 'CF
Public BDSIS_TBNFP_CPCST As Field 'ST
Public BDSIS_TBNFP_CPUNI As Field 'Unidade
Public BDSIS_TBNFP_CPQUA As Field 'Quantidade
Public BDSIS_TBNFP_CPPUN As Field 'Preço Unitário
Public BDSIS_TBNFP_CPPTO As Field 'Preço Total
Public BDSIS_TBNFP_CPBCI As Field 'Base Cálculo ICMS
Public BDSIS_TBNFP_CPAIC As Field 'Alíquota ICMS
Public BDSIS_TBNFP_CPVIC As Field 'Valor ICMS
Public BDSIS_TBNFP_CPAIP As Field 'Alíquota IPI
Public BDSIS_TBNFP_CPVIP As Field 'Valor IPI
Public BDSIS_TBNFP_CPPPA As Field 'Peso Parcial
'Campos da Tabela Nota Fiscal - Declarações
Public BDSIS_TBNFD_CPNNF As Field 'Número da Nota Fiscal
Public BDSIS_TBNFD_CPDEC As Field 'Declaração
Public BDSIS_TBNFD_CPLIN As Field 'NumLin
'Campos da Tabela Nota Fiscal - Bancos
Public BDSIS_TBNFB_CPNNF As Field 'Número da Nota Fiscal
Public BDSIS_TBNFB_CPCON As Field 'Nome Conta
Public BDSIS_TBNFB_CPLIN As Field 'NumLin
'Campos da Tabela Nota Fiscal - Comentários
Public BDSIS_TBNFC_CPNNF As Field 'Número da Nota Fiscal
Public BDSIS_TBNFC_CPCOM As Field 'Comentários
Public BDSIS_TBNFC_CPLIN As Field 'NumLin
'Campos da Tabela de Estoque - Grupos
Public BDSIS_TBGRU_CPGRU As Field 'Grupo
Public BDSIS_TBGRU_CPDES As Field 'Descricao
Public BDSIS_TBGRU_CPVAL As Field 'Valor
Public BDSIS_TBGRU_CPTIP As Field 'Tipo
' Campos da Tabela de Estoque
Public BDSIS_TBEST_CPFIC As Field 'Índice de Ficha
Public BDSIS_TBEST_CPFIG As Field 'Figura
Public BDSIS_TBEST_CPBIT As Field 'Bitola
Public BDSIS_TBEST_CPMAT As Field 'Material
Public BDSIS_TBEST_CPVUN As Field 'Valor Unitário
Public BDSIS_TBEST_CPVMI As Field 'Valor Mínimo
Public BDSIS_TBEST_CPVCU As Field 'Valor Custo
Public BDSIS_TBEST_CPPUN As Field 'Peso Unitário
Public BDSIS_TBEST_CPEST As Field 'Estoque
Public BDSIS_TBEST_CPETM As Field 'Estoque Mínimo
Public BDSIS_TBEST_CPVEN As Field 'Vendidas
Public BDSIS_TBEST_CPCOT As Field 'Cotadas
Public BDSIS_TBEST_CPINQ As Field 'Índice MP de Quantidades
Public BDSIS_TBEST_CPINP As Field 'Índice MP de Peças
Public BDSIS_TBEST_CPINN As Field 'Índice MP de Nomes
Public BDSIS_TBEST_CPINB As Field 'Índice MP de Bitolas
Public BDSIS_TBEST_CPINM As Field 'Índice MP de Materiais
Public BDSIS_TBEST_CPNDC As Field 'DC
'Campos da Tabela Estoque - Índice
Public BDSIS_TBEID_CPIFI As Field 'Índice Figura
Public BDSIS_TBEID_CPDRE As Field 'DesRed
Public BDSIS_TBEID_CPTRE As Field 'TipRed
Public BDSIS_TBEID_CPDNO As Field 'DesNor
Public BDSIS_TBEID_CPTNO As Field 'TipNor
Public BDSIS_TBEID_CPDCO As Field 'DesCom
Public BDSIS_TBEID_CPTCO As Field 'DesCom
Public BDSIS_TBEID_CPGPE As Field 'Grupo Peça
Public BDSIS_TBEID_CPGMP As Field 'Grupo Matéria-Prima
Public BDSIS_TBEID_CPGBI As Field 'Grupo Bitola
Public BDSIS_TBEID_CPGIN As Field 'Grupo Internos
Public BDSIS_TBEID_CPGEX As Field 'Grupo Extremidades
Public BDSIS_TBEID_CPGCL As Field 'Grupo Classes
Public BDSIS_TBEID_CPGMA As Field 'Grupo Materiais
'Campos da Tabela Estoque - Figuras
Public BDSIS_TBEFG_CPFIG As Field 'Figura
Public BDSIS_TBEFG_CPIFG As Field 'Índice Figura
Public BDSIS_TBEFG_CPGIN As Field 'Grupo Internos
Public BDSIS_TBEFG_CPGEX As Field 'Grupo Extremidade
Public BDSIS_TBEFG_CPGCL As Field 'Grupo Classe
Public BDSIS_TBEFG_CPCOM As Field 'Grupo Complemento
'Campos da Tabela Bancos
Public BDSIS_TBBAN_CPNMC As Field 'Nome da Conta
Public BDSIS_TBBAN_CPNMB As Field 'Nome do Banco
Public BDSIS_TBBAN_CPBAI As Field 'Bairro
Public BDSIS_TBBAN_CPAGE As Field 'Agência
Public BDSIS_TBBAN_CPCON As Field 'Conta Corrente
'Campos da Tabela de Estoque - CF e ST
Public BDSIS_TBCFS_CPFIG As Field 'Figura
Public BDSIS_TBCFS_CPGMT As Field 'Grupo Material
Public BDSIS_TBCFS_CPGCF As Field 'Grupo CF
Public BDSIS_TBCFS_CPGST As Field 'Grupo ST
'Campos da Tabela de Estoque - Alíquotas
Public BDSIS_TBEAL_CPALI As Field 'Alíquota
Public BDSIS_TBEAL_CPEST As Field 'Estado
Public BDSIS_TBEAL_CPCCF As Field 'CF
Public BDSIS_TBEAL_CPPOR As Field 'Porcentagem
Public BDSIS_TBEAL_CPRBC As Field 'RBC
'Campos da Tabela de Estoque - Declarações
Public BDSIS_TBEDC_CPCCF As Field 'CF
Public BDSIS_TBEDC_CPDSP As Field 'DecSP
Public BDSIS_TBEDC_CPDOU As Field 'DecOutros
'Campos da Tabela de Estoque - Declarações
Public BDSIS_TBCTG_CPTEL As Field 'Tela
Public BDSIS_TBCTG_CPCOM As Field 'Combo
Public BDSIS_TBCTG_CPVAL As Field 'Valor
Public BDSIS_TBCTG_CPINI As Field 'Valor
'Campos da Tabela de Mapa - Faturamento do Mês
Public BDSIS_TBMFA_CPNNF As Field 'Número da Nota Fiscal
Public BDSIS_TBMFA_CPDAT As Field 'Data
Public BDSIS_TBMFA_CPVAL As Field 'Valor
Public BDSIS_TBMFA_CPEMP As Field 'Empresa
Public BDSIS_TBMFA_CPDEV As Field 'Devolução
'Campos da Tabela de Mapa - Contas à Receber
Public BDSIS_TBMCR_CPDEM As Field 'Data Emissão
Public BDSIS_TBMCR_CPEMP As Field 'Empresa
Public BDSIS_TBMCR_CPNDP As Field 'Número Duplicata
Public BDSIS_TBMCR_CPVAL As Field 'Valor
Public BDSIS_TBMCR_CPDVE As Field 'Data Vencimento
'Campos da Tabela de Mapa - Impostos
Public BDSIS_TBMIM_CPIMP As Field 'Imposto
Public BDSIS_TBMIM_CPMES As Field 'Mês
Public BDSIS_TBMIM_CPANO As Field 'Ano
Public BDSIS_TBMIM_CPVAL As Field 'Valor
'Campos da Tabela Configurações da Nota Fiscal
Public BDSIS_TBCNF_CPITE As Field 'Item
Public BDSIS_TBCNF_CPDES As Field 'Descrição
Public BDSIS_TBCNF_CPEXE As Field 'Exemplo
Public BDSIS_TBCNF_CPIND As Field 'Índice
Public BDSIS_TBCNF_CPCOL As Field 'Col
Public BDSIS_TBCNF_CPLIN As Field 'Lin
Public BDSIS_TBCNF_CPFON As Field 'Fonte
Public BDSIS_TBCNF_CPTAM As Field 'Tam
Public BDSIS_TBCNF_CPNEG As Field 'Negrito
Public BDSIS_TBCNF_CPITA As Field 'Italico
Public BDSIS_TBCNF_CPSUB As Field 'Sublinhado
'Campos da Tabela Usuários
Public BDSIS_TBUSU_CPUSU As Field 'Usuário
Public BDSIS_TBUSU_CPADM As Field 'Administrador
Public BDSIS_TBUSU_CPNOM As Field 'Nome
Public BDSIS_TBUSU_CPSEN As Field 'Senha
Public BDSIS_TBUSU_CPPER As Field 'Permissões
'Campos da Tabela Usuários - Computadores
Public BDSIS_TBUCO_CPCOM As Field 'Computador
Public BDSIS_TBUCO_CPDWI As Field 'DirWin
Public BDSIS_TBUCO_CPDSY As Field 'DirSys
Public BDSIS_TBUCO_CPDTP As Field 'DirTem
Public BDSIS_TBUCO_CPOBU As Field 'OSBuild
Public BDSIS_TBUCO_CPOCS As Field 'OSCSD
Public BDSIS_TBUCO_CPOVE As Field 'OSVersao
Public BDSIS_TBUCO_CPOPL As Field 'OSPlat
Public BDSIS_TBUCO_CPSMP As Field 'SisMascProc
Public BDSIS_TBUCO_CPSOE As Field 'SisOEM
Public BDSIS_TBUCO_CPSTP As Field 'SisTipProc
Public BDSIS_TBUCO_CPRAM As Field 'RAM
Public BDSIS_TBUCO_CPCTA As Field 'TamC
Public BDSIS_TBUCO_CPCSN As Field 'SNC
'Campos da Tabela Usuários - Menus
'Public BDSIS_TBUME_CPIND As Field 'Índice
Public BDSIS_TBUME_CPMEN As Field 'Menu
'Public BDSIS_TBUME_CPPRI As Field 'Primeiro
'Campos da Tabela Configurações da Nota Fiscal - Imagem
Public BDSIS_TBINF_CPIMG As Field 'Imagem
'Campos da Tabela Cotações
Public BDSIS_TBCOT_CPIND As Field 'Índice
Public BDSIS_TBCOT_CPDAT As Field 'Data Emissão
Public BDSIS_TBCOT_CPHOR As Field 'Hora
Public BDSIS_TBCOT_CPEMP As Field 'Empresa
Public BDSIS_TBCOT_CPCON As Field 'Contato
Public BDSIS_TBCOT_CPCPG As Field 'Condição de Pagamento
Public BDSIS_TBCOT_CPTRA As Field 'Transportadora
Public BDSIS_TBCOT_CPITE As Field 'Ítens
Public BDSIS_TBCOT_CPVAL As Field 'Valor
Public BDSIS_TBCOT_CPIVE As Field 'Índice Vendedor
Public BDSIS_TBCOT_CPLIQ As Field 'Liquidado
'Campos da Tabela Cotações - Ítens
Public BDSIS_TBCTI_CPIND As Field 'Índice dos Ítens
Public BDSIS_TBCTI_CPNCO As Field 'Número Cotação
Public BDSIS_TBCTI_CPINF As Field 'Índice de Ficha
Public BDSIS_TBCTI_CPQUA As Field 'Quantidade
Public BDSIS_TBCTI_CPDES As Field 'Descrição
Public BDSIS_TBCTI_CPCOM As Field 'Complemento
Public BDSIS_TBCTI_CPPRE As Field 'Preço Unitário
Public BDSIS_TBCTI_CPPRA As Field 'Prazo de Entrega
Public BDSIS_TBCTI_CPLIQ As Field 'Liquidado
'Campos da Tabela Empresas - Contatos
Public BDSIS_TBECO_CPEMP As Field 'Empresa
Public BDSIS_TBECO_CPCON As Field 'Contato
'Campos da Tabela Configurações de Telas para Impressoras
Public BDSIS_TBTIM_CPTEL As Field 'Tela
Public BDSIS_TBTIM_CPNOM As Field 'Nome
Public BDSIS_TBTIM_CPIMP As Field 'Impressora
'Campos da Tabela Impressoras
Public BDSIS_TBIMP_CPIMP As Field 'Impressora
Public BDSIS_TBIMP_CPNOM As Field 'Nome
Public BDSIS_TBIMP_CPDRI As Field 'Driver
'Campos da Log de Eventos
Public BDSIS_TBGEV_CPDAT As Field 'Data
Public BDSIS_TBGEV_CPHOR As Field 'Hora
Public BDSIS_TBGEV_CPUSU As Field 'Usuário
Public BDSIS_TBGEV_CPCOM As Field 'Computador
Public BDSIS_TBGEV_CPEVE As Field 'Evento
Public BDSIS_TBGEV_CPVAL As Field 'Valor
'Campos da Log de Erros
Public BDSIS_TBGRR_CPDAT As Field 'Data
Public BDSIS_TBGRR_CPHOR As Field 'Hora
Public BDSIS_TBGRR_CPUSU As Field 'Usuário
Public BDSIS_TBGRR_CPCOM As Field 'Computador
Public BDSIS_TBGRR_CPAPL As Field 'Aplicativo
Public BDSIS_TBGRR_CPENU As Field 'Erro - Número
Public BDSIS_TBGRR_CPEDE As Field 'Erro - Descrição
Public BDSIS_TBGRR_CPEOR As Field 'Erro - Origem
Public BDSIS_TBGRR_CPEJA As Field 'Erro - Arquivo Ajuda
Public BDSIS_TBGRR_CPEJC As Field 'Erro - Contexto Ajuda
'Campos de Matéria-Prima - Índice de Peças
Public BDSIS_TBMPP_CPINP As Field 'Índice de Peças
Public BDSIS_TBMPP_CPPEC As Field 'Peças
'Campos de Matéria-Prima - Índice de Bitolas
Public BDSIS_TBMPB_CPINB As Field 'Índice de Bitolas
Public BDSIS_TBMPB_CPBIT As Field 'Bitolas
'Campos de Matéria-Prima - Índice de Materiais
Public BDSIS_TBMPM_CPINM As Field 'Índice de Materiais
Public BDSIS_TBMPM_CPMAT As Field 'Materiais
'Campos de Matéria-Prima - Índice de Quantidades
Public BDSIS_TBMPQ_CPINQ As Field 'Índice de Quantidades
Public BDSIS_TBMPQ_CPQUA As Field 'Quantidades
'Campos de Matéria-Prima - Índice de Nomes
Public BDSIS_TBMPN_CPINN As Field 'Índice de Nomes
Public BDSIS_TBMPN_CPNOM As Field 'Nomes
'Campos de Matéria-Prima - Relação de Materiais
Public BDSIS_TBMPR_CPFIG As Field 'Figura
Public BDSIS_TBMPR_CPMFG As Field 'Material da Figura
Public BDSIS_TBMPR_CPMMP As Field 'Material da Matéria-Prima
'Campos de Entrada de Peças - Índice
Public BDSIS_TBEPI_CPNEP As Field 'Entrada
Public BDSIS_TBEPI_CPDAT As Field 'Data
Public BDSIS_TBEPI_CPHOR As Field 'Hora
Public BDSIS_TBEPI_CPIND As Field 'Índice
'Campos de Entrada de Peças
Public BDSIS_TBEPE_CPIND As Field 'Índice
Public BDSIS_TBEPE_CPQUA As Field 'Quantidade
Public BDSIS_TBEPE_CPIFI As Field 'Índice de Ficha
Public BDSIS_TBEPE_CPFIN As Field 'Finalidade
'Campos de Estoque - Complementos
Public BDSIS_TBECM_CPIND As Field 'Índice do Complemento
Public BDSIS_TBECM_CPCOM As Field 'Complemento
'Campos de Mapa - Cotações
Public BDSIS_TBMCO_CPMEA As Field 'MesAno
Public BDSIS_TBMCO_CPVAL As Field 'Valor
'Campos da Tabela Pedidos
Public BDSIS_TBPED_CPIND As Field 'Índice de Pedido
Public BDSIS_TBPED_CPDAT As Field 'Data Emissão
Public BDSIS_TBPED_CPHOR As Field 'Hora
Public BDSIS_TBPED_CPEMP As Field 'Empresa
Public BDSIS_TBPED_CPCON As Field 'Contato
Public BDSIS_TBPED_CPCPG As Field 'Condição de Pagamento
Public BDSIS_TBPED_CPTRA As Field 'Transportadora
Public BDSIS_TBPED_CPNSP As Field 'Seu Pedido
Public BDSIS_TBPED_CPITE As Field 'Índice dos Ítens
Public BDSIS_TBPED_CPVAL As Field 'Valor
Public BDSIS_TBPED_CPIVE As Field 'Índice Vendedor
Public BDSIS_TBPED_CPIFR As Field 'Índice Frete
Public BDSIS_TBPED_CPOUD As Field 'Outras Despesas
Public BDSIS_TBPED_CPDAD As Field 'Dados Adicionais
Public BDSIS_TBPED_CPOBS As Field 'Observações
Public BDSIS_TBPED_CPLIQ As Field 'Liquidado
Public BDSIS_TBPED_CPABE As Field 'Em Aberto
'Campos da Tabela Pedidos - Ítens
Public BDSIS_TBPIT_CPIND As Field 'Índice do Ítem
Public BDSIS_TBPIT_CPNPE As Field 'Número do Pedido
Public BDSIS_TBPIT_CPINF As Field 'Índice de Ficha
Public BDSIS_TBPIT_CPQUA As Field 'Quantidade
Public BDSIS_TBPIT_CPDES As Field 'Descrição
Public BDSIS_TBPIT_CPCOM As Field 'Complemento
Public BDSIS_TBPIT_CPPRE As Field 'Preço Unitário
Public BDSIS_TBPIT_CPPRA As Field 'Prazo de Entrega
Public BDSIS_TBPIT_CPLIQ As Field 'Liquidado
'Campos de Mapa - Pedidos
Public BDSIS_TBMPE_CPMEA As Field 'MesAno
Public BDSIS_TBMPE_CPVAL As Field 'Valor
'Campos de Ordem de Fabricação
Public BDSIS_TBODF_CPNOF As Field 'OF
Public BDSIS_TBODF_CPDAE As Field 'Data Emissão
Public BDSIS_TBODF_CPINF As Field 'Índice de Ficha
Public BDSIS_TBODF_CPDES As Field 'Descrição
Public BDSIS_TBODF_CPNDC As Field 'DC
Public BDSIS_TBODF_CPQES As Field 'Quantidade Estipulada
Public BDSIS_TBODF_CPQPR As Field 'Quantidade Produzida
Public BDSIS_TBODF_CPCOR As Field 'Corridas
Public BDSIS_TBODF_CPDAL As Field 'Data Liquidação
Public BDSIS_TBODF_CPINP As Field 'Índices de Processos
Public BDSIS_TBODF_CPLIQ As Field 'Liquidado
'Campos de Ordem de Fabricação - Processos
Public BDSIS_TBOFP_CPIND As Field 'Índice do Processo
Public BDSIS_TBOFP_CPNOF As Field 'OF
Public BDSIS_TBOFP_CPNET As Field 'Etapa
Public BDSIS_TBOFP_CPNMA As Field 'Máquina
Public BDSIS_TBOFP_CPDAI As Field 'Data Início
Public BDSIS_TBOFP_CPDAF As Field 'Data Fim
Public BDSIS_TBOFP_CPNHE As Field 'Hora Extra
'Campos de Ordem de Montagem
Public BDSIS_TBODM_CPNOM As Field 'OM
Public BDSIS_TBODM_CPDAT As Field 'Data
Public BDSIS_TBODM_CPHOR As Field 'Hora
Public BDSIS_TBODM_CPINF As Field 'Índice de Ficha
Public BDSIS_TBODM_CPDES As Field 'Descrição
Public BDSIS_TBODM_CPNDC As Field 'DC
Public BDSIS_TBODM_CPQES As Field 'Quantidade Estipulada
Public BDSIS_TBODM_CPQMO As Field 'Quantidade Montada
Public BDSIS_TBODM_CPDIN As Field 'Data Início
Public BDSIS_TBODM_CPHIN As Field 'Hora Início
Public BDSIS_TBODM_CPDFI As Field 'Data Fim
Public BDSIS_TBODM_CPHFI As Field 'Hora Fim
Public BDSIS_TBODM_CPNCA As Field 'NCA
Public BDSIS_TBODM_CPNCB As Field 'NCB
Public BDSIS_TBODM_CPNCC As Field 'NCC
Public BDSIS_TBODM_CPNCD As Field 'NCD
Public BDSIS_TBODM_CPNCE As Field 'NCE
Public BDSIS_TBODM_CPLIQ As Field 'Liquidado
'Campos de Ordem de Expedição
Public BDSIS_TBODE_CPNOE As Field 'OE
Public BDSIS_TBODE_CPDAT As Field 'Data
Public BDSIS_TBODE_CPHOR As Field 'Hora
Public BDSIS_TBODE_CPEMP As Field 'Empresa
Public BDSIS_TBODE_CPNPE As Field 'PE
Public BDSIS_TBODE_CPCFS As Field 'Código Funcionário - Separado
Public BDSIS_TBODE_CPCFC As Field 'Código Funcionário - Conferido
Public BDSIS_TBODE_CPCFV As Field 'Código Funcionário - Visto
Public BDSIS_TBODE_CPETP As Field 'Embalagem - Tipo
Public BDSIS_TBODE_CPENU As Field 'Embalagem - Números
Public BDSIS_TBODE_CPPBR As Field 'Peso Bruto
Public BDSIS_TBODE_CPPLI As Field 'Peso Líquido
Public BDSIS_TBODE_CPTRA As Field 'Transportadora
Public BDSIS_TBODE_CPCOR As Field 'Corridas
Public BDSIS_TBODE_CPOBS As Field 'Observações
Public BDSIS_TBODE_CPLIQ As Field 'Liquidado
'Campos de Funcionários
Public BDSIS_TBFUN_CPIND As Field 'Índice
Public BDSIS_TBFUN_CPFUN As Field 'Funcionários
Public BDSIS_TBFUN_CPREG As Field 'Número do Registro
Public BDSIS_TBFUN_CPTIP As Field 'Tipo
Public BDSIS_TBFUN_CPSET As Field 'Setor
'Campos de Contas à Pagar e Receber
Public BDSIS_TBCPR_CPIND As Field 'Índice do Movimento
Public BDSIS_TBCPR_CPEMI As Field 'Data de Emissão
Public BDSIS_TBCPR_CPVEN As Field 'Data de Vencimento
Public BDSIS_TBCPR_CPMOV As Field 'Movimento
Public BDSIS_TBCPR_CPVAL As Field 'Valor
Public BDSIS_TBCPR_CPORI As Field 'Origem e Destino
Public BDSIS_TBCPR_CPNDO As Field 'Número Documento
Public BDSIS_TBCPR_CPNNU As Field 'Nosso Número
Public BDSIS_TBCPR_CPSNU As Field 'Seu Número
Public BDSIS_TBCPR_CPTIP As Field 'Tipo
Public BDSIS_TBCPR_CPBAN As Field 'Banco
Public BDSIS_TBCPR_CPCAR As Field 'Em Carteira
Public BDSIS_TBCPR_CPCOB As Field 'Em Cobrança
Public BDSIS_TBCPR_CPOBS As Field 'Observações
'Campos de Imagens
Public BDSIS_TBIMA_CPIND As Field 'Índice da Imagem
Public BDSIS_TBIMA_CPNOM As Field 'Nome
Public BDSIS_TBIMA_CPTIP As Field 'Tipo da Imagem
Public BDSIS_TBIMA_CPIMA As Field 'Imagem
'Campos de Equipamentos
Public BDSIS_TBEQU_CPIND As Field 'Índice
Public BDSIS_TBEQU_CPEQU As Field 'Equipamento
Public BDSIS_TBEQU_CPNOM As Field 'Nome
Public BDSIS_TBEQU_CPMOD As Field 'Modelo
Public BDSIS_TBEQU_CPFAB As Field 'Fabricante
Public BDSIS_TBEQU_CPANO As Field 'Ano
Public BDSIS_TBEQU_CPVAL As Field 'Valor
Public BDSIS_TBEQU_CPFOT As Field 'Nome da Foto
Public BDSIS_TBEQU_CPCHO As Field 'Custo da Hora
Public BDSIS_TBEQU_CPPRO As Field 'Processo
Public BDSIS_TBEQU_CPFUN As Field 'Índice do Funcionário
'Campos de Equipamentos - Manutenção
Public BDSIS_TBEMA_CPIND As Field 'Índice
Public BDSIS_TBEMA_CPEQU As Field 'Equipamento
Public BDSIS_TBEMA_CPMAN As Field 'Manutenção
Public BDSIS_TBEMA_CPDUL As Field 'Data da Última
Public BDSIS_TBEMA_CPDPR As Field 'Data da Próxima
Public BDSIS_TBEMA_CPVAL As Field 'Valor
Public BDSIS_TBEMA_CPEMP As Field 'Empresa
Public BDSIS_TBEMA_CPCON As Field 'Contato
Public BDSIS_TBEMA_CPTIP As Field 'Tipo
Public BDSIS_TBEMA_CPAVI As Field 'Aviso
Public BDSIS_TBEMA_CPUSU As Field 'Usuário
'Campos de Avisos
Public BDSIS_TBAVI_CPIND As Field 'Índice
Public BDSIS_TBAVI_CPREM As Field 'Remetente
Public BDSIS_TBAVI_CPEMI As Field 'Data de Emissão
Public BDSIS_TBAVI_CPDES As Field 'Destinatário
Public BDSIS_TBAVI_CPVEN As Field 'Data do Aviso
Public BDSIS_TBAVI_CPTIP As Field 'Tipo
Public BDSIS_TBAVI_CPAVI As Field 'Aviso
'Campos de CNC - Funções
Public BDSIS_TBCFU_CPIND As Field 'Índice
Public BDSIS_TBCFU_CPCNC As Field 'CNC
Public BDSIS_TBCFU_CPFUN As Field 'Função
Public BDSIS_TBCFU_CPAPL As Field 'Aplicação
Public BDSIS_TBCFU_CPVAR As Field 'Variáveis
'Campos de CNC - Máquinas
Public BDSIS_TBCMA_CPIND As Field 'Índice
Public BDSIS_TBCMA_CPNNE As Field 'Índice NE
'Campos de CNC - Programas
Public BDSIS_TBCNC_CPIND As Field 'Índice
Public BDSIS_TBCNC_CPNOM As Field 'Nome do Programa
Public BDSIS_TBCNC_CPCNC As Field 'Nome no CNC
Public BDSIS_TBCNC_CPDAT As Field 'Data Criação
Public BDSIS_TBCNC_CPULT As Field 'Data Último Acesso
Public BDSIS_TBCNC_CPTIP As Field 'Tipo do CNC
Public BDSIS_TBCNC_CPMAQ As Field 'Índice da Máquina
Public BDSIS_TBCNC_CPPRO As Field 'Programa
'Campos de Layouts
Public BDSIS_TBLAY_CPIND As Field 'Número NL
Public BDSIS_TBLAY_CPLAY As Field 'Layout
Public BDSIS_TBLAY_CPMAQ As Field 'Máquina
Public BDSIS_TBLAY_CPNEL As Field 'Número EL
Public BDSIS_TBLAY_CPNDL As Field 'Número DL
Public BDSIS_TBLAY_CPDCR As Field 'Data de Criação
Public BDSIS_TBLAY_CPCOM As Field 'Comentário
Public BDSIS_TBLAY_CPNRE As Field 'Número da Revisão
Public BDSIS_TBLAY_CPDRE As Field 'Data da Revisão
Public BDSIS_TBLAY_CPCFU As Field 'Código do Funcionário
Public BDSIS_TBLAY_CPFOT As Field 'Foto
'Campos de Layouts - Esquemas
Public BDSIS_TBLES_CPIND As Field 'Número EL
Public BDSIS_TBLES_CPESQ As Field 'Esquema
Public BDSIS_TBLES_CPMAQ As Field 'Máquina
Public BDSIS_TBLES_CPVAR As Field 'Variáveis
Public BDSIS_TBLES_CPDIS As Field 'Dispositivos
Public BDSIS_TBLES_CPFER As Field 'Ferramentas
'Campos de Layouts - Dimensões
Public BDSIS_TBLDI_CPIND As Field 'Número DL
Public BDSIS_TBLDI_CPNEL As Field 'Número EL
Public BDSIS_TBLDI_CPVAR As Field 'Variáveis
Public BDSIS_TBLDI_CPDIS As Field 'Dispositivos
Public BDSIS_TBLDI_CPFER As Field 'Ferramentas
Public BDSIS_TBLDI_CPOBS As Field 'Observações



'*****************************************************
'    B A N C O   D E   D A D O S   C L I E N T E
'*****************************************************
Public BDCLI As Database 'Cliente

' ******************   TABELAS    ******************
Public BDCLI_TBDPR As Recordset 'Diretório do Programa
Public BDCLI_TBDBD As Recordset 'Diretório do Banco de Dados
Public BDCLI_TBIMP As Recordset 'Impressoras
Public BDCLI_TBULG As Recordset 'Último Logon


' ******************   CAMPOS    *******************
' Campos da Tabela Diretorio do Programa
Public BDCLI_TBDPR_CPCAM As Field 'Caminho
' Campos da Tabela Diretorio do Banco de Dados
Public BDCLI_TBDBD_CPCAM As Field 'Caminho
' Campos da Tabela Impressoras
Public BDCLI_TBIMP_CPIMP As Field 'Impressora
Public BDCLI_TBIMP_CPNOM As Field 'Nome
Public BDCLI_TBIMP_CPDRI As Field 'Driver
' Campos da Tabela Último Logon
Public BDCLI_TBULG_CPDAT As Field 'Data
Public BDCLI_TBULG_CPHOR As Field 'Hora
Public BDCLI_TBULG_CPCOM As Field 'Computador
Public BDCLI_TBULG_CPUSU As Field 'Usuário




'*****************************************************
'                    F U N Ç Õ E S
'*****************************************************
Public Function IniciaSistema(X, Y, Z) As Boolean
    If X <> App.ProductName And _
       Y <> App.EXEName And _
       Z <> App.LegalCopyright Then
       IniciaSistema = False
       Exit Function
    End If

    IniciaSistema = False
    Dir_EXE = "c:\Arquivos de Programas\Siscoval"
    Dir_BD = "\\Servidor1\Sistema"
    'Lê registro do Siscoval'
GOTO_PEGADIREXE: 'Pega diretório do Siscoval.exe
    On Error GoTo ERRO_DIREXE
    'Dir_EXE = GetSetting("Siscoval", "Diretório", "Programa")
    Tela_CD.Visible = False
    Tela_CD.LA.Visible = False
    Tela_CD.LA.Path = Dir_EXE
    Tela_CD.LA.Pattern = "Siscoval.exe"
    If Tela_CD.LA.ListCount < 1 Then GoTo ERRO_DIREXE
    Unload Tela_CD
GOTO_PEGADIRBD: 'Pega diretório do Siscoval.scv
    On Error GoTo ERRO_DIRBD
    'Dir_BD = GetSetting("Siscoval", "Diretório", "Banco de Dados")
    Tela_CD.Visible = True
    Tela_CD.LA.Visible = True
    Tela_CD.LA.Path = Dir_BD
    Tela_CD.LA.Pattern = "Siscoval.scv"
    If Tela_CD.LA.ListCount < 1 Then GoTo ERRO_DIRBD
    Unload Tela_CD
' GOTO_PEGADIRCLI: 'Verifica arquivo Cliente.scv
'    On Error GoTo ERRO_DIREXE
'    'Dir_EXE = GetSetting("Siscoval", "Diretório", "Programa")
'    Tela_CD.Visible = False
'    Tela_CD.LA.Visible = False
'    Tela_CD.LA.Path = Dir_EXE
'    Tela_CD.LA.Pattern = "Cliente.scv"
'    If Tela_CD.LA.ListCount < 1 Then GoTo ERRO_CANCELA
    Unload Tela_CD

GOTO_ARRUMA_BDSIS: 'Repara e compacta BD Siscoval
    On Error Resume Next
    'Arruma BDSIS
GOTO_ARRUMA_BDCLI: 'Repara e compacta BD Cliente
    On Error Resume Next
    'Arruma BDCLI
    
    'Se chegar aqui ocorreu tudo bem
    IniciaSistema = True
    Exit Function


ERRO_CANCELA: 'Se entrar aqui, ocorreu algum problema que encerrará a abertura
    IniciaSistema = False
    Exit Function
ERRO_DIREXE:
    Tela_CD.Visible = False
    Tela_CD.CD.DialogTitle = "Indique o caminho do arquivo Siscoval.exe"
    Tela_CD.CD.FileName = ""
    Tela_CD.CD.Filter = "Siscoval.exe|Siscoval.exe"
    Tela_CD.CD.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames
    Tela_CD.CD.ShowOpen
    If Tela_CD.CD.FileName <> "" Then SaveSetting "Siscoval", "Diretório", "Programa", RetornaDir(Tela_CD.CD.FileName)
    Unload Tela_CD
    GoTo GOTO_PEGADIREXE
ERRO_DIRBD:
    Tela_CD.Visible = False
    Tela_CD.CD.DialogTitle = "Indique o caminho do arquivo Siscoval.scv"
    Tela_CD.CD.FileName = ""
    Tela_CD.CD.Filter = "Siscoval.scv|Siscoval.scv"
    Tela_CD.CD.Flags = cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames
    Tela_CD.CD.ShowOpen
    If Tela_CD.CD.FileName <> "" Then SaveSetting "Siscoval", "Diretório", "Banco de Dados", RetornaDir(Tela_CD.CD.FileName)
    Unload Tela_CD
    GoTo GOTO_PEGADIRBD
End Function
Private Function RetornaDir(DirInicial As String) As String
    Dim DirFinal As String, I As Integer
    DirFinal = ""
    For I = Len(DirInicial) To 1 Step -1
        If Mid(DirInicial, I, 1) = "\" Then
            DirFinal = Left(DirInicial, Len(DirInicial) - (Len(DirInicial) - I) - 1)
            Exit For
        End If
    Next I
    RetornaDir = DirFinal
End Function
Public Function AbreBD(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set WKSP = DBEngine.Workspaces(0)
    Dir_EXE = GetSetting("Siscoval", "Diretório", "Dir_EXE")
    Dir_BD = GetSetting("Siscoval", "Diretório", "Dir_BD")
    'Dir_EXE = "C:\Arquivos de Programas\Siscoval"
    'Dir_BD = "\\VENDAS\Sistema"
    On Error GoTo ERRO_SCV
    Set BDSIS = WKSP.OpenDatabase(Dir_BD & "\SISCOVAL.SCV", , False)
    'On Error GoTo ERRO_CLI
    'Set BDCLI = WKSP.OpenDatabase(Dir_EXE & "\Cliente.scv", , False)
    AbreBD = True
    Exit Function
ERRO_SCV:
    MsgBox "erro no siscoval"
ERRO_CLI:
    MsgBox "erro no cliente"
ERRO_ABERTURA:
    AbreBD = False
End Function
Public Function FechaBD(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    BDSIS.Close
    'BDCLI.Close
    Set BDSIS = Nothing
    'Set BDCLI = Nothing
    FechaBD = True
    Exit Function
ERRO_FINALIZACAO:
    FechaBD = False
End Function




'**********************************************************************************************************
'
'           B A N C O     D E     D A D O S     S I S C O V A L
'
'**********************************************************************************************************





'*****************************************************
'     F U N Ç Õ E S     A B R E     T A B E L A S
'*****************************************************
Public Function AbreTabela_Usuarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBUSU = BDSIS.OpenRecordset("Usuários")
    BDSIS_TBUSU.Index = "Usuário"
    AbreTabela_Usuarios = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Usuarios = False
End Function
Public Function AbreTabela_UsuariosComputadores(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBUCO = BDSIS.OpenRecordset("Usuários - Computadores")
    BDSIS_TBUCO.Index = "Computador"
    AbreTabela_UsuariosComputadores = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_UsuariosComputadores = False
End Function
Public Function AbreTabela_UsuariosMenus(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBUME = BDSIS.OpenRecordset("Usuários - Menus")
    BDSIS_TBUME.Index = "Menu"
    AbreTabela_UsuariosMenus = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_UsuariosMenus = False
End Function
Public Function AbreTabela_Grupos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBGRU = BDSIS.OpenRecordset("Grupos")
    BDSIS_TBGRU.Index = "Grupo"
    AbreTabela_Grupos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Grupos = False
End Function
Public Function AbreTabela_Estoque(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEST = BDSIS.OpenRecordset("Estoque")
    BDSIS_TBEST.Index = "Estoque"
    AbreTabela_Estoque = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Estoque = False
End Function
Public Function AbreTabela_EstoqueIndice(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEID = BDSIS.OpenRecordset("Estoque - Índice")
    BDSIS_TBEID.Index = "Índice"
    AbreTabela_EstoqueIndice = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EstoqueIndice = False
End Function
Public Function AbreTabela_EstoqueFiguras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEFG = BDSIS.OpenRecordset("Estoque - Figuras")
    BDSIS_TBEFG.Index = "Figura"
    AbreTabela_EstoqueFiguras = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EstoqueFiguras = False
End Function
Public Function AbreTabela_Empresas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEMP = BDSIS.OpenRecordset("Empresas")
    BDSIS_TBEMP.Index = "Apelido"
    AbreTabela_Empresas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Empresas = False
End Function
Public Function AbreTabela_EstoqueAliquotas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEAL = BDSIS.OpenRecordset("Estoque - Alíquotas")
    BDSIS_TBEAL.Index = "ProcAli"
    AbreTabela_EstoqueAliquotas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EstoqueAliquotas = False
End Function
Public Function AbreTabela_Bancos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBBAN = BDSIS.OpenRecordset("Bancos")
    BDSIS_TBBAN.Index = "Nome da Conta"
    AbreTabela_Bancos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Bancos = False
End Function
Public Function AbreTabela_CodigosFiscais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCDF = BDSIS.OpenRecordset("Códigos Fiscais")
    'BDSIS_TBCDF.Index = ""
    AbreTabela_CodigosFiscais = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_CodigosFiscais = False
End Function
Public Function AbreTabela_EstoqueCFeST(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCFS = BDSIS.OpenRecordset("Estoque - CF e ST")
    BDSIS_TBCFS.Index = "FigMat"
    AbreTabela_EstoqueCFeST = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EstoqueCFeST = False
End Function
Public Function AbreTabela_EstoqueDeclaracoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEDC = BDSIS.OpenRecordset("Estoque - Declarações")
    BDSIS_TBEDC.Index = "CF"
    AbreTabela_EstoqueDeclaracoes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EstoqueDeclaracoes = False
End Function
Public Function AbreTabela_ConfiguracoesNotaFiscal(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCNF = BDSIS.OpenRecordset("Configurações da Nota Fiscal")
    BDSIS_TBCNF.Index = "Item"
    AbreTabela_ConfiguracoesNotaFiscal = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_ConfiguracoesNotaFiscal = False
End Function
Public Function AbreTabela_ConfiguracoesNotaFiscalImagem(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBINF = BDSIS.OpenRecordset("Configurações da Nota Fiscal - Imagem")
    AbreTabela_ConfiguracoesNotaFiscalImagem = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_ConfiguracoesNotaFiscalImagem = False
End Function
Public Function AbreTabela_NotaFiscal(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNTF = BDSIS.OpenRecordset("Nota Fiscal")
    BDSIS_TBNTF.Index = "NumNF"
    AbreTabela_NotaFiscal = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_NotaFiscal = False
End Function
Public Function AbreTabela_NotaFiscalProdutos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFP = BDSIS.OpenRecordset("Nota Fiscal - Produtos")
    BDSIS_TBNFP.Index = "Número da Nota Fiscal"
    AbreTabela_NotaFiscalProdutos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_NotaFiscalProdutos = False
End Function
Public Function AbreTabela_NotaFiscalDeclaracoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFD = BDSIS.OpenRecordset("Nota Fiscal - Declarações")
    AbreTabela_NotaFiscalDeclaracoes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_NotaFiscalDeclaracoes = False
End Function
Public Function AbreTabela_NotaFiscalBancos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFB = BDSIS.OpenRecordset("Nota Fiscal - Bancos")
    BDSIS_TBNFB.Index = "NumNF"
    AbreTabela_NotaFiscalBancos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_NotaFiscalBancos = False
End Function
Public Function AbreTabela_NotaFiscalComentarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFC = BDSIS.OpenRecordset("Nota Fiscal - Comentários")
    AbreTabela_NotaFiscalComentarios = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_NotaFiscalComentarios = False
End Function
Public Function AbreTabela_MapaFaturamentoMes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMFA = BDSIS.OpenRecordset("Mapa - Faturamento do Mês")
    BDSIS_TBMFA.Index = "NumNF"
    AbreTabela_MapaFaturamentoMes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MapaFaturamentoMes = False
End Function
Public Function AbreTabela_MapaContasReceber(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMCR = BDSIS.OpenRecordset("Mapa - Contas à Receber")
    BDSIS_TBMCR.Index = "DataEmpDP"
    AbreTabela_MapaContasReceber = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MapaContasReceber = False
End Function
Public Function AbreTabela_MapaImpostos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMIM = BDSIS.OpenRecordset("Mapa - Impostos")
    BDSIS_TBMIM.Index = "ImpMesAno"
    AbreTabela_MapaImpostos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MapaImpostos = False
End Function
Public Function AbreTabela_Cotacoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCOT = BDSIS.OpenRecordset("Cotações")
    BDSIS_TBCOT.Index = "Índice de Cotação"
    AbreTabela_Cotacoes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Cotacoes = False
End Function
Public Function AbreTabela_CotacoesItens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCTI = BDSIS.OpenRecordset("Cotações - Ítens")
    BDSIS_TBCTI.Index = "Ítem"
    AbreTabela_CotacoesItens = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_CotacoesItens = False
End Function
Public Function AbreTabela_EmpresasContatos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBECO = BDSIS.OpenRecordset("Empresas - Contatos")
    BDSIS_TBECO.Index = "Contato"
    AbreTabela_EmpresasContatos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EmpresasContatos = False
End Function
Public Function AbreTabela_ConfiguracoesTelasImpressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBTIM = BDSIS.OpenRecordset("Configurações de Telas para Impressoras")
    BDSIS_TBTIM.Index = "Tela"
    AbreTabela_ConfiguracoesTelasImpressoras = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_ConfiguracoesTelasImpressoras = False
End Function
Public Function AbreTabela_Impressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBIMP = BDSIS.OpenRecordset("Impressoras")
    BDSIS_TBIMP.Index = "Impressora"
    AbreTabela_Impressoras = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Impressoras = False
End Function
Public Function AbreTabela_LogDeEventos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBGEV = BDSIS.OpenRecordset("Log de Eventos")
    'BDSIS_TBGEV.Index = ""
    AbreTabela_LogDeEventos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_LogDeEventos = False
End Function
Public Function AbreTabela_LogDeErros(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBGRR = BDSIS.OpenRecordset("Log de Erros")
    'BDSIS_TBGEV.Index = ""
    AbreTabela_LogDeErros = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_LogDeErros = False
End Function
Public Function AbreTabela_EntradaPecasIndice(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEPI = BDSIS.OpenRecordset("Entrada de Peças - Índice")
    BDSIS_TBEPI.Index = "Entrada"
    AbreTabela_EntradaPecasIndice = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EntradaPecasIndice = False
End Function
Public Function AbreTabela_EntradaPecas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEPE = BDSIS.OpenRecordset("Entrada de Peças")
    BDSIS_TBEPE.Index = "Índice"
    AbreTabela_EntradaPecas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EntradaPecas = False
End Function
Public Function AbreTabela_MateriaPrimaIndiceDePecas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPP = BDSIS.OpenRecordset("Matéria-Prima - Índice de Peças")
    BDSIS_TBMPP.Index = "Peças"
    AbreTabela_MateriaPrimaIndiceDePecas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MateriaPrimaIndiceDePecas = False
End Function
Public Function AbreTabela_MateriaPrimaIndiceDeBitolas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPB = BDSIS.OpenRecordset("Matéria-Prima - Índice de Bitolas")
    BDSIS_TBMPB.Index = "Bitolas"
    AbreTabela_MateriaPrimaIndiceDeBitolas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MateriaPrimaIndiceDeBitolas = False
End Function
Public Function AbreTabela_MateriaPrimaIndiceDeMateriais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPM = BDSIS.OpenRecordset("Matéria-Prima - Índice de Materiais")
    BDSIS_TBMPM.Index = "Materiais"
    AbreTabela_MateriaPrimaIndiceDeMateriais = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MateriaPrimaIndiceDeMateriais = False
End Function
Public Function AbreTabela_MateriaPrimaIndiceDeQuantidades(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPQ = BDSIS.OpenRecordset("Matéria-Prima - Índice de Quantidades")
    BDSIS_TBMPQ.Index = "Quantidades"
    AbreTabela_MateriaPrimaIndiceDeQuantidades = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MateriaPrimaIndiceDeQuantidades = False
End Function
Public Function AbreTabela_MateriaPrimaIndiceDeNomes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPN = BDSIS.OpenRecordset("Matéria-Prima - Índice de Nomes")
    'BDSIS_TBMPN.Index = "Nomes"
    AbreTabela_MateriaPrimaIndiceDeNomes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MateriaPrimaIndiceDeNomes = False
End Function
Public Function AbreTabela_MateriaPrimaRelacaoMateriais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPR = BDSIS.OpenRecordset("Matéria-Prima - Relação de Materiais")
    BDSIS_TBMPR.Index = "FigMat"
    AbreTabela_MateriaPrimaRelacaoMateriais = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MateriaPrimaRelacaoMateriais = False
End Function
Public Function AbreTabela_EstoqueComplementos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBECM = BDSIS.OpenRecordset("Estoque - Complementos")
    BDSIS_TBECM.Index = "Índice do Complemento"
    AbreTabela_EstoqueComplementos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EstoqueComplementos = False
End Function
Public Function AbreTabela_MapaCotacoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMCO = BDSIS.OpenRecordset("Mapa - Cotações")
    BDSIS_TBMCO.Index = "MêsAno"
    AbreTabela_MapaCotacoes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MapaCotacoes = False
End Function
Public Function AbreTabela_MapaPedidos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPE = BDSIS.OpenRecordset("Mapa - Pedidos")
    BDSIS_TBMPE.Index = "MêsAno"
    AbreTabela_MapaPedidos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_MapaPedidos = False
End Function
Public Function AbreTabela_OrdemFabricacao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBODF = BDSIS.OpenRecordset("Ordem de Fabricação")
    BDSIS_TBODF.Index = "OF"
    AbreTabela_OrdemFabricacao = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_OrdemFabricacao = False
End Function
Public Function AbreTabela_OrdemFabricacaoProcessos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBOFP = BDSIS.OpenRecordset("Ordem de Fabricação - Processos")
    BDSIS_TBOFP.Index = "Índice do Processo"
    AbreTabela_OrdemFabricacaoProcessos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_OrdemFabricacaoProcessos = False
End Function
Public Function AbreTabela_OrdemMontagem(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBODM = BDSIS.OpenRecordset("Ordem de Montagem")
    BDSIS_TBODF.Index = "OF"
    AbreTabela_OrdemMontagem = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_OrdemMontagem = False
End Function
Public Function AbreTabela_OrdemExpedicao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBODE = BDSIS.OpenRecordset("Ordem de Expedição")
    BDSIS_TBODE.Index = "OE"
    AbreTabela_OrdemExpedicao = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_OrdemExpedicao = False
End Function
Public Function AbreTabela_Funcionarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBFUN = BDSIS.OpenRecordset("Funcionários")
    BDSIS_TBFUN.Index = "Índice"
    AbreTabela_Funcionarios = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Funcionarios = False
End Function
Public Function AbreTabela_ContasPagarReceber(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCPR = BDSIS.OpenRecordset("Contas à Pagar e Receber")
    BDSIS_TBCPR.Index = "Índice do Movimento"
    AbreTabela_ContasPagarReceber = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_ContasPagarReceber = False
End Function
Public Function AbreTabela_Imagens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBIMA = BDSIS.OpenRecordset("Imagens")
    BDSIS_TBIMA.Index = "Índice da Imagem"
    AbreTabela_Imagens = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Imagens = False
End Function
Public Function AbreTabela_Equipamentos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEQU = BDSIS.OpenRecordset("Equipamentos")
    BDSIS_TBEQU.Index = "Índice do Equipamento"
    AbreTabela_Equipamentos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Equipamentos = False
End Function
Public Function AbreTabela_EquipamentosManutencao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEMA = BDSIS.OpenRecordset("Equipamentos - Manutenção")
    BDSIS_TBEMA.Index = "Índice"
    AbreTabela_EquipamentosManutencao = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_EquipamentosManutencao = False
End Function
Public Function AbreTabela_Avisos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBAVI = BDSIS.OpenRecordset("Avisos")
    BDSIS_TBAVI.Index = "Índice"
    AbreTabela_Avisos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Avisos = False
End Function
Public Function AbreTabela_CNCFuncoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCFU = BDSIS.OpenRecordset("CNC - Funções")
    BDSIS_TBCFU.Index = "Índice"
    AbreTabela_CNCFuncoes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_CNCFuncoes = False
End Function
Public Function AbreTabela_CNCMaquinas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCMA = BDSIS.OpenRecordset("CNC - Máquinas")
    'BDSIS_TBCMA.Index = "Índice"
    AbreTabela_CNCMaquinas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_CNCMaquinas = False
End Function
Public Function AbreTabela_CNCProgramas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCNC = BDSIS.OpenRecordset("CNC - Programas")
    BDSIS_TBCNC.Index = "Índice"
    AbreTabela_CNCProgramas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_CNCProgramas = False
End Function
Public Function AbreTabela_Layouts(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBLAY = BDSIS.OpenRecordset("Layouts")
    BDSIS_TBLAY.Index = "Índice"
    AbreTabela_Layouts = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Layouts = False
End Function
Public Function AbreTabela_LayoutsEsquemas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBLES = BDSIS.OpenRecordset("Layouts - Esquemas")
    BDSIS_TBLES.Index = "Índice"
    AbreTabela_LayoutsEsquemas = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_LayoutsEsquemas = False
End Function
Public Function AbreTabela_LayoutsDimensoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBLDI = BDSIS.OpenRecordset("Layouts - Dimensões")
    BDSIS_TBLDI.Index = "Índice"
    AbreTabela_LayoutsDimensoes = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_LayoutsDimensoes = False
End Function



'*****************************************************
'     F U N Ç Õ E S     F E C H A     T A B E L A S
'*****************************************************
Public Function FechaTabela_Usuarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBUSU_CPUSU = Nothing
    Set BDSIS_TBUSU_CPADM = Nothing
    Set BDSIS_TBUSU_CPNOM = Nothing
    Set BDSIS_TBUSU_CPSEN = Nothing
    Set BDSIS_TBUSU_CPPER = Nothing
    BDSIS_TBUSU.Close
    Set BDSIS_TBUSU = Nothing
    FechaTabela_Usuarios = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Usuarios = False
End Function
Public Function FechaTabela_UsuariosComputadores(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBUCO_CPCOM = Nothing
    Set BDSIS_TBUCO_CPDWI = Nothing
    Set BDSIS_TBUCO_CPDSY = Nothing
    Set BDSIS_TBUCO_CPDTP = Nothing
    Set BDSIS_TBUCO_CPOBU = Nothing
    Set BDSIS_TBUCO_CPOCS = Nothing
    Set BDSIS_TBUCO_CPOVE = Nothing
    Set BDSIS_TBUCO_CPOPL = Nothing
    Set BDSIS_TBUCO_CPSMP = Nothing
    Set BDSIS_TBUCO_CPSOE = Nothing
    Set BDSIS_TBUCO_CPSTP = Nothing
    Set BDSIS_TBUCO_CPRAM = Nothing
    Set BDSIS_TBUCO_CPCTA = Nothing
    Set BDSIS_TBUCO_CPCSN = Nothing
    BDSIS_TBUCO.Close
    Set BDSIS_TBUCO = Nothing
    FechaTabela_UsuariosComputadores = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_UsuariosComputadores = False
End Function
Public Function FechaTabela_UsuariosMenus(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    'Set BDSIS_TBUME_CPIND = Nothing
    Set BDSIS_TBUME_CPMEN = Nothing
    'Set BDSIS_TBUME_CPPRI = Nothing
    BDSIS_TBUME.Close
    Set BDSIS_TBUME = Nothing
    FechaTabela_UsuariosMenus = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_UsuariosMenus = False
End Function
Public Function FechaTabela_Grupos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBGRU_CPGRU = Nothing
    Set BDSIS_TBGRU_CPDES = Nothing
    Set BDSIS_TBGRU_CPVAL = Nothing
    Set BDSIS_TBGRU_CPTIP = Nothing
    BDSIS_TBGRU.Close
    Set BDSIS_TBGRU = Nothing
    FechaTabela_Grupos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Grupos = False
End Function
Public Function FechaTabela_Estoque(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEST_CPFIC = Nothing
    Set BDSIS_TBEST_CPFIG = Nothing
    Set BDSIS_TBEST_CPBIT = Nothing
    Set BDSIS_TBEST_CPMAT = Nothing
    Set BDSIS_TBEST_CPVUN = Nothing
    Set BDSIS_TBEST_CPVMI = Nothing
    Set BDSIS_TBEST_CPVCU = Nothing
    Set BDSIS_TBEST_CPPUN = Nothing
    Set BDSIS_TBEST_CPEST = Nothing
    Set BDSIS_TBEST_CPETM = Nothing
    Set BDSIS_TBEST_CPVEN = Nothing
    Set BDSIS_TBEST_CPCOT = Nothing
    Set BDSIS_TBEST_CPINP = Nothing
    Set BDSIS_TBEST_CPINB = Nothing
    Set BDSIS_TBEST_CPINM = Nothing
    Set BDSIS_TBEST_CPNDC = Nothing
    BDSIS_TBEST.Close
    Set BDSIS_TBEST = Nothing
    FechaTabela_Estoque = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Estoque = False
End Function
Public Function FechaTabela_EstoqueIndice(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEID_CPIFI = Nothing
    Set BDSIS_TBEID_CPDRE = Nothing
    Set BDSIS_TBEID_CPTRE = Nothing
    Set BDSIS_TBEID_CPDNO = Nothing
    Set BDSIS_TBEID_CPTNO = Nothing
    Set BDSIS_TBEID_CPDCO = Nothing
    Set BDSIS_TBEID_CPTCO = Nothing
    Set BDSIS_TBEID_CPGPE = Nothing
    Set BDSIS_TBEID_CPGMP = Nothing
    Set BDSIS_TBEID_CPGBI = Nothing
    Set BDSIS_TBEID_CPGIN = Nothing
    Set BDSIS_TBEID_CPGEX = Nothing
    Set BDSIS_TBEID_CPGCL = Nothing
    Set BDSIS_TBEID_CPGMA = Nothing
    BDSIS_TBEID.Close
    Set BDSIS_TBEID = Nothing
    FechaTabela_EstoqueIndice = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EstoqueIndice = False
End Function
Public Function FechaTabela_EstoqueFiguras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEFG_CPFIG = Nothing
    Set BDSIS_TBEFG_CPIFG = Nothing
    Set BDSIS_TBEFG_CPGIN = Nothing
    Set BDSIS_TBEFG_CPGEX = Nothing
    Set BDSIS_TBEFG_CPGCL = Nothing
    Set BDSIS_TBEFG_CPCOM = Nothing
    BDSIS_TBEFG.Close
    Set BDSIS_TBEFG = Nothing
    FechaTabela_EstoqueFiguras = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EstoqueFiguras = False
End Function
Public Function FechaTabela_Empresas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEMP_CPAPE = Nothing
    Set BDSIS_TBEMP_CPEMP = Nothing
    Set BDSIS_TBEMP_CPCGC = Nothing
    Set BDSIS_TBEMP_CPINE = Nothing
    Set BDSIS_TBEMP_CPEND = Nothing
    Set BDSIS_TBEMP_CPBAI = Nothing
    Set BDSIS_TBEMP_CPPRA = Nothing
    Set BDSIS_TBEMP_CPCID = Nothing
    Set BDSIS_TBEMP_CPEST = Nothing
    Set BDSIS_TBEMP_CPCEP = Nothing
    Set BDSIS_TBEMP_CPFON = Nothing
    Set BDSIS_TBEMP_CPFAX = Nothing
    Set BDSIS_TBEMP_CPTIP = Nothing
    Set BDSIS_TBEMP_CPCOM = Nothing
    Set BDSIS_TBEMP_CPOBS = Nothing
    Set BDSIS_TBEMP_CPTRA = Nothing
    Set BDSIS_TBEMP_CPEMA = Nothing
    Set BDSIS_TBEMP_CPSIT = Nothing
    BDSIS_TBEMP.Close
    Set BDSIS_TBEMP = Nothing
    FechaTabela_Empresas = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Empresas = False
End Function
Public Function FechaTabela_EmpresasContatos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBECO_CPEMP = Nothing
    Set BDSIS_TBECO_CPCON = Nothing
    BDSIS_TBECO.Close
    Set BDSIS_TBECO = Nothing
    FechaTabela_EmpresasContatos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EmpresasContatos = False
End Function
Public Function FechaTabela_EstoqueAliquotas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEAL_CPALI = Nothing
    Set BDSIS_TBEAL_CPEST = Nothing
    Set BDSIS_TBEAL_CPCCF = Nothing
    Set BDSIS_TBEAL_CPPOR = Nothing
    Set BDSIS_TBEAL_CPRBC = Nothing
    BDSIS_TBEAL.Close
    Set BDSIS_TBEAL = Nothing
    FechaTabela_EstoqueAliquotas = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EstoqueAliquotas = False
End Function
Public Function FechaTabela_Bancos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBBAN_CPNMC = Nothing
    Set BDSIS_TBBAN_CPNMB = Nothing
    Set BDSIS_TBBAN_CPBAI = Nothing
    Set BDSIS_TBBAN_CPAGE = Nothing
    Set BDSIS_TBBAN_CPCON = Nothing
    BDSIS_TBBAN.Close
    Set BDSIS_TBBAN = Nothing
    FechaTabela_Bancos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Bancos = False
End Function
Public Function FechaTabela_CodigosFiscais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCDF_CPNRG = Nothing
    Set BDSIS_TBCDF_CPNTO = Nothing
    Set BDSIS_TBCDF_CPCFO = Nothing
    Set BDSIS_TBCDF_CPTIP = Nothing
    BDSIS_TBCDF.Close
    Set BDSIS_TBCDF = Nothing
    FechaTabela_CodigosFiscais = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_CodigosFiscais = False
End Function
Public Function FechaTabela_EstoqueCFeST(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCFS_CPFIG = Nothing
    Set BDSIS_TBCFS_CPGMT = Nothing
    Set BDSIS_TBCFS_CPGCF = Nothing
    Set BDSIS_TBCFS_CPGST = Nothing
    BDSIS_TBCFS.Close
    Set BDSIS_TBCFS = Nothing
    FechaTabela_EstoqueCFeST = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EstoqueCFeST = False
End Function
Public Function FechaTabela_EstoqueDeclaracoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEDC_CPCCF = Nothing
    Set BDSIS_TBEDC_CPDSP = Nothing
    Set BDSIS_TBEDC_CPDOU = Nothing
    BDSIS_TBEDC.Close
    Set BDSIS_TBEDC = Nothing
    FechaTabela_EstoqueDeclaracoes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EstoqueDeclaracoes = False
End Function
Public Function FechaTabela_ConfiguracoesNotaFiscal(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCNF_CPITE = Nothing
    Set BDSIS_TBCNF_CPDES = Nothing
    Set BDSIS_TBCNF_CPEXE = Nothing
    Set BDSIS_TBCNF_CPIND = Nothing
    Set BDSIS_TBCNF_CPCOL = Nothing
    Set BDSIS_TBCNF_CPLIN = Nothing
    Set BDSIS_TBCNF_CPFON = Nothing
    Set BDSIS_TBCNF_CPTAM = Nothing
    Set BDSIS_TBCNF_CPNEG = Nothing
    Set BDSIS_TBCNF_CPITA = Nothing
    Set BDSIS_TBCNF_CPSUB = Nothing
    BDSIS_TBCNF.Close
    Set BDSIS_TBCNF = Nothing
    FechaTabela_ConfiguracoesNotaFiscal = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_ConfiguracoesNotaFiscal = False
End Function
Public Function FechaTabela_ConfiguracoesNotaFiscalImagem(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBINF_CPIMG = Nothing
    BDSIS_TBINF.Close
    Set BDSIS_TBINF = Nothing
    FechaTabela_ConfiguracoesNotaFiscalImagem = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_ConfiguracoesNotaFiscalImagem = False
End Function
Public Function FechaTabela_Cotacoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCOT_CPIND = Nothing
    Set BDSIS_TBCOT_CPDAT = Nothing
    Set BDSIS_TBCOT_CPHOR = Nothing
    Set BDSIS_TBCOT_CPEMP = Nothing
    Set BDSIS_TBCOT_CPCON = Nothing
    Set BDSIS_TBCOT_CPCPG = Nothing
    Set BDSIS_TBCOT_CPTRA = Nothing
    Set BDSIS_TBCOT_CPITE = Nothing
    Set BDSIS_TBCOT_CPIVE = Nothing
    Set BDSIS_TBCOT_CPLIQ = Nothing
    BDSIS_TBCOT.Close
    Set BDSIS_TBCOT = Nothing
    FechaTabela_Cotacoes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Cotacoes = False
End Function
Public Function FechaTabela_CotacoesItens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCTI_CPIND = Nothing
    Set BDSIS_TBCTI_CPNCO = Nothing
    Set BDSIS_TBCTI_CPINF = Nothing
    Set BDSIS_TBCTI_CPQUA = Nothing
    Set BDSIS_TBCTI_CPDES = Nothing
    Set BDSIS_TBCTI_CPCOM = Nothing
    Set BDSIS_TBCTI_CPPRE = Nothing
    Set BDSIS_TBCTI_CPPRA = Nothing
    Set BDSIS_TBCTI_CPLIQ = Nothing
    BDSIS_TBCTI.Close
    Set BDSIS_TBCTI = Nothing
    FechaTabela_CotacoesItens = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_CotacoesItens = False
End Function
Public Function FechaTabela_MapaFaturamentoMes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMFA_CPNNF = Nothing
    Set BDSIS_TBMFA_CPDAT = Nothing
    Set BDSIS_TBMFA_CPVAL = Nothing
    Set BDSIS_TBMFA_CPEMP = Nothing
    Set BDSIS_TBMFA_CPDEV = Nothing
    BDSIS_TBMFA.Close
    Set BDSIS_TBMFA = Nothing
    FechaTabela_MapaFaturamentoMes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MapaFaturamentoMes = False
End Function
Public Function FechaTabela_MapaContasReceber(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMCR_CPDEM = Nothing
    Set BDSIS_TBMCR_CPEMP = Nothing
    Set BDSIS_TBMCR_CPNDP = Nothing
    Set BDSIS_TBMCR_CPVAL = Nothing
    Set BDSIS_TBMCR_CPDVE = Nothing
    BDSIS_TBMCR.Close
    Set BDSIS_TBMCR = Nothing
    FechaTabela_MapaContasReceber = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MapaContasReceber = False
End Function
Public Function FechaTabela_MapaImpostos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMIM_CPIMP = Nothing
    Set BDSIS_TBMIM_CPMES = Nothing
    Set BDSIS_TBMIM_CPANO = Nothing
    Set BDSIS_TBMIM_CPVAL = Nothing
    BDSIS_TBMIM.Close
    Set BDSIS_TBMIM = Nothing
    FechaTabela_MapaImpostos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MapaImpostos = False
End Function
Public Function FechaTabela_NotaFiscal(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBNTF_CPNNF = Nothing
    Set BDSIS_TBNTF_CPDEM = Nothing
    Set BDSIS_TBNTF_CPHEM = Nothing
    Set BDSIS_TBNTF_CPTIP = Nothing
    Set BDSIS_TBNTF_CPNOP = Nothing
    Set BDSIS_TBNTF_CPCFO = Nothing
    Set BDSIS_TBNTF_CPEMP = Nothing
    Set BDSIS_TBNTF_CPDSA = Nothing
    Set BDSIS_TBNTF_CPHSA = Nothing
    Set BDSIS_TBNTF_CPVAL = Nothing
    Set BDSIS_TBNTF_CPNDP = Nothing
    Set BDSIS_TBNTF_CPVEN = Nothing
    Set BDSIS_TBNTF_CPNPI = Nothing
    Set BDSIS_TBNTF_CPNSP = Nothing
    Set BDSIS_TBNTF_CPOPE = Nothing
    Set BDSIS_TBNTF_CPDEC = Nothing
    Set BDSIS_TBNTF_CPBAN = Nothing
    Set BDSIS_TBNTF_CPCOM = Nothing
    Set BDSIS_TBNTF_CPBCI = Nothing
    Set BDSIS_TBNTF_CPVIC = Nothing
    Set BDSIS_TBNTF_CPBIS = Nothing
    Set BDSIS_TBNTF_CPVIS = Nothing
    Set BDSIS_TBNTF_CPVTP = Nothing
    Set BDSIS_TBNTF_CPVFR = Nothing
    Set BDSIS_TBNTF_CPVSE = Nothing
    Set BDSIS_TBNTF_CPODA = Nothing
    Set BDSIS_TBNTF_CPVIP = Nothing
    Set BDSIS_TBNTF_CPVTN = Nothing
    Set BDSIS_TBNTF_CPTRA = Nothing
    Set BDSIS_TBNTF_CPFCO = Nothing
    Set BDSIS_TBNTF_CPPVE = Nothing
    Set BDSIS_TBNTF_CPEPV = Nothing
    Set BDSIS_TBNTF_CPQUA = Nothing
    Set BDSIS_TBNTF_CPESP = Nothing
    Set BDSIS_TBNTF_CPMAR = Nothing
    Set BDSIS_TBNTF_CPNVO = Nothing
    Set BDSIS_TBNTF_CPPBR = Nothing
    Set BDSIS_TBNTF_CPPLI = Nothing
    Set BDSIS_TBNTF_CPVIN = Nothing
    Set BDSIS_TBNTF_CPVEX = Nothing
    Set BDSIS_TBNTF_CPSET = Nothing
    Set BDSIS_TBNTF_CPCFI = Nothing
    Set BDSIS_TBNTF_CPCFJ = Nothing
    Set BDSIS_TBNTF_CPVLA = Nothing
    Set BDSIS_TBNTF_CPVCA = Nothing
    Set BDSIS_TBNTF_CPVLB = Nothing
    Set BDSIS_TBNTF_CPVCB = Nothing
    Set BDSIS_TBNTF_CPVLC = Nothing
    Set BDSIS_TBNTF_CPVCC = Nothing
    Set BDSIS_TBNTF_CPVLD = Nothing
    Set BDSIS_TBNTF_CPVCD = Nothing
    BDSIS_TBNTF.Close
    Set BDSIS_TBNTF = Nothing
    FechaTabela_NotaFiscal = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_NotaFiscal = False
End Function
Public Function FechaTabela_NotaFiscalProdutos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBNFP_CPNNF = Nothing
    Set BDSIS_TBNFP_CPFIG = Nothing
    Set BDSIS_TBNFP_CPBIT = Nothing
    Set BDSIS_TBNFP_CPMAT = Nothing
    Set BDSIS_TBNFP_CPDES = Nothing
    Set BDSIS_TBNFP_CPCCF = Nothing
    Set BDSIS_TBNFP_CPCST = Nothing
    Set BDSIS_TBNFP_CPUNI = Nothing
    Set BDSIS_TBNFP_CPQUA = Nothing
    Set BDSIS_TBNFP_CPPUN = Nothing
    Set BDSIS_TBNFP_CPPTO = Nothing
    Set BDSIS_TBNFP_CPBCI = Nothing
    Set BDSIS_TBNFP_CPAIC = Nothing
    Set BDSIS_TBNFP_CPVIC = Nothing
    Set BDSIS_TBNFP_CPAIP = Nothing
    Set BDSIS_TBNFP_CPVIP = Nothing
    Set BDSIS_TBNFP_CPPPA = Nothing
    BDSIS_TBNFP.Close
    Set BDSIS_TBNFP = Nothing
    FechaTabela_NotaFiscalProdutos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_NotaFiscalProdutos = False
End Function
Public Function FechaTabela_NotaFiscalDeclaracoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBNFD_CPNNF = Nothing
    Set BDSIS_TBNFD_CPDEC = Nothing
    Set BDSIS_TBNFD_CPLIN = Nothing
    BDSIS_TBNFD.Close
    Set BDSIS_TBNFD = Nothing
    FechaTabela_NotaFiscalDeclaracoes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_NotaFiscalDeclaracoes = False
End Function
Public Function FechaTabela_NotaFiscalBancos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBNFB_CPNNF = Nothing
    Set BDSIS_TBNFB_CPCON = Nothing
    Set BDSIS_TBNFB_CPLIN = Nothing
    BDSIS_TBNFB.Close
    Set BDSIS_TBNFB = Nothing
    FechaTabela_NotaFiscalBancos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_NotaFiscalBancos = False
End Function
Public Function FechaTabela_NotaFiscalComentarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBNFC_CPNNF = Nothing
    Set BDSIS_TBNFC_CPCOM = Nothing
    Set BDSIS_TBNFC_CPLIN = Nothing
    BDSIS_TBNFC.Close
    Set BDSIS_TBNFC = Nothing
    FechaTabela_NotaFiscalComentarios = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_NotaFiscalComentarios = False
End Function
Public Function FechaTabela_ConfiguracoesTelasImpressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBTIM_CPTEL = Nothing
    Set BDSIS_TBTIM_CPNOM = Nothing
    Set BDSIS_TBTIM_CPIMP = Nothing
    BDSIS_TBTIM.Close
    Set BDSIS_TBTIM = Nothing
    FechaTabela_ConfiguracoesTelasImpressoras = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_ConfiguracoesTelasImpressoras = False
End Function
Public Function FechaTabela_Impressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBIMP_CPIMP = Nothing
    Set BDSIS_TBIMP_CPNOM = Nothing
    Set BDSIS_TBIMP_CPDRI = Nothing
    BDSIS_TBIMP.Close
    Set BDSIS_TBIMP = Nothing
    FechaTabela_Impressoras = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Impressoras = False
End Function
Public Function FechaTabela_LogDeEventos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBGEV_CPDAT = Nothing
    Set BDSIS_TBGEV_CPHOR = Nothing
    Set BDSIS_TBGEV_CPUSU = Nothing
    Set BDSIS_TBGEV_CPCOM = Nothing
    Set BDSIS_TBGEV_CPEVE = Nothing
    Set BDSIS_TBGEV_CPVAL = Nothing
    BDSIS_TBGEV.Close
    Set BDSIS_TBGEV = Nothing
    FechaTabela_LogDeEventos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_LogDeEventos = False
End Function
Public Function FechaTabela_LogDeErros(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBGRR_CPDAT = Nothing
    Set BDSIS_TBGRR_CPHOR = Nothing
    Set BDSIS_TBGRR_CPUSU = Nothing
    Set BDSIS_TBGRR_CPCOM = Nothing
    Set BDSIS_TBGRR_CPAPL = Nothing
    Set BDSIS_TBGRR_CPENU = Nothing
    Set BDSIS_TBGRR_CPEDE = Nothing
    Set BDSIS_TBGRR_CPEOR = Nothing
    Set BDSIS_TBGRR_CPEJA = Nothing
    Set BDSIS_TBGRR_CPEJC = Nothing
    BDSIS_TBGRR.Close
    Set BDSIS_TBGRR = Nothing
    FechaTabela_LogDeErros = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_LogDeErros = False
End Function
Public Function FechaTabela_EntradaPecasIndice(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEPI_CPNEP = Nothing
    Set BDSIS_TBEPI_CPDAT = Nothing
    Set BDSIS_TBEPI_CPHOR = Nothing
    Set BDSIS_TBEPI_CPIND = Nothing
    BDSIS_TBEPI.Close
    Set BDSIS_TBEPI = Nothing
    FechaTabela_EntradaPecasIndice = True
    Exit Function
ERRO_ABERTURA:
    FechaTabela_EntradaPecasIndice = False
End Function
Public Function FechaTabela_EntradaPecas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEPE_CPIND = Nothing
    Set BDSIS_TBEPE_CPQUA = Nothing
    Set BDSIS_TBEPE_CPIFI = Nothing
    Set BDSIS_TBEPE_CPFIN = Nothing
    BDSIS_TBEPE.Close
    Set BDSIS_TBEPE = Nothing
    FechaTabela_EntradaPecas = True
    Exit Function
ERRO_ABERTURA:
    FechaTabela_EntradaPecas = False
End Function
Public Function FechaTabela_MateriaPrimaIndiceDePecas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMPP_CPINP = Nothing
    Set BDSIS_TBMPP_CPPEC = Nothing
    BDSIS_TBMPP.Close
    Set BDSIS_TBMPP = Nothing
    FechaTabela_MateriaPrimaIndiceDePecas = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MateriaPrimaIndiceDePecas = False
End Function
Public Function FechaTabela_MateriaPrimaIndiceDeBitolas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMPB_CPINB = Nothing
    Set BDSIS_TBMPB_CPBIT = Nothing
    BDSIS_TBMPB.Close
    Set BDSIS_TBMPB = Nothing
    FechaTabela_MateriaPrimaIndiceDeBitolas = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MateriaPrimaIndiceDeBitolas = False
End Function
Public Function FechaTabela_MateriaPrimaIndiceDeMateriais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMPM_CPINM = Nothing
    Set BDSIS_TBMPM_CPMAT = Nothing
    BDSIS_TBMPM.Close
    Set BDSIS_TBMPM = Nothing
    FechaTabela_MateriaPrimaIndiceDeMateriais = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MateriaPrimaIndiceDeMateriais = False
End Function
Public Function FechaTabela_MateriaPrimaIndiceDeQuantidades(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMPQ_CPINQ = Nothing
    Set BDSIS_TBMPQ_CPQUA = Nothing
    BDSIS_TBMPQ.Close
    Set BDSIS_TBMPM = Nothing
    FechaTabela_MateriaPrimaIndiceDeQuantidades = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MateriaPrimaIndiceDeQuantidades = False
End Function
Public Function FechaTabela_MateriaPrimaIndiceDeNomes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMPN_CPINN = Nothing
    Set BDSIS_TBMPN_CPNOM = Nothing
    BDSIS_TBMPN.Close
    Set BDSIS_TBMPN = Nothing
    FechaTabela_MateriaPrimaIndiceDeNomes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MateriaPrimaIndiceDeNomes = False
End Function
Public Function FechaTabela_MateriaPrimaRelacaoMateriais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMPR_CPFIG = Nothing
    Set BDSIS_TBMPR_CPMFG = Nothing
    Set BDSIS_TBMPR_CPMMP = Nothing
    BDSIS_TBMPR.Close
    Set BDSIS_TBMPR = Nothing
    FechaTabela_MateriaPrimaRelacaoMateriais = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MateriaPrimaRelacaoMateriais = False
End Function
Public Function FechaTabela_EstoqueComplementos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBECM_CPIND = Nothing
    Set BDSIS_TBECM_CPCOM = Nothing
    BDSIS_TBECM.Close
    Set BDSIS_TBECM = Nothing
    FechaTabela_EstoqueComplementos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EstoqueComplementos = False
End Function
Public Function FechaTabela_MapaCotacoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMCO_CPMEA = Nothing
    Set BDSIS_TBMCO_CPVAL = Nothing
    BDSIS_TBMCO.Close
    Set BDSIS_TBMCO = Nothing
    FechaTabela_MapaCotacoes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MapaCotacoes = False
End Function
Public Function FechaTabela_Pedidos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBPED_CPIND = Nothing
    Set BDSIS_TBPED_CPDAT = Nothing
    Set BDSIS_TBPED_CPHOR = Nothing
    Set BDSIS_TBPED_CPEMP = Nothing
    Set BDSIS_TBPED_CPCON = Nothing
    Set BDSIS_TBPED_CPCPG = Nothing
    Set BDSIS_TBPED_CPTRA = Nothing
    Set BDSIS_TBPED_CPNSP = Nothing
    Set BDSIS_TBPED_CPITE = Nothing
    Set BDSIS_TBPED_CPIVE = Nothing
    Set BDSIS_TBPED_CPLIQ = Nothing
    Set BDSIS_TBPED_CPABE = Nothing
    Set BDSIS_TBPED_CPIFR = Nothing
    Set BDSIS_TBPED_CPOUD = Nothing
    Set BDSIS_TBPED_CPDAD = Nothing
    Set BDSIS_TBPED_CPOBS = Nothing
    BDSIS_TBPED.Close
    Set BDSIS_TBPED = Nothing
    FechaTabela_Pedidos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Pedidos = False
End Function
Public Function FechaTabela_PedidosItens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBPIT_CPIND = Nothing
    Set BDSIS_TBPIT_CPNPE = Nothing
    Set BDSIS_TBPIT_CPINF = Nothing
    Set BDSIS_TBPIT_CPQUA = Nothing
    Set BDSIS_TBPIT_CPDES = Nothing
    Set BDSIS_TBPIT_CPCOM = Nothing
    Set BDSIS_TBPIT_CPPRE = Nothing
    Set BDSIS_TBPIT_CPPRA = Nothing
    Set BDSIS_TBPIT_CPLIQ = Nothing
    BDSIS_TBPIT.Close
    Set BDSIS_TBPIT = Nothing
    FechaTabela_PedidosItens = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_PedidosItens = False
End Function
Public Function FechaTabela_MapaPedidos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBMPE_CPMEA = Nothing
    Set BDSIS_TBMPE_CPVAL = Nothing
    BDSIS_TBMPE.Close
    Set BDSIS_TBMPE = Nothing
    FechaTabela_MapaPedidos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_MapaPedidos = False
End Function
Public Function FechaTabela_OrdemFabricacao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBODF_CPNOF = Nothing
    Set BDSIS_TBODF_CPDAE = Nothing
    Set BDSIS_TBODF_CPINF = Nothing
    Set BDSIS_TBODF_CPDES = Nothing
    Set BDSIS_TBODF_CPNDC = Nothing
    Set BDSIS_TBODF_CPQES = Nothing
    Set BDSIS_TBODF_CPQPR = Nothing
    Set BDSIS_TBODF_CPCOR = Nothing
    Set BDSIS_TBODF_CPDAL = Nothing
    Set BDSIS_TBODF_CPLIQ = Nothing
    BDSIS_TBODF.Close
    Set BDSIS_TBODF = Nothing
    FechaTabela_OrdemFabricacao = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_OrdemFabricacao = False
End Function
Public Function FechaTabela_OrdemFabricacaoProcessos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBOFP_CPIND = Nothing
    Set BDSIS_TBOFP_CPNOF = Nothing
    Set BDSIS_TBOFP_CPNET = Nothing
    Set BDSIS_TBOFP_CPNMA = Nothing
    Set BDSIS_TBOFP_CPDAI = Nothing
    Set BDSIS_TBOFP_CPDAF = Nothing
    Set BDSIS_TBOFP_CPNHE = Nothing
    BDSIS_TBOFP.Close
    Set BDSIS_TBOFP = Nothing
    FechaTabela_OrdemFabricacaoProcessos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_OrdemFabricacaoProcessos = False
End Function
Public Function FechaTabela_OrdemMontagem(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBODM_CPNOM = Nothing
    Set BDSIS_TBODM_CPDAT = Nothing
    Set BDSIS_TBODM_CPHOR = Nothing
    Set BDSIS_TBODM_CPINF = Nothing
    Set BDSIS_TBODM_CPDES = Nothing
    Set BDSIS_TBODM_CPNDC = Nothing
    Set BDSIS_TBODM_CPQES = Nothing
    Set BDSIS_TBODM_CPQMO = Nothing
    Set BDSIS_TBODM_CPDIN = Nothing
    Set BDSIS_TBODM_CPHIN = Nothing
    Set BDSIS_TBODM_CPDFI = Nothing
    Set BDSIS_TBODM_CPHFI = Nothing
    Set BDSIS_TBODM_CPNCA = Nothing
    Set BDSIS_TBODM_CPNCB = Nothing
    Set BDSIS_TBODM_CPNCC = Nothing
    Set BDSIS_TBODM_CPNCD = Nothing
    Set BDSIS_TBODM_CPNCE = Nothing
    Set BDSIS_TBODM_CPLIQ = Nothing
    BDSIS_TBODM.Close
    Set BDSIS_TBODM = Nothing
    FechaTabela_OrdemMontagem = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_OrdemMontagem = False
End Function
Public Function FechaTabela_OrdemExpedicao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBODE_CPNOE = Nothing
    Set BDSIS_TBODE_CPDAT = Nothing
    Set BDSIS_TBODE_CPHOR = Nothing
    Set BDSIS_TBODE_CPEMP = Nothing
    Set BDSIS_TBODE_CPNPE = Nothing
    Set BDSIS_TBODE_CPCFS = Nothing
    Set BDSIS_TBODE_CPCFC = Nothing
    Set BDSIS_TBODE_CPCFV = Nothing
    Set BDSIS_TBODE_CPETP = Nothing
    Set BDSIS_TBODE_CPENU = Nothing
    Set BDSIS_TBODE_CPPBR = Nothing
    Set BDSIS_TBODE_CPPLI = Nothing
    Set BDSIS_TBODE_CPTRA = Nothing
    Set BDSIS_TBODE_CPCOR = Nothing
    Set BDSIS_TBODE_CPOBS = Nothing
    Set BDSIS_TBODE_CPLIQ = Nothing
    BDSIS_TBODE.Close
    Set BDSIS_TBODE = Nothing
    FechaTabela_OrdemExpedicao = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_OrdemExpedicao = False
End Function
Public Function FechaTabela_Funcionarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBFUN_CPIND = Nothing
    Set BDSIS_TBFUN_CPFUN = Nothing
    Set BDSIS_TBFUN_CPREG = Nothing
    Set BDSIS_TBFUN_CPTIP = Nothing
    Set BDSIS_TBFUN_CPSET = Nothing
    BDSIS_TBFUN.Close
    Set BDSIS_TBFUN = Nothing
    FechaTabela_Funcionarios = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Funcionarios = False
End Function
Public Function FechaTabela_ContasPagarReceber(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCPR_CPIND = Nothing
    Set BDSIS_TBCPR_CPEMI = Nothing
    Set BDSIS_TBCPR_CPVEN = Nothing
    Set BDSIS_TBCPR_CPMOV = Nothing
    Set BDSIS_TBCPR_CPVAL = Nothing
    Set BDSIS_TBCPR_CPORI = Nothing
    Set BDSIS_TBCPR_CPNDO = Nothing
    Set BDSIS_TBCPR_CPNNU = Nothing
    Set BDSIS_TBCPR_CPSNU = Nothing
    Set BDSIS_TBCPR_CPTIP = Nothing
    Set BDSIS_TBCPR_CPBAN = Nothing
    Set BDSIS_TBCPR_CPCAR = Nothing
    Set BDSIS_TBCPR_CPCOB = Nothing
    Set BDSIS_TBCPR_CPOBS = Nothing
    BDSIS_TBCPR.Close
    Set BDSIS_TBCPR = Nothing
    FechaTabela_ContasPagarReceber = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_ContasPagarReceber = False
End Function
Public Function FechaTabela_Imagens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBIMA_CPIND = Nothing
    Set BDSIS_TBIMA_CPNOM = Nothing
    Set BDSIS_TBIMA_CPTIP = Nothing
    Set BDSIS_TBIMA_CPIMA = Nothing
    BDSIS_TBCPR.Close
    Set BDSIS_TBCPR = Nothing
    FechaTabela_Imagens = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Imagens = False
End Function
Public Function FechaTabela_Equipamentos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEQU_CPIND = Nothing
    Set BDSIS_TBEQU_CPEQU = Nothing
    Set BDSIS_TBEQU_CPNOM = Nothing
    Set BDSIS_TBEQU_CPMOD = Nothing
    Set BDSIS_TBEQU_CPFAB = Nothing
    Set BDSIS_TBEQU_CPANO = Nothing
    Set BDSIS_TBEQU_CPVAL = Nothing
    Set BDSIS_TBEQU_CPFOT = Nothing
    Set BDSIS_TBEQU_CPCHO = Nothing
    Set BDSIS_TBEQU_CPPRO = Nothing
    Set BDSIS_TBEQU_CPFUN = Nothing
    BDSIS_TBEQU.Close
    Set BDSIS_TBEQU = Nothing
    FechaTabela_Equipamentos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Equipamentos = False
End Function
Public Function FechaTabela_EquipamentosManutencao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBEMA_CPIND = Nothing
    Set BDSIS_TBEMA_CPEQU = Nothing
    Set BDSIS_TBEMA_CPMAN = Nothing
    Set BDSIS_TBEMA_CPDUL = Nothing
    Set BDSIS_TBEMA_CPDPR = Nothing
    Set BDSIS_TBEMA_CPVAL = Nothing
    Set BDSIS_TBEMA_CPEMP = Nothing
    Set BDSIS_TBEMA_CPCON = Nothing
    Set BDSIS_TBEMA_CPTIP = Nothing
    Set BDSIS_TBEMA_CPAVI = Nothing
    Set BDSIS_TBEMA_CPUSU = Nothing
    BDSIS_TBEMA.Close
    Set BDSIS_TBEMA = Nothing
    FechaTabela_EquipamentosManutencao = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_EquipamentosManutencao = False
End Function
Public Function FechaTabela_Avisos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBAVI_CPIND = Nothing
    Set BDSIS_TBAVI_CPREM = Nothing
    Set BDSIS_TBAVI_CPEMI = Nothing
    Set BDSIS_TBAVI_CPDES = Nothing
    Set BDSIS_TBAVI_CPVEN = Nothing
    Set BDSIS_TBAVI_CPTIP = Nothing
    Set BDSIS_TBAVI_CPAVI = Nothing
    BDSIS_TBAVI.Close
    Set BDSIS_TBAVI = Nothing
    FechaTabela_Avisos = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Avisos = False
End Function
Public Function FechaTabela_CNCFuncoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCFU_CPIND = Nothing
    Set BDSIS_TBCFU_CPCNC = Nothing
    Set BDSIS_TBCFU_CPFUN = Nothing
    Set BDSIS_TBCFU_CPAPL = Nothing
    Set BDSIS_TBCFU_CPVAR = Nothing
    BDSIS_TBCFU.Close
    Set BDSIS_TBCFU = Nothing
    FechaTabela_CNCFuncoes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_CNCFuncoes = False
End Function
Public Function FechaTabela_CNCMaquinas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCMA_CPIND = Nothing
    Set BDSIS_TBCMA_CPNNE = Nothing
    BDSIS_TBCMA.Close
    Set BDSIS_TBCMA = Nothing
    FechaTabela_CNCMaquinas = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_CNCMaquinas = False
End Function
Public Function FechaTabela_CNCProgramas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBCNC_CPIND = Nothing
    Set BDSIS_TBCNC_CPNOM = Nothing
    Set BDSIS_TBCNC_CPCNC = Nothing
    Set BDSIS_TBCNC_CPDAT = Nothing
    Set BDSIS_TBCNC_CPULT = Nothing
    Set BDSIS_TBCNC_CPULT = Nothing
    Set BDSIS_TBCNC_CPTIP = Nothing
    Set BDSIS_TBCNC_CPMAQ = Nothing
    Set BDSIS_TBCNC_CPPRO = Nothing
    BDSIS_TBCNC.Close
    Set BDSIS_TBCNC = Nothing
    FechaTabela_CNCProgramas = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_CNCProgramas = False
End Function
Public Function FechaTabela_Layouts(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBLAY_CPIND = Nothing
    Set BDSIS_TBLAY_CPLAY = Nothing
    Set BDSIS_TBLAY_CPMAQ = Nothing
    Set BDSIS_TBLAY_CPNEL = Nothing
    Set BDSIS_TBLAY_CPNDL = Nothing
    Set BDSIS_TBLAY_CPDCR = Nothing
    Set BDSIS_TBLAY_CPCOM = Nothing
    Set BDSIS_TBLAY_CPNRE = Nothing
    Set BDSIS_TBLAY_CPDRE = Nothing
    Set BDSIS_TBLAY_CPCFU = Nothing
    Set BDSIS_TBLAY_CPFOT = Nothing
    BDSIS_TBLAY.Close
    Set BDSIS_TBLAY = Nothing
    FechaTabela_Layouts = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_Layouts = False
End Function
Public Function FechaTabela_LayoutsEsquemas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBLES_CPIND = Nothing
    Set BDSIS_TBLES_CPESQ = Nothing
    Set BDSIS_TBLES_CPMAQ = Nothing
    Set BDSIS_TBLES_CPVAR = Nothing
    Set BDSIS_TBLES_CPDIS = Nothing
    Set BDSIS_TBLES_CPFER = Nothing
    BDSIS_TBLES.Close
    Set BDSIS_TBLES = Nothing
    FechaTabela_LayoutsEsquemas = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_LayoutsEsquemas = False
End Function
Public Function FechaTabela_LayoutsDimensoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    Set BDSIS_TBLDI_CPIND = Nothing
    Set BDSIS_TBLDI_CPNEL = Nothing
    Set BDSIS_TBLDI_CPVAR = Nothing
    Set BDSIS_TBLDI_CPDIS = Nothing
    Set BDSIS_TBLDI_CPFER = Nothing
    Set BDSIS_TBLDI_CPOBS = Nothing
    BDSIS_TBLDI.Close
    Set BDSIS_TBLDI = Nothing
    FechaTabela_LayoutsDimensoes = True
    Exit Function
ERRO_FINALIZACAO:
    FechaTabela_LayoutsDimensoes = False
End Function



'*****************************************************
'        F U N Ç Õ E S     A B R E     C A M P O S
'*****************************************************
Public Function AbreCampos_Usuarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBUSU_CPUSU = BDSIS_TBUSU.Fields("Usuário")
    Set BDSIS_TBUSU_CPADM = BDSIS_TBUSU.Fields("Administrador")
    Set BDSIS_TBUSU_CPNOM = BDSIS_TBUSU.Fields("Nome")
    Set BDSIS_TBUSU_CPSEN = BDSIS_TBUSU.Fields("Senha")
    Set BDSIS_TBUSU_CPPER = BDSIS_TBUSU.Fields("Permissões")
    AbreCampos_Usuarios = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Usuarios = False
End Function
Public Function AbreCampos_UsuariosComputadores(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBUCO_CPCOM = BDSIS_TBUCO.Fields("Computador")
    Set BDSIS_TBUCO_CPDWI = BDSIS_TBUCO.Fields("DirWin")
    Set BDSIS_TBUCO_CPDSY = BDSIS_TBUCO.Fields("DirSys")
    Set BDSIS_TBUCO_CPDTP = BDSIS_TBUCO.Fields("DirTem")
    Set BDSIS_TBUCO_CPOBU = BDSIS_TBUCO.Fields("OSBuild")
    Set BDSIS_TBUCO_CPOCS = BDSIS_TBUCO.Fields("OSCSD")
    Set BDSIS_TBUCO_CPOVE = BDSIS_TBUCO.Fields("OSVersao")
    Set BDSIS_TBUCO_CPOPL = BDSIS_TBUCO.Fields("OSPlat")
    Set BDSIS_TBUCO_CPSMP = BDSIS_TBUCO.Fields("SisMascProc")
    Set BDSIS_TBUCO_CPSOE = BDSIS_TBUCO.Fields("SisOEM")
    Set BDSIS_TBUCO_CPSTP = BDSIS_TBUCO.Fields("SisTipProc")
    Set BDSIS_TBUCO_CPRAM = BDSIS_TBUCO.Fields("RAM")
    Set BDSIS_TBUCO_CPCTA = BDSIS_TBUCO.Fields("TamC")
    Set BDSIS_TBUCO_CPCSN = BDSIS_TBUCO.Fields("SNC")
    AbreCampos_UsuariosComputadores = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_UsuariosComputadores = False
End Function
Public Function AbreCampos_UsuariosMenus(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDSIS_TBUME_CPIND = BDSIS_TBUME.Fields("Índice")
    Set BDSIS_TBUME_CPMEN = BDSIS_TBUME.Fields("Menu")
    'Set BDSIS_TBUME_CPPRI = BDSIS_TBUME.Fields("Primeiro")
    AbreCampos_UsuariosMenus = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_UsuariosMenus = False
End Function
Public Function AbreCampos_Grupos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBGRU_CPGRU = BDSIS_TBGRU.Fields("Grupo")
    Set BDSIS_TBGRU_CPDES = BDSIS_TBGRU.Fields("Descrição")
    Set BDSIS_TBGRU_CPVAL = BDSIS_TBGRU.Fields("Valor")
    Set BDSIS_TBGRU_CPTIP = BDSIS_TBGRU.Fields("Tipo")
    AbreCampos_Grupos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Grupos = False
End Function
Public Function AbreCampos_Estoque(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEST_CPFIC = BDSIS_TBEST.Fields("Índice de Ficha")
    Set BDSIS_TBEST_CPFIG = BDSIS_TBEST.Fields("Figura")
    Set BDSIS_TBEST_CPBIT = BDSIS_TBEST.Fields("Bitola")
    Set BDSIS_TBEST_CPMAT = BDSIS_TBEST.Fields("Material")
    Set BDSIS_TBEST_CPVUN = BDSIS_TBEST.Fields("Valor Unitário")
    Set BDSIS_TBEST_CPVMI = BDSIS_TBEST.Fields("Valor Mínimo")
    Set BDSIS_TBEST_CPVCU = BDSIS_TBEST.Fields("Valor Custo")
    Set BDSIS_TBEST_CPPUN = BDSIS_TBEST.Fields("Peso Unitário")
    Set BDSIS_TBEST_CPEST = BDSIS_TBEST.Fields("Estoque")
    Set BDSIS_TBEST_CPETM = BDSIS_TBEST.Fields("Estoque Mínimo")
    Set BDSIS_TBEST_CPVEN = BDSIS_TBEST.Fields("Vendidas")
    Set BDSIS_TBEST_CPCOT = BDSIS_TBEST.Fields("Cotadas")
    Set BDSIS_TBEST_CPINQ = BDSIS_TBEST.Fields("Índice MP de Quantidades")
    Set BDSIS_TBEST_CPINP = BDSIS_TBEST.Fields("Índice MP de Peças")
    Set BDSIS_TBEST_CPINN = BDSIS_TBEST.Fields("Índice MP de Nomes")
    Set BDSIS_TBEST_CPINB = BDSIS_TBEST.Fields("Índice MP de Bitolas")
    Set BDSIS_TBEST_CPINM = BDSIS_TBEST.Fields("Índice MP de Materiais")
    Set BDSIS_TBEST_CPNDC = BDSIS_TBEST.Fields("DC")
    AbreCampos_Estoque = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Estoque = False
End Function
Public Function AbreCampos_EstoqueIndice(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEID_CPIFI = BDSIS_TBEID.Fields("Índice de Figura")
    Set BDSIS_TBEID_CPDRE = BDSIS_TBEID.Fields("DesRed")
    Set BDSIS_TBEID_CPTRE = BDSIS_TBEID.Fields("TipRed")
    Set BDSIS_TBEID_CPDNO = BDSIS_TBEID.Fields("DesNor")
    Set BDSIS_TBEID_CPTNO = BDSIS_TBEID.Fields("TipNor")
    Set BDSIS_TBEID_CPDCO = BDSIS_TBEID.Fields("DesCom")
    Set BDSIS_TBEID_CPTCO = BDSIS_TBEID.Fields("TipCom")
    Set BDSIS_TBEID_CPGPE = BDSIS_TBEID.Fields("Grupo Peça")
    Set BDSIS_TBEID_CPGMP = BDSIS_TBEID.Fields("Grupo Matéria-Prima")
    Set BDSIS_TBEID_CPGBI = BDSIS_TBEID.Fields("Grupo Bitola")
    Set BDSIS_TBEID_CPGIN = BDSIS_TBEID.Fields("Grupo Internos")
    Set BDSIS_TBEID_CPGEX = BDSIS_TBEID.Fields("Grupo Extremidades")
    Set BDSIS_TBEID_CPGCL = BDSIS_TBEID.Fields("Grupo Classes")
    Set BDSIS_TBEID_CPGMA = BDSIS_TBEID.Fields("Grupo Materiais")
    AbreCampos_EstoqueIndice = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EstoqueIndice = False
End Function
Public Function AbreCampos_EstoqueFiguras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEFG_CPFIG = BDSIS_TBEFG.Fields("Figura")
    Set BDSIS_TBEFG_CPIFG = BDSIS_TBEFG.Fields("Índice de Figura")
    Set BDSIS_TBEFG_CPGIN = BDSIS_TBEFG.Fields("Grupo Internos")
    Set BDSIS_TBEFG_CPGEX = BDSIS_TBEFG.Fields("Grupo Extremidade")
    Set BDSIS_TBEFG_CPGCL = BDSIS_TBEFG.Fields("Grupo Classe")
    Set BDSIS_TBEFG_CPCOM = BDSIS_TBEFG.Fields("Complemento")
    AbreCampos_EstoqueFiguras = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EstoqueFiguras = False
End Function
Public Function AbreCampos_Empresas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEMP_CPAPE = BDSIS_TBEMP.Fields("Apelido")
    Set BDSIS_TBEMP_CPEMP = BDSIS_TBEMP.Fields("Empresa")
    Set BDSIS_TBEMP_CPCGC = BDSIS_TBEMP.Fields("CGC")
    Set BDSIS_TBEMP_CPINE = BDSIS_TBEMP.Fields("Inscrição Estadual")
    Set BDSIS_TBEMP_CPEND = BDSIS_TBEMP.Fields("Endereço")
    Set BDSIS_TBEMP_CPBAI = BDSIS_TBEMP.Fields("Bairro")
    Set BDSIS_TBEMP_CPPRA = BDSIS_TBEMP.Fields("Praça de Pagamento")
    Set BDSIS_TBEMP_CPCID = BDSIS_TBEMP.Fields("Cidade")
    Set BDSIS_TBEMP_CPEST = BDSIS_TBEMP.Fields("Estado")
    Set BDSIS_TBEMP_CPCEP = BDSIS_TBEMP.Fields("CEP")
    Set BDSIS_TBEMP_CPFON = BDSIS_TBEMP.Fields("Fone")
    Set BDSIS_TBEMP_CPFAX = BDSIS_TBEMP.Fields("Fax")
    Set BDSIS_TBEMP_CPTIP = BDSIS_TBEMP.Fields("Tipo")
    Set BDSIS_TBEMP_CPCOM = BDSIS_TBEMP.Fields("Comentário de Nota Fiscal")
    Set BDSIS_TBEMP_CPOBS = BDSIS_TBEMP.Fields("Observações")
    Set BDSIS_TBEMP_CPTRA = BDSIS_TBEMP.Fields("Transportadora")
    Set BDSIS_TBEMP_CPEMA = BDSIS_TBEMP.Fields("E-mail")
    Set BDSIS_TBEMP_CPSIT = BDSIS_TBEMP.Fields("Site")
    AbreCampos_Empresas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Empresas = False
End Function
Public Function AbreCampos_EmpresasContatos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBECO_CPEMP = BDSIS_TBECO.Fields("Empresa")
    Set BDSIS_TBECO_CPCON = BDSIS_TBECO.Fields("Contato")
    AbreCampos_EmpresasContatos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EmpresasContatos = False
End Function
Public Function AbreCampos_EstoqueAliquotas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEAL_CPALI = BDSIS_TBEAL.Fields("Alíquota")
    Set BDSIS_TBEAL_CPEST = BDSIS_TBEAL.Fields("Estado")
    Set BDSIS_TBEAL_CPCCF = BDSIS_TBEAL.Fields("CF")
    Set BDSIS_TBEAL_CPPOR = BDSIS_TBEAL.Fields("Porcentagem")
    Set BDSIS_TBEAL_CPRBC = BDSIS_TBEAL.Fields("RBC")
    AbreCampos_EstoqueAliquotas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EstoqueAliquotas = False
End Function
Public Function AbreCampos_Bancos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBBAN_CPNMC = BDSIS_TBBAN.Fields("Nome da Conta")
    Set BDSIS_TBBAN_CPNMB = BDSIS_TBBAN.Fields("Nome do Banco")
    Set BDSIS_TBBAN_CPBAI = BDSIS_TBBAN.Fields("Bairro")
    Set BDSIS_TBBAN_CPAGE = BDSIS_TBBAN.Fields("Agência")
    Set BDSIS_TBBAN_CPCON = BDSIS_TBBAN.Fields("Conta Corrente")
    AbreCampos_Bancos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Bancos = False
End Function
Public Function AbreCampos_CodigosFiscais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCDF_CPNRG = BDSIS_TBCDF.Fields("Número do Registro")
    Set BDSIS_TBCDF_CPNTO = BDSIS_TBCDF.Fields("Natureza da Operação")
    Set BDSIS_TBCDF_CPCFO = BDSIS_TBCDF.Fields("CFOP")
    Set BDSIS_TBCDF_CPTIP = BDSIS_TBCDF.Fields("Tipo")
    AbreCampos_CodigosFiscais = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_CodigosFiscais = False
End Function
Public Function AbreCampos_EstoqueCFeST(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCFS_CPFIG = BDSIS_TBCFS.Fields("Figura")
    Set BDSIS_TBCFS_CPGMT = BDSIS_TBCFS.Fields("Grupo Material")
    Set BDSIS_TBCFS_CPGCF = BDSIS_TBCFS.Fields("Grupo CF")
    Set BDSIS_TBCFS_CPGST = BDSIS_TBCFS.Fields("Grupo ST")
    AbreCampos_EstoqueCFeST = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EstoqueCFeST = False
End Function
Public Function AbreCampos_EstoqueDeclaracoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEDC_CPCCF = BDSIS_TBEDC.Fields("CF")
    Set BDSIS_TBEDC_CPDSP = BDSIS_TBEDC.Fields("DecSP")
    Set BDSIS_TBEDC_CPDOU = BDSIS_TBEDC.Fields("DecOutros")
    AbreCampos_EstoqueDeclaracoes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EstoqueDeclaracoes = False
End Function
Public Function AbreCampos_ConfiguracoesNotaFiscal(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCNF_CPITE = BDSIS_TBCNF.Fields("Item")
    Set BDSIS_TBCNF_CPDES = BDSIS_TBCNF.Fields("Descrição")
    Set BDSIS_TBCNF_CPEXE = BDSIS_TBCNF.Fields("Exemplo")
    Set BDSIS_TBCNF_CPIND = BDSIS_TBCNF.Fields("Indice")
    Set BDSIS_TBCNF_CPCOL = BDSIS_TBCNF.Fields("Col")
    Set BDSIS_TBCNF_CPLIN = BDSIS_TBCNF.Fields("Lin")
    Set BDSIS_TBCNF_CPFON = BDSIS_TBCNF.Fields("Fonte")
    Set BDSIS_TBCNF_CPTAM = BDSIS_TBCNF.Fields("Tamanho")
    Set BDSIS_TBCNF_CPNEG = BDSIS_TBCNF.Fields("Negrito")
    Set BDSIS_TBCNF_CPITA = BDSIS_TBCNF.Fields("Italico")
    Set BDSIS_TBCNF_CPSUB = BDSIS_TBCNF.Fields("Sublinhado")
    AbreCampos_ConfiguracoesNotaFiscal = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_ConfiguracoesNotaFiscal = False
End Function
Public Function AbreCampos_ConfiguracoesNotaFiscalImagem(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBINF_CPIMG = BDSIS_TBINF.Fields("Imagem")
    AbreCampos_ConfiguracoesNotaFiscalImagem = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_ConfiguracoesNotaFiscalImagem = False
End Function
Public Function AbreCampos_Cotacoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCOT_CPIND = BDSIS_TBCOT.Fields("Índice de Cotação")
    Set BDSIS_TBCOT_CPDAT = BDSIS_TBCOT.Fields("Data Emissão")
    Set BDSIS_TBCOT_CPHOR = BDSIS_TBCOT.Fields("Hora")
    Set BDSIS_TBCOT_CPEMP = BDSIS_TBCOT.Fields("Empresa")
    Set BDSIS_TBCOT_CPCON = BDSIS_TBCOT.Fields("Contato")
    Set BDSIS_TBCOT_CPCPG = BDSIS_TBCOT.Fields("Condição de Pagamento")
    Set BDSIS_TBCOT_CPTRA = BDSIS_TBCOT.Fields("Transportadora")
    Set BDSIS_TBCOT_CPITE = BDSIS_TBCOT.Fields("Índice dos Ítens")
    Set BDSIS_TBCOT_CPVAL = BDSIS_TBCOT.Fields("Valor")
    Set BDSIS_TBCOT_CPIVE = BDSIS_TBCOT.Fields("Índice Vendedor")
    Set BDSIS_TBCOT_CPLIQ = BDSIS_TBCOT.Fields("Liquidado")
    AbreCampos_Cotacoes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Cotacoes = False
End Function
Public Function AbreCampos_CotacoesItens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCTI_CPIND = BDSIS_TBCTI.Fields("Índice do Ítem")
    Set BDSIS_TBCTI_CPNCO = BDSIS_TBCTI.Fields("Número da Cotação")
    Set BDSIS_TBCTI_CPINF = BDSIS_TBCTI.Fields("Índice de Ficha")
    Set BDSIS_TBCTI_CPQUA = BDSIS_TBCTI.Fields("Quantidade")
    Set BDSIS_TBCTI_CPDES = BDSIS_TBCTI.Fields("Descrição")
    Set BDSIS_TBCTI_CPCOM = BDSIS_TBCTI.Fields("Complemento")
    Set BDSIS_TBCTI_CPPRE = BDSIS_TBCTI.Fields("Preço Unitário")
    Set BDSIS_TBCTI_CPPRA = BDSIS_TBCTI.Fields("Prazo de Entrega")
    Set BDSIS_TBCTI_CPLIQ = BDSIS_TBCTI.Fields("Liquidado")
    AbreCampos_CotacoesItens = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_CotacoesItens = False
End Function
Public Function AbreCampos_MapaFaturamentoMes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMFA_CPNNF = BDSIS_TBMFA.Fields("Número da Nota Fiscal")
    Set BDSIS_TBMFA_CPDAT = BDSIS_TBMFA.Fields("Data")
    Set BDSIS_TBMFA_CPVAL = BDSIS_TBMFA.Fields("Valor")
    Set BDSIS_TBMFA_CPEMP = BDSIS_TBMFA.Fields("Empresa")
    Set BDSIS_TBMFA_CPDEV = BDSIS_TBMFA.Fields("Devolução")
    AbreCampos_MapaFaturamentoMes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MapaFaturamentoMes = False
End Function
Public Function AbreCampos_MapaContasReceber(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMCR_CPDEM = BDSIS_TBMCR.Fields("Data Emissão")
    Set BDSIS_TBMCR_CPEMP = BDSIS_TBMCR.Fields("Empresa")
    Set BDSIS_TBMCR_CPNDP = BDSIS_TBMCR.Fields("Número Duplicata")
    Set BDSIS_TBMCR_CPVAL = BDSIS_TBMCR.Fields("Valor")
    Set BDSIS_TBMCR_CPDVE = BDSIS_TBMCR.Fields("Data Vencimento")
    AbreCampos_MapaContasReceber = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MapaContasReceber = False
End Function
Public Function AbreCampos_MapaImpostos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMIM_CPIMP = BDSIS_TBMIM.Fields("Imposto")
    Set BDSIS_TBMIM_CPMES = BDSIS_TBMIM.Fields("Mês")
    Set BDSIS_TBMIM_CPANO = BDSIS_TBMIM.Fields("Ano")
    Set BDSIS_TBMIM_CPVAL = BDSIS_TBMIM.Fields("Valor")
    AbreCampos_MapaImpostos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MapaImpostos = False
End Function
Public Function AbreCampos_NotaFiscal(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNTF_CPNNF = BDSIS_TBNTF.Fields("Número da Nota Fiscal")
    Set BDSIS_TBNTF_CPDEM = BDSIS_TBNTF.Fields("Data Emissão")
    Set BDSIS_TBNTF_CPHEM = BDSIS_TBNTF.Fields("Hora Emissão")
    Set BDSIS_TBNTF_CPTIP = BDSIS_TBNTF.Fields("Tipo")
    Set BDSIS_TBNTF_CPNOP = BDSIS_TBNTF.Fields("Natureza da Operação")
    Set BDSIS_TBNTF_CPCFO = BDSIS_TBNTF.Fields("CFPO")
    Set BDSIS_TBNTF_CPEMP = BDSIS_TBNTF.Fields("Empresa")
    Set BDSIS_TBNTF_CPDSA = BDSIS_TBNTF.Fields("Data Saída")
    Set BDSIS_TBNTF_CPHSA = BDSIS_TBNTF.Fields("Hora Saída")
    Set BDSIS_TBNTF_CPVAL = BDSIS_TBNTF.Fields("Valor")
    Set BDSIS_TBNTF_CPNDP = BDSIS_TBNTF.Fields("Número Duplicata")
    Set BDSIS_TBNTF_CPVEN = BDSIS_TBNTF.Fields("Vencimento")
    Set BDSIS_TBNTF_CPNPI = BDSIS_TBNTF.Fields("Número Pedido Interno")
    Set BDSIS_TBNTF_CPNSP = BDSIS_TBNTF.Fields("Número Seu Pedido")
    Set BDSIS_TBNTF_CPOPE = BDSIS_TBNTF.Fields("Operação")
    Set BDSIS_TBNTF_CPDEC = BDSIS_TBNTF.Fields("Declarações")
    Set BDSIS_TBNTF_CPBAN = BDSIS_TBNTF.Fields("Banco")
    Set BDSIS_TBNTF_CPCOM = BDSIS_TBNTF.Fields("Comentários")
    Set BDSIS_TBNTF_CPBCI = BDSIS_TBNTF.Fields("Base Cálculo ICMS")
    Set BDSIS_TBNTF_CPVIC = BDSIS_TBNTF.Fields("Valor ICMS")
    Set BDSIS_TBNTF_CPBIS = BDSIS_TBNTF.Fields("Base Cálculo ICMS Substituição")
    Set BDSIS_TBNTF_CPVIS = BDSIS_TBNTF.Fields("Valor ICMS Substituição")
    Set BDSIS_TBNTF_CPVTP = BDSIS_TBNTF.Fields("Valor Total dos Produtos")
    Set BDSIS_TBNTF_CPVFR = BDSIS_TBNTF.Fields("Valor do Frete")
    Set BDSIS_TBNTF_CPVSE = BDSIS_TBNTF.Fields("Valor do Seguro")
    Set BDSIS_TBNTF_CPODA = BDSIS_TBNTF.Fields("Outras Despesas Acessórias")
    Set BDSIS_TBNTF_CPVIP = BDSIS_TBNTF.Fields("Valor Total do IPI")
    Set BDSIS_TBNTF_CPVTN = BDSIS_TBNTF.Fields("Valor Total da Nota")
    Set BDSIS_TBNTF_CPTRA = BDSIS_TBNTF.Fields("Transportadora")
    Set BDSIS_TBNTF_CPFCO = BDSIS_TBNTF.Fields("Frete por Conta")
    Set BDSIS_TBNTF_CPPVE = BDSIS_TBNTF.Fields("Placa Veículo")
    Set BDSIS_TBNTF_CPEPV = BDSIS_TBNTF.Fields("Estado Placa Veículo")
    Set BDSIS_TBNTF_CPQUA = BDSIS_TBNTF.Fields("Quantidade")
    Set BDSIS_TBNTF_CPESP = BDSIS_TBNTF.Fields("Espécie")
    Set BDSIS_TBNTF_CPMAR = BDSIS_TBNTF.Fields("Marca")
    Set BDSIS_TBNTF_CPNVO = BDSIS_TBNTF.Fields("Número")
    Set BDSIS_TBNTF_CPPBR = BDSIS_TBNTF.Fields("Peso Bruto")
    Set BDSIS_TBNTF_CPPLI = BDSIS_TBNTF.Fields("Peso Líquido")
    Set BDSIS_TBNTF_CPVIN = BDSIS_TBNTF.Fields("Vendedor Interno")
    Set BDSIS_TBNTF_CPVEX = BDSIS_TBNTF.Fields("Vendedor Externo")
    Set BDSIS_TBNTF_CPSET = BDSIS_TBNTF.Fields("Setor")
    Set BDSIS_TBNTF_CPCFI = BDSIS_TBNTF.Fields("CF I")
    Set BDSIS_TBNTF_CPCFJ = BDSIS_TBNTF.Fields("CF J")
    Set BDSIS_TBNTF_CPVLA = BDSIS_TBNTF.Fields("Valor A")
    Set BDSIS_TBNTF_CPVCA = BDSIS_TBNTF.Fields("Vencimento A")
    Set BDSIS_TBNTF_CPVLB = BDSIS_TBNTF.Fields("Valor B")
    Set BDSIS_TBNTF_CPVCB = BDSIS_TBNTF.Fields("Vencimento B")
    Set BDSIS_TBNTF_CPVLC = BDSIS_TBNTF.Fields("Valor C")
    Set BDSIS_TBNTF_CPVCC = BDSIS_TBNTF.Fields("Vencimento C")
    Set BDSIS_TBNTF_CPVLD = BDSIS_TBNTF.Fields("Valor D")
    Set BDSIS_TBNTF_CPVCD = BDSIS_TBNTF.Fields("Vencimento D")
    AbreCampos_NotaFiscal = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_NotaFiscal = False
End Function
Public Function AbreCampos_NotaFiscalProdutos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFP_CPNNF = BDSIS_TBNFP.Fields("Número da Nota Fiscal")
    Set BDSIS_TBNFP_CPFIG = BDSIS_TBNFP.Fields("Figura")
    Set BDSIS_TBNFP_CPBIT = BDSIS_TBNFP.Fields("Bitola")
    Set BDSIS_TBNFP_CPMAT = BDSIS_TBNFP.Fields("Material")
    Set BDSIS_TBNFP_CPDES = BDSIS_TBNFP.Fields("Descrição")
    Set BDSIS_TBNFP_CPCCF = BDSIS_TBNFP.Fields("CF")
    Set BDSIS_TBNFP_CPCST = BDSIS_TBNFP.Fields("ST")
    Set BDSIS_TBNFP_CPUNI = BDSIS_TBNFP.Fields("Unidade")
    Set BDSIS_TBNFP_CPQUA = BDSIS_TBNFP.Fields("Quantidade")
    Set BDSIS_TBNFP_CPPUN = BDSIS_TBNFP.Fields("Preço Unitário")
    Set BDSIS_TBNFP_CPPTO = BDSIS_TBNFP.Fields("Preço Total")
    Set BDSIS_TBNFP_CPBCI = BDSIS_TBNFP.Fields("Base Cálculo ICMS")
    Set BDSIS_TBNFP_CPAIC = BDSIS_TBNFP.Fields("Alíquota ICMS")
    Set BDSIS_TBNFP_CPVIC = BDSIS_TBNFP.Fields("Valor ICMS")
    Set BDSIS_TBNFP_CPAIP = BDSIS_TBNFP.Fields("Alíquota IPI")
    Set BDSIS_TBNFP_CPVIP = BDSIS_TBNFP.Fields("Valor IPI")
    Set BDSIS_TBNFP_CPPPA = BDSIS_TBNFP.Fields("Peso Parcial")
    AbreCampos_NotaFiscalProdutos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_NotaFiscalProdutos = False
End Function
Public Function AbreCampos_NotaFiscalDeclaracoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFD_CPNNF = BDSIS_TBNFD.Fields("Número da Nota Fiscal")
    Set BDSIS_TBNFD_CPDEC = BDSIS_TBNFD.Fields("Declaração")
    Set BDSIS_TBNFD_CPLIN = BDSIS_TBNFD.Fields("NumLin")
    AbreCampos_NotaFiscalDeclaracoes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_NotaFiscalDeclaracoes = False
End Function
Public Function AbreCampos_NotaFiscalBancos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFB_CPNNF = BDSIS_TBNFB.Fields("Número da Nota Fiscal")
    Set BDSIS_TBNFB_CPCON = BDSIS_TBNFB.Fields("Nome Conta")
    Set BDSIS_TBNFB_CPLIN = BDSIS_TBNFB.Fields("NumLin")
    AbreCampos_NotaFiscalBancos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_NotaFiscalBancos = False
End Function
Public Function AbreCampos_NotaFiscalComentarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBNFC_CPNNF = BDSIS_TBNFC.Fields("Número da Nota Fiscal")
    Set BDSIS_TBNFC_CPCOM = BDSIS_TBNFC.Fields("Comentários")
    Set BDSIS_TBNFC_CPLIN = BDSIS_TBNFC.Fields("NumLin")
    AbreCampos_NotaFiscalComentarios = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_NotaFiscalComentarios = False
End Function
Public Function AbreCampos_ConfiguracoesTelasImpressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBTIM_CPTEL = BDSIS_TBTIM.Fields("Tela")
    Set BDSIS_TBTIM_CPNOM = BDSIS_TBTIM.Fields("Nome")
    Set BDSIS_TBTIM_CPIMP = BDSIS_TBTIM.Fields("Impressora")
    AbreCampos_ConfiguracoesTelasImpressoras = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_ConfiguracoesTelasImpressoras = False
End Function
Public Function AbreCampos_Impressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBIMP_CPIMP = BDSIS_TBIMP.Fields("Impressora")
    Set BDSIS_TBIMP_CPNOM = BDSIS_TBIMP.Fields("Nome")
    Set BDSIS_TBIMP_CPDRI = BDSIS_TBIMP.Fields("Driver")
    AbreCampos_Impressoras = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Impressoras = False
End Function
Public Function AbreCampos_LogDeEventos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBGEV_CPDAT = BDSIS_TBGEV.Fields("Data")
    Set BDSIS_TBGEV_CPHOR = BDSIS_TBGEV.Fields("Hora")
    Set BDSIS_TBGEV_CPUSU = BDSIS_TBGEV.Fields("Usuário")
    Set BDSIS_TBGEV_CPCOM = BDSIS_TBGEV.Fields("Computador")
    Set BDSIS_TBGEV_CPEVE = BDSIS_TBGEV.Fields("Evento")
    Set BDSIS_TBGEV_CPVAL = BDSIS_TBGEV.Fields("Valor")
    AbreCampos_LogDeEventos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_LogDeEventos = False
End Function
Public Function AbreCampos_LogDeErros(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBGRR_CPDAT = BDSIS_TBGRR.Fields("Data")
    Set BDSIS_TBGRR_CPHOR = BDSIS_TBGRR.Fields("Hora")
    Set BDSIS_TBGRR_CPUSU = BDSIS_TBGRR.Fields("Usuário")
    Set BDSIS_TBGRR_CPCOM = BDSIS_TBGRR.Fields("Computador")
    Set BDSIS_TBGRR_CPAPL = BDSIS_TBGRR.Fields("Aplicativo")
    Set BDSIS_TBGRR_CPENU = BDSIS_TBGRR.Fields("Erro - Número")
    Set BDSIS_TBGRR_CPEDE = BDSIS_TBGRR.Fields("Erro - Descrição")
    Set BDSIS_TBGRR_CPEOR = BDSIS_TBGRR.Fields("Erro - Origem")
    Set BDSIS_TBGRR_CPEJA = BDSIS_TBGRR.Fields("Erro - Arquivo Ajuda")
    Set BDSIS_TBGRR_CPEJC = BDSIS_TBGRR.Fields("Erro - Contexto Ajuda")
    AbreCampos_LogDeErros = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_LogDeErros = False
End Function
Public Function AbreCampos_EntradaPecasIndice(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEPI_CPNEP = BDSIS_TBEPI.Fields("Entrada")
    Set BDSIS_TBEPI_CPDAT = BDSIS_TBEPI.Fields("Data")
    Set BDSIS_TBEPI_CPHOR = BDSIS_TBEPI.Fields("Hora")
    Set BDSIS_TBEPI_CPIND = BDSIS_TBEPI.Fields("Índice")
    AbreCampos_EntradaPecasIndice = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EntradaPecasIndice = False
End Function
Public Function AbreCampos_EntradaPecas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEPE_CPIND = BDSIS_TBEPE.Fields("Índice")
    Set BDSIS_TBEPE_CPQUA = BDSIS_TBEPE.Fields("Quantidade")
    Set BDSIS_TBEPE_CPIFI = BDSIS_TBEPE.Fields("Índice de Ficha")
    Set BDSIS_TBEPE_CPFIN = BDSIS_TBEPE.Fields("Finalidade")
    AbreCampos_EntradaPecas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EntradaPecas = False
End Function
Public Function AbreCampos_MateriaPrimaIndiceDePecas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPP_CPINP = BDSIS_TBMPP.Fields("Índice de Peças")
    Set BDSIS_TBMPP_CPPEC = BDSIS_TBMPP.Fields("Peças")
    AbreCampos_MateriaPrimaIndiceDePecas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MateriaPrimaIndiceDePecas = False
End Function
Public Function AbreCampos_MateriaPrimaIndiceDeBitolas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPB_CPINB = BDSIS_TBMPB.Fields("Índice de Bitolas")
    Set BDSIS_TBMPB_CPBIT = BDSIS_TBMPB.Fields("Bitolas")
    AbreCampos_MateriaPrimaIndiceDeBitolas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MateriaPrimaIndiceDeBitolas = False
End Function
Public Function AbreCampos_MateriaPrimaIndiceDeMateriais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPM_CPINM = BDSIS_TBMPM.Fields("Índice de Materiais")
    Set BDSIS_TBMPM_CPMAT = BDSIS_TBMPM.Fields("Materiais")
    AbreCampos_MateriaPrimaIndiceDeMateriais = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MateriaPrimaIndiceDeMateriais = False
End Function
Public Function AbreCampos_MateriaPrimaIndiceDeQuantidades(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPQ_CPINQ = BDSIS_TBMPQ.Fields("Índice de Quantidades")
    Set BDSIS_TBMPQ_CPQUA = BDSIS_TBMPQ.Fields("Quantidades")
    AbreCampos_MateriaPrimaIndiceDeQuantidades = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MateriaPrimaIndiceDeQuantidades = False
End Function
Public Function AbreCampos_MateriaPrimaIndiceDeNomes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPN_CPINN = BDSIS_TBMPN.Fields("Índice de Nomes")
    Set BDSIS_TBMPN_CPNOM = BDSIS_TBMPN.Fields("Nomes")
    AbreCampos_MateriaPrimaIndiceDeNomes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MateriaPrimaIndiceDeNomes = False
End Function
Public Function AbreCampos_MateriaPrimaRelacaoMateriais(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPR_CPFIG = BDSIS_TBMPR.Fields("Figura")
    Set BDSIS_TBMPR_CPMFG = BDSIS_TBMPR.Fields("Material da Figura")
    Set BDSIS_TBMPR_CPMMP = BDSIS_TBMPR.Fields("Material da Matéria-Prima")
    AbreCampos_MateriaPrimaRelacaoMateriais = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MateriaPrimaRelacaoMateriais = False
End Function
Public Function AbreCampos_EstoqueComplementos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBECM_CPIND = BDSIS_TBECM.Fields("Índice do Complemento")
    Set BDSIS_TBECM_CPCOM = BDSIS_TBECM.Fields("Complemento")
    AbreCampos_EstoqueComplementos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EstoqueComplementos = False
End Function
Public Function AbreCampos_MapaCotacoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMCO_CPMEA = BDSIS_TBMCO.Fields("MêsAno")
    Set BDSIS_TBMCO_CPVAL = BDSIS_TBMCO.Fields("Valor")
    AbreCampos_MapaCotacoes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MapaCotacoes = False
End Function
Public Function AbreCampos_Pedidos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBPED_CPIND = BDSIS_TBPED.Fields("Índice de Pedido")
    Set BDSIS_TBPED_CPDAT = BDSIS_TBPED.Fields("Data Emissão")
    Set BDSIS_TBPED_CPHOR = BDSIS_TBPED.Fields("Hora")
    Set BDSIS_TBPED_CPEMP = BDSIS_TBPED.Fields("Empresa")
    Set BDSIS_TBPED_CPCON = BDSIS_TBPED.Fields("Contato")
    Set BDSIS_TBPED_CPCPG = BDSIS_TBPED.Fields("Condição de Pagamento")
    Set BDSIS_TBPED_CPTRA = BDSIS_TBPED.Fields("Transportadora")
    Set BDSIS_TBPED_CPNSP = BDSIS_TBPED.Fields("Seu Pedido")
    Set BDSIS_TBPED_CPITE = BDSIS_TBPED.Fields("Índice dos Ítens")
    Set BDSIS_TBPED_CPVAL = BDSIS_TBPED.Fields("Valor")
    Set BDSIS_TBPED_CPIVE = BDSIS_TBPED.Fields("Índice Vendedor")
    Set BDSIS_TBPED_CPLIQ = BDSIS_TBPED.Fields("Liquidado")
    Set BDSIS_TBPED_CPABE = BDSIS_TBPED.Fields("Em Aberto")
    Set BDSIS_TBPED_CPIFR = BDSIS_TBPED.Fields("Índice Frete")
    Set BDSIS_TBPED_CPOUD = BDSIS_TBPED.Fields("Outras Despesas")
    Set BDSIS_TBPED_CPDAD = BDSIS_TBPED.Fields("Dados Adicionais")
    Set BDSIS_TBPED_CPOBS = BDSIS_TBPED.Fields("Observações")
    AbreCampos_Pedidos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Pedidos = False
End Function
Public Function AbreCampos_PedidosItens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBPIT_CPIND = BDSIS_TBPIT.Fields("Índice do Ítem")
    Set BDSIS_TBPIT_CPNPE = BDSIS_TBPIT.Fields("Número do Pedido")
    Set BDSIS_TBPIT_CPINF = BDSIS_TBPIT.Fields("Índice de Ficha")
    Set BDSIS_TBPIT_CPQUA = BDSIS_TBPIT.Fields("Quantidade")
    Set BDSIS_TBPIT_CPDES = BDSIS_TBPIT.Fields("Descrição")
    Set BDSIS_TBPIT_CPCOM = BDSIS_TBPIT.Fields("Complemento")
    Set BDSIS_TBPIT_CPPRE = BDSIS_TBPIT.Fields("Preço Unitário")
    Set BDSIS_TBPIT_CPPRA = BDSIS_TBPIT.Fields("Prazo de Entrega")
    Set BDSIS_TBPIT_CPLIQ = BDSIS_TBPIT.Fields("Liquidado")
    AbreCampos_PedidosItens = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_PedidosItens = False
End Function
Public Function AbreCampos_MapaPedidos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBMPE_CPMEA = BDSIS_TBMPE.Fields("MêsAno")
    Set BDSIS_TBMPE_CPVAL = BDSIS_TBMPE.Fields("Valor")
    AbreCampos_MapaPedidos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_MapaPedidos = False
End Function
Public Function AbreCampos_OrdemFabricacao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBODF_CPNOF = BDSIS_TBODF.Fields("OF")
    Set BDSIS_TBODF_CPDAE = BDSIS_TBODF.Fields("Data Emissão")
    Set BDSIS_TBODF_CPINF = BDSIS_TBODF.Fields("Índice de Ficha")
    Set BDSIS_TBODF_CPDES = BDSIS_TBODF.Fields("Descrição")
    Set BDSIS_TBODF_CPNDC = BDSIS_TBODF.Fields("DC")
    Set BDSIS_TBODF_CPQES = BDSIS_TBODF.Fields("Quantidade Estipulada")
    Set BDSIS_TBODF_CPQPR = BDSIS_TBODF.Fields("Quantidade Produzida")
    Set BDSIS_TBODF_CPCOR = BDSIS_TBODF.Fields("Corridas")
    Set BDSIS_TBODF_CPDAL = BDSIS_TBODF.Fields("Data Liquidação")
    Set BDSIS_TBODF_CPINP = BDSIS_TBODF.Fields("Índices de Processos")
    Set BDSIS_TBODF_CPLIQ = BDSIS_TBODF.Fields("Liquidado")
    AbreCampos_OrdemFabricacao = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_OrdemFabricacao = False
End Function
Public Function AbreCampos_OrdemFabricacaoProcessos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBOFP_CPIND = BDSIS_TBOFP.Fields("Índice do Processo")
    Set BDSIS_TBOFP_CPNOF = BDSIS_TBOFP.Fields("OF")
    Set BDSIS_TBOFP_CPNET = BDSIS_TBOFP.Fields("Etapa")
    Set BDSIS_TBOFP_CPNMA = BDSIS_TBOFP.Fields("Máquina")
    Set BDSIS_TBOFP_CPDAI = BDSIS_TBOFP.Fields("Data Início")
    Set BDSIS_TBOFP_CPDAF = BDSIS_TBOFP.Fields("Data Fim")
    Set BDSIS_TBOFP_CPNHE = BDSIS_TBOFP.Fields("Hora Extra")
    AbreCampos_OrdemFabricacaoProcessos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_OrdemFabricacaoProcessos = False
End Function
Public Function AbreCampos_OrdemMontagem(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBODM_CPNOM = BDSIS_TBODM.Fields("OM")
    Set BDSIS_TBODM_CPDAT = BDSIS_TBODM.Fields("Data")
    Set BDSIS_TBODM_CPHOR = BDSIS_TBODM.Fields("Hora")
    Set BDSIS_TBODM_CPINF = BDSIS_TBODM.Fields("Índice de Ficha")
    Set BDSIS_TBODM_CPDES = BDSIS_TBODM.Fields("Descrição")
    Set BDSIS_TBODM_CPNDC = BDSIS_TBODM.Fields("DC")
    Set BDSIS_TBODM_CPQES = BDSIS_TBODM.Fields("Quantidade Estipulada")
    Set BDSIS_TBODM_CPQMO = BDSIS_TBODM.Fields("Quantidade Montada")
    Set BDSIS_TBODM_CPDIN = BDSIS_TBODM.Fields("Data Início")
    Set BDSIS_TBODM_CPHIN = BDSIS_TBODM.Fields("Hora Início")
    Set BDSIS_TBODM_CPDFI = BDSIS_TBODM.Fields("Data Fim")
    Set BDSIS_TBODM_CPHFI = BDSIS_TBODM.Fields("Hora Fim")
    Set BDSIS_TBODM_CPNCA = BDSIS_TBODM.Fields("NCA")
    Set BDSIS_TBODM_CPNCB = BDSIS_TBODM.Fields("NCB")
    Set BDSIS_TBODM_CPNCC = BDSIS_TBODM.Fields("NCC")
    Set BDSIS_TBODM_CPNCD = BDSIS_TBODM.Fields("NCD")
    Set BDSIS_TBODM_CPNCE = BDSIS_TBODM.Fields("NCE")
    Set BDSIS_TBODM_CPLIQ = BDSIS_TBODM.Fields("Liquidado")
    AbreCampos_OrdemMontagem = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_OrdemMontagem = False
End Function
Public Function AbreCampos_OrdemExpedicao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBODE_CPNOE = BDSIS_TBODE.Fields("OE")
    Set BDSIS_TBODE_CPDAT = BDSIS_TBODE.Fields("Data")
    Set BDSIS_TBODE_CPHOR = BDSIS_TBODE.Fields("Hora")
    Set BDSIS_TBODE_CPEMP = BDSIS_TBODE.Fields("Empresa")
    Set BDSIS_TBODE_CPNPE = BDSIS_TBODE.Fields("PE")
    Set BDSIS_TBODE_CPCFS = BDSIS_TBODE.Fields("Código Funcionário - Separado")
    Set BDSIS_TBODE_CPCFC = BDSIS_TBODE.Fields("Código Funcionário - Conferido")
    Set BDSIS_TBODE_CPCFV = BDSIS_TBODE.Fields("Código Funcionário - Visto")
    Set BDSIS_TBODE_CPETP = BDSIS_TBODE.Fields("Embalagem - Tipo")
    Set BDSIS_TBODE_CPENU = BDSIS_TBODE.Fields("Embalagem - Números")
    Set BDSIS_TBODE_CPPBR = BDSIS_TBODE.Fields("Peso Bruto")
    Set BDSIS_TBODE_CPPLI = BDSIS_TBODE.Fields("Peso Bruto")
    Set BDSIS_TBODE_CPTRA = BDSIS_TBODE.Fields("Transportadora")
    Set BDSIS_TBODE_CPCOR = BDSIS_TBODE.Fields("Corridas")
    Set BDSIS_TBODE_CPOBS = BDSIS_TBODE.Fields("Observações")
    Set BDSIS_TBODE_CPLIQ = BDSIS_TBODE.Fields("Liquidado")
    AbreCampos_OrdemExpedicao = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_OrdemExpedicao = False
End Function
Public Function AbreCampos_Funcionarios(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBFUN_CPIND = BDSIS_TBFUN.Fields("Índice")
    Set BDSIS_TBFUN_CPFUN = BDSIS_TBFUN.Fields("Funcionário")
    Set BDSIS_TBFUN_CPREG = BDSIS_TBFUN.Fields("Número do Registro")
    Set BDSIS_TBFUN_CPTIP = BDSIS_TBFUN.Fields("Tipo")
    Set BDSIS_TBFUN_CPSET = BDSIS_TBFUN.Fields("Setor")
    AbreCampos_Funcionarios = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Funcionarios = False
End Function
Public Function AbreCampos_ContasPagarReceber(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCPR_CPIND = BDSIS_TBCPR.Fields("Índice do Movimento")
    Set BDSIS_TBCPR_CPEMI = BDSIS_TBCPR.Fields("Data de Emissão")
    Set BDSIS_TBCPR_CPVEN = BDSIS_TBCPR.Fields("Data de Vencimento")
    Set BDSIS_TBCPR_CPMOV = BDSIS_TBCPR.Fields("Movimento")
    Set BDSIS_TBCPR_CPVAL = BDSIS_TBCPR.Fields("Valor")
    Set BDSIS_TBCPR_CPORI = BDSIS_TBCPR.Fields("Origem e Destino")
    Set BDSIS_TBCPR_CPNDO = BDSIS_TBCPR.Fields("Número Documento")
    Set BDSIS_TBCPR_CPNNU = BDSIS_TBCPR.Fields("Nosso Número")
    Set BDSIS_TBCPR_CPSNU = BDSIS_TBCPR.Fields("Seu Número")
    Set BDSIS_TBCPR_CPTIP = BDSIS_TBCPR.Fields("Tipo")
    Set BDSIS_TBCPR_CPBAN = BDSIS_TBCPR.Fields("Banco")
    Set BDSIS_TBCPR_CPCAR = BDSIS_TBCPR.Fields("Em Carteira")
    Set BDSIS_TBCPR_CPCOB = BDSIS_TBCPR.Fields("Em Cobrança")
    Set BDSIS_TBCPR_CPOBS = BDSIS_TBCPR.Fields("Observações")
    AbreCampos_ContasPagarReceber = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_ContasPagarReceber = False
End Function
Public Function AbreCampos_Imagens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBIMA_CPIND = BDSIS_TBIMA.Fields("Índice da Imagem")
    Set BDSIS_TBIMA_CPNOM = BDSIS_TBIMA.Fields("Nome")
    Set BDSIS_TBIMA_CPTIP = BDSIS_TBIMA.Fields("Tipo da Imagem")
    Set BDSIS_TBIMA_CPIMA = BDSIS_TBIMA.Fields("Imagem")
    AbreCampos_Imagens = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Imagens = False
End Function
Public Function AbreCampos_Equipamentos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEQU_CPIND = BDSIS_TBEQU.Fields("Índice")
    Set BDSIS_TBEQU_CPEQU = BDSIS_TBEQU.Fields("Equipamento")
    Set BDSIS_TBEQU_CPNOM = BDSIS_TBEQU.Fields("Nome")
    Set BDSIS_TBEQU_CPMOD = BDSIS_TBEQU.Fields("Modelo")
    Set BDSIS_TBEQU_CPFAB = BDSIS_TBEQU.Fields("Fabricante")
    Set BDSIS_TBEQU_CPANO = BDSIS_TBEQU.Fields("Ano")
    Set BDSIS_TBEQU_CPVAL = BDSIS_TBEQU.Fields("Valor")
    Set BDSIS_TBEQU_CPFOT = BDSIS_TBEQU.Fields("Nome da Foto")
    Set BDSIS_TBEQU_CPCHO = BDSIS_TBEQU.Fields("Custo da Hora")
    Set BDSIS_TBEQU_CPPRO = BDSIS_TBEQU.Fields("Processo")
    Set BDSIS_TBEQU_CPFUN = BDSIS_TBEQU.Fields("Índice do Funcionário")
    AbreCampos_Equipamentos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Equipamentos = False
End Function
Public Function AbreCampos_EquipamentosManutencao(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBEMA_CPIND = BDSIS_TBEMA.Fields("Índice")
    Set BDSIS_TBEMA_CPEQU = BDSIS_TBEMA.Fields("Equipamento")
    Set BDSIS_TBEMA_CPMAN = BDSIS_TBEMA.Fields("Manutenção")
    Set BDSIS_TBEMA_CPDUL = BDSIS_TBEMA.Fields("Data da Última")
    Set BDSIS_TBEMA_CPDPR = BDSIS_TBEMA.Fields("Data da Próxima")
    Set BDSIS_TBEMA_CPVAL = BDSIS_TBEMA.Fields("Valor")
    Set BDSIS_TBEMA_CPEMP = BDSIS_TBEMA.Fields("Empresa")
    Set BDSIS_TBEMA_CPCON = BDSIS_TBEMA.Fields("Contato")
    Set BDSIS_TBEMA_CPTIP = BDSIS_TBEMA.Fields("Tipo")
    Set BDSIS_TBEMA_CPAVI = BDSIS_TBEMA.Fields("Aviso")
    Set BDSIS_TBEMA_CPUSU = BDSIS_TBEMA.Fields("Usuário")
    AbreCampos_EquipamentosManutencao = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_EquipamentosManutencao = False
End Function
Public Function AbreCampos_Avisos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBAVI_CPIND = BDSIS_TBAVI.Fields("Índice")
    Set BDSIS_TBAVI_CPREM = BDSIS_TBAVI.Fields("Remetente")
    Set BDSIS_TBAVI_CPEMI = BDSIS_TBAVI.Fields("Data de Emissão")
    Set BDSIS_TBAVI_CPDES = BDSIS_TBAVI.Fields("Destinatário")
    Set BDSIS_TBAVI_CPVEN = BDSIS_TBAVI.Fields("Data do Aviso")
    Set BDSIS_TBAVI_CPTIP = BDSIS_TBAVI.Fields("Tipo")
    Set BDSIS_TBAVI_CPAVI = BDSIS_TBAVI.Fields("Aviso")
    AbreCampos_Avisos = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Avisos = False
End Function
Public Function AbreCampos_CNCFuncoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCFU_CPIND = BDSIS_TBCFU.Fields("Índice")
    Set BDSIS_TBCFU_CPCNC = BDSIS_TBCFU.Fields("CNC")
    Set BDSIS_TBCFU_CPFUN = BDSIS_TBCFU.Fields("Função")
    Set BDSIS_TBCFU_CPAPL = BDSIS_TBCFU.Fields("Aplicação")
    Set BDSIS_TBCFU_CPVAR = BDSIS_TBCFU.Fields("Variáveis")
    AbreCampos_CNCFuncoes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_CNCFuncoes = False
End Function
Public Function AbreCampos_CNCMaquinas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCMA_CPIND = BDSIS_TBCMA.Fields("Índice")
    Set BDSIS_TBCMA_CPNNE = BDSIS_TBCMA.Fields("Índice NE")
    AbreCampos_CNCMaquinas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_CNCMaquinas = False
End Function
Public Function AbreCampos_CNCProgramas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBCNC_CPIND = BDSIS_TBCNC.Fields("Índice")
    Set BDSIS_TBCNC_CPNOM = BDSIS_TBCNC.Fields("Nome do Programa")
    Set BDSIS_TBCNC_CPCNC = BDSIS_TBCNC.Fields("Nome no CNC")
    Set BDSIS_TBCNC_CPDAT = BDSIS_TBCNC.Fields("Data Criação")
    Set BDSIS_TBCNC_CPULT = BDSIS_TBCNC.Fields("Data Último Acesso")
    Set BDSIS_TBCNC_CPTIP = BDSIS_TBCNC.Fields("Tipo do CNC")
    Set BDSIS_TBCNC_CPMAQ = BDSIS_TBCNC.Fields("Índice da Máquina")
    Set BDSIS_TBCNC_CPPRO = BDSIS_TBCNC.Fields("Programa")
    AbreCampos_CNCProgramas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_CNCProgramas = False
End Function
Public Function AbreCampos_Layouts(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBLAY_CPIND = BDSIS_TBLAY.Fields("Número NL")
    Set BDSIS_TBLAY_CPLAY = BDSIS_TBLAY.Fields("Layout")
    Set BDSIS_TBLAY_CPMAQ = BDSIS_TBLAY.Fields("Máquina")
    Set BDSIS_TBLAY_CPNEL = BDSIS_TBLAY.Fields("Número EL")
    Set BDSIS_TBLAY_CPNDL = BDSIS_TBLAY.Fields("Número DL")
    Set BDSIS_TBLAY_CPDCR = BDSIS_TBLAY.Fields("Data de Criação")
    Set BDSIS_TBLAY_CPCOM = BDSIS_TBLAY.Fields("Comentário")
    Set BDSIS_TBLAY_CPNRE = BDSIS_TBLAY.Fields("Número da Revisão")
    Set BDSIS_TBLAY_CPDRE = BDSIS_TBLAY.Fields("Data da Revisão")
    Set BDSIS_TBLAY_CPCFU = BDSIS_TBLAY.Fields("Código do Funcionário")
    Set BDSIS_TBLAY_CPFOT = BDSIS_TBLAY.Fields("Foto")
    AbreCampos_Layouts = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_Layouts = False
End Function
Public Function AbreCampos_LayoutsEsquemas(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBLES_CPIND = BDSIS_TBLES.Fields("Número EL")
    Set BDSIS_TBLES_CPESQ = BDSIS_TBLES.Fields("Esquema")
    Set BDSIS_TBLES_CPMAQ = BDSIS_TBLES.Fields("Máquina")
    Set BDSIS_TBLES_CPVAR = BDSIS_TBLES.Fields("Variáveis")
    Set BDSIS_TBLES_CPDIS = BDSIS_TBLES.Fields("Dispositivos")
    Set BDSIS_TBLES_CPFER = BDSIS_TBLES.Fields("Ferramentas")
    AbreCampos_LayoutsEsquemas = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_LayoutsEsquemas = False
End Function
Public Function AbreCampos_LayoutsDimensoes(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBLDI_CPIND = BDSIS_TBLDI.Fields("Número DL")
    Set BDSIS_TBLDI_CPNEL = BDSIS_TBLDI.Fields("Número EL")
    Set BDSIS_TBLDI_CPVAR = BDSIS_TBLDI.Fields("Variáveis")
    Set BDSIS_TBLDI_CPDIS = BDSIS_TBLDI.Fields("Dispositivos")
    Set BDSIS_TBLDI_CPFER = BDSIS_TBLDI.Fields("Ferramentas")
    Set BDSIS_TBLDI_CPOBS = BDSIS_TBLDI.Fields("Observações")
    AbreCampos_LayoutsDimensoes = True
    Exit Function
ERRO_ABERTURA:
    AbreCampos_LayoutsDimensoes = False
End Function



'**********************************************************************************************************
'
'           B A N C O     D E     D A D O S     C L I E N T E
'
'**********************************************************************************************************


'*****************************************************
'     F U N Ç Õ E S     A B R E     T A B E L A S
'*****************************************************
Public Function Cliente_AbreTabela_DirPrg(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBDPR = BDCLI.OpenRecordset("Diretório do Programa")
    'BDCLI_TBDPR.Index = ""
    Cliente_AbreTabela_DirPrg = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreTabela_DirPrg = False
End Function
Public Function Cliente_AbreTabela_DirBD(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBDBD = BDCLI.OpenRecordset("Diretório do Banco de Dados")
    'BDCLI_TBDBD.Index = ""
    Cliente_AbreTabela_DirBD = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreTabela_DirBD = False
End Function
Public Function Cliente_AbreTabela_Impressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBIMP = BDCLI.OpenRecordset("Impressoras")
    'BDCLI_TBIMP.Index = ""
    Cliente_AbreTabela_Impressoras = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreTabela_Impressoras = False
End Function
Public Function Cliente_AbreTabela_UltimoLogon(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBULG = BDCLI.OpenRecordset("Último Logon")
    'BDCLI_TBULG.Index = ""
    Cliente_AbreTabela_UltimoLogon = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreTabela_UltimoLogon = False
End Function
Public Function AbreTabela_Pedidos(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBPED = BDSIS.OpenRecordset("Pedidos")
    BDSIS_TBPED.Index = "Índice de Pedido"
    AbreTabela_Pedidos = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_Pedidos = False
End Function
Public Function AbreTabela_PedidosItens(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    Set BDSIS_TBPIT = BDSIS.OpenRecordset("Pedidos - Ítens")
    BDSIS_TBPIT.Index = "Ítem"
    AbreTabela_PedidosItens = True
    Exit Function
ERRO_ABERTURA:
    AbreTabela_PedidosItens = False
End Function



'*****************************************************
'     F U N Ç Õ E S     F E C H A     T A B E L A S
'*****************************************************
Public Function Cliente_FechaTabela_DirPrg(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    'Set BDCLI_TBDPR_CPCAM = Nothing
    'BDCLI_TBDPR.Close
    'Set BDCLI_TBDPR = Nothing
    Cliente_FechaTabela_DirPrg = True
    Exit Function
ERRO_FINALIZACAO:
    Cliente_FechaTabela_DirPrg = False
End Function
Public Function Cliente_FechaTabela_DirBD(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    'Set BDCLI_TBDBD_CPCAM = Nothing
    'BDCLI_TBDBD.Close
    'Set BDCLI_TBDBD = Nothing
    Cliente_FechaTabela_DirBD = True
    Exit Function
ERRO_FINALIZACAO:
    Cliente_FechaTabela_DirBD = False
End Function
Public Function Cliente_FechaTabela_Impressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    'Set BDCLI_TBIMP_CPIMP = Nothing
    'Set BDCLI_TBIMP_CPNOM = Nothing
    'Set BDCLI_TBIMP_CPDRI = Nothing
    'BDCLI_TBIMP.Close
    'Set BDCLI_TBIMP = Nothing
    Cliente_FechaTabela_Impressoras = True
    Exit Function
ERRO_FINALIZACAO:
    Cliente_FechaTabela_Impressoras = False
End Function
Public Function Cliente_FechaTabela_UltimoLogon(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_FINALIZACAO
    On Error GoTo ERRO_FINALIZACAO
    'Set BDCLI_TBULG_CPDAT = Nothing
    'Set BDCLI_TBULG_CPHOR = Nothing
    'Set BDCLI_TBULG_CPCOM = Nothing
    'Set BDCLI_TBULG_CPUSU = Nothing
    'BDCLI_TBULG.Close
    'Set BDCLI_TBULG = Nothing
    Cliente_FechaTabela_UltimoLogon = True
    Exit Function
ERRO_FINALIZACAO:
    Cliente_FechaTabela_UltimoLogon = False
End Function


'*****************************************************
'        F U N Ç Õ E S     A B R E     C A M P O S
'*****************************************************
Public Function Cliente_AbreCampos_DirPrg(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBDPR_CPCAM = BDCLI_TBDPR.Fields("Caminho")
    Cliente_AbreCampos_DirPrg = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreCampos_DirPrg = False
End Function
Public Function Cliente_AbreCampos_DirBD(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBDBD_CPCAM = BDCLI_TBDBD.Fields("Caminho")
    Cliente_AbreCampos_DirBD = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreCampos_DirBD = False
End Function
Public Function Cliente_AbreCampos_Impressoras(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBIMP_CPIMP = BDCLI_TBIMP.Fields("Impressora")
    'Set BDCLI_TBIMP_CPNOM = BDCLI_TBIMP.Fields("Nome")
    'Set BDCLI_TBIMP_CPDRI = BDCLI_TBIMP.Fields("Driver")
    Cliente_AbreCampos_Impressoras = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreCampos_Impressoras = False
End Function
Public Function Cliente_AbreCampos_UltimoLogon(X, Y, Z) As Boolean
    If X <> App.ProductName And Y <> App.EXEName And Z <> App.LegalCopyright Then GoTo ERRO_ABERTURA
    On Error GoTo ERRO_ABERTURA
    'Set BDCLI_TBULG_CPDAT = BDCLI_TBULG.Fields("Data")
    'Set BDCLI_TBULG_CPHOR = BDCLI_TBULG.Fields("Hora")
    'Set BDCLI_TBULG_CPCOM = BDCLI_TBULG.Fields("Computador")
    'Set BDCLI_TBULG_CPUSU = BDCLI_TBULG.Fields("Usuário")
    Cliente_AbreCampos_UltimoLogon = True
    Exit Function
ERRO_ABERTURA:
    Cliente_AbreCampos_UltimoLogon = False
End Function
