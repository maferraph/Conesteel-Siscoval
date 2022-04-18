Attribute VB_Name = "Modulo_NF"
Option Explicit
' ****************** VARIÁVEIS DLL's ****************
Global DLL_BD As Scvbd.Classe_Scvbd
Global DLL_CARGA As Scvcarr.Classe_Scvcarr
Global DLL_FUNCS As Scvfunc.Classe_Scvfunc
Global DLL_ASFIG As Assfig.Classe_Assfig
Global DLL_CADEMP As Cademp.Classe_Cademp
Global DLL_IMP As Impform.Classe_Impform


' ****************** DECLARAÇÕES ****************
Const NOMEAPLIC As String = "Assistente de Nota Fiscal"
Global CD_1 As CommonDialog
Global PesoItem As Double
Global DescricaoItem As String
Global DesRed, DesNor, DesCom As String
Global RBC_ICMS As Double
Global I, J, K, NumLinha As Integer
Global Diretorio As String
Global InseriuPedido As Boolean
Global RespMsg
Global NumLB As Integer
Global ValRed As Double
Global CF_I As String, CF_J As String, ESTIND As String
Public bPedLiq As Boolean
