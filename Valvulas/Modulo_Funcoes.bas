Attribute VB_Name = "Modulo_Funcoes"
Option Explicit
'***********************************
'            VARIAVEIS
'***********************************


'***********************************
'          SETA MATRIZES
'***********************************

Private Function CodigoBruto(PECA As String, BIT As String, MAT As String) As String
    Dim tmp As String
    tmp = "A" 'BRUTO
    'PECA
    Select Case PECA
        Case "CORPO GAVETA"
            tmp = tmp & "A"
        Case "CASTELO"
            tmp = tmp & "B"
        Case "PREME"
            tmp = tmp & "C"
        Case "TAMPA"
            tmp = tmp & "D"
        Case "CUNHA"
            tmp = tmp & "E"
        Case "CONTRA-SEDE"
            tmp = tmp & "F"
        Case "PÊNDULO"
            tmp = tmp & "G"
        Case "FLANGE"
            tmp = tmp & "H"
        Case "VOLANTE GAVETA"
            tmp = tmp & "I"
        Case "VOLANTE GLOBO"
            tmp = tmp & "J"
        Case "REDONDO"
            tmp = tmp & "K"
        Case "SEXTAVADO"
            tmp = tmp & "L"
        Case "QUADRADO"
            tmp = tmp & "M"
        Case "SOLDA"
            tmp = tmp & "N"
        Case "REVESTIMENTO"
            tmp = tmp & "O"
    End Select
    'BITOLA
    Select Case BIT
        Case "1/2" & Chr(34)
            tmp = tmp & "A"
        Case "3/4" & Chr(34)
            tmp = tmp & "B"
        Case "1" & Chr(34)
            tmp = tmp & "C"
        Case "1.1/2" & Chr(34)
            tmp = tmp & "D"
        Case "2" & Chr(34)
            tmp = tmp & "E"
        Case "1/2" & Chr(34) & " e 3/4" & Chr(34)
            tmp = tmp & "F"
        Case "1.1/2" & Chr(34) & " e 2" & Chr(34)
            tmp = tmp & "G"
        Case "FLANGE"
            tmp = tmp & "H"
        Case "VOLANTE GAVETA"
            tmp = tmp & "I"
        Case "VOLANTE GLOBO"
            tmp = tmp & "J"
        Case "REDONDO"
            tmp = tmp & "K"
        Case "SEXTAVADO"
            tmp = tmp & "L"
        Case "QUADRADO"
            tmp = tmp & "M"
        Case "SOLDA"
            tmp = tmp & "N"
        Case "REVESTIMENTO"
            tmp = tmp & "O"
    End Select
    
    CodigoBruto = tmp
End Function


