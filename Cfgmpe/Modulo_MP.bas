Attribute VB_Name = "Modulo_MP"
Option Explicit

Dim I As Integer, RespMsg

Dim kBIT18 As String
Dim kBIT14 As String
Dim kBIT38 As String
Dim kBIT12 As String
Dim kBIT34 As String
Dim kBIT1 As String
Dim kBIT114 As String
Dim kBIT112 As String
Dim kBIT2 As String
Dim kBIT12E34 As String
Dim kBIT12E34E1 As String
Dim kBIT112E2 As String
Dim kBIT316 As String

Dim aFIG As Variant, aQUA As Variant, aNOM As Variant, aBIT As Variant
Dim sBITTMP As String, aBitFor As Variant, aBit1 As Variant, aBit2 As Variant, sFT As String, sBT As String



'*****************************************************************
' RELACAO DE FUNCOES PARA CONFIGURAR MP
'*****************************************************************
Public Static Function ProcuraMaterialMP(vPECAS As Variant, sMat As String) As Variant
    If Not IsArray(vPECAS) Then Exit Function
    Dim aMATERIAL As Variant
    aMATERIAL = Array()
    For I = 0 To UBound(vPECAS) - 1
        'procura informações de materiais sobre MP
        Tela_Cfg_MateriaPrima.DLL_BD.BDSIS_TBMPR.Seek "=", vPECAS(I), sMat
        ReDim Preserve aMATERIAL(UBound(aMATERIAL) + 1)
        If Tela_Cfg_MateriaPrima.DLL_BD.BDSIS_TBMPR.NoMatch Then
            aMATERIAL(UBound(aMATERIAL)) = ""
        Else
            aMATERIAL(UBound(aMATERIAL)) = Tela_Cfg_MateriaPrima.DLL_BD.BDSIS_TBMPR_CPMMP.Value
        End If
    Next I
    For I = 0 To UBound(aMATERIAL) - 1
        If aMATERIAL(I) = "" Then
            RespMsg = InputBox("Existe uma peça/componente que não foi possível localizar seu material:" & vbCr & vbCr & "Peça: " & vPECAS(I) & vbCr & "Material da Figura: " & sMat & vbCr & vbCr & "Por favor, indique qual material deverá ser usado.", "Falta material")
            If RespMsg <> "" Then aMATERIAL(I) = RespMsg
        End If
    Next I
    ProcuraMaterialMP = aMATERIAL
End Function




'*****************************************************************
' RELACAO DE FUNCOES AUXILIARES PARA CONFIGURAR MP
'*****************************************************************
Public Static Sub LE_BITOLAS()
    kBIT18 = "1/8" & Chr(34)
    kBIT14 = "1/4" & Chr(34)
    kBIT38 = "3/8" & Chr(34)
    kBIT12 = "1/2" & Chr(34)
    kBIT34 = "3/4" & Chr(34)
    kBIT1 = "1" & Chr(34)
    kBIT114 = "1.1/4" & Chr(34)
    kBIT112 = "1.1/2" & Chr(34)
    kBIT2 = "2" & Chr(34)
    kBIT12E34 = "1/2" & Chr(34) & " E " & "3/4" & Chr(34)
    kBIT12E34E1 = "1/2" & Chr(34) & " E " & "3/4" & Chr(34) & " E " & "1" & Chr(34)
    kBIT112E2 = "1.1/2" & Chr(34) & " E " & "2" & Chr(34)
    kBIT316 = "3/16" & Chr(34)
    aFIG = ""
    aQUA = ""
    aNOM = ""
    aBIT = ""
End Sub
Public Static Sub GAV800(sBit As String)
    If sBit = "1/8" & Chr(34) Then
        aBIT = Array(kBIT18, kBIT12E34, kBIT12E34, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12)
    ElseIf sBit = "1/4" & Chr(34) Then
        aBIT = Array(kBIT14, kBIT12E34, kBIT12E34, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12)
    ElseIf sBit = "3/8" & Chr(34) Then
        aBIT = Array(kBIT38, kBIT12E34, kBIT12E34, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12)
    ElseIf sBit = "1/2" & Chr(34) Then
        aBIT = Array(kBIT12, kBIT12E34, kBIT12E34, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12)
    ElseIf sBit = "3/4" & Chr(34) Then
        aBIT = Array(kBIT34, kBIT12E34, kBIT12E34, kBIT34, kBIT34, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT34, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT34)
    ElseIf sBit = "1" & Chr(34) Then
        aBIT = Array(kBIT1, kBIT1, kBIT1, kBIT1, kBIT1, kBIT1, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT1, kBIT1, kBIT1, kBIT12E34E1, kBIT1)
    ElseIf sBit = "1.1/4" & Chr(34) Then
        aBIT = Array(kBIT114, kBIT112E2, kBIT112E2, kBIT112, kBIT112, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "1.1/2" & Chr(34) Then
        aBIT = Array(kBIT112, kBIT112E2, kBIT112E2, kBIT112, kBIT112, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "2" & Chr(34) Then
        aBIT = Array(kBIT2, kBIT112E2, kBIT112E2, kBIT112, kBIT112, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    End If
End Sub
Public Static Sub GAV1500(sBit As String)
    If sBit = "1/8" & Chr(34) Then
        aBIT = Array(kBIT18, kBIT12E34, kBIT1, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "1/4" & Chr(34) Then
        aBIT = Array(kBIT14, kBIT12E34, kBIT1, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "3/8" & Chr(34) Then
        aBIT = Array(kBIT38, kBIT12E34, kBIT1, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "1/2" & Chr(34) Then
        aBIT = Array(kBIT12, kBIT12E34, kBIT1, kBIT12, kBIT12, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "3/4" & Chr(34) Then
        aBIT = Array(kBIT34, kBIT12E34, kBIT1, kBIT34, kBIT34, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT34, kBIT1, kBIT1, kBIT12E34E1, kBIT34)
    ElseIf sBit = "1" & Chr(34) Then
        aBIT = Array(kBIT1, kBIT1, kBIT112E2, kBIT1, kBIT1, kBIT1, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT1, kBIT112E2, kBIT112E2, kBIT112E2, kBIT1)
    ElseIf sBit = "1.1/4" & Chr(34) Then
        aBIT = Array(kBIT114, kBIT112E2, kBIT112E2, kBIT112, kBIT112, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "1.1/2" & Chr(34) Then
        aBIT = Array(kBIT112, kBIT112E2, kBIT112E2, kBIT112, kBIT112, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "2" & Chr(34) Then
        aBIT = Array(kBIT2, kBIT112E2, kBIT112E2, kBIT112, kBIT112, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    End If
End Sub
Public Static Sub GLO800(sBit As String)
    If sBit = "1/8" & Chr(34) Then
        aBIT = Array(kBIT18, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34E1, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "1/4" & Chr(34) Then
        aBIT = Array(kBIT14, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34E1, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "3/8" & Chr(34) Then
        aBIT = Array(kBIT38, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34E1, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "1/2" & Chr(34) Then
        aBIT = Array(kBIT12, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT12E34E1, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT12)
    ElseIf sBit = "3/4" & Chr(34) Then
        aBIT = Array(kBIT34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT12E34, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT34, kBIT12E34E1, kBIT12E34, kBIT12E34E1, kBIT12E34E1, kBIT34)
    ElseIf sBit = "1" & Chr(34) Then
        aBIT = Array(kBIT1, kBIT1, kBIT1, kBIT1, kBIT1, kBIT1, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT1, kBIT12E34E1, kBIT1, kBIT12E34E1, kBIT112E2, kBIT1)
    ElseIf sBit = "1.1/4" & Chr(34) Then
        aBIT = Array(kBIT114, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "1.1/2" & Chr(34) Then
        aBIT = Array(kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "2" & Chr(34) Then
        aBIT = Array(kBIT2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    End If
End Sub
Public Static Sub GLO1500(sBit As String)
    If sBit = "1/8" & Chr(34) Then
        aBIT = Array(kBIT18, kBIT12E34, kBIT1, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT112E2, kBIT12)
    ElseIf sBit = "1/4" & Chr(34) Then
        aBIT = Array(kBIT14, kBIT12E34, kBIT1, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT112E2, kBIT12)
    ElseIf sBit = "3/8" & Chr(34) Then
        aBIT = Array(kBIT38, kBIT12E34, kBIT1, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT112E2, kBIT12)
    ElseIf sBit = "1/2" & Chr(34) Then
        aBIT = Array(kBIT12, kBIT12E34, kBIT1, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT12, kBIT1, kBIT1, kBIT12E34E1, kBIT112E2, kBIT12)
    ElseIf sBit = "3/4" & Chr(34) Then
        aBIT = Array(kBIT34, kBIT12E34, kBIT1, kBIT12E34, kBIT12E34, kBIT12E34, kBIT18, kBIT1, kBIT1, kBIT12E34E1, kBIT12E34E1, kBIT34, kBIT1, kBIT1, kBIT12E34E1, kBIT112E2, kBIT34)
    ElseIf sBit = "1" & Chr(34) Then
        aBIT = Array(kBIT1, kBIT1, kBIT112E2, kBIT1, kBIT1, kBIT1, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT1, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT1)
    ElseIf sBit = "1.1/4" & Chr(34) Then
        aBIT = Array(kBIT114, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "1.1/2" & Chr(34) Then
        aBIT = Array(kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    ElseIf sBit = "2" & Chr(34) Then
        aBIT = Array(kBIT2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT316, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112E2, kBIT112)
    End If
End Sub
Public Static Sub BitFor1(sFIG As String, sBit As String)
    'QUANTIDADE
    aQUA = Array("1")
    'BITOLA
    aBitFor = Array("3/8" & Chr(34), "1/2" & Chr(34), "3/4" & Chr(34), "1" & Chr(34), "1.1/4" & Chr(34), "1.1/2" & Chr(34), "2" & Chr(34), "2" & Chr(34) & " 3000#", "2.1/2" & Chr(34), "3" & Chr(34))
    Dim nN As Integer
    nN = Len(sFIG)
    If (Mid(sFIG, nN - 1, 1) = "N" Or Mid(sFIG, nN - 1, 1) = "B" Or Mid(sFIG, nN - 1, 1) = "T") Then
        If (Len(sBit) = 4 And Left(sBit, 3) = "1/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(0))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(1))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(1))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(2))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/2") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(1))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(2))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(3))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(2))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(3))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(4))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "1") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "1" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(3))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(4))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(5))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/4") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(4))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(5))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(6))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(5))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(6))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(7))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "2") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(6))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(7))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(8))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "2.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "2.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(8))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(9))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array("-")
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "3") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "3" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array(aBitFor(9))
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array("-")
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array("-")
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "4") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "2" Then aBIT = Array("-")
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array("-")
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array("-")
        End If
    ElseIf Mid(sFIG, nN - 1, 1) = "S" Then
        If (Len(sBit) = 4 And Left(sBit, 3) = "1/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(0))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(1))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(0))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(1))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(2))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/2") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(1))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(2))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(3))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(2))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(3))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(4))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "1") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "1" & Chr(34) & " X") Then
            aQUA = Array("1")
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(3))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(4))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(5))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/4") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(4))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(5))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(6))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(5))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(6))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(7))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "2") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(6))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(7))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(8))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "2.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "2.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(8))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBitFor(9))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBitFor(0))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "3") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "3" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBitFor(9))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array("-")
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array("-")
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "4") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array("-")
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array("-")
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array("-")
        End If
    End If
End Sub
Public Static Sub BitLam1(sFIG As String, sBit As String)
    'BITOLA
    aBit1 = Array("5/8" & Chr(34), "3/4" & Chr(34), "7/8" & Chr(34), "1.1/8" & Chr(34), "1.3/8" & Chr(34), "1.3/4" & Chr(34), "2.1/4" & Chr(34), "2.1/2" & Chr(34), "3" & Chr(34), "3.5/8" & Chr(34), "4.1/4" & Chr(34), "5.1/2" & Chr(34))
    aBit2 = Array("3/4" & Chr(34), "7/8" & Chr(34), "1" & Chr(34), "1.1/4" & Chr(34), "1.1/2" & Chr(34), "1.3/4" & Chr(34), "2.1/4" & Chr(34), "2.1/2" & Chr(34), "3" & Chr(34), "3.5/8" & Chr(34), "4.1/4" & Chr(34), "5.1/2" & Chr(34))
    Dim nN As Integer
    nN = Len(sFIG)
    If (Mid(sFIG, nN - 1, 1) = "N" Or Mid(sFIG, nN - 1, 1) = "B" Or Mid(sFIG, nN - 1, 1) = "T") Then
        If (Len(sBit) = 4 And Left(sBit, 3) = "1/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(0))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(1))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(1))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(2))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(2))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(3))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/2") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(3))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(4))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(4))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(5))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "1") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "1" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(5))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(6))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/4") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(6))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(7))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(7))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(8))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "2") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(8))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(9))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "2.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "2.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(9))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(10))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "3") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "3" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(10))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit1(11))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "4") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit1(8))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array("-")
        End If
    ElseIf Mid(sFIG, nN - 1, 1) = "S" Then
        If (Len(sBit) = 4 And Left(sBit, 3) = "1/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(0))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(1))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(2))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(1))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(2))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(3))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/8") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/8" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(2))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(3))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(4))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "1/2") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(3))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(4))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(5))
        ElseIf (Len(sBit) = 4 And Left(sBit, 3) = "3/4") Or _
           (Len(sBit) > 6 And Left(sBit, 6) = "3/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(4))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(5))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(6))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "1") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "1" & Chr(34) & " X") Then
            aQUA = Array("1")
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(5))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(6))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(7))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/4") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(6))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(7))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(8))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "1.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "1.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(7))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(8))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(9))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "2") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(8))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(9))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(10))
        ElseIf (Len(sBit) = 6 And Left(sBit, 5) = "2.1/2") Or _
           (Len(sBit) > 8 And Left(sBit, 8) = "2.1/2" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(9))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(10))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array(aBit2(11))
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "3") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "3" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(10))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array(aBit2(11))
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array("-")
        ElseIf (Len(sBit) = 2 And Left(sBit, 1) = "4") Or _
           (Len(sBit) > 4 And Left(sBit, 4) = "4" & Chr(34) & " X") Then
            If Mid(sFIG, nN, 1) = "3" Then aBIT = Array(aBit2(11))
            If Mid(sFIG, nN, 1) = "6" Then aBIT = Array("-")
            If Mid(sFIG, nN, 1) = "9" Then aBIT = Array("-")
        End If
    End If
End Sub
















'*****************************************************************
' RELACAO DE FUNCOES DE CADA INDICE PARA CONFIGURAR MP
'*****************************************************************
Public Static Function MP_Gaveta(sFIGURA As Variant, sBITOLA As Variant) As Variant
    LE_BITOLAS
    
    'QUANTIDADE
    aQUA = Array("1", "1", "1", "1", "2", "1", "0,030", "4", "4", "2", "4", "1", "1", "1", "1", "1")
    sBITTMP = Left(sBITOLA, Len(sBITOLA) - 1)
    If sBITTMP = "1" Or sBITTMP = "1.1/4" Or sBITTMP = "1.1/2" Or sBITTMP = "2" Then
        aQUA = Array("1", "1", "1", "1", "2", "1", "0,050", "4", "4", "2", "4", "1", "1", "1", "1", "1")
    End If
    
    'FIGURA
    aFIG = Array("corpo", "castelo", "CP-PREME", "CP-CUNHA", "CP-ANEL", "junta", "CP-GAXETA", "CP-PRICOR", "CP-PORCOR", "CP-PRIPRE", "CP-PORPRE", "HASTE", "CP-BMOVGAV", "CP-VOLGAV", "CP-PORVOLGAV", "CP-PLAIDE-GAV")
    
    'CORPO
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Then
        aFIG(0) = "CP-GAV-CORAPA-N8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "B" Then
        aFIG(0) = "CP-GAV-CORAPA-B8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        aFIG(0) = "CP-GAV-CORAPA-S8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        aFIG(0) = "CP-GAV-CORAPA-S5"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "T" Then
        aFIG(0) = "CP-GAV-CORAPA-T8"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Then
        aFIG(0) = "CP-GAV-CORAPA-N5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "B5" Then
        aFIG(0) = "CP-GAV-CORAPA-B5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "T5" Then
        aFIG(0) = "CP-GAV-CORAPA-T5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        aFIG(0) = "CP-GAV-CORAPA-S5"
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        aFIG(0) = "CP-GAV-CORAPA-S5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        aFIG(0) = "CP-GAV-CORAPA-S8"
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        aFIG(0) = "CP-GAV-CORAPA-S8"
    End If
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Or _
       Mid(sFIGURA, 5, 1) = "B" Or _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "T" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        aFIG(1) = "CP-GAV-CASAPA-8"
        aFIG(5) = "CP-JUNESP"
        aFIG(11) = "CP-HASGAV-8"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        aFIG(1) = "CP-GAV-CASAPA-8"
        aFIG(5) = "CP-JUNESP"
        aFIG(11) = "CP-HASGAV-8"
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        aFIG(1) = "CP-GAV-CASAPA-8"
        aFIG(5) = "CP-JUNESP"
        aFIG(11) = "CP-HASGAV-8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        aFIG(1) = "CP-GAV-CASAPA-5"
        aFIG(5) = "CP-ANEL-RTJ"
        aFIG(11) = "CP-HASGAV-5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Or _
       Mid(sFIGURA, 5, 2) = "B5" Or _
       Mid(sFIGURA, 5, 2) = "T5" Or _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        aFIG(1) = "CP-GAV-CASAPA-5"
        aFIG(5) = "CP-ANEL-RTJ"
        aFIG(11) = "CP-HASGAV-5"
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        aFIG(1) = "CP-GAV-CASAPA-5"
        aFIG(5) = "CP-ANEL-RTJ"
        aFIG(11) = "CP-HASGAV-5"
    End If
    
    'NOMES
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Or _
       Mid(sFIGURA, 5, 1) = "B" Or _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "T" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Cunha", "Anel", "Junta Espirotálica", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Cunha", "Anel", "Junta Espirotálica", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Cunha", "Anel", "Junta Espirotálica", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Cunha", "Anel", "Anel RTJ", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Or _
       Mid(sFIGURA, 5, 2) = "B5" Or _
       Mid(sFIGURA, 5, 2) = "T5" Or _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Cunha", "Anel", "Anel RTJ", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Cunha", "Anel", "Anel RTJ", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Placa de Identificação")
    End If
    
    'BITOLAS
    sBITTMP = Trim(sBITOLA)
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Or _
       Mid(sFIGURA, 5, 1) = "B" Or _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "T" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        GAV800 sBITTMP
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        GAV800 sBITTMP
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        GAV800 sBITTMP
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        GAV1500 sBITTMP
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Or _
       Mid(sFIGURA, 5, 2) = "B5" Or _
       Mid(sFIGURA, 5, 2) = "T5" Or _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        GAV1500 sBITTMP
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        GAV1500 sBITTMP
    End If
    
    'INTERNOS
    If Mid(sFIGURA, 3, 1) = "1" Or _
       Mid(sFIGURA, 3, 1) = "5" Or _
       Mid(sFIGURA, 3, 1) = "7" Then
        aFIG(3) = "CP-CUNHA-XU"
    ElseIf Mid(sFIGURA, 3, 1) = "2" Or _
       Mid(sFIGURA, 3, 1) = "8" Or _
       Mid(sFIGURA, 3, 1) = "9" Then
        aFIG(3) = "CP-CUNHA-XU"
        aFIG(4) = "CP-ANEL-XU"
    End If
    If Mid(sFIGURA, 3, 1) = "1" Or _
       Mid(sFIGURA, 3, 1) = "5" Or _
       Mid(sFIGURA, 3, 1) = "7" Or _
       Mid(sFIGURA, 3, 1) = "2" Or _
       Mid(sFIGURA, 3, 1) = "8" Or _
       Mid(sFIGURA, 3, 1) = "9" Then
        ReDim Preserve aQUA(UBound(aQUA) + 1)
        ReDim Preserve aFIG(UBound(aFIG) + 1)
        ReDim Preserve aNOM(UBound(aNOM) + 1)
        ReDim Preserve aBIT(UBound(aBIT) + 1)
        aQUA(UBound(aQUA)) = "1"
        aBIT(UBound(aBIT)) = "2,5 mm"
        If sBITTMP = "1" Or sBITTMP = "1.1/4" Or sBITTMP = "1.1/2" Or sBITTMP = "2" Then
            aBIT(UBound(aBIT)) = "3,25 mm"
        End If
        aNOM(UBound(aNOM)) = "Solda do Revestimento"
        aFIG(UBound(aFIG)) = "MP-SOLDA-REVE"
    End If
    
    'verifica se é alguma valvula com flange ou bw
    If (Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or Mid(sFIGURA, 5, 1) = "6" Or _
       Mid(sFIGURA, 5, 1) = "5" Or Mid(sFIGURA, 5, 1) = "9") _
       Or _
       (Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Or Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6") _
       Or _
       (Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80") _
       Or _
       (Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS") Then
        'redimensionas os vetores
        ReDim Preserve aQUA(UBound(aQUA) + 2)
        ReDim Preserve aFIG(UBound(aFIG) + 2)
        ReDim Preserve aNOM(UBound(aNOM) + 2)
        ReDim Preserve aBIT(UBound(aBIT) + 2)
        'altera valores dos vetores
        aQUA(UBound(aQUA) - 1) = "2" 'flanges ou pontas
        aBIT(UBound(aBIT) - 1) = sBITTMP
        aQUA(UBound(aQUA)) = "2" 'solda
        aBIT(UBound(aBIT)) = "1 mm"
        aNOM(UBound(aNOM)) = "Solda da Adaptação"
        aFIG(UBound(aFIG)) = "MP-SOLDA-ADAP"
        'verifica se é bw ou flangeada
        If (Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W40" Or Mid(sFIGURA, 5, 3) = "W80") Or _
           (Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "W160" Or Mid(sFIGURA, 5, 4) = "WXXS") Then            'bw
            aNOM(UBound(aNOM) - 1) = "Ponta biselada"
            If Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W40" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GAV-40"
            ElseIf Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W80" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GAV-80"
            ElseIf Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "W160" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GAV-160"
            ElseIf Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "WXXS" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GAV-XXS"
            End If
        Else 'flangeada
            aNOM(UBound(aNOM) - 1) = "Flange"
            If Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "1" Then
                aFIG(UBound(aFIG) - 1) = "50-1/S40"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "3" Then
                aFIG(UBound(aFIG) - 1) = "50-3/S40"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "6" Then
                aFIG(UBound(aFIG) - 1) = "50-6/S80"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "9" Then
                aFIG(UBound(aFIG) - 1) = "50-9/S160"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "5" Then
                aFIG(UBound(aFIG) - 1) = "50-5/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F5" Then
                aFIG(UBound(aFIG) - 1) = "50-F5/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F9" Then
                aFIG(UBound(aFIG) - 1) = "50-F9/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F6" Then
                aFIG(UBound(aFIG) - 1) = "50-F6/S80"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "R5" Then
                aFIG(UBound(aFIG) - 1) = "50-R5/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "R9" Then
                aFIG(UBound(aFIG) - 1) = "50-R9/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "R6" Then
                aFIG(UBound(aFIG) - 1) = "50-R6/S80"
            End If
        End If
    End If
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_Gaveta = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_Gaveta = ""
    End If
End Function
Public Static Function MP_Globo(sFIGURA As Variant, sBITOLA As Variant) As Variant
    LE_BITOLAS
    
    'QUANTIDADE
    aQUA = Array("1", "1", "1", "1", "2", "1", "0,030", "4", "4", "2", "4", "1", "1", "1", "1", "1", "1")
    sBITTMP = Left(sBITOLA, Len(sBITOLA) - 1)
    If sBITTMP = "1" Or sBITTMP = "1.1/4" Or sBITTMP = "1.1/2" Or sBITTMP = "2" Then
        aQUA = Array("1", "1", "1", "1", "2", "1", "0,050", "4", "4", "2", "4", "1", "1", "1", "1", "1", "1")
    End If
    
    'FIGURA
    aFIG = Array("corpo", "castelo", "CP-PREME", "CP-CONSED", "CP-SEDE", "junta", "CP-GAXETA", "CP-PRICOR", "CP-PORCOR", "CP-PRIPRE", "CP-PORPRE", "haste", "CP-BMOVGLO", "CP-VOLGLO", "CP-PORVOLGLO", "CP-ARRVOLGLO", "CP-PLAIDE-GLO")
    
    'CORPO
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Then
        aFIG(0) = "CP-GLO-CORAPA-N8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "B" Then
        aFIG(0) = "CP-GLO-CORAPA-B8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        aFIG(0) = "CP-GLO-CORAPA-S8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        aFIG(0) = "CP-GLO-CORAPA-S5"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "T" Then
        aFIG(0) = "CP-GLO-CORAPA-T8"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Then
        aFIG(0) = "CP-GLO-CORAPA-N5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "B5" Then
        aFIG(0) = "CP-GLO-CORAPA-B5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "T5" Then
        aFIG(0) = "CP-GLO-CORAPA-T5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        aFIG(0) = "CP-GLO-CORAPA-S5"
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        aFIG(0) = "CP-GLO-CORAPA-S5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        aFIG(0) = "CP-GLO-CORAPA-S8"
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        aFIG(0) = "CP-GLO-CORAPA-S8"
    End If
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Or _
       Mid(sFIGURA, 5, 1) = "B" Or _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "T" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        aFIG(1) = "CP-GLO-CASAPA-8"
        aFIG(5) = "CP-JUNESP"
        aFIG(11) = "CP-HASGLO-8"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        aFIG(1) = "CP-GLO-CASAPA-8"
        aFIG(5) = "CP-JUNESP"
        aFIG(11) = "CP-HASGLO-8"
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        aFIG(1) = "CP-GLO-CASAPA-8"
        aFIG(5) = "CP-JUNESP"
        aFIG(11) = "CP-HASGLO-8"
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        aFIG(1) = "CP-GLO-CASAPA-5"
        aFIG(5) = "CP-ANEL-RTJ"
        aFIG(11) = "CP-HASGLO-5"
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Or _
       Mid(sFIGURA, 5, 2) = "B5" Or _
       Mid(sFIGURA, 5, 2) = "T5" Or _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        aFIG(1) = "CP-GLO-CASAPA-5"
        aFIG(5) = "CP-ANEL-RTJ"
        aFIG(11) = "CP-HASGLO-5"
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        aFIG(1) = "CP-GLO-CASAPA-5"
        aFIG(5) = "CP-ANEL-RTJ"
        aFIG(11) = "CP-HASGLO-5"
    End If
    
    'NOMES
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Or _
       Mid(sFIGURA, 5, 1) = "B" Or _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "T" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Contra-Sede", "Sede", "Junta Espirotálica", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Arruela Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Contra-Sede", "Sede", "Junta Espirotálica", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Arruela Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Contra-Sede", "Sede", "Junta Espirotálica", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Arruela Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Contra-Sede", "Sede", "Anel RTJ", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Arruela Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Or _
       Mid(sFIGURA, 5, 2) = "B5" Or _
       Mid(sFIGURA, 5, 2) = "T5" Or _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Contra-Sede", "Sede", "Anel RTJ", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Arruela Porca Volante", "Placa de Identificação")
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        aNOM = Array("Corpo Aparafusado", "Castelo Aparafusado", "Preme", "Contra-Sede", "Sede", "Anel RTJ", "Gaxeta", "Prisioneiro Corpo", "Porca Corpo", "Prisioneiro Preme", "Porca Preme", "Haste", "Bucha Movimento", "Volante", "Porca Volante", "Arruela Porca Volante", "Placa de Identificação")
    End If
    
    'BITOLAS
    sBITTMP = Trim(sBITOLA)
    If Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "N" Or _
       Mid(sFIGURA, 5, 1) = "B" Or _
       Mid(sFIGURA, 5, 1) = "S" Or _
       Mid(sFIGURA, 5, 1) = "T" Or _
       Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or _
       Mid(sFIGURA, 5, 1) = "6" Then
        GLO800 sBITTMP
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6" Then
        GLO800 sBITTMP
    ElseIf Len(sFIGURA) = 7 And _
       Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80" Then
        GLO800 sBITTMP
    ElseIf Len(sFIGURA) = 5 And _
       Mid(sFIGURA, 5, 1) = "5" Or _
       Mid(sFIGURA, 5, 1) = "9" Then
        GLO1500 sBITTMP
    ElseIf Len(sFIGURA) = 6 And _
       Mid(sFIGURA, 5, 2) = "N5" Or _
       Mid(sFIGURA, 5, 2) = "B5" Or _
       Mid(sFIGURA, 5, 2) = "T5" Or _
       Mid(sFIGURA, 5, 2) = "S5" Or _
       Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or _
       Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Then
        GLO1500 sBITTMP
    ElseIf Len(sFIGURA) = 8 And _
       Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS" Then
        GLO1500 sBITTMP
    End If
    
    'INTERNOS
    If Mid(sFIGURA, 3, 1) = "1" Or _
       Mid(sFIGURA, 3, 1) = "5" Or _
       Mid(sFIGURA, 3, 1) = "7" Then
        aFIG(3) = "CP-CONSED-XU"
    ElseIf Mid(sFIGURA, 3, 1) = "2" Or _
       Mid(sFIGURA, 3, 1) = "8" Or _
       Mid(sFIGURA, 3, 1) = "9" Then
        aFIG(3) = "CP-CONSED-XU"
        aFIG(4) = "CP-SEDE-XU"
    End If
    If Mid(sFIGURA, 3, 1) = "1" Or _
       Mid(sFIGURA, 3, 1) = "5" Or _
       Mid(sFIGURA, 3, 1) = "7" Or _
       Mid(sFIGURA, 3, 1) = "2" Or _
       Mid(sFIGURA, 3, 1) = "8" Or _
       Mid(sFIGURA, 3, 1) = "9" Then
        ReDim Preserve aQUA(UBound(aQUA) + 1)
        ReDim Preserve aFIG(UBound(aFIG) + 1)
        ReDim Preserve aNOM(UBound(aNOM) + 1)
        ReDim Preserve aBIT(UBound(aBIT) + 1)
        aQUA(UBound(aQUA)) = "1"
        aBIT(UBound(aBIT)) = "2,5 mm"
        If sBITTMP = "1" Or sBITTMP = "1.1/4" Or sBITTMP = "1.1/2" Or sBITTMP = "2" Then
            aBIT(UBound(aBIT)) = "3,25 mm"
        End If
        aNOM(UBound(aNOM)) = "Solda do Revestimento"
        aFIG(UBound(aFIG)) = "MP-SOLDA-REVE"
    End If
    
    'verifica se é alguma valvula com flange ou bw
    If (Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "1" Or _
       Mid(sFIGURA, 5, 1) = "3" Or Mid(sFIGURA, 5, 1) = "6" Or _
       Mid(sFIGURA, 5, 1) = "5" Or Mid(sFIGURA, 5, 1) = "9") _
       Or _
       (Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F5" Or _
       Mid(sFIGURA, 5, 2) = "F9" Or Mid(sFIGURA, 5, 2) = "R5" Or _
       Mid(sFIGURA, 5, 2) = "R9" Or Mid(sFIGURA, 5, 2) = "F6" Or _
       Mid(sFIGURA, 5, 2) = "R6") _
       Or _
       (Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W40" Or _
       Mid(sFIGURA, 5, 3) = "W80") _
       Or _
       (Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "W160" Or _
       Mid(sFIGURA, 5, 4) = "WXXS") Then
        'redimensionas os vetores
        ReDim Preserve aQUA(UBound(aQUA) + 2)
        ReDim Preserve aFIG(UBound(aFIG) + 2)
        ReDim Preserve aNOM(UBound(aNOM) + 2)
        ReDim Preserve aBIT(UBound(aBIT) + 2)
        'altera valores dos vetores
        aQUA(UBound(aQUA) - 1) = "2" 'flanges ou pontas
        aBIT(UBound(aBIT) - 1) = sBITTMP
        aQUA(UBound(aQUA)) = "2" 'solda
        aBIT(UBound(aBIT)) = "1 mm"
        aNOM(UBound(aNOM)) = "Solda da Adaptação"
        aFIG(UBound(aFIG)) = "MP-SOLDA-ADAP"
        'verifica se é bw ou flangeada
        If (Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W40" Or Mid(sFIGURA, 5, 3) = "W80") Or _
           (Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "W160" Or Mid(sFIGURA, 5, 4) = "WXXS") Then            'bw
            aNOM(UBound(aNOM) - 1) = "Ponta biselada"
            If Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W40" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GLO-40"
            ElseIf Len(sFIGURA) = 7 And Mid(sFIGURA, 5, 3) = "W80" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GLO-80"
            ElseIf Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "W160" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GLO-160"
            ElseIf Len(sFIGURA) = 8 And Mid(sFIGURA, 5, 4) = "WXXS" Then
                aFIG(UBound(aFIG) - 1) = "CP-PONBIS-GLO-XXS"
            End If
        Else 'flangeada
            aNOM(UBound(aNOM) - 1) = "Flange"
            If Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "1" Then
                aFIG(UBound(aFIG) - 1) = "50-1/S40"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "3" Then
                aFIG(UBound(aFIG) - 1) = "50-3/S40"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "6" Then
                aFIG(UBound(aFIG) - 1) = "50-6/S80"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "9" Then
                aFIG(UBound(aFIG) - 1) = "50-9/S160"
            ElseIf Len(sFIGURA) = 5 And Mid(sFIGURA, 5, 1) = "5" Then
                aFIG(UBound(aFIG) - 1) = "50-5/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F5" Then
                aFIG(UBound(aFIG) - 1) = "50-F5/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F9" Then
                aFIG(UBound(aFIG) - 1) = "50-F9/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "F6" Then
                aFIG(UBound(aFIG) - 1) = "50-F6/S80"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "R5" Then
                aFIG(UBound(aFIG) - 1) = "50-R5/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "R9" Then
                aFIG(UBound(aFIG) - 1) = "50-R9/S160"
            ElseIf Len(sFIGURA) = 6 And Mid(sFIGURA, 5, 2) = "R6" Then
                aFIG(UBound(aFIG) - 1) = "50-R6/S80"
            End If
        End If
    End If
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_Globo = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_Globo = ""
    End If
End Function
Public Static Function MP_BUCHA(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    If Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then
        aQUA = Array("0,019")
        aBIT = Array("5/8" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then
        aQUA = Array("0,021")
        aBIT = Array("11/16" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then
        aQUA = Array("0,023")
        aBIT = Array("7/8" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then
        aQUA = Array("0,027")
        aBIT = Array("1.1/16" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("1" & Chr(34)) Then
        aQUA = Array("0,030")
        aBIT = Array("1.3/8" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then
        aQUA = Array("0,033")
        aBIT = Array("1.3/4" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then
        aQUA = Array("0,034")
        aBIT = Array("2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("2" & Chr(34)) Then
        aQUA = Array("0,036")
        aBIT = Array("2.1/2" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then
        aQUA = Array("0,042")
        aBIT = Array("3" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("3" & Chr(34)) Then
        aQUA = Array("0,044")
        aBIT = Array("3.1/2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("4" & Chr(34)) Then
        aQUA = Array("0,050")
        aBIT = Array("4" & Chr(34))
    End If
    
    'FIGURA
    aFIG = Array("MP-SEX")
    'NOME
    aNOM = Array("Sextavado")
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_BUCHA = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_BUCHA = ""
    End If
End Function
Public Static Function MP_CAPS(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitLam1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-RED")
    'NOME
    aNOM = Array("Redondo")
    'QUANTIDADE
    Dim nN As Integer
    nN = Len(sFIGURA)
    If (Mid(sFIGURA, nN - 1, 1) = "N" Or Mid(sFIGURA, nN - 1, 1) = "B" Or Mid(sFIGURA, nN - 1, 1) = "T") Then
        If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then aQUA = Array("0,022")
        If Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then aQUA = Array("0,028")
        If Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then aQUA = Array("0,028")
        If Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then aQUA = Array("0,035")
        If Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then aQUA = Array("0,040")
        If Left(sBITOLA, 2) = ("1" & Chr(34)) Then aQUA = Array("0,044")
        If Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then aQUA = Array("0,047")
        If Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then aQUA = Array("0,047")
        If Left(sBITOLA, 2) = ("2" & Chr(34)) Then aQUA = Array("0,051")
        If Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then aQUA = Array("0,063")
        If Left(sBITOLA, 2) = ("3" & Chr(34)) Then aQUA = Array("0,068")
        If Left(sBITOLA, 2) = ("4" & Chr(34)) Then aQUA = Array("0,071")
    ElseIf Mid(sFIGURA, nN - 1, 1) = "S" Then
        If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then aQUA = Array("0,019")
        If Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then aQUA = Array("0,020")
        If Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then aQUA = Array("0,020")
        If Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then aQUA = Array("0,022")
        If Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then aQUA = Array("0,025")
        If Left(sBITOLA, 2) = ("1" & Chr(34)) Then aQUA = Array("0,028")
        If Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then aQUA = Array("0,028")
        If Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then aQUA = Array("0,029")
        If Left(sBITOLA, 2) = ("2" & Chr(34)) Then aQUA = Array("0,034")
        If Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then aQUA = Array("0,038")
        If Left(sBITOLA, 2) = ("3" & Chr(34)) Then aQUA = Array("0,041")
        If Left(sBITOLA, 2) = ("4" & Chr(34)) Then aQUA = Array("0,048")
    End If
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_CAPS = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_CAPS = ""
    End If
End Function
Public Static Function MP_COT90(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitFor1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-CT9")
    'NOME
    aNOM = Array("Cotovelo 90º")
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_COT90 = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_COT90 = ""
    End If
End Function
Public Static Function MP_COTMF(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitFor1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-CTM")
    'NOME
    aNOM = Array("Cotovelo 90º M/F")
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_COTMF = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_COTMF = ""
    End If
End Function
Public Static Function MP_COT45(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitFor1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-CT4")
    'NOME
    aNOM = Array("Cotovelo 45º")
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_COT45 = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_COT45 = ""
    End If
End Function
Public Static Function MP_LUVA(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitLam1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-RED")
    'NOME
    aNOM = Array("Redondo")
    'QUANTIDADE
    Dim nN As Integer
    nN = Len(sFIGURA)
    If (Mid(sFIGURA, nN - 1, 1) = "N" Or Mid(sFIGURA, nN - 1, 1) = "B" Or Mid(sFIGURA, nN - 1, 1) = "T") Then
        If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then aQUA = Array("0,035")
        If sBITOLA = "1/4" & Chr(34) Then aQUA = Array("0,038")
        If Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then aQUA = Array("0,041")
        If Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then aQUA = Array("0,051")
        If Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then aQUA = Array("0,054")
        If Left(sBITOLA, 2) = ("1" & Chr(34)) Then aQUA = Array("0,063")
        If Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then aQUA = Array("0,070")
        If Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then aQUA = Array("0,082")
        If Left(sBITOLA, 2) = ("2" & Chr(34)) Then aQUA = Array("0,089")
        If Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then aQUA = Array("0,095")
        If Left(sBITOLA, 2) = ("3" & Chr(34)) Then aQUA = Array("0,111")
        If Left(sBITOLA, 2) = ("4" & Chr(34)) Then aQUA = Array("0,124")
    ElseIf Mid(sFIGURA, nN - 1, 1) = "S" Then
        If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then aQUA = Array("0,030")
        If Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then aQUA = Array("0,030")
        If Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then aQUA = Array("0,030")
        If Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then aQUA = Array("0,033")
        If Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then aQUA = Array("0,039")
        If Left(sBITOLA, 2) = ("1" & Chr(34)) Then aQUA = Array("0,042")
        If Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then aQUA = Array("0,042")
        If Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then aQUA = Array("0,042")
        If Left(sBITOLA, 2) = ("2" & Chr(34)) Then aQUA = Array("0,054")
        If Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then aQUA = Array("0,054")
        If Left(sBITOLA, 2) = ("3" & Chr(34)) Then aQUA = Array("0,054")
        If Left(sBITOLA, 2) = ("4" & Chr(34)) Then aQUA = Array("0,061")
    End If
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_LUVA = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_LUVA = ""
    End If
End Function
Public Static Function MP_MEIALUVA(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitLam1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-RED")
    'NOME
    aNOM = Array("Redondo")
    'QUANTIDADE
    Dim nN As Integer
    nN = Len(sFIGURA)
    If (Mid(sFIGURA, nN - 1, 1) = "N" Or Mid(sFIGURA, nN - 1, 1) = "B" Or Mid(sFIGURA, nN - 1, 1) = "T") Then
        If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then aQUA = Array("0,019")
        If Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then aQUA = Array("0,021")
        If Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then aQUA = Array("0,022")
        If Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then aQUA = Array("0,027")
        If Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then aQUA = Array("0,029")
        If Left(sBITOLA, 2) = ("1" & Chr(34)) Then aQUA = Array("0,033")
        If Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then aQUA = Array("0,037")
        If Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then aQUA = Array("0,043")
        If Left(sBITOLA, 2) = ("2" & Chr(34)) Then aQUA = Array("0,046")
        If Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then aQUA = Array("0,049")
        If Left(sBITOLA, 2) = ("3" & Chr(34)) Then aQUA = Array("0,057")
        If Left(sBITOLA, 2) = ("4" & Chr(34)) Then aQUA = Array("0,064")
    ElseIf Mid(sFIGURA, nN - 1, 1) = "S" Then
        If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then aQUA = Array("0,029")
        If Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then aQUA = Array("0,029")
        If Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then aQUA = Array("0,029")
        If Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then aQUA = Array("0,035")
        If Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then aQUA = Array("0,040")
        If Left(sBITOLA, 2) = ("1" & Chr(34)) Then aQUA = Array("0,044")
        If Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then aQUA = Array("0,046")
        If Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then aQUA = Array("0,047")
        If Left(sBITOLA, 2) = ("2" & Chr(34)) Then aQUA = Array("0,060")
        If Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then aQUA = Array("0,061")
        If Left(sBITOLA, 2) = ("3" & Chr(34)) Then aQUA = Array("0,063")
        If Left(sBITOLA, 2) = ("4" & Chr(34)) Then aQUA = Array("0,070")
    End If
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_MEIALUVA = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_MEIALUVA = ""
    End If
End Function
Public Static Function MP_NIPLE(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then
        aQUA = Array("0,029")
        aBIT = Array("7/16" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then
        aQUA = Array("0,030")
        aBIT = Array("5/8" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then
        aQUA = Array("0,033")
        aBIT = Array("3/4" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then
        aQUA = Array("0,042")
        aBIT = Array("7/8" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then
        aQUA = Array("0,044")
        aBIT = Array("1.1/16" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("1" & Chr(34)) Then
        aQUA = Array("0,049")
        aBIT = Array("1.3/8" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then
        aQUA = Array("0,053")
        aBIT = Array("1.3/4" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then
        aQUA = Array("0,054")
        aBIT = Array("2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("2" & Chr(34)) Then
        aQUA = Array("0,059")
        aBIT = Array("2.1/2" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then
        aQUA = Array("0,069")
        aBIT = Array("3" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("3" & Chr(34)) Then
        aQUA = Array("0,073")
        aBIT = Array("3.1/2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("4" & Chr(34)) Then
        aQUA = Array("0,082")
        aBIT = Array("4.5/8" & Chr(34))
    End If
    
    'FIGURA
    aFIG = Array("MP-SEX")
    'NOME
    aNOM = Array("Sextavado")
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_NIPLE = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_NIPLE = ""
    End If
End Function
Public Static Function MP_PRED(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then
        aQUA = Array("0,039")
        aBIT = Array("7/16" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then
        aQUA = Array("0,045")
        aBIT = Array("9/16" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then
        aQUA = Array("0,045")
        aBIT = Array("11/16" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then
        aQUA = Array("0,048")
        aBIT = Array("7/8" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then
        aQUA = Array("0,048")
        aBIT = Array("1.1/16" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("1" & Chr(34)) Then
        aQUA = Array("0,055")
        aBIT = Array("1.5/16" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then
        aQUA = Array("0,055")
        aBIT = Array("1.11/16" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then
        aQUA = Array("0,055")
        aBIT = Array("2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("2" & Chr(34)) Then
        aQUA = Array("0,068")
        aBIT = Array("2.3/8" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then
        aQUA = Array("0,075")
        aBIT = Array("2.7/8" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("3" & Chr(34)) Then
        aQUA = Array("0,075")
        aBIT = Array("3.1/2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("4" & Chr(34)) Then
        aQUA = Array("0,080")
        aBIT = Array("4.1/2" & Chr(34))
    End If
    
    'FIGURA
    aFIG = Array("MP-RED")
    'NOME
    aNOM = Array("Redondo")
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_PRED = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_PRED = ""
    End If
End Function
Public Static Function MP_PQUA(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE
    aQUA = Array("1")
    'BITOLA
    If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then aBIT = Array("1/8" & Chr(34))
    If Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then aBIT = Array("1/4" & Chr(34))
    If Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then aBIT = Array("3/8" & Chr(34))
    If Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then aBIT = Array("1/2" & Chr(34))
    If Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then aBIT = Array("3/4" & Chr(34))
    If Left(sBITOLA, 2) = ("1" & Chr(34)) Then aBIT = Array("1" & Chr(34))
    If Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then aBIT = Array("1.1/4" & Chr(34))
    If Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then aBIT = Array("1.1/2" & Chr(34))
    If Left(sBITOLA, 2) = ("2" & Chr(34)) Then aBIT = Array("2" & Chr(34))
    If Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then aBIT = Array("2.1/2" & Chr(34))
    If Left(sBITOLA, 2) = ("3" & Chr(34)) Then aBIT = Array("3" & Chr(34))
    If Left(sBITOLA, 2) = ("4" & Chr(34)) Then aBIT = Array("4" & Chr(34))
    
    'FIGURA
    aFIG = Array("MP-PQD")
    'NOME
    aNOM = Array("Plug Quadrado")
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_PQUA = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_PQUA = ""
    End If
End Function
Public Static Function MP_PSEX(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    If Left(sBITOLA, 4) = ("1/8" & Chr(34)) Then
        aQUA = Array("0,020")
        aBIT = Array("7/16" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("1/4" & Chr(34)) Then
        aQUA = Array("0,022")
        aBIT = Array("5/8" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/8" & Chr(34)) Then
        aQUA = Array("0,025")
        aBIT = Array("11/16" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("1/2" & Chr(34)) Then
        aQUA = Array("0,027")
        aBIT = Array("7/8" & Chr(34))
    ElseIf Left(sBITOLA, 4) = ("3/4" & Chr(34)) Then
        aQUA = Array("0,031")
        aBIT = Array("1.1/16" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("1" & Chr(34)) Then
        aQUA = Array("0,034")
        aBIT = Array("1.3/8" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/4" & Chr(34)) Then
        aQUA = Array("0,040")
        aBIT = Array("1.3/4" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("1.1/2" & Chr(34)) Then
        aQUA = Array("0,042")
        aBIT = Array("2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("2" & Chr(34)) Then
        aQUA = Array("0,044")
        aBIT = Array("2.1/2" & Chr(34))
    ElseIf Left(sBITOLA, 6) = ("2.1/2" & Chr(34)) Then
        aQUA = Array("0,051")
        aBIT = Array("3" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("3" & Chr(34)) Then
        aQUA = Array("0,057")
        aBIT = Array("3.1/2" & Chr(34))
    ElseIf Left(sBITOLA, 2) = ("4" & Chr(34)) Then
        aQUA = Array("0,063")
        aBIT = Array("4.5/8" & Chr(34))
    End If
    
    'FIGURA
    aFIG = Array("MP-SEX")
    'NOME
    aNOM = Array("Sextavado")
    
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_PSEX = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_PSEX = ""
    End If
End Function
Public Static Function MP_TEE90(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitFor1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-TE9")
    'NOME
    aNOM = Array("Tê 90º")
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_TEE90 = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_TEE90 = ""
    End If
End Function
Public Static Function MP_CRUZETA(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitFor1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-CZT")
    'NOME
    aNOM = Array("Cruzeta")
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_CRUZETA = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_CRUZETA = ""
    End If
End Function
Public Static Function MP_TEETUB(sFIGURA As Variant, sBITOLA As Variant) As Variant
    'QUANTIDADE E BITOLA
    sFT = sFIGURA
    sBT = sBITOLA
    BitFor1 sFT, sBT
    'FIGURA
    aFIG = Array("MP-TE9")
    'NOME
    aNOM = Array("Tê 90º")
    'retorna vetores
    If IsArray(aFIG) And IsArray(aQUA) And IsArray(aNOM) And IsArray(aBIT) Then
        MP_TEETUB = Array(aQUA, aFIG, aNOM, aBIT)
    Else
        MP_TEETUB = ""
    End If
End Function
Public Static Sub ConfigCP()
    Dim NovaFig As String, aMP As Variant, sNome As String
    Tela_Cfg_MateriaPrima.TelaEmEspera True
    With Tela_Cfg_MateriaPrima.DLL_BD
        .BDSIS_TBEST.MoveFirst
        Tela_Cfg_MateriaPrima.BS.SimpleText = ""
        Tela_Cfg_MateriaPrima.BP.Max = .BDSIS_TBEST.RecordCount
        Tela_Cfg_MateriaPrima.BP.Value = 0
        Do While Not .BDSIS_TBEST.EOF
            Tela_Cfg_MateriaPrima.BS.SimpleText = Trim(.BDSIS_TBEST_CPFIG.Value) & " de " & Trim(.BDSIS_TBEST_CPBIT.Value) & " em " & Trim(.BDSIS_TBEST_CPMAT.Value)
            If .BDSIS_TBEST_CPINQ.Value = 0 And _
               .BDSIS_TBEST_CPINP.Value = 0 And _
               .BDSIS_TBEST_CPINN.Value = 0 And _
               .BDSIS_TBEST_CPINB.Value = 0 And _
               .BDSIS_TBEST_CPINM.Value = 0 Then
                If Left(.BDSIS_TBEST_CPFIG.Value, 2) = "CP" Then
                    NovaFig = "PA" & Mid(.BDSIS_TBEST_CPFIG.Value, 3, Len(.BDSIS_TBEST_CPFIG.Value))
                    sNome = Trim(.BDSIS_TBEID_CPDRE.Value) & " " & Trim(.BDSIS_TBEID_CPTRE.Value)
                    aMP = Tela_Cfg_MateriaPrima.ProcuraIndicesMP("1", NovaFig, sNome, .BDSIS_TBEST_CPBIT.Value, .BDSIS_TBEST_CPMAT.Value)
                    'altera indices da MP
                    .BDSIS_TBEST.Edit
                    .BDSIS_TBEST_CPINQ.Value = aMP(0)
                    .BDSIS_TBEST_CPINP.Value = aMP(1)
                    .BDSIS_TBEST_CPINN.Value = aMP(2)
                    .BDSIS_TBEST_CPINB.Value = aMP(3)
                    .BDSIS_TBEST_CPINM.Value = aMP(4)
                    .BDSIS_TBEST.Update
                End If
            End If
            Tela_Cfg_MateriaPrima.BP.Value = Tela_Cfg_MateriaPrima.BP.Value + 1
            .BDSIS_TBEST.MoveNext
        Loop
    End With
    Tela_Cfg_MateriaPrima.BS.SimpleText = ""
    Tela_Cfg_MateriaPrima.BP.Value = 0
    Tela_Cfg_MateriaPrima.TelaEmEspera False
End Sub
Public Static Sub ConfigPA()
    Dim NovaFig As String, aMP As Variant, sNome As String, aCMP As Variant, sPeca As String, sQuan As String, sBit As String
    Tela_Cfg_MateriaPrima.TelaEmEspera True
    With Tela_Cfg_MateriaPrima.DLL_BD
        .BDSIS_TBEST.MoveFirst
        Tela_Cfg_MateriaPrima.BS.SimpleText = ""
        Tela_Cfg_MateriaPrima.BP.Max = .BDSIS_TBEST.RecordCount
        Tela_Cfg_MateriaPrima.BP.Value = 0
        LE_BITOLAS
        Do While Not .BDSIS_TBEST.EOF
            Tela_Cfg_MateriaPrima.BS.SimpleText = Trim(.BDSIS_TBEST_CPFIG.Value) & " de " & Trim(.BDSIS_TBEST_CPBIT.Value) & " em " & Trim(.BDSIS_TBEST_CPMAT.Value)
            If .BDSIS_TBEST_CPINQ.Value = 0 And _
               .BDSIS_TBEST_CPINP.Value = 0 And _
               .BDSIS_TBEST_CPINN.Value = 0 And _
               .BDSIS_TBEST_CPINB.Value = 0 And _
               .BDSIS_TBEST_CPINM.Value = 0 Then
                If Left(.BDSIS_TBEST_CPFIG.Value, 2) = "PA" Then
                    aCMP = ConfigPA_Aux1(.BDSIS_TBEST_CPFIG.Value, .BDSIS_TBEST_CPBIT.Value)
                    If aCMP(0) <> "" And aCMP(1) <> "" And aCMP(2) <> "" Then
                        sNome = Trim(.BDSIS_TBEID_CPDNO.Value) & " em bruto"
                        sQuan = aCMP(0)
                        sPeca = aCMP(1)
                        sBit = aCMP(2)
                        aMP = Tela_Cfg_MateriaPrima.ProcuraIndicesMP(sQuan, sPeca, sNome, sBit, .BDSIS_TBEST_CPMAT.Value)
                        'altera indices da MP
                        .BDSIS_TBEST.Edit
                        .BDSIS_TBEST_CPINQ.Value = aMP(0)
                        .BDSIS_TBEST_CPINP.Value = aMP(1)
                        .BDSIS_TBEST_CPINN.Value = aMP(2)
                        .BDSIS_TBEST_CPINB.Value = aMP(3)
                        .BDSIS_TBEST_CPINM.Value = aMP(4)
                        .BDSIS_TBEST.Update
                    End If
                End If
            End If
            Tela_Cfg_MateriaPrima.BP.Value = Tela_Cfg_MateriaPrima.BP.Value + 1
            .BDSIS_TBEST.MoveNext
        Loop
    End With
    Tela_Cfg_MateriaPrima.BS.SimpleText = ""
    Tela_Cfg_MateriaPrima.BP.Value = 0
    Tela_Cfg_MateriaPrima.TelaEmEspera False
End Sub
Private Static Function ConfigPA_Aux1(COM As String, BIT As String) As Variant
    Dim sPec As String, sBit As String, sQua As String
    sPec = ""
    sBit = ""
    sQua = ""
    If COM = "" Or BIT = "" Then
        ConfigPA_Aux1 = Array(sQua, sPec, sBit)
        Exit Function
    End If
    'corpo valvula
    If Len(COM) >= 10 And ( _
       Left(COM, 10) = "PA-GAV-COR" Or _
       Left(COM, 10) = "PA-GLO-COR" Or _
       Left(COM, 10) = "PA-PIS-COR" Or _
       Left(COM, 10) = "PA-POR-COR") Then
        sPec = "MP-COR"
        sQua = 1
        If Right(COM, 1) = "8" Then
            If BIT = kBIT14 Or BIT = kBIT38 Or BIT = kBIT12 Then
                sBit = kBIT12
            ElseIf BIT = kBIT114 Then
                sBit = kBIT112
            Else
                sBit = BIT
            End If
        ElseIf Right(COM, 1) = "5" Then
            If BIT = kBIT14 Or BIT = kBIT38 Then
                sBit = kBIT34
            ElseIf BIT = kBIT12 Or BIT = kBIT34 Then
                sBit = kBIT1
            ElseIf BIT = kBIT1 Or BIT = kBIT114 Then
                sBit = kBIT112
            ElseIf BIT = kBIT112 Then
                sBit = kBIT2
            End If
        End If
    'castelo
    ElseIf Len(COM) >= 10 And ( _
       Left(COM, 10) = "PA-GAV-CAS" Or _
       Left(COM, 10) = "PA-GLO-CAS") Then
        sPec = "MP-CAS"
        sQua = 1
        If Right(COM, 1) = "8" Then
            If BIT = kBIT14 Or BIT = kBIT38 Or BIT = kBIT12 Or BIT = kBIT34 Then
                sBit = kBIT12E34
            ElseIf BIT = kBIT1 Then
                sBit = kBIT1
            Else
                sBit = kBIT112E2
            End If
        ElseIf Right(COM, 1) = "5" Then
            If BIT = kBIT14 Or BIT = kBIT38 Then
                sBit = kBIT12E34
            ElseIf BIT = kBIT12 Or BIT = kBIT34 Then
                sBit = kBIT1
            Else
                sBit = kBIT112E2
            End If
        End If
    'tampa
    ElseIf Len(COM) >= 10 And ( _
       Left(COM, 10) = "PA-PIS-TAM" Or _
       Left(COM, 10) = "PA-POR-TAM") Then
        sPec = "MP-TAM"
        sQua = 1
        If Right(COM, 1) = "8" Then
            If BIT = kBIT14 Or BIT = kBIT38 Or BIT = kBIT12 Or BIT = kBIT34 Then
                sBit = kBIT12E34
            ElseIf BIT = kBIT1 Then
                sBit = kBIT1
            Else
                sBit = kBIT112E2
            End If
        ElseIf Right(COM, 1) = "5" Then
            If BIT = kBIT14 Or BIT = kBIT38 Then
                sBit = kBIT12E34
            ElseIf BIT = kBIT12 Or BIT = kBIT34 Then
                sBit = kBIT1
            Else
                sBit = kBIT112E2
            End If
        End If
    'cunha
    ElseIf Len(COM) >= 8 And Left(COM, 8) = "PA-CUNHA" Then
        sPec = "MP-CNH"
        sQua = 1
        If BIT = kBIT14 Or BIT = kBIT38 Or BIT = kBIT12 Then
            sBit = kBIT12
        ElseIf BIT = kBIT114 Or BIT = kBIT112 Then
            sBit = kBIT112
        Else
            sBit = BIT
        End If
    'contra-sede
    ElseIf Len(COM) >= 8 And Left(COM, 8) = "PA-CUNHA" Then
        sPec = "MP-CON"
        sQua = 1
        If BIT = kBIT14 Or BIT = kBIT38 Or BIT = kBIT12 Or BIT = kBIT34 Then
            sBit = kBIT12E34
        ElseIf BIT = kBIT1 Then
            sBit = kBIT1
        Else
            sBit = kBIT112E2
        End If
    End If
    'retorna valores configurados
    ConfigPA_Aux1 = Array(sQua, sPec, sBit)
End Function
