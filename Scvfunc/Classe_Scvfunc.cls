VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Classe_Scvfunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' ****************** VARIÁVEIS DLL's ****************

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function GetFocus Lib "user32" () As Long

Private Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

Private Type SYSTEM_INFO
        dwOemID As Long
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type

' ****************** DECLARAÇÕES ****************
Private INF_OS As OSVERSIONINFO
Dim INF_SYS As SYSTEM_INFO
Dim INF_MEM As MEMORYSTATUS
Private cValor As String, nTipo As Integer, I As Integer


' ****************** FUNÇÕES API ****************
Public Function PegaNomeComputador() As String
    Dim sBuffer As String, lBufSize As Long, lStatus As Long
    Dim sGetComputerName As String
    lBufSize = 255
    sBuffer = String$(lBufSize, " ")
    lStatus = GetComputerName(sBuffer, lBufSize)
    sGetComputerName = ""
    If lStatus <> 0 Then
        sGetComputerName = Left(sBuffer, lBufSize)
        PegaNomeComputador = sGetComputerName
    Else
        PegaNomeComputador = ""
    End If
End Function
Public Function OS_Plataforma() As String
    INF_OS.dwOSVersionInfoSize = Len(INF_OS)
    GetVersionEx INF_OS
    Select Case INF_OS.dwPlatformId
    Case 0
        OS_Plataforma = "Windows " + CStr(INF_OS.dwMajorVersion) + "." + CStr(INF_OS.dwMinorVersion)
    Case 1
        If INF_OS.dwMinorVersion >= 10 Then
            OS_Plataforma = "Windows 98"
        Else
            OS_Plataforma = "Windows 95"
        End If
    Case 2
        OS_Plataforma = "Windows NT"
    End Select
End Function
Public Function OS_CSDVersao() As String
    INF_OS.dwOSVersionInfoSize = Len(INF_OS)
    GetVersionEx INF_OS
    Dim nNull
    nNull = InStr(INF_OS.szCSDVersion, vbNullChar)
    If nNull > 1 Then
        OS_CSDVersao = Left(INF_OS.szCSDVersion, nNull - 1)
    ElseIf nNull = 1 Then
        OS_CSDVersao = ""
    End If
End Function
Public Function OS_Build() As String
    INF_OS.dwOSVersionInfoSize = Len(INF_OS)
    GetVersionEx INF_OS
    'Transforma numero (nao sei como) - vide API
    If (INF_OS.dwBuildNumber And &HFFFF&) > &H7FFF Then
       OS_Build = CStr((INF_OS.dwBuildNumber And &HFFFF&) - &H10000)
    Else
       OS_Build = CStr(INF_OS.dwBuildNumber And &HFFFF&)
    End If
End Function
Public Function OS_NumVersao() As String
    INF_OS.dwOSVersionInfoSize = Len(INF_OS)
    GetVersionEx INF_OS
    OS_NumVersao = CStr(INF_OS.dwMajorVersion) & "." & CStr(INF_OS.dwMinorVersion)
End Function
Public Function SIS_MascaraProcessador() As String
    GetSystemInfo INF_SYS
    SIS_MascaraProcessador = CStr(INF_SYS.dwActiveProcessorMask)
End Function
Public Function SIS_OEMID() As String
    GetSystemInfo INF_SYS
    SIS_OEMID = CStr(INF_SYS.dwOemID)
End Function
Public Function SIS_TipoProcessador() As String
    GetSystemInfo INF_SYS
    SIS_TipoProcessador = CStr(INF_SYS.dwProcessorType)
End Function
Public Function MEM_TotRAM() As String
    GlobalMemoryStatus INF_MEM
    MEM_TotRAM = CLng(INF_MEM.dwTotalPhys) / 1024 & " KB"
End Function
Public Function DiretorioWindows() As String
    Dim Bufstr As String
    Bufstr = Space$(50)
    If GetWindowsDirectory(Bufstr, 50) > 0 Then
        DiretorioWindows = Bufstr
        DiretorioWindows = RTrim(DiretorioWindows)
        DiretorioWindows = RetiraChr0(DiretorioWindows)
        If Right$(DiretorioWindows, 1) <> "\" Then
            DiretorioWindows = DiretorioWindows & "\"
        End If
    Else
        DiretorioWindows = ""
    End If
End Function
Public Function DiretorioSystem() As String
    Dim Bufstr As String
    Bufstr = Space$(50)
    If GetSystemDirectory(Bufstr, 50) > 0 Then
        DiretorioSystem = Bufstr
        DiretorioSystem = RTrim(DiretorioSystem)
        DiretorioSystem = RetiraChr0(DiretorioSystem)
        If Right$(DiretorioSystem, 1) <> "\" Then
            DiretorioSystem = DiretorioSystem & "\"
        End If
    Else
        DiretorioSystem = ""
    End If
End Function
Public Function DiretorioTemporario() As String
    Dim Bufstr As String
    Bufstr = Space$(50)
    If GetTempPath(50, Bufstr) > 0 Then
        DiretorioTemporario = Bufstr
        DiretorioTemporario = RTrim(DiretorioTemporario)
        DiretorioTemporario = RetiraChr0(DiretorioTemporario)
        If Right$(DiretorioTemporario, 1) <> "\" Then
            DiretorioTemporario = DiretorioTemporario & "\"
        End If
    Else
        DiretorioTemporario = ""
    End If
End Function
Private Function RetiraChr0(ByVal strString As String) As String
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        RetiraChr0 = Left$(strString, intZeroPos - 1)
    Else
        RetiraChr0 = strString
    End If
End Function
Public Function PegaSerialC() As String
    Dim VolLabel As String, VolSize As Long
    Dim SerNum As Long, MaxLen As Long
    Dim Flags As Long, name As String
    Dim NameSize As Long, Check As String
    If GetVolumeInformation("C:\", VolLabel, VolSize, SerNum, MaxLen, Flags, name, NameSize) Then
        Check = Format(Hex(SerNum), "00000000")
        PegaSerialC = Left$(Check, 4) + "-" + Right$(Check, 4)
    Else
        PegaSerialC = "0000-0000"
    End If
End Function
Public Function PegaTamanhoC() As String
    Dim TotalClusters As Long, FreeClusters As Long
    Dim SectorsPerCluster As Long, BytesPerSector As Long
    If GetDiskFreeSpace("C:\", SectorsPerCluster, BytesPerSector, FreeClusters, TotalClusters) Then
        PegaTamanhoC = (SectorsPerCluster * BytesPerSector * TotalClusters) / 1024 & " KB"
    Else
        PegaTamanhoC = ""
    End If
End Function
Public Function MoveArquivo(DirVelho As String, DirNovo As String)
    If MoveFile(DirVelho, DirNovo) Then Beep
End Function
Public Function PegaFoco() As Long
    PegaFoco = GetFocus
End Function



'***************************************************
'             FUNÇÕES E ROTINAS MINHAS
'***************************************************

Public Function ValidaTexto(KeyAscii As Integer) As Integer
    If Chr(KeyAscii) <> "0" And _
       Chr(KeyAscii) <> "1" And _
       Chr(KeyAscii) <> "2" And _
       Chr(KeyAscii) <> "3" And _
       Chr(KeyAscii) <> "4" And _
       Chr(KeyAscii) <> "5" And _
       Chr(KeyAscii) <> "6" And _
       Chr(KeyAscii) <> "7" And _
       Chr(KeyAscii) <> "8" And _
       Chr(KeyAscii) <> "9" And _
       Chr(KeyAscii) <> "," And _
       Chr(KeyAscii) <> "/" And _
       KeyAscii > 30 Then
        KeyAscii = 27
    End If
    ValidaTexto = KeyAscii
End Function
Public Function ValorExtenso(Num As String) As String
    If Len(Num) = 0 Then Exit Function
        
    '1.000.000,00
    Dim cValTmp As String, cResNum As String
    Dim cTmp As String
    Dim nCarVir As Integer, nPos As Integer
    
    nCarVir = 0
    nPos = 0
    cValTmp = ""
    cResNum = ""
    cTmp = ""
    Num = Format(Num, "###############.00")
    'pega centavos
    For I = Len(CStr(Num)) To 1 Step -1
        If Mid(CStr(Num), I, 1) = "," Then
            nCarVir = I
            Exit For
        End If
    Next I
    cResNum = Num
    If nCarVir <> 0 Then
        If Len(CStr(Num)) >= nCarVir + 1 Then
            If Mid(Num, nCarVir + 2, 1) = "" Then Num = CStr(Num) & "0"
            cValTmp = Aux_ValorExtenso_PegaNumero(Mid(Num, nCarVir + 1, 2))
            If cValTmp <> "" Then
                If cValTmp = "um" Then
                    cValor = cValTmp & " centavo"
                Else
                    cValor = cValTmp & " centavos"
                End If
            End If
            cResNum = Left(Num, nCarVir - 1)
        End If
    End If

    'pega unidade
    
    cTmp = ""
    Do While Len(cResNum) > 3
        cTmp = Right(cResNum, 3)
        cValTmp = Aux_ValorExtenso_PegaNumero(cTmp)
        Aux_ValorExtenso_PegaValor (cValTmp)
        nTipo = nTipo + 1
        cResNum = Left(cResNum, Len(cResNum) - 3)
    Loop
    cValTmp = Aux_ValorExtenso_PegaNumero(cResNum)
    Aux_ValorExtenso_PegaValor (cValTmp)
    
    'Verifica se é UM REAIS
    If Left(cValor, 8) = "um reais" And Len(cValor) = 8 Then
        cValor = "um real"
    ElseIf Left(cValor, 8) = "um reais" And Len(cValor) > 8 Then
        cValor = "um real" & Right(cValor, Len(cValor) - 8)
    End If
    'Transf. UM p/ HUM
    If Left(cValor, 2) = "um" Then cValor = "hum" & Mid(cValor, 3, Len(cValor))
    'Transf. 1ª letra em maiuscula
    cValor = UCase(Left(cValor, 1)) & Mid(cValor, 2, Len(cValor))
    'Retorna valor
    ValorExtenso = cValor
End Function
Private Function Aux_ValorExtenso_PegaNumero(Num As String) As String
    If Len(Num) = 0 Then Exit Function
    Dim cTmpNum As String
    
    If Len(Num) = 1 Then 'se o Num for 1 caracter
        If Num = "1" Then
            Aux_ValorExtenso_PegaNumero = "um"
        ElseIf Num = "2" Then
            Aux_ValorExtenso_PegaNumero = "dois"
        ElseIf Num = "3" Then
            Aux_ValorExtenso_PegaNumero = "três"
        ElseIf Num = "4" Then
            Aux_ValorExtenso_PegaNumero = "quatro"
        ElseIf Num = "5" Then
            Aux_ValorExtenso_PegaNumero = "cinco"
        ElseIf Num = "6" Then
            Aux_ValorExtenso_PegaNumero = "seis"
        ElseIf Num = "7" Then
            Aux_ValorExtenso_PegaNumero = "sete"
        ElseIf Num = "8" Then
            Aux_ValorExtenso_PegaNumero = "oito"
        ElseIf Num = "9" Then
            Aux_ValorExtenso_PegaNumero = "nove"
        End If
    
    ElseIf Len(Num) = 2 Then 'se o Num forem 2 caracteres
        If Left(Num, 1) = "0" Then 'se a 1ª casa for ZERO
            If Right(Num, 1) = "1" Then
                Aux_ValorExtenso_PegaNumero = "um"
            ElseIf Right(Num, 1) = "2" Then
                Aux_ValorExtenso_PegaNumero = "dois"
            ElseIf Right(Num, 1) = "3" Then
                Aux_ValorExtenso_PegaNumero = "três"
            ElseIf Right(Num, 1) = "4" Then
                Aux_ValorExtenso_PegaNumero = "quatro"
            ElseIf Right(Num, 1) = "5" Then
                Aux_ValorExtenso_PegaNumero = "cinco"
            ElseIf Right(Num, 1) = "6" Then
                Aux_ValorExtenso_PegaNumero = "seis"
            ElseIf Right(Num, 1) = "7" Then
                Aux_ValorExtenso_PegaNumero = "sete"
            ElseIf Right(Num, 1) = "8" Then
                Aux_ValorExtenso_PegaNumero = "oito"
            ElseIf Right(Num, 1) = "9" Then
                Aux_ValorExtenso_PegaNumero = "nove"
            End If
        ElseIf Left(Num, 1) = "1" Then 'se a 1ª casa for UM
            If Num = "10" Then
                Aux_ValorExtenso_PegaNumero = "dez"
            ElseIf Num = "11" Then
                Aux_ValorExtenso_PegaNumero = "onze"
            ElseIf Num = "12" Then
                Aux_ValorExtenso_PegaNumero = "doze"
            ElseIf Num = "13" Then
                Aux_ValorExtenso_PegaNumero = "treze"
            ElseIf Num = "14" Then
                Aux_ValorExtenso_PegaNumero = "quatorze"
            ElseIf Num = "15" Then
                Aux_ValorExtenso_PegaNumero = "quinze"
            ElseIf Num = "16" Then
                Aux_ValorExtenso_PegaNumero = "dezesseis"
            ElseIf Num = "17" Then
                Aux_ValorExtenso_PegaNumero = "dezessete"
            ElseIf Num = "18" Then
                Aux_ValorExtenso_PegaNumero = "dezoito"
            ElseIf Num = "19" Then
                Aux_ValorExtenso_PegaNumero = "dezenove"
            End If
        Else 'Se a 1ª casa for numero maior que UM
            cTmpNum = ""
            'pega a dezena
            If Left(Num, 1) = "2" Then
                cTmpNum = "vinte"
            ElseIf Left(Num, 1) = "3" Then
                cTmpNum = "trinta"
            ElseIf Left(Num, 1) = "4" Then
                cTmpNum = "quarenta"
            ElseIf Left(Num, 1) = "5" Then
                cTmpNum = "cinquenta"
            ElseIf Left(Num, 1) = "6" Then
                cTmpNum = "sessenta"
            ElseIf Left(Num, 1) = "7" Then
                cTmpNum = "setenta"
            ElseIf Left(Num, 1) = "8" Then
                cTmpNum = "oitenta"
            ElseIf Left(Num, 1) = "9" Then
                cTmpNum = "noventa"
            End If
            'pega a unidade
            If Right(Num, 1) = "1" Then
                cTmpNum = cTmpNum & " e um"
            ElseIf Right(Num, 1) = "2" Then
                cTmpNum = cTmpNum & " e dois"
            ElseIf Right(Num, 1) = "3" Then
                cTmpNum = cTmpNum & " e três"
            ElseIf Right(Num, 1) = "4" Then
                cTmpNum = cTmpNum & " e quatro"
            ElseIf Right(Num, 1) = "5" Then
                cTmpNum = cTmpNum & " e cinco"
            ElseIf Right(Num, 1) = "6" Then
                cTmpNum = cTmpNum & " e seis"
            ElseIf Right(Num, 1) = "7" Then
                cTmpNum = cTmpNum & " e sete"
            ElseIf Right(Num, 1) = "8" Then
                cTmpNum = cTmpNum & " e oito"
            ElseIf Right(Num, 1) = "9" Then
                cTmpNum = cTmpNum & " e nove"
            End If
            Aux_ValorExtenso_PegaNumero = cTmpNum
        End If
    ElseIf Len(Num) = 3 Then 'se o Num forem 3 caracteres
        If Left(Num, 2) = "00" Then
            If Right(Num, 1) = "1" Then
                Aux_ValorExtenso_PegaNumero = "um"
            ElseIf Right(Num, 1) = "2" Then
                Aux_ValorExtenso_PegaNumero = "dois"
            ElseIf Right(Num, 1) = "3" Then
                Aux_ValorExtenso_PegaNumero = "três"
            ElseIf Right(Num, 1) = "4" Then
                Aux_ValorExtenso_PegaNumero = "quatro"
            ElseIf Right(Num, 1) = "5" Then
                Aux_ValorExtenso_PegaNumero = "cinco"
            ElseIf Right(Num, 1) = "6" Then
                Aux_ValorExtenso_PegaNumero = "seis"
            ElseIf Right(Num, 1) = "7" Then
                Aux_ValorExtenso_PegaNumero = "sete"
            ElseIf Right(Num, 1) = "8" Then
                Aux_ValorExtenso_PegaNumero = "oito"
            ElseIf Right(Num, 1) = "9" Then
                Aux_ValorExtenso_PegaNumero = "nove"
            ElseIf Right(Num, 1) = "0" Then
                Aux_ValorExtenso_PegaNumero = "reais"
            End If
        ElseIf Left(Num, 1) = "0" And Mid(Num, 2, 1) <> "0" Then
            cTmpNum = ""
            'pega a dezena
            If Int(Right(Num, 2)) > 19 Then
                If Mid(Num, 2, 1) = "2" Then
                    cTmpNum = "vinte"
                ElseIf Mid(Num, 2, 1) = "3" Then
                    cTmpNum = "trinta"
                ElseIf Mid(Num, 2, 1) = "4" Then
                    cTmpNum = "quarenta"
                ElseIf Mid(Num, 2, 1) = "5" Then
                    cTmpNum = "cinquenta"
                ElseIf Mid(Num, 2, 1) = "6" Then
                    cTmpNum = "sessenta"
                ElseIf Mid(Num, 2, 1) = "7" Then
                    cTmpNum = "setenta"
                ElseIf Mid(Num, 2, 1) = "8" Then
                    cTmpNum = "oitenta"
                ElseIf Mid(Num, 2, 1) = "9" Then
                    cTmpNum = "noventa"
                End If
                'pega a unidade
                If Right(Num, 1) = "1" Then
                    cTmpNum = cTmpNum & " e um"
                ElseIf Right(Num, 1) = "2" Then
                    cTmpNum = cTmpNum & " e dois"
                ElseIf Right(Num, 1) = "3" Then
                    cTmpNum = cTmpNum & " e três"
                ElseIf Right(Num, 1) = "4" Then
                    cTmpNum = cTmpNum & " e quatro"
                ElseIf Right(Num, 1) = "5" Then
                    cTmpNum = cTmpNum & " e cinco"
                ElseIf Right(Num, 1) = "6" Then
                    cTmpNum = cTmpNum & " e seis"
                ElseIf Right(Num, 1) = "7" Then
                    cTmpNum = cTmpNum & " e sete"
                ElseIf Right(Num, 1) = "8" Then
                    cTmpNum = cTmpNum & " e oito"
                ElseIf Right(Num, 1) = "9" Then
                    cTmpNum = cTmpNum & " e nove"
                End If
            Else 'é de 10 à 19
                If Right(Num, 2) = "10" Then
                    cTmpNum = "dez"
                ElseIf Right(Num, 2) = "11" Then
                    cTmpNum = "onze"
                ElseIf Right(Num, 2) = "12" Then
                    cTmpNum = "doze"
                ElseIf Right(Num, 2) = "13" Then
                    cTmpNum = "treze"
                ElseIf Right(Num, 2) = "14" Then
                    cTmpNum = "quatorze"
                ElseIf Right(Num, 2) = "15" Then
                    cTmpNum = "quinze"
                ElseIf Right(Num, 2) = "16" Then
                    cTmpNum = "dezesseis"
                ElseIf Right(Num, 2) = "17" Then
                    cTmpNum = "dezessete"
                ElseIf Right(Num, 2) = "18" Then
                    cTmpNum = "dezoito"
                ElseIf Right(Num, 2) = "19" Then
                    cTmpNum = "dezenove"
                End If
            End If
            Aux_ValorExtenso_PegaNumero = cTmpNum
        ElseIf Left(Num, 1) <> "0" Then
            cTmpNum = ""
            'pega a centena
            If Left(Num, 1) = "1" Then
                cTmpNum = "cento"
            ElseIf Left(Num, 1) = "2" Then
                cTmpNum = "duzentos"
            ElseIf Left(Num, 1) = "3" Then
                cTmpNum = "trezentos"
            ElseIf Left(Num, 1) = "4" Then
                cTmpNum = "quatrocentos"
            ElseIf Left(Num, 1) = "5" Then
                cTmpNum = "quinhentos"
            ElseIf Left(Num, 1) = "6" Then
                cTmpNum = "seiscentos"
            ElseIf Left(Num, 1) = "7" Then
                cTmpNum = "setecentos"
            ElseIf Left(Num, 1) = "8" Then
                cTmpNum = "oitocentos"
            ElseIf Left(Num, 1) = "9" Then
                cTmpNum = "novecentos"
            End If
            'nao tem dezena
            If Mid(Num, 2, 1) = "0" Then
                If Right(Num, 1) = "1" Then
                    cTmpNum = cTmpNum & " e um"
                ElseIf Right(Num, 1) = "2" Then
                    cTmpNum = cTmpNum & " e dois"
                ElseIf Right(Num, 1) = "3" Then
                    cTmpNum = cTmpNum & " e três"
                ElseIf Right(Num, 1) = "4" Then
                    cTmpNum = cTmpNum & " e quatro"
                ElseIf Right(Num, 1) = "5" Then
                    cTmpNum = cTmpNum & " e cinco"
                ElseIf Right(Num, 1) = "6" Then
                    cTmpNum = cTmpNum & " e seis"
                ElseIf Right(Num, 1) = "7" Then
                    cTmpNum = cTmpNum & " e sete"
                ElseIf Right(Num, 1) = "8" Then
                    cTmpNum = cTmpNum & " e oito"
                ElseIf Right(Num, 1) = "9" Then
                    cTmpNum = cTmpNum & " e nove"
                End If
            End If
            'pega dezena do 10 ao 19
            If Mid(Num, 2, 2) = "10" Then
                cTmpNum = cTmpNum & " e dez"
            ElseIf Mid(Num, 2, 2) = "11" Then
                cTmpNum = cTmpNum & " e onze"
            ElseIf Mid(Num, 2, 2) = "12" Then
                cTmpNum = cTmpNum & " e doze"
            ElseIf Mid(Num, 2, 2) = "13" Then
                cTmpNum = cTmpNum & " e treze"
            ElseIf Mid(Num, 2, 2) = "14" Then
                cTmpNum = cTmpNum & " e quatorze"
            ElseIf Mid(Num, 2, 2) = "15" Then
                cTmpNum = cTmpNum & " e quinze"
            ElseIf Mid(Num, 2, 2) = "16" Then
                cTmpNum = cTmpNum & " e dezesseis"
            ElseIf Mid(Num, 2, 2) = "17" Then
                cTmpNum = cTmpNum & " e dezessete"
            ElseIf Mid(Num, 2, 2) = "18" Then
                cTmpNum = cTmpNum & " e dezoito"
            ElseIf Mid(Num, 2, 2) = "19" Then
                cTmpNum = cTmpNum & " e dezenove"
            'pega dezena acima do 20
            ElseIf Mid(Num, 2, 1) <> "0" And Mid(Num, 2, 1) <> "1" Then
                'pega dezena
                If Mid(Num, 2, 1) = "2" Then
                    cTmpNum = cTmpNum & " e vinte"
                ElseIf Mid(Num, 2, 1) = "3" Then
                    cTmpNum = cTmpNum & " e trinta"
                ElseIf Mid(Num, 2, 1) = "4" Then
                    cTmpNum = cTmpNum & " e quarenta"
                ElseIf Mid(Num, 2, 1) = "5" Then
                    cTmpNum = cTmpNum & " e cinquenta"
                ElseIf Mid(Num, 2, 1) = "6" Then
                    cTmpNum = cTmpNum & " e sessenta"
                ElseIf Mid(Num, 2, 1) = "7" Then
                    cTmpNum = cTmpNum & " e setenta"
                ElseIf Mid(Num, 2, 1) = "8" Then
                    cTmpNum = cTmpNum & " e oitenta"
                ElseIf Mid(Num, 2, 1) = "9" Then
                    cTmpNum = cTmpNum & " e noventa"
                End If
                'pega unidade
                If Right(Num, 1) = "1" Then
                    cTmpNum = cTmpNum & " e um"
                ElseIf Right(Num, 1) = "2" Then
                    cTmpNum = cTmpNum & " e dois"
                ElseIf Right(Num, 1) = "3" Then
                    cTmpNum = cTmpNum & " e três"
                ElseIf Right(Num, 1) = "4" Then
                    cTmpNum = cTmpNum & " e quatro"
                ElseIf Right(Num, 1) = "5" Then
                    cTmpNum = cTmpNum & " e cinco"
                ElseIf Right(Num, 1) = "6" Then
                    cTmpNum = cTmpNum & " e seis"
                ElseIf Right(Num, 1) = "7" Then
                    cTmpNum = cTmpNum & " e sete"
                ElseIf Right(Num, 1) = "8" Then
                    cTmpNum = cTmpNum & " e oito"
                ElseIf Right(Num, 1) = "9" Then
                    cTmpNum = cTmpNum & " e nove"
                End If
            End If
            Aux_ValorExtenso_PegaNumero = cTmpNum
        End If
    End If
End Function
Private Sub Aux_ValorExtenso_PegaValor(Valor As String)
    If Valor <> "" Then
        If nTipo = 0 Then
            If Valor = "reais" Then 'Se for X.000,XX
                If cValor = "" Then
                    cValor = "reais"
                Else
                    cValor = "reais e " & cValor
                End If
                Exit Sub
            End If
            If Valor = "um" Then 'Verifica se é UM real
                If cValor = "" Then
                    cValor = Valor & " real"
                Else
                    cValor = Valor & " real e " & cValor
                End If
            ElseIf Valor = "cento" Then 'Verifica se é CENTO
                If cValor = "" Then
                    cValor = "cem" & " reais"
                Else
                    cValor = "cem" & " reais e " & cValor
                End If
            Else
                If cValor = "" Then
                    cValor = Valor & " reais"
                Else
                    cValor = Valor & " reais e " & cValor
                End If
            End If
        ElseIf nTipo = 1 Then
            If Left(cValor, 5) = "reais" Then 'Se for X.000,XX
                If Len(cValor) > 5 Then
                    If Valor <> "reais" And Valor <> "cento" Then cValor = Valor & " mil " & cValor
                    If Valor <> "reais" And Valor = "cento" Then cValor = "cem mil " & cValor
                Else
                    If Valor <> "reais" And Valor <> "cento" Then cValor = Valor & " mil " & cValor
                    If Valor <> "reais" And Valor = "cento" Then cValor = "cem mil " & cValor
                End If
            ElseIf Left(cValor, 7) = "um real" Then
                If Valor <> "reais" And Valor <> "cento" Then cValor = Valor & " mil e um reais " & Right(cValor, Len(cValor) - 7)
                If Valor <> "reais" And Valor = "cento" Then cValor = "cem mil e um reais " & Right(cValor, Len(cValor) - 7)
            Else
                If Valor <> "reais" And Valor <> "cento" Then cValor = Valor & " mil, " & cValor
                If Valor <> "reais" And Valor = "cento" Then cValor = "cem mil, " & cValor
            End If
        ElseIf nTipo = 2 Then
            If Left(cValor, 5) = "reais" Then
                If Valor = "um" Then
                    cValor = Valor & " milhão de " & cValor
                ElseIf Valor = "cento" Then
                    cValor = "cem milhões de " & cValor
                Else
                    cValor = Valor & " milhões de " & cValor
                End If
            Else
                If Valor = "um" Then
                    cValor = Valor & " milhão, " & cValor
                ElseIf Valor = "cento" Then
                    cValor = "cem milhões, " & cValor
                Else
                    cValor = Valor & " milhões, " & cValor
                End If
            End If
        End If
    End If
End Sub
Public Function MultiString(Texto As String, Multi As Integer) As String
    If Texto = "" Then Exit Function
    If IsNumeric(Multi) = False Then Exit Function
    If Multi < 0 Then Exit Function
    
    Dim TextoFinal As String
    TextoFinal = ""
    For I = 1 To Multi
        TextoFinal = TextoFinal & Texto
    Next I
    MultiString = TextoFinal
End Function
Public Static Function NomeImpressora(Tela As String) As String
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    On Error GoTo ERRO_NOMEIMPRESSORA
    NomeImpressora = ""
    If Tela = "" Then Exit Function
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreTabela_ConfiguracoesTelasImpressoras App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreTabela_Impressoras App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreCampos_ConfiguracoesTelasImpressoras App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreCampos_Impressoras App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.BDSIS_TBTIM.Seek "=", Tela
    If DLL_BD.BDSIS_TBTIM.NoMatch Then
        MsgBox "Não foi possível localizar as configurações desta tela para impressoras no banco de dados.", vbCritical + vbOKOnly, "Falha na procura"
        GoTo ERRO_NOMEIMPRESSORA
    End If
    DLL_BD.BDSIS_TBIMP.Seek "=", DLL_BD.BDSIS_TBTIM_CPIMP.Value
    If DLL_BD.BDSIS_TBIMP.NoMatch Then
        MsgBox "Não foi possível localizar as configurações da impressoras no banco de dados.", vbCritical + vbOKOnly, "Falha na procura"
        GoTo ERRO_NOMEIMPRESSORA
    End If
    NomeImpressora = DLL_BD.BDSIS_TBIMP_CPNOM.Value
ERRO_NOMEIMPRESSORA:
    DLL_BD.FechaTabela_ConfiguracoesTelasImpressoras App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaTabela_Impressoras App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Function
Public Static Sub SelecionaImpressora(Impressora As String)
    'COLOQUEI TRES VEZES PQ ESSA PORRA NAO QUERIA MUDAR AUTOMATICO
    For I = 0 To Printers.Count - 1
        If Printers(I).DeviceName = Impressora Then
            Set Printer = Printers(I)
            Exit For
        End If
    Next I
    For I = 0 To Printers.Count - 1
        If Printers(I).DeviceName = Impressora Then
            Set Printer = Printers(I)
            Exit For
        End If
    Next I
    For I = 0 To Printers.Count - 1
        If Printers(I).DeviceName = Impressora Then
            Set Printer = Printers(I)
            Exit For
        End If
    Next I
End Sub
Public Static Function LogDeEventos(Usuario As String, Computador As String, Evento As String, Valor As String) As Boolean
    LogDeEventos = False
    If Usuario = "" Or _
       Computador = "" Or _
       Evento = "" Or _
       Valor = "" Then Exit Function
    'Abre BD
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    On Error GoTo ERRO_LOGDEEVENTOS
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreTabela_LogDeEventos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreCampos_LogDeEventos App.ProductName, "Scvbd", App.LegalCopyright
    'Registra evento
    DLL_BD.BDSIS_TBGEV.AddNew
    DLL_BD.BDSIS_TBGEV_CPDAT.Value = Format(Date, "dd/mm/yyyy")
    DLL_BD.BDSIS_TBGEV_CPHOR.Value = Format(Time, "hh:mm:ss")
    DLL_BD.BDSIS_TBGEV_CPUSU.Value = Usuario
    DLL_BD.BDSIS_TBGEV_CPCOM.Value = Computador
    DLL_BD.BDSIS_TBGEV_CPEVE.Value = Evento
    DLL_BD.BDSIS_TBGEV_CPVAL.Value = Valor
    DLL_BD.BDSIS_TBGEV.Update
    'Finaliza
    LogDeEventos = True
ERRO_LOGDEEVENTOS:
    DLL_BD.FechaTabela_LogDeEventos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Function
Public Static Function MensagemErro(Usuario As String, Computador As String, ErroNum, ErroDes, ErroOri, Aplicativo, ErroAjuArq, ErroAjuCon) As Boolean
    MensagemErro = False
    Dim ErroAk As Boolean
    ErroAk = False
    'If Usuario = "" Or _
    '   Computador = "" Or _
    '   ErroNum = "" Or _
    '   Aplicativo = "" Then
    '    ErroAk = True
    '    GoTo ERRO_LOGDEERROS
    'End If
    'Exibe msgbox
    MsgBox "Ocorreu um erro de execução nesta operação. Tente mais tarde e/ou avise o administrador do sistema." & vbCr & vbCr & "Descrição: " & ErroDes & vbCr & "Código: " & ErroNum & vbCr & "Origem: " & ErroOri & vbCr, vbCritical + vbOKOnly + vbMsgBoxHelpButton, Aplicativo, ErroAjuArq, ErroAjuCon
    
    'Abre BD
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    On Error GoTo ERRO_LOGDEERROS
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreTabela_LogDeErros App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreCampos_LogDeErros App.ProductName, "Scvbd", App.LegalCopyright
    'Registra Erro
    DLL_BD.BDSIS_TBGRR.AddNew
    DLL_BD.BDSIS_TBGRR_CPDAT.Value = Format(Date, "dd/mm/yyyy")
    DLL_BD.BDSIS_TBGRR_CPHOR.Value = Format(Time, "hh:mm:ss")
    DLL_BD.BDSIS_TBGRR_CPUSU.Value = Usuario
    DLL_BD.BDSIS_TBGRR_CPCOM.Value = Computador
    DLL_BD.BDSIS_TBGRR_CPAPL.Value = Aplicativo
    DLL_BD.BDSIS_TBGRR_CPENU.Value = ErroNum
    DLL_BD.BDSIS_TBGRR_CPEDE.Value = ErroDes
    DLL_BD.BDSIS_TBGRR_CPEOR.Value = ErroOri
    DLL_BD.BDSIS_TBGRR_CPEJA.Value = ErroAjuArq
    DLL_BD.BDSIS_TBGRR_CPEJC.Value = ErroAjuCon
    DLL_BD.BDSIS_TBGRR.Update
    'Finaliza
    MensagemErro = True
ERRO_LOGDEERROS:
    If ErroAk = True Or Err Then MsgBox "Não foi possível registrar as mensagens de erro. Encerre o sistema.", vbCritical + vbOKOnly, "Erro crítico"
    DLL_BD.FechaTabela_LogDeErros App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Function
Public Static Sub RegistraUsuario(Usuario As String, Computador As String)
    If Usuario = "" And Computador = "" Then Exit Sub
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    On Error GoTo ERRO_REGISTROUSUARIO
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.Cliente_AbreTabela_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.Cliente_AbreCampos_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    If DLL_BD.BDCLI_TBULG.RecordCount < 1 Then
        DLL_BD.BDCLI_TBULG.AddNew
    Else
        DLL_BD.BDCLI_TBULG.MoveFirst
        Do While Not DLL_BD.BDCLI_TBULG.EOF
            DLL_BD.BDCLI_TBULG.Delete
        Loop
        DLL_BD.BDCLI_TBULG.AddNew
    End If
    DLL_BD.BDCLI_TBULG_CPDAT.Value = Format(Date, "dd/mm/yyyy")
    DLL_BD.BDCLI_TBULG_CPHOR.Value = Format(Time, "hh:mm:ss")
    DLL_BD.BDCLI_TBULG_CPUSU.Value = Usuario
    DLL_BD.BDCLI_TBULG_CPCOM.Value = Computador
    DLL_BD.BDCLI_TBULG.Update
    'atualiza registro
    SaveSetting "Siscoval", "Logon", "Nome", Usuario
    SaveSetting "Siscoval", "Logon", "Computador", Computador
    SaveSetting "Siscoval", "Logon", "Data", Format(Date, "dd/mm/yyyy")
    SaveSetting "Siscoval", "Logon", "Hora", Format(Time, "hh:mm:ss")
ERRO_REGISTROUSUARIO:
    DLL_BD.Cliente_FechaTabela_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Sub
Public Static Function PegaUsuario() As String
    PegaUsuario = ""
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    On Error GoTo ERRO_REGISTROUSUARIO
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.Cliente_AbreTabela_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.Cliente_AbreCampos_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    If DLL_BD.BDCLI_TBULG.RecordCount < 1 Then
        GoTo ERRO_REGISTROUSUARIO
    Else
        DLL_BD.BDCLI_TBULG.MoveFirst
        PegaUsuario = DLL_BD.BDCLI_TBULG_CPUSU.Value
    End If
ERRO_REGISTROUSUARIO:
    DLL_BD.Cliente_FechaTabela_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Function
Public Static Function RegistraEvento(Evento As String, Valor As String) As Boolean
    RegistraEvento = False
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    On Error GoTo ERRO_REGISTROUSUARIO
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreTabela_LogDeEventos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.Cliente_AbreTabela_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreCampos_LogDeEventos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.Cliente_AbreCampos_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    'Pega nome usuario
    Dim Usuario As String, Computador As String
    DLL_BD.BDCLI_TBULG.MoveLast
    Usuario = Trim(DLL_BD.BDCLI_TBULG_CPUSU.Value)
    Computador = Trim(DLL_BD.BDCLI_TBULG_CPCOM.Value)
    If Usuario = "" Or Computador = "" Then GoTo ERRO_REGISTROUSUARIO
    'Registra Evento
    DLL_BD.BDSIS_TBGEV.AddNew
    DLL_BD.BDSIS_TBGEV_CPUSU.Value = Usuario
    DLL_BD.BDSIS_TBGEV_CPCOM.Value = Computador
    DLL_BD.BDSIS_TBGEV_CPDAT.Value = Format(Date, "dd/mm/yyyy")
    DLL_BD.BDSIS_TBGEV_CPHOR.Value = Format(Time, "hh:mm:ss")
    DLL_BD.BDSIS_TBGEV_CPEVE.Value = Evento
    DLL_BD.BDSIS_TBGEV_CPVAL.Value = Valor
    DLL_BD.BDSIS_TBGEV.Update
    RegistraEvento = True
ERRO_REGISTROUSUARIO:
    DLL_BD.FechaTabela_LogDeEventos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.Cliente_FechaTabela_UltimoLogon App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Function
Public Static Function ProcuraGrupo(Texto As String) As String
    On Error GoTo ERRO_PROCURAGRUPO
    ProcuraGrupo = ""
    If Texto = "" Then Exit Function
    'Abre BD
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreTabela_Grupos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreCampos_Grupos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.BDSIS_TBGRU.Seek "=", Texto
    If DLL_BD.BDSIS_TBGRU.NoMatch Then
        GoTo ERRO_PROCURAGRUPO
    Else
        ProcuraGrupo = DLL_BD.BDSIS_TBGRU_CPVAL.Value
    End If
ERRO_PROCURAGRUPO:
    DLL_BD.FechaTabela_Grupos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Function
Public Static Function ProcuraValorGrupo(Texto As String, TipoGrupo As String) As String
    On Error GoTo ERRO_PROCURAVALORGRUPO
    ProcuraValorGrupo = ""
    If Texto = "" And TipoGrupo = "" Then GoTo ERRO_PROCURAVALORGRUPO
    'Abre BD
    Dim DLL_BD As Scvbd.Classe_Scvbd
    Set DLL_BD = New Scvbd.Classe_Scvbd
    DLL_BD.AbreBD App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreTabela_Grupos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.AbreCampos_Grupos App.ProductName, "Scvbd", App.LegalCopyright
    'inicia procura
    DLL_BD.BDSIS_TBGRU.MoveFirst
    Do While Not DLL_BD.BDSIS_TBGRU.EOF
        If Texto = DLL_BD.BDSIS_TBGRU_CPVAL.Value And _
           TipoGrupo = DLL_BD.BDSIS_TBGRU_CPTIP.Value Then
            ProcuraValorGrupo = DLL_BD.BDSIS_TBGRU_CPGRU.Value
            GoTo ERRO_PROCURAVALORGRUPO
        End If
        DLL_BD.BDSIS_TBGRU.MoveNext
    Loop
ERRO_PROCURAVALORGRUPO:
    DLL_BD.FechaTabela_Grupos App.ProductName, "Scvbd", App.LegalCopyright
    DLL_BD.FechaBD App.ProductName, "Scvbd", App.LegalCopyright
    Set DLL_BD = Nothing
End Function
Public Static Function IB(EstadoInicial As Boolean) As Boolean
    If EstadoInicial = True Then
        IB = False
    Else
        IB = True
    End If
End Function
Public Static Function NomeMes(Mes As Integer) As String
    If Mes < 1 Or Mes > 12 Then Exit Function
    If Mes = 1 Then
        NomeMes = "Janeiro"
    ElseIf Mes = 2 Then
        NomeMes = "Fevereiro"
    ElseIf Mes = 3 Then
        NomeMes = "Março"
    ElseIf Mes = 4 Then
        NomeMes = "Abril"
    ElseIf Mes = 5 Then
        NomeMes = "Maio"
    ElseIf Mes = 6 Then
        NomeMes = "Junho"
    ElseIf Mes = 7 Then
        NomeMes = "Julho"
    ElseIf Mes = 8 Then
        NomeMes = "Agosto"
    ElseIf Mes = 9 Then
        NomeMes = "Setembro"
    ElseIf Mes = 10 Then
        NomeMes = "Outubro"
    ElseIf Mes = 11 Then
        NomeMes = "Novembro"
    ElseIf Mes = 12 Then
        NomeMes = "Dezembro"
    End If
End Function
Public Static Function Porcentagem(Valor As Currency, PorcentagemDe As Long) As Currency
    If PorcentagemDe < 1 Then
        Porcentagem = 0
        Exit Function
    End If
    If Len(CStr(PorcentagemDe)) = 1 Then
        Porcentagem = Valor * Val("0.0" & CStr(PorcentagemDe))
    ElseIf Len(CStr(PorcentagemDe)) = 2 Then
        Porcentagem = Valor * Val("0." & CStr(PorcentagemDe))
    End If
End Function
Public Static Function PegaNumeroItem(Indice As Integer) As String
    If Len(CStr(Indice)) = 1 Then
        PegaNumeroItem = "0" & CStr(Indice)
    Else
        PegaNumeroItem = Indice
    End If
End Function
Public Static Function PegaUnidade(Valor As Double, Unidade As Integer) As String
    PegaUnidade = ""
    If Valor < 0 Then Exit Function
    If Unidade < 0 Then Exit Function
    'Unidades:
    '0 - pçs
    Dim UniS As String, UniP As String
    If Unidade = 0 Then 'pçs
        UniS = "pç."
        UniP = "pçs."
    End If
    If Valor <= 1 Then
        PegaUnidade = Trim(CStr(Valor)) & " " & UniS
    Else
        PegaUnidade = Trim(CStr(Valor)) & " " & UniP
    End If
End Function
Public Static Function ArquivoExiste(Arquivo As String) As Boolean
    On Error GoTo ERRO
    Open Arquivo For Input As #1
    Close #1
    ArquivoExiste = True
    Exit Function
ERRO:
    ArquivoExiste = False
End Function
Public Static Function Teste_Caracter(Caracter As String) As Boolean
    If (AscB(Caracter) >= 65 And AscB(Caracter) <= 90) Or _
       (AscB(Caracter) >= 97 And AscB(Caracter) <= 122) Then 'A-Z ou a-z
        Teste_Caracter = True
    Else
        Teste_Caracter = False
    End If
End Function
Public Static Function Teste_Numero(Caracter As String) As Boolean
    If (AscB(Caracter) >= 48 And AscB(Caracter) <= 57) Then '0-9
        Teste_Numero = True
    Else
        Teste_Numero = False
    End If
End Function


