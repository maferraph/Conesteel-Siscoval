Attribute VB_Name = "API"
'Constantes que indicam o que ocorreu sobre
'o icone na Barra de Tarefas do Windows
Public Const WM_MOUSEISMOVING = &H200 ' Mouse is moving
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As enm_NIM_Shell, pnid As NOTIFYICONDATA) As Boolean

'Tipos usados pela funcao Shell_NotifyIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
    Tela As String
End Type
Public Enum enm_NIM_Shell
    NIM_ADD = &H0
    NIM_MODIFY = &H1
    NIM_DELETE = &H2
    NIF_MESSAGE = &H1
    NIF_ICON = &H2
    NIF_TIP = &H4
    WM_MOUSEMOVE = &H200
End Enum
Public IconeTela As NOTIFYICONDATA

'Funcoes API para de menus
Public Declare Function GetMenu Lib "user32" _
    (ByVal hwnd As Long) As Long

Public Declare Function GetSubMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPos As Long) As Long

Public Declare Function SetMenuItemBitmaps Lib _
    "user32" (ByVal hMenu As Long, ByVal nPosition _
    As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked _
    As Long, ByVal hBitmapChecked As Long) As Long

Declare Function DeleteMenu Lib "user32" _
    (ByVal hMenu As Long, ByVal nPosition As Long, _
    ByVal wFlags As Long) As Long

'Constante de posicao para inserir figura no menu
Public Const MF_BYPOSITION = &H400&
Public Const MF_BYCOMMAND = &H0&

Public msg As Long


