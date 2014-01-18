Attribute VB_Name = "ModuloConfigInicial"
Option Explicit

Public Const MIN_SOCKETS_REQD As Long = 1
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const SOCKET_ERROR As Long = -1
Public Const ERROR_SUCCESS As Long = 0
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Type WSAData
   wVersion As Integer
   wHighVersion As Integer
   szDescription(0 To MAX_WSADescription) As Byte
   szSystemStatus(0 To MAX_WSASYSStatus) As Byte
   wMaxSockets As Integer
   wMaxUDPDG As Integer
   dwVendorInfo As Long
End Type
Type WSADataInfo
   wVersion As Integer
   wHighVersion As Integer
   szDescription As String * WSADESCRIPTION_LEN
   szSystemStatus As String * WSASYS_STATUS_LEN
   iMaxSockets As Integer
   iMaxUdpDg As Integer
   lpVendorInfo As String
End Type
Public Type HOSTENT
   hName As Long
   hAliases As Long
   hAddrType As Integer
   hLen As Integer
   hAddrList As Long
End Type

Declare Function WSAStartupInfo Lib "WSOCK32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSADATA As WSADataInfo) As Long
Declare Function WSACleanup Lib "WSOCK32" () As Long
Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Declare Function gethostname Lib "WSOCK32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Declare Function gethostbyname Lib "WSOCK32" (ByVal szHost As String) As Long
Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Type WKSTA_INFO_101
   wki101_platform_id As Long
   wki101_computername As Long
   wki101_langroup As Long
   wki101_ver_major As Long
   wki101_ver_minor As Long
   wki101_lanroot As Long
End Type

Type WKSTA_USER_INFO_1
   wkui1_username As Long
   wkui1_logon_domain As Long
   wkui1_logon_server As Long
   wkui1_oth_domains As Long
End Type

Declare Function WNetGetUser& Lib "Mpr" Alias "WNetGetUserA" _
   (lpName As Any, ByVal lpUserName$, lpnLength&)
Declare Function NetWkstaGetInfo& Lib "netapi32" _
   (strServer As Any, ByVal lLevel&, pbBuffer As Any)
Declare Function NetWkstaUserGetInfo& Lib "netapi32" _
   (reserved As Any, ByVal lLevel&, pbBuffer As Any)
Declare Sub lstrcpyW Lib "kernel32" (dest As Any, ByVal src As Any)
Declare Sub lstrcpy Lib "kernel32" (dest As Any, ByVal src As Any)
Declare Sub RtlMoveMemory Lib "kernel32" _
   (dest As Any, src As Any, ByVal size&)
Declare Function NetApiBufferFree& Lib "netapi32" (ByVal buffer&)


Private Const NERR_SUCCESS As Long = 0&
'tipos compartidos
Private Const STYPE_ALL As Long = -1 'note: my const
Private Const STYPE_DISKTREE As Long = 0
Private Const STYPE_PRINTQ As Long = 1
Private Const STYPE_DEVICE As Long = 2
Private Const STYPE_IPC As Long = 3
Private Const STYPE_SPECIAL As Long = &H80000000
'permisos
Private Const ACCESS_READ As Long = &H1
Private Const ACCESS_WRITE As Long = &H2
Private Const ACCESS_CREATE As Long = &H4
Private Const ACCESS_EXEC As Long = &H8
Private Const ACCESS_DELETE As Long = &H10
Private Const ACCESS_ATRIB As Long = &H20
Private Const ACCESS_PERM As Long = &H40
Private Const ACCESS_ALL As Long = ACCESS_READ Or _
ACCESS_WRITE Or _
ACCESS_CREATE Or _
ACCESS_EXEC Or _
ACCESS_DELETE Or _
ACCESS_ATRIB Or _
ACCESS_PERM
Private Type SHARE_INFO_2
shi2_netname As Long
shi2_type As Long
shi2_remark As Long
shi2_permissions As Long
shi2_max_uses As Long
shi2_current_uses As Long
shi2_path As Long
shi2_passwd As Long
End Type

Private Declare Function NetShareAdd Lib "netapi32" _
(ByVal servername As Long, _
ByVal level As Long, _
buf As Any, _
parmerr As Long) As Long

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
Const APPLICATION As String = "tbrEG"

Dim m_Left As Single
Dim m_Top As Single
Dim m_Width As Single
Dim m_Height As Single

Dim Path_Archivo_Ini As String

''Función api que recupera un valor-dato de un archivo Ini
'Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
'    ByVal lpApplicationName As String, _
'    ByVal lpKeyName As String, _
'    ByVal lpDefault As String, _
'    ByVal lpReturnedString As String, _
'    ByVal nSize As Long, _
'    ByVal lpFileName As String) As Long
'
''Función api que Escribe un valor - dato en un archivo Ini
'Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
'    ByVal lpApplicationName As String, _
'    ByVal lpKeyName As String, _
'    ByVal lpString As String, _
'    ByVal lpFileName As String) As Long
'

''Lee un dato _
'-----------------------------
''Recibe la ruta del archivo, la clave a leer y _
' el valor por defecto en caso de que la Key no exista
'Public Function Leer_Ini(Path_INI As String, Key As String, Default As Variant) As String
'
'Dim bufer As String * 256
'Dim Len_Value As Long
'
'        Len_Value = GetPrivateProfileString(APPLICATION, _
'                                         Key, _
'                                         Default, _
'                                         bufer, _
'                                         Len(bufer), _
'                                         Path_INI)
'
'        Leer_Ini = Left$(bufer, Len_Value)
'
'End Function
'
''Escribe un dato en el INI _
'-----------------------------
''Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave
'
'Public Function Grabar_Ini(Path_INI As String, Key As String, Valor As Variant) As String
'
'    WritePrivateProfileString APPLICATION, _
'                                         Key, _
'                                         Valor, _
'                                         Path_INI
'
'End Function


'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Public Function ShareAdd(sServer As String, sSharePath As String, sShareName As String, sShareRemark As String, sSharePw As String) As Long
Dim dwServer As Long
Dim dwNetname As Long
Dim dwPath As Long
Dim dwRemark As Long
Dim dwPw As Long
Dim parmerr As Long
Dim si2 As SHARE_INFO_2

'Obtiene los punteros del servidor
dwServer = StrPtr(sServer)
dwNetname = StrPtr(sShareName)
dwPath = StrPtr(sSharePath)

If Len(sShareRemark) > 0 Then
dwRemark = StrPtr(sShareRemark)
End If

If Len(sSharePw) > 0 Then
dwPw = StrPtr(sSharePw)
End If

'Estructura SHARE_INFO_2
With si2
.shi2_netname = dwNetname
.shi2_path = dwPath
.shi2_remark = dwRemark
.shi2_type = STYPE_DISKTREE
.shi2_permissions = ACCESS_ALL
.shi2_max_uses = -1
.shi2_passwd = dwPw
End With

'Añadir recurso
ShareAdd = NetShareAdd(dwServer, _
2, _
si2, _
parmerr)

End Function

Function GetWorkstationInfo()
   Dim ret As Long, buffer(512) As Byte, I As Integer
   Dim wk101 As WKSTA_INFO_101, pwk101 As Long
   Dim wk1 As WKSTA_USER_INFO_1, pwk1 As Long
   Dim cbusername As Long, username As String
   Dim computername As String, langroup As String, logondomain As _
      String

   ' Clear all of the display values.
   computername = "": langroup = "": username = "": logondomain = ""

   ' Windows 95 or NT - call WNetGetUser to get the name of the user.
   username = Space(256)
   cbusername = Len(username)
   ret = WNetGetUser(ByVal 0&, username, cbusername)
   If ret = 0 Then
      ' Success - strip off the null.
      username = Left(username, InStr(username, Chr(0)) - 1)
   Else
      username = ""
   End If

'==================================================================
' The following section works only under Windows NT or Windows 2000
'==================================================================

   'NT only - call NetWkstaGetInfo to get computer name and lan group
   ret = NetWkstaGetInfo(ByVal 0&, 101, pwk101)
   RtlMoveMemory wk101, ByVal pwk101, Len(wk101)
   lstrcpyW buffer(0), wk101.wki101_computername
   ' Get every other byte from Unicode string.
   I = 0
   Do While buffer(I) <> 0
      computername = computername & Chr(buffer(I))
      I = I + 2
   Loop
   lstrcpyW buffer(0), wk101.wki101_langroup
   I = 0
   Do While buffer(I) <> 0
      langroup = langroup & Chr(buffer(I))
      I = I + 2
   Loop
   ret = NetApiBufferFree(pwk101)

   ' NT only - call NetWkstaUserGetInfo.
   ret = NetWkstaUserGetInfo(ByVal 0&, 1, pwk1)
   RtlMoveMemory wk1, ByVal pwk1, Len(wk1)
   lstrcpyW buffer(0), wk1.wkui1_logon_domain
   I = 0
   Do While buffer(I) <> 0
      logondomain = logondomain & Chr(buffer(I))
      I = I + 2
   Loop
   ret = NetApiBufferFree(pwk1)

'================================================================
'End NT/Windows 2000-specific section
'================================================================

   Debug.Print computername, langroup, username, logondomain
End Function

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

Public Function GetIPAddress() As String
   Dim sHostName As String * 256
   Dim lpHost As Long
   Dim HOST As HOSTENT
   Dim dwIPAddr As Long
   Dim tmpIPAddr() As Byte
   Dim I As Integer
   Dim sIPAddr As String
   If Not SocketsInitialize() Then
       GetIPAddress = ""
       Exit Function
   End If
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
       GetIPAddress = ""
       MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
       SocketsCleanup
       Exit Function
   End If
   sHostName = Trim$(sHostName)
   lpHost = gethostbyname(sHostName)
   If lpHost = 0 Then
       GetIPAddress = ""
       MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
       SocketsCleanup
       Exit Function
   End If
   CopyMemoryIP HOST, lpHost, Len(HOST)
   CopyMemoryIP dwIPAddr, HOST.hAddrList, 4
   ReDim tmpIPAddr(1 To HOST.hLen)
   CopyMemoryIP tmpIPAddr(1), dwIPAddr, HOST.hLen
   For I = 1 To HOST.hLen
       sIPAddr = sIPAddr & tmpIPAddr(I) & "."
   Next
   GetIPAddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
   SocketsCleanup
End Function
Public Function GetIPHostName() As String
   Dim sHostName As String * 256
   If Not SocketsInitialize() Then
       GetIPHostName = ""
       Exit Function
   End If
   If gethostname(sHostName, 256) = SOCKET_ERROR Then
       GetIPHostName = ""
       MsgBox "Windows Sockets error " & Str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
       SocketsCleanup
       Exit Function
   End If
   GetIPHostName = Left$(sHostName, InStr(sHostName, Chr(0)) - 1)
   SocketsCleanup
End Function
Public Function HiByte(ByVal wParam As Integer)
   HiByte = wParam \ &H100 And &HFF&
End Function
Public Function LoByte(ByVal wParam As Integer)
   LoByte = wParam And &HFF&
End Function
Public Sub SocketsCleanup()
   If WSACleanup() <> ERROR_SUCCESS Then
       MsgBox "Socket error occurred in Cleanup."
   End If
End Sub
Public Function SocketsInitialize() As Boolean
   Dim WSAD As WSAData
   Dim sLoByte As String
   Dim sHiByte As String
   If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
       MsgBox "The 32-bit Windows Socket is not responding."
       SocketsInitialize = False
       Exit Function
   End If
   If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
       MsgBox "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
       SocketsInitialize = False
       Exit Function
   End If
   If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
       sHiByte = CStr(HiByte(WSAD.wVersion))
       sLoByte = CStr(LoByte(WSAD.wVersion))
       MsgBox "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
       SocketsInitialize = False
       Exit Function
   End If
   'must be OK, so lets do it
   SocketsInitialize = True
End Function

