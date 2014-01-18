VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Red"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum eModo
    emCliente = 1
    emServidor = 2
End Enum

Public Enum eTipoEvento
    Alta = 1
    Baja = 2
    Modificacion = 3
End Enum

Public Event RecibirEventoABM(pTipo As eTipoEvento, pTabla As String, Id As Long)
Public Event StateChanged(pEstado As String)
Public Event UserNameRequest(outUserName As String)

Private mModo As eModo

Private WithEvents mServer As Servidor
Attribute mServer.VB_VarHelpID = -1
Private WithEvents mCliente As Conexion
Attribute mCliente.VB_VarHelpID = -1

Private mNombre As String
Private mRemoteIp As String

Public Property Get HostName() As String
    If mModo = emCliente Then HostName = mCliente.Nombre
End Property

Public Property Let HostName(pHostName As String)
    If mModo = emCliente Then
        mCliente.Nombre = pHostName
        mNombre = pHostName
    End If
End Property

Public Property Let RemoteIP(pIp As String)
    If mModo = emCliente Then
        mCliente.RemoteIP = pIp
        mRemoteIp = pIp
    End If
End Property

Public Property Get RemoteIP() As String
If mModo = emCliente Then RemoteIP = mCliente.RemoteIP
End Property

Public Function EnviarEventoABM(pTipo As eTipoEvento, pTabla As String, Id As Long)
If Modo = emCliente Then
    mCliente.EnviarEventoABM CInt(pTipo), pTabla, Id
Else
    mServer.EnviarEventoABM CInt(pTipo), pTabla, Id
End If
End Function

Public Function Conectar(pModo As eModo)
'setear modo antes de llamar esta function
mModo = pModo
Load frm
If mModo = emCliente Then
    Set mCliente = New Conexion
    mCliente.Nombre = mNombre
    mCliente.RemoteIP = mRemoteIp
    Set frm.Cliente = mCliente
    mCliente.Conectar
Else
    Set mServer = New Servidor
    Set frm.Server = mServer
    mServer.Conectar
End If
End Function

Public Function Desconectar()
    If mModo = emCliente Then
        mCliente.Desconectar
    Else
        Unload frm
        Load frm
        'Set frm.Server = Nothing
    End If
End Function

Public Property Get Modo() As eModo
    Modo = mModo
End Property

Public Sub VerEstado()
    PedirNombres
    frm.Show
End Sub

Private Sub mCliente_EventoABM(pTipo As Integer, pTabla As String, Id As Long)
    RaiseEvent RecibirEventoABM(Choose(pTipo, eTipoEvento.Alta, eTipoEvento.Baja, eTipoEvento.Modificacion), pTabla, Id)
End Sub

Private Sub mCliente_PedirNombreUsuario(outNombreUsuario As String)
    RaiseEvent UserNameRequest(outNombreUsuario)
End Sub

Private Sub mServer_RecibirEventoABM(pTipo As Integer, pTabla As String, Id As Long)
    RaiseEvent RecibirEventoABM(Choose(pTipo, eTipoEvento.Alta, eTipoEvento.Baja, eTipoEvento.Modificacion), pTabla, Id)
End Sub

Public Sub PedirNombres()
    If mModo = emServidor Then mServer.PedirNombresHost
End Sub

'aviso a la capa superior q me cambio el estado
Private Sub mServer_ServerStateChanged(pState As String)
    RaiseEvent StateChanged(pState)
End Sub

Private Sub mCliente_ConnectionStateChanged(pState As String)
    RaiseEvent StateChanged(pState)
End Sub

