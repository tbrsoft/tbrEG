VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conexion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Enum eTipoMensaje
    eEventoABM = 1
    eHostNameRequest = 2
    eHostNameResponse = 3
End Enum

Public Event EventoABM(pTipo As Integer, pTabla As String, Id As Long)
Public Event ConnectionStateChanged(pState As String)
Public Event PedirNombreUsuario(outNombreUsuario As String)

Private myCol As ConexionManager

Private bCliente As Boolean
Private mId As Integer
Private mName As String
Private mIp As String
Private mUsuario As String

Public Property Get Usuario() As String
    Usuario = mUsuario
End Property

Public Property Let Usuario(pUsuario As String)
    mUsuario = pUsuario
End Property

Public Property Get Nombre() As String
    Nombre = mName
End Property

Public Property Let Nombre(pName As String)
    mName = pName
End Property

Public Property Get RemoteIP() As String
    RemoteIP = mIp
End Property

Public Property Let RemoteIP(pIp As String)
    mIp = pIp
End Property

Friend Sub Inicializar(pId As Integer, pCol As ConexionManager)
    Set myCol = pCol
    mId = pId
    bCliente = False
    Me.PedirNombre
End Sub

Public Sub enviarDatos(datos As String)
    Dim DatosaEnviar As String
    DatosaEnviar = datos
    wSock.SendData DatosaEnviar
End Sub

Private Sub Class_Terminate()
    Set myCol = Nothing
End Sub

'el frm me pasa los eventos a traves de estas funciones porq no puedo agarrar los eventos
'Friend Sub wSock_Close()
'    myCol.CerrarConexion wSock.SocketHandle
'End Sub
 
Friend Sub wSock_DataArrival(pData As String)
    'RaiseEvent DatosRecibidos(DatosRecibidos)
    Debug.Print pData
    Dim aux() As String
    Dim mTipoMensaje As eTipoMensaje
    aux = Split(pData, "|")
    mTipoMensaje = Val(Trim(aux(0)))
    Select Case mTipoMensaje
        'modelo de mensaje
        'tipoMensaje|tipoEvento|tabla|id
        Case eTipoMensaje.eEventoABM
            If bCliente Then
                RaiseEvent EventoABM(Val(aux(1)), aux(2), CLng(aux(3)))
            Else
                myCol.EventoABM mId, CInt(aux(1)), aux(2), CLng(aux(3))
            End If
        Case eTipoMensaje.eHostNameRequest
            RaiseEvent PedirNombreUsuario(mUsuario)  'aca le asigno un nonmbre de usuario
                                    'tipomensaje=hostanameresponse
            frm.wskCliente.SendData "3|" + mName + "|" + mUsuario
        Case eTipoMensaje.eHostNameResponse
            Nombre = aux(1)
            Usuario = aux(2)
            myCol.RecibirNombre
    End Select
End Sub

Friend Sub wSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    wSock.Close
  '  If Servidor Then esperarConexion Else Conectar
End Sub

Public Function EnviarEventoABM(pTipo As Integer, pTabla As String, Id As Long)
    On Error GoTo erman:
    If bCliente Then
                                'tipomensaje
        frm.wskCliente.SendData "1|" + Trim(Str(pTipo)) + "|" + pTabla + "|" + Trim(Str(Id))
    Else
                                'tipomensaje
        frm.wskServer(mId).SendData "1|" + Trim(Str(pTipo)) + "|" + pTabla + "|" + Trim(Str(Id))
     
    End If
Exit Function
erman:
If Err.Number = 40006 Then
    If bCliente Then
        Conectar
    Else
        frm.wskServer_Close mId
        myCol.Remove mId
    End If
End If
End Function

Public Function Conectar()
    bCliente = True
    frm.wskCliente.Close
    If mIp <> "" Then frm.wskCliente.RemoteHost = mIp
    frm.wskCliente.Connect
    frm.lblEstado = "Intentando conectar..."
    RaiseEvent ConnectionStateChanged("Intentando conectar...")
    frm.lblModo = "Cliente"
End Function

Public Sub Desconectar()
    RaiseEvent ConnectionStateChanged("Desconectado")
    If bCliente Then
        frm.wskCliente.Close
        Set frm.Cliente = Nothing
    End If
End Sub

Public Property Get Id() As Integer
    Id = mId
End Property

Friend Sub PedirNombre()
    frm.wskServer(mId).SendData "2|"
End Sub

Friend Sub CambiarEstado(pEstado As String)
    RaiseEvent ConnectionStateChanged(pEstado)
End Sub
