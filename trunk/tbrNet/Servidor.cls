VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Servidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event RecibirEventoABM(pTipo As Integer, pTabla As String, Id As Long)
Public Event RefrescarConexiones()
Public Event ServerStateChanged(pState As String)

Private WithEvents mConexiones As ConexionManager
Attribute mConexiones.VB_VarHelpID = -1

Public Function Conectar()
    On Error Resume Next
    Load frm.wskServer(1) 'mmmm olor a negrada  ... (encapsulamiento de la clase = -1)
    frm.wskServer(1).Listen
    frm.lblEstado = "Esperando conexiones..."
    RaiseEvent ServerStateChanged("Esperando conexiones...")
    frm.lblModo = "Servidor"
End Function

Public Function EnviarEventoABM(pTipo As Integer, pTabla As String, Id As Long)
    mConexiones.EnviarEventoABM pTipo, pTabla, Id
End Function

Friend Sub AddNewConexion(Id As Integer)
mConexiones.Add Id
RaiseEvent ServerStateChanged("Esperando conexiones...")
End Sub

Private Sub Class_Initialize()
Set mConexiones = New ConexionManager
End Sub

Public Property Get Conexiones() As ConexionManager
    Set Conexiones = mConexiones
End Property

Private Sub mConexiones_HostNameResponse()
    RaiseEvent RefrescarConexiones
End Sub

Private Sub mConexiones_RecibirEventoABM(pTipo As Integer, pTabla As String, Id As Long)
    RaiseEvent RecibirEventoABM(pTipo, pTabla, Id)
End Sub

Private Sub mConexiones_Refresh()
    RaiseEvent RefrescarConexiones
End Sub

Friend Function PedirNombresHost()
    mConexiones.PedirNombres
End Function
