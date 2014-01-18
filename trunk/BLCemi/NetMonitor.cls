VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NetMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim WithEvents mRed As Red
Attribute mRed.VB_VarHelpID = -1
Dim WithEvents myConfig As tbrconfig.clsConfiguracion
Attribute myConfig.VB_VarHelpID = -1

Public Event AtencionesChanged()
Public Event CuotasChanged()
Public Event NetStatusChanged(pState As String)
Public Event PedirNombreUsuario(pNombreUsuario As String)

Private Sub Class_Initialize()
    Set mRed = New Red
    Set myConfig = CCFFGG.Configuracion
    mRed.RemoteIP = CCFFGG.Configuracion.Red.DirIPRemota 'se asigna solo si es cliente, sino la ignora
    mRed.Conectar IIf(CCFFGG.Configuracion.Red.ModoServer, 2, 1)
    If Not CCFFGG.Configuracion.Red.ModoServer Then mRed.HostName = CCFFGG.Configuracion.Red.Nombre

End Sub

Private Sub mRed_RecibirEventoABM(pTipo As tbrNet.eTipoEvento, pTabla As String, Id As Long)
    Dim tabla As String
    tabla = LCase(pTabla)
    Select Case tabla
        Case "atencion", "atencionb"
            RaiseEvent AtencionesChanged
        Case "cuota"
            RaiseEvent CuotasChanged
    End Select
End Sub

Public Property Get MiRedLocal() As Red
    Set MiRedLocal = mRed
End Property

Private Sub mRed_StateChanged(pEstado As String)
    RaiseEvent NetStatusChanged(pEstado)
End Sub

Private Sub mRed_UserNameRequest(outUserName As String)
    RaiseEvent PedirNombreUsuario(outUserName)
End Sub

Private Sub myConfig_ConfigChanged()
    mRed.Desconectar
    mRed.HostName = CCFFGG.Configuracion.Red.Nombre
    mRed.RemoteIP = CCFFGG.Configuracion.Red.DirIPRemota
    mRed.Conectar IIf(CCFFGG.Configuracion.Red.ModoServer, 2, 1)
End Sub

Public Sub ForzarConexion()
    mRed.RemoteIP = CCFFGG.Configuracion.Red.DirIPRemota 'se asigna solo si es cliente, sino la ignora
    mRed.Conectar IIf(CCFFGG.Configuracion.Red.ModoServer, 2, 1)
End Sub
