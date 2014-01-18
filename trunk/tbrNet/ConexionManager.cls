VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConexionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event RecibirEventoABM(pTipo As Integer, pTabla As String, Id As Long)
Public Event HostNameResponse()
Public Event Refresh()

Private mCol As Collection

'Friend Sub CerrarConexion(idConexion As Long)
'    mCol.Remove "k" + Str(idConexion)
'End Sub

Friend Sub Add(Id As Integer)
    Dim c As New Conexion
    c.Inicializar Id, Me
    mCol.Add c, "k" + Str(Id)
    RaiseEvent Refresh
End Sub

Friend Sub Remove(Id As Integer)
    On Error Resume Next
    mCol.Remove "k" + Str(Id)
    RaiseEvent Refresh
End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get Item(Id As Integer) As Conexion
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(Id))
End Property

'para recibir un mensaje de una conexion
Friend Sub EventoABM(idConexion As Integer, pTipo As Integer, pTabla As String, Id As Long)
    'aca envio este mensaje a todas las conexiones excepto a la q me la mando
    EnviarEventoABM pTipo, pTabla, Id, idConexion
    
    RaiseEvent RecibirEventoABM(pTipo, pTabla, Id)
    
End Sub

Friend Sub EnviarEventoABM(pTipo As Integer, pTabla As String, Id As Long, Optional idExceptoA As Integer = -1)
Dim c As Conexion
For Each c In Me
    If c.Id <> idExceptoA Then c.EnviarEventoABM pTipo, pTabla, Id
Next
End Sub

'para recibir el nombre de una conexion
Friend Sub RecibirNombre()
    RaiseEvent HostNameResponse
End Sub

Friend Sub PedirNombres()
Dim c As Conexion
For Each c In Me
     c.PedirNombre
Next
End Sub
