VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetalleSegManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection
'todos los detalles de esta coleccion son de la misma atencion
Private mIdAtencion As Long

Public Property Get idAtencion() As Long
    idAtencion = mIdAtencion
End Property

Public Property Let idAtencion(pId As Long)
    mIdAtencion = pId
End Property

'##ModelId=46A8DFD8000F
Public Sub AddItem(pDetalle As DetalleSeguimiento)
    On Error GoTo AddItemErr

    mCol.Add pDetalle ', "k" + Str(pDetalle.id)

    Exit Sub
AddItemErr:
If err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "DetSegManager", "AddItem", err.Description
End If
End Sub

Friend Function Add(pId As Long, pCampo As String, pValorAnterior As String, pValorNuevo As String, pEmpleado As Empleado, pFecha As Date, pHora As String) As DetalleSeguimiento
    Dim c As New DetalleSeguimiento
    c.Id = pId
    c.campo = pCampo
    Set c.Empleado = pEmpleado
    c.Fecha = pFecha
    c.Hora = pHora
    c.ValorAnterior = pValorAnterior
    c.ValorNuevo = pValorNuevo
    'es comun para todos los items de esta atencion
    c.idAtencion = mIdAtencion
    AddItem c
    Set Add = c
    Set c = Nothing
End Function

Public Function Nuevo(pCampo As String, pValorAnterior As String, pValorNuevo As String, pEmpleado As Empleado, pFecha As Date, pHora As String) As DetalleSeguimiento
    Dim c As New DetalleSeguimiento
    c.campo = pCampo
    Set c.Empleado = pEmpleado
    c.Fecha = pFecha
    c.Hora = pHora
    c.ValorAnterior = pValorAnterior
    c.ValorNuevo = pValorNuevo
    'es comun para todos los items de esta atencion
    c.idAtencion = mIdAtencion
    AddItem c
    Set Nuevo = c
    Set c = Nothing
End Function

Public Sub CargarXAtencion(pIdAtencion)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

mIdAtencion = pIdAtencion

Dim rs
Set rs = BD.SelectSQL("Select * from detalleseguimiento where idAtencion=" + Str(pIdAtencion))
While Not rs.EOF
        
    Add rs!Id, rs!campo, rs!ValorAnterior, rs!ValorNuevo, EmpleadosLocal.Item(rs!IdEmpleado), rs!Fecha, rs!Hora
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "DetalleSegManager", "cargarTodos ", err.Description
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

Public Property Get Item(pId As Long) As DetalleSeguimiento
    On Error Resume Next
    Dim d As DetalleSeguimiento
    For Each d In Me
        If d.Id = pId Then
            Set Item = d
            Exit For
        End If
    Next
    'Set Item =mCol.Item("k" + Str(pId))
End Property

Public Sub Save()

On Error GoTo errman
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    Dim d As DetalleSeguimiento
    
    For Each d In Me
        d.Save trans
    Next

    trans.Execute
    Exit Sub
errman:
    ErrorLog "DetallesegManager", "Save", err.Description
End Sub
