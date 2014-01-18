VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EmpleadoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A8157E009C"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Event ItemAdded(pEmpleado As Empleado)
Public Event ItemRemoved(pEmpleado As Empleado)
Public Event ItemModified(pEmpleado As Empleado)

Private mCol As Collection
Private mColOriginal As Collection

Private mColEliminados As Collection

'##ModelId=46A8158C01B5
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr
    
    RaiseEvent ItemRemoved(Item(pId))
    mCol.Remove "k" + Str(pId)
        
    Exit Sub
RemoveErr:
    Call ErrorLog("EmpleadoManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A8158E00EA
Public Sub AddItem(pEmpleado As Empleado)
    On Error GoTo AddItemErr

    mCol.Add pEmpleado, "k" + Str(pEmpleado.Id)
    RaiseEvent ItemAdded(pEmpleado)
    
    Exit Sub
AddItemErr:
If err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "EmpleadoManager", "AddItem", err.Description
End If
End Sub

Public Property Get Item(pId As Long) As Empleado
    On Error GoTo errman    'si salta un error probablemente este eliminado
    Set Item = mCol.Item("k" + Str(pId))
    Exit Property
errman:
    On Error Resume Next 'si salta un error aca, me esta pidiendo un empleado q no tengo
    Set Item = mColEliminados.Item("k" + Str(pId))
End Property

Public Function Nuevo(pApellido As String, pNombre As String, pCargos As CargoManager, pDireccion As Direccion, pFechaNac As Date, pLogin As String, pPass As String, pTipoDoc As TipoDoc, pNrodoc As Long, pPermisos As PermisoManager, pTelefonos As TelefonoManager, pMp As String) As Empleado
    Dim e As New Empleado
    
   'el id se lo asigno en empleado.guardar 'e.id = BD.getIdDisponible(tEmpleado)

    e.Apellido = pApellido
    Set e.Cargos = pCargos
    Set e.Direccion = pDireccion
    e.FechaNacimiento = pFechaNac
    e.Login = pLogin
    e.Nombre = pNombre
    e.NroDoc = pNrodoc
    e.Pass = pPass
    Set e.Permisos = pPermisos
    Set e.Telefonos = pTelefonos
    Set e.TipoDoc = pTipoDoc
    e.MP = pMp
    
    e.Guardar
    
    mCol.Add e, "k" + Str(e.Id)
    Set Nuevo = e
    
End Function

Public Function GetByCargoFijo(eCargo As eCargosFijos) As EmpleadoManager
Dim em As EmpleadoManager
Set em = New EmpleadoManager
Dim e As Empleado

For Each e In Me
    If Not e.Cargos.Item(eCargo) Is Nothing Then em.AddItem e
Next
Set GetByCargoFijo = em
End Function

Public Function GetByCargos(pCargos As CargoManager) As EmpleadoManager
Dim em As EmpleadoManager
Set em = New EmpleadoManager
Dim e As Empleado
Dim c As Cargo

For Each e In Me
    For Each c In pCargos
        If Not e.Cargos.Item(c.Id) Is Nothing Then
            em.AddItem e
            Exit For
        End If
    Next
Next
Set GetByCargos = em
End Function


Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mColEliminados = New Collection
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

Friend Sub Add(pId As Long, pNombre As String, pApellido As String, pFechaNac As Date, pNrodoc As Long, pTipoDoc As TipoDoc, pIdDireccion As Long, pPermisos As String, pLogin As String, pPass As String, pEsSuperUsuario As Boolean, pEliminado As Boolean, pMp As String)
Dim e As Empleado
Set e = New Empleado
e.Apellido = pApellido
e.Cargos.CargarPorEmpleado (pId)
e.Direccion.Cargar pIdDireccion
e.FechaNacimiento = pFechaNac
e.Id = pId
e.Login = pLogin
e.Nombre = pNombre
e.NroDoc = pNrodoc
e.Pass = pPass
e.Permisos.CadenaDePermisos = pPermisos
e.Permisos.EsSuperUsuario = pEsSuperUsuario
e.MP = pMp
'e.Telefonos
Set e.TipoDoc = pTipoDoc

If Not pEliminado Then
    mCol.Add e, "k" + Str(e.Id)
Else
    mColEliminados.Add e, "k" + Str(e.Id)
End If

End Sub

Friend Sub CargarEmpleados()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tEmpleado)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre, rs!Apellido, rs!FechaNacimiento, rs!NroDoc, TiposDocumentoLocal.Item(rs!TipoDoc), rs!idDireccion, rs!Permisos, rs!Login, rs!Pass, IIf(rs!EsSuperUsuario = 1, True, False), CBool(rs!eliminado), "" & rs!MP
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "EmpleadosManager", "cargarempleados ", err.Description
End Sub

Friend Sub CargarEmpleadosXEquipo(pIdEquipo As Long)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("SELECT idEmpleado FROM EmpleadoXEquipo WHERE idEquipo=" + Str(pIdEquipo))
While Not rs.EOF
        
    AddItem EmpleadosLocal.Item(rs!IdEmpleado)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "EmpleadosManager", "cargarempleadosxequipo ", err.Description
End Sub

Friend Sub GuardarXEquipo(pTrans As Transaccion, idEquipo As Long)

Dim e As Empleado
Dim params As ParametrosManager
For Each e In Me
    Set params = New ParametrosManager
    params.Add "idEmpleado", e.Id
    params.Add "idEquipo", idEquipo
    pTrans.AddInsertCommand tEmpleadoXEquipo, params
Next
End Sub

Friend Sub GuardarCambios(pTrans As Transaccion, idEquipo As Long)
    Dim params As New ParametrosManager
    params.Add "idEquipo", idEquipo, True
    pTrans.AddDeleteCommand tEmpleadoXEquipo, params
    GuardarXEquipo pTrans, idEquipo
    
End Sub

Public Function GetEliminados() As EmpleadoManager
    Dim em As New EmpleadoManager
    Dim e As Empleado
    For Each e In mColEliminados
        em.AddItem e
    Next
    Set GetEliminados = em
End Function

Public Sub DarItemDeBaja(pId As Long)
   'esto es porq la unica coleccion q tiene los elementos eliminados es empleadosLocal
   EmpleadosLocal.DarItemDeBajaLocal pId
   RaiseEvent ItemRemoved(EmpleadosLocal.Item(pId))
End Sub

Public Sub DarItemDeBajaLocal(pId As Long)
    Dim e As Empleado
    Set e = Me.Item(pId)
    e.DarDeBaja
    mColEliminados.Add e, "k" + Str(e.Id)
    mCol.Remove "k" + Str(e.Id)
End Sub

Public Sub RestaurarItem(pId As Long)
    'esto es porq la unica coleccion q tiene los elementos eliminados es empleadosLocal
    EmpleadosLocal.RestaurarItemLocal pId
End Sub

Friend Sub RestaurarItemLocal(pId As Long)
    Dim e As Empleado
    Set e = Me.Item(pId) 'en item tambien devuelvo los eliminados
    e.Restaurar
    mColEliminados.Remove "k" + Str(e.Id)
    AddItem e
    
End Sub

Public Sub BeginChanges()
    Set mColOriginal = New Collection
    Dim e As Empleado
    For Each e In Me
        mColOriginal.Add e, "k" + Str(e.Id)
    Next
End Sub

Public Sub CancelChanges()
    If Not mColOriginal Is Nothing Then
        Set mCol = mColOriginal
        Set mColOriginal = Nothing
    End If
End Sub
