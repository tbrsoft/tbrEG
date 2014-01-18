VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AfiliadoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A80F7000DA"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Event ItemAdded(pAfiliado As Afiliado)
Public Event ItemRemoved(pAfiliado As Afiliado)
Public Event ItemModified(pAfiliado As Afiliado)

Private mCol As Collection
Private mColEliminados As Collection

'##ModelId=46A80F8302AF
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr
    
    RaiseEvent ItemRemoved(Me.Item(pId))
    mCol.Remove "k" + Str(pId)
    
    Exit Sub
RemoveErr:
    ErrorLog "AfiliadoManager", "Remove", Err.Description
End Sub

'##ModelId=46A80F8501B5
Public Sub AddItem(pAfiliado As Afiliado)
    On Error GoTo AddItemErr

    mCol.Add pAfiliado, "k" + Str(pAfiliado.Id)

    Exit Sub
AddItemErr:
    ErrorLog "AfiliadoManager", "AddItem", Err.Description
End Sub

Public Function Nuevo(pApellido As String, pCobrador As Empleado, pDireccion As Direccion, pEstadoCivil As EstadoCivil, pFechaInscripcion As Date, pFechaNacimiento As Date, pId As Long, pInicioPrestacion As Date, pNombre As String, pNrodoc As Long, pObraSocial As ObraSocial, pObservaciones As String, pOcupacion As Ocupacion, pPersonasACargo As AfiliadoManager, pSexo As Integer, pTelefonos As TelefonoManager, pTipoDoc As TipoDoc, pTopeAtenciones As Integer, pVehiculo As Vehiculo, pHistoriaClinica As HistoriaClinica, pImporte As Currency)
On Error GoTo errman:

Dim a As New Afiliado

a.TipoAfiliado = eTitular

a.Apellido = pApellido
'a.Atenciones un afiliado nuevo todavia no tiene atenciones
Set a.Cobrador = pCobrador
Set a.Direccion = pDireccion
Set a.EstadoCivil = pEstadoCivil
a.FechaInscripcion = pFechaInscripcion
a.FechaNacimiento = pFechaNacimiento
Set a.HistoriaClinica = pHistoriaClinica
a.Id = pId
a.Importe = pImporte
a.InicioPrestacion = pInicioPrestacion
a.Nombre = pNombre
a.NroDoc = pNrodoc
Set a.ObraSocial = pObraSocial
a.Observaciones = pObservaciones
Set a.Ocupacion = pOcupacion
'a.Pagos un afiliado nuevo todavia no tiene pagos
'a.Parent un afiliado titular no tiene parent
'a.Parentezco un afiliado titular no tiene parentezco
Set a.PersonasACargo = pPersonasACargo
a.Sexo = pSexo
Set a.Telefonos = pTelefonos
Set a.TipoDoc = pTipoDoc
a.TopeAtenciones = pTopeAtenciones
Set a.Vehiculo = pVehiculo

a.InsertTitular

AddItem a

Set Nuevo = a

Exit Function
errman:
ErrorLog "AfiliadoManager", "Nuevo", Err.Description
End Function

Public Function NuevoACargo(pApellido As String, pDireccion As Direccion, pEstadoCivil As EstadoCivil, pFechaInscripcion As Date, pFechaNacimiento As Date, pId As Long, pInicioPrestacion As Date, pNombre As String, pNrodoc As Long, pObraSocial As ObraSocial, pObservaciones As String, pOcupacion As Ocupacion, pSexo As Integer, pTelefonos As TelefonoManager, pTipoDoc As TipoDoc, pTopeAtenciones As Integer, pParentezco As Parentezco, pHistoriaClinica As HistoriaClinica, pParent As Afiliado)
On Error GoTo errman

Dim a As New Afiliado

a.TipoAfiliado = eACargo

a.Apellido = pApellido
'a.Atenciones un afiliado nuevo todavia no tiene atenciones
'a.Cobrador = pCobrador es el mismo q en parent
Set a.Direccion = pDireccion
Set a.EstadoCivil = pEstadoCivil
a.FechaInscripcion = pFechaInscripcion
a.FechaNacimiento = pFechaNacimiento
Set a.HistoriaClinica = pHistoriaClinica
a.Id = pId
a.InicioPrestacion = pInicioPrestacion
a.Nombre = pNombre
a.NroDoc = pNrodoc
Set a.ObraSocial = pObraSocial
a.Observaciones = pObservaciones
Set a.Ocupacion = pOcupacion
'a.Pagos un afiliado nuevo todavia no tiene pagos
'a.Parent se lo seteo en guardar porq todavia no lo tengo creado
Set a.Parentezco = pParentezco
'a.PersonasACargo = pPersonasACargo
a.Sexo = pSexo
Set a.Telefonos = pTelefonos
Set a.TipoDoc = pTipoDoc
a.TopeAtenciones = pTopeAtenciones
'a.Vehiculo = pVehiculo

If Not pParent Is Nothing Then
    Dim t As Transaccion
    Set t = BD.Transaccion
    a.InsertACargo t, pParent
    t.Execute
End If
'a.InsertACargo  eACargo no va porq no lo guardo hasta q no guarde el afiliado titular

AddItem a
Set NuevoACargo = a

Exit Function
errman:
ErrorLog "AfiliadoManager", "NuevoACargo", Err.Description
End Function

Friend Function Add(pId As Long, pApellido As String, pNombre As String, pTipoDoc As TipoDoc, pNrodoc As Long, pIdDireccion As Long, pEstadoCivil As EstadoCivil, pFechaInscripcion As Date, pFechaNacimiento As Date, pInicioPrestacion As Date, pObraSocial As ObraSocial, pObservaciones As String, pOcupacion As Ocupacion, pSexo As Integer, pTopeAtenciones As Integer, pParentezco As Parentezco, pCobrador As Empleado, pImporte As Currency, pParent As Afiliado, pIdVehiculo As Long, pEliminado As Boolean)
On Error GoTo errman:

Dim a As New Afiliado

a.TipoAfiliado = IIf(pParent Is Nothing, eTitular, eACargo) '  ver

a.Apellido = pApellido
'a.Atenciones ver
Set a.Cobrador = pCobrador
a.Direccion.Cargar pIdDireccion
Set a.EstadoCivil = pEstadoCivil
a.FechaInscripcion = pFechaInscripcion
a.FechaNacimiento = pFechaNacimiento
'Set a.HistoriaClinica = pHistoriaClinica  on demand
a.Id = pId
a.InicioPrestacion = pInicioPrestacion
a.Importe = pImporte
a.Nombre = pNombre
a.NroDoc = pNrodoc
Set a.ObraSocial = pObraSocial
a.Observaciones = pObservaciones
Set a.Ocupacion = pOcupacion
'a.Pagos un afiliado nuevo todavia no tiene pagos
Set a.Parent = pParent
Set a.Parentezco = pParentezco
'a.PersonasACargo = pPersonasACargo
a.Sexo = pSexo
'Set a.Telefonos = pTelefonos on demand
Set a.TipoDoc = pTipoDoc
a.TopeAtenciones = pTopeAtenciones
If pIdVehiculo <> 0 Then
    Set a.Vehiculo = New Vehiculo
    a.Vehiculo.Cargar pIdVehiculo
End If

If Not pEliminado Then
    AddItem a
Else
    mColEliminados.Add a, "k" + Str(pId)
End If

Set Add = a

Exit Function
errman:
ErrorLog "AfiliadoManager", "Add", Err.Description

End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mColEliminados = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
    Set mColEliminados = Nothing
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get Item(pId As Long) As Afiliado
On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
If Err.Number <> 0 Then 'si no lo encuentra se fija en eliminados
    Err.Clear
    Set Item = mColEliminados.Item("k" + Str(pId))
    If Err.Number <> 0 Then
        Set Item = Nothing 'si tampoco esta es porq no existe
    End If
End If
End Property

Public Function GetAfiliadosTitulares() As AfiliadoManager
On Error GoTo errMaq

Dim am As New AfiliadoManager
Dim af As Afiliado
For Each af In Me
    If af.TipoAfiliado = eTitular Then
        am.AddItem af
    End If
Next
Set GetAfiliadosTitulares = am
Exit Function
errMaq:
ErrorLog "AfiliadosManager", "GetAfiliadosTitulares", Err.Description
End Function

Public Function GetAfiliadosACargo(pAfiliadoTitular As Afiliado) As AfiliadoManager
On Error GoTo errMaq
Dim am As New AfiliadoManager
Dim af As Afiliado
For Each af In Me
    If af.TipoAfiliado = eACargo Then 'si es a cargo me fijo el parent
        If af.Parent.Id = pAfiliadoTitular.Id Then
            am.AddItem af
        End If
    End If
Next
Set GetAfiliadosACargo = am

Exit Function
errMaq:
ErrorLog "AfiliadosManager", " GetAfiliadosAcargo", Err.Description
End Function

Public Function GetEliminados() As AfiliadoManager
    Dim am As New AfiliadoManager
    Dim a As Afiliado
    For Each a In mColEliminados
        am.AddItem a
    Next
    Set GetEliminados = am
End Function

Public Sub DarItemDeBaja(pId As Long)
   'esto es porq la unica coleccion q tiene los elementos eliminados es empleadosLocal
   AfiliadosLocal.DarItemDeBajaLocal pId
   RaiseEvent ItemRemoved(AfiliadosLocal.Item(pId))
End Sub

Public Sub DarItemDeBajaLocal(pId As Long)
    Dim e As Afiliado
    Set e = Me.Item(pId)
    If Not e Is Nothing Then
        e.DarDeBaja
        mColEliminados.Add e, "k" + Str(e.Id)
        mCol.Remove "k" + Str(e.Id)
    End If
End Sub

Public Sub RestaurarItem(pId As Long)
    'esto es porq la unica coleccion q tiene los elementos eliminados es empleadosLocal
    AfiliadosLocal.RestaurarItemLocal pId
End Sub

Friend Sub RestaurarItemLocal(pId As Long)
    Dim e As Afiliado
    Set e = Me.Item(pId) 'en item tambien devuelvo los eliminados
    e.Restaurar
    mColEliminados.Remove "k" + Str(e.Id)
    AddItem e
    
End Sub

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
'cargo todos, incluso los eliminados... despues los filtro en el add...
On Error GoTo errMaq
'puede no tener obra social
Dim mObraSocialAux As ObraSocial

Dim rs
'tengo q cargar primero los titulares y despues los acargo
Set rs = BD.SelectSQL("select * from afiliado where idafiliadotitular is null") ' and eliminado like 'false'")
While Not rs.EOF
    
    If Not IsNull(rs!idObraSocial) Then
        Set mObraSocialAux = ObrasSocialesLocal.Item(rs!idObraSocial)
    Else
        Set mObraSocialAux = Nothing
    End If
    
    Add rs!Id, rs!Apellido, rs!Nombre, TiposDocumentoLocal.Item(rs!TipoDoc), rs!NroDoc, rs!idDireccion, EstadosCivilesLocal.Item(rs!EstadoCivil), rs!FechaInscripcion, rs!FechaNacimiento, rs!InicioPrestacion, mObraSocialAux, rs!Observaciones, OcupacionesLocal.Item(rs!idocupacion), rs!Sexo, rs!TopeAtenciones, Nothing, EmpleadosLocal.Item(rs!IdEmpleado), rs!Importe, Nothing, rs!idVehiculo, CBool(rs!eliminado)
    rs.MoveNext
    
Wend
'si idaftitular no es null esta a cargo de alguien...

Set rs = BD.SelectSQL("select * from afiliado where idafiliadotitular  is not null")
While Not rs.EOF
    
    If Not IsNull(rs!idObraSocial) Then
        Set mObraSocialAux = ObrasSocialesLocal.Item(rs!idObraSocial)
    Else
        Set mObraSocialAux = Nothing
    End If                                                                                                                                                                                                                                                                                                            'cero porque no tiene tope de atenciones
    
    Add rs!Id, rs!Apellido, rs!Nombre, TiposDocumentoLocal.Item(rs!TipoDoc), rs!NroDoc, rs!idDireccion, EstadosCivilesLocal.Item(rs!EstadoCivil), rs!FechaInscripcion, rs!FechaNacimiento, rs!InicioPrestacion, mObraSocialAux, rs!Observaciones, OcupacionesLocal.Item(rs!idocupacion), rs!Sexo, 0, ParentezcoLocal.Item(rs!idParentezco), Nothing, rs!Importe, Me.Item(rs!idafiliadotitular), 0, CBool(rs!eliminado)
    rs.MoveNext
    
Wend

Exit Sub
errMaq:
ErrorLog "AfiliadosManager", "cargarTodos ", Err.Description
End Sub

Friend Sub Guardar(trans As Transaccion, pParent As Afiliado)
    Dim af As Afiliado
    For Each af In Me
        af.InsertACargo trans, pParent
    Next
End Sub

Public Function ExisteId(pId As Long) As Boolean
Dim aux As Boolean
Dim rs
Set rs = BD.SelectSQL("select id from afiliado where id=" + Str(pId))
aux = Not (rs.EOF)
If Not aux Then 'si no existe en la bd me fijo si tengo alguno con ese mismo id
    Dim af As Afiliado
    For Each af In Me
        If af.Id = pId Then
            aux = True
            Exit For
        End If
    Next
End If
ExisteId = aux
Set rs = Nothing
End Function

Public Function GetUltimoIdTitular() As Long
    'esta funcion devuelve el ultimo id de la bd,
    'de todas maneras antes de hacer un insert chequeo q no exista el id
    Dim rs
        
    Set rs = BD.SelectSQL("Select max(id) from afiliado")
    Dim idAux As Long
    idAux = CLng(IIf(IsNull(rs.Fields(0).value), 0, rs.Fields(0).value))
    If idAux = 0 Then
        GetUltimoIdTitular = 0
    Else
        GetUltimoIdTitular = Int(idAux / 1000)
    End If
    
End Function

Public Function GetUltimoIdACargo(pIdParent As Long) As Long
    'esta funcion devuelve el ultimo id de la bd,
    'de todas maneras antes de hacer un insert chequeo q no exista el id
    Dim rs
        
    Set rs = BD.SelectSQL("Select max(id) from afiliado where idafiliadotitular between " + Str(pIdParent) + " and " + Str(pIdParent + 1000))
    Dim idAux As Long
    idAux = CLng(IIf(IsNull(rs.Fields(0).value), 0, rs.Fields(0).value))
    If idAux = 0 Then
        GetUltimoIdACargo = 0
    Else
        GetUltimoIdACargo = idAux - Int(idAux / 1000) * 1000
    End If
    
End Function

