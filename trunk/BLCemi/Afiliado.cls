VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Afiliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FDBF0000"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Enum eTipoAfiliado
    eTitular = 1
    eACargo = 2
End Enum

Private mId As Long
Private mNombre As String
Private mApellido As String
Private mTipoDoc As TipoDoc
Private mNroDoc As Long
Private mFechaNacimiento As Date
Private mObraSocial As ObraSocial
Private mTelefonos As TelefonoManager
Private mVehiculo As Vehiculo
Private mOcupacion As Ocupacion
Private mSexo As Integer
Private mEstadoCivil As EstadoCivil
Private mFechaInscripcion As Date
Private mInicioPrestacion As Date
Private mTopeAtenciones As Integer
Private mObservaciones As String
Private mCobrador As Empleado
Private mDireccion As Direccion
Private mPersonasACargo As AfiliadoManager
Private mParent As Afiliado
Private mParentezco As Parentezco
Private mHistoriaClinica As HistoriaClinica
Private mTipoAfiliado As eTipoAfiliado
Private mAtenciones As AtencionManager
Private mCuotas As CuotaManager
Private mImporte As Currency

Private eTipo As eTipoAfiliado

Private bPersonasACargoCargados As Boolean

'##ModelId=46A91EC1037A
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Afiliado", "Id Property Get", Err.Description)
End Property

'##ModelId=46A91EC1037B
'Public porq se puede asignar a mano...
Public Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Afiliado", "Id Property Let", Err.Description)
End Property

'##ModelId=46A91EC1037F
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Afiliado", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A91EC10380
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Afiliado", "Nombre Property Let", Err.Description)
End Property

'##ModelId=46A91EC10384
Public Property Get Apellido() As String
    On Error GoTo ApellidoErr

    '## Generated default body ...
    Apellido = mApellido

    Exit Property
ApellidoErr:
    Call ErrorLog("Afiliado", "Apellido Property Get", Err.Description)
End Property

'##ModelId=46A91EC10385
Public Property Let Apellido(ByVal vNewValue As String)
    On Error GoTo ApellidoErr

    '## Generated default body ...
    mApellido = vNewValue

    Exit Property
ApellidoErr:
    Call ErrorLog("Afiliado", "Apellido Property Let", Err.Description)
End Property

'##ModelId=46A91EC10389
Public Property Get TipoDoc() As TipoDoc
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set TipoDoc = mTipoDoc

    Exit Property
TipoDocErr:
    Call ErrorLog("Afiliado", "TipoDoc Property Get", Err.Description)
End Property

'##ModelId=46A91EC1038A
Public Property Set TipoDoc(ByVal vNewValue As TipoDoc)
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set mTipoDoc = vNewValue

    Exit Property
TipoDocErr:
    Call ErrorLog("Afiliado", "TipoDoc Property Let", Err.Description)
End Property

'##ModelId=46A91EC1038E
Public Property Get NroDoc() As Long
    On Error GoTo NroDocErr

    '## Generated default body ...
    NroDoc = mNroDoc

    Exit Property
NroDocErr:
    Call ErrorLog("Afiliado", "NroDoc Property Get", Err.Description)
End Property

'##ModelId=46A91EC1038F
Public Property Let NroDoc(ByVal vNewValue As Long)
    On Error GoTo NroDocErr

    '## Generated default body ...
    mNroDoc = vNewValue

    Exit Property
NroDocErr:
    Call ErrorLog("Afiliado", "NroDoc Property Let", Err.Description)
End Property

'##ModelId=46A91EC10393
Public Property Get FechaNacimiento() As Date
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    FechaNacimiento = mFechaNacimiento

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("Afiliado", "FechaNacimiento Property Get", Err.Description)
End Property

'##ModelId=46A91EC10394
Public Property Let FechaNacimiento(ByVal vNewValue As Date)
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    mFechaNacimiento = vNewValue

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("Afiliado", "FechaNacimiento Property Let", Err.Description)
End Property

'##ModelId=46A91EC10398
Public Property Get ObraSocial() As ObraSocial
    On Error GoTo ObraSocialErr

    '## Generated default body ...
    Set ObraSocial = mObraSocial

    Exit Property
ObraSocialErr:
    Call ErrorLog("Afiliado", "ObraSocial Property Get", Err.Description)
End Property

'##ModelId=46A91EC1039B
Public Property Set ObraSocial(ByVal vNewValue As ObraSocial)
    On Error GoTo ObraSocialErr

    '## Generated default body ...
    Set mObraSocial = vNewValue

    Exit Property
ObraSocialErr:
    Call ErrorLog("Afiliado", "ObraSocial Property Set", Err.Description)
End Property

'##ModelId=46A91EC1039D
Public Property Get Telefonos() As TelefonoManager
    On Error GoTo TelefonosErr
    
    If mTelefonos Is Nothing Then
        Set mTelefonos = TelefonosLocal.GetByOwner(Me.Id, eOTAfiliado)
    End If
    
    Set Telefonos = mTelefonos
    
    Exit Property
TelefonosErr:
    Call ErrorLog("Afiliado", "Telefonos Property Get", Err.Description)
End Property

'##ModelId=46A91EC103A0
Public Property Set Telefonos(ByVal vNewValue As TelefonoManager)
    On Error GoTo TelefonosErr

    Set mTelefonos = vNewValue
    If Not mTelefonos Is Nothing Then
        mTelefonos.OwnerId = Me.Id
        mTelefonos.OwnerType = eOTAfiliado
    End If
    
    Exit Property
TelefonosErr:
    Call ErrorLog("Afiliado", "Telefonos Property Set", Err.Description)
End Property

'##ModelId=46A91EC103A2
Public Property Get Vehiculo() As Vehiculo
    On Error GoTo VehiculoErr

    '## Generated default body ...
    Set Vehiculo = mVehiculo

    Exit Property
VehiculoErr:
    Call ErrorLog("Afiliado", "Vehiculo Property Get", Err.Description)
End Property

'##ModelId=46A91EC103A5
Public Property Set Vehiculo(ByVal vNewValue As Vehiculo)
    On Error GoTo VehiculoErr

    '## Generated default body ...
    Set mVehiculo = vNewValue

    Exit Property
VehiculoErr:
    Call ErrorLog("Afiliado", "Vehiculo Property Set", Err.Description)
End Property

'##ModelId=46A91EC103A7
Public Property Get Ocupacion() As Ocupacion
    On Error GoTo OcupacionErr

    '## Generated default body ...
    Set Ocupacion = mOcupacion

    Exit Property
OcupacionErr:
    Call ErrorLog("Afiliado", "Ocupacion Property Get", Err.Description)
End Property

'##ModelId=46A91EC103AA
Public Property Set Ocupacion(ByVal vNewValue As Ocupacion)
    On Error GoTo OcupacionErr

    '## Generated default body ...
    Set mOcupacion = vNewValue

    Exit Property
OcupacionErr:
    Call ErrorLog("Afiliado", "Ocupacion Property Set", Err.Description)
End Property

'##ModelId=46A91EC103AC
Public Property Get Sexo() As Integer
    On Error GoTo SexoErr

    '## Generated default body ...
    Sexo = mSexo

    Exit Property
SexoErr:
    Call ErrorLog("Afiliado", "Sexo Property Get", Err.Description)
End Property

'##ModelId=46A91EC103AD
Public Property Let Sexo(ByVal vNewValue As Integer)
    On Error GoTo SexoErr

    '## Generated default body ...
    mSexo = vNewValue

    Exit Property
SexoErr:
    Call ErrorLog("Afiliado", "Sexo Property Let", Err.Description)
End Property

'##ModelId=46A91EC103B1
Public Property Get EstadoCivil() As EstadoCivil
    On Error GoTo EstadoCivilErr

    '## Generated default body ...
    Set EstadoCivil = mEstadoCivil

    Exit Property
EstadoCivilErr:
    Call ErrorLog("Afiliado", "EstadoCivil Property Get", Err.Description)
End Property

'##ModelId=46A91EC103B2
Public Property Set EstadoCivil(ByVal vNewValue As EstadoCivil)
    On Error GoTo EstadoCivilErr

    '## Generated default body ...
    Set mEstadoCivil = vNewValue

    Exit Property
EstadoCivilErr:
    Call ErrorLog("Afiliado", "EstadoCivil Property Let", Err.Description)
End Property

'##ModelId=46A91EC103B6
Public Property Get FechaInscripcion() As Date
    On Error GoTo FechaInscripcionErr

    '## Generated default body ...
    FechaInscripcion = mFechaInscripcion

    Exit Property
FechaInscripcionErr:
    Call ErrorLog("Afiliado", "FechaInscripcion Property Get", Err.Description)
End Property

'##ModelId=46A91EC103B7
Public Property Let FechaInscripcion(ByVal vNewValue As Date)
    On Error GoTo FechaInscripcionErr

    '## Generated default body ...
    mFechaInscripcion = vNewValue

    Exit Property
FechaInscripcionErr:
    Call ErrorLog("Afiliado", "FechaInscripcion Property Let", Err.Description)
End Property

'##ModelId=46A91EC103BB
Public Property Get InicioPrestacion() As Date
    On Error GoTo InicioPrestacionErr

    '## Generated default body ...
    InicioPrestacion = mInicioPrestacion

    Exit Property
InicioPrestacionErr:
    Call ErrorLog("Afiliado", "InicioPrestacion Property Get", Err.Description)
End Property

'##ModelId=46A91EC103BC
Public Property Let InicioPrestacion(ByVal vNewValue As Date)
    On Error GoTo InicioPrestacionErr

    '## Generated default body ...
    mInicioPrestacion = vNewValue

    Exit Property
InicioPrestacionErr:
    Call ErrorLog("Afiliado", "InicioPrestacion Property Let", Err.Description)
End Property

'##ModelId=46A91EC103C0
Public Property Get TopeAtenciones() As Integer
    On Error GoTo TopeAtencionesErr

    '## Generated default body ...
    TopeAtenciones = mTopeAtenciones

    Exit Property
TopeAtencionesErr:
    Call ErrorLog("Afiliado", "TopeAtenciones Property Get", Err.Description)
End Property

'##ModelId=46A91EC103C1
Public Property Let TopeAtenciones(ByVal vNewValue As Integer)
    On Error GoTo TopeAtencionesErr

    '## Generated default body ...
    mTopeAtenciones = vNewValue

    Exit Property
TopeAtencionesErr:
    Call ErrorLog("Afiliado", "TopeAtenciones Property Let", Err.Description)
End Property

'##ModelId=46A91EC103CF
Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("Afiliado", "Observaciones Property Get", Err.Description)
End Property

'##ModelId=46A91EC103D0
Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("Afiliado", "Observaciones Property Let", Err.Description)
End Property

'##ModelId=46A91EC103D4
Public Property Get Atenciones() As AtencionManager
    On Error GoTo AtencionesErr

    'para q se mantenga actualizado lo pido cada vez q lo necesito
    Set mAtenciones = AtencionesLocal.GetByTipo(eAfiliadoPropio, Me.Id)
   
    Set Atenciones = mAtenciones

    Exit Function
AtencionesErr:
    Call ErrorLog("Afiliado", "Atenciones Property Get", Err.Description)
End Function

'##ModelId=46A91EC103D7
Public Property Set Atenciones(ByVal vNewValue As AtencionManager)
    On Error GoTo AtencionesErr

    '## Generated default body ...
    Set mAtenciones = vNewValue

    Exit Property
AtencionesErr:
    Call ErrorLog("Afiliado", "Atenciones Property Set", Err.Description)
End Property

'##ModelId=46A91EC103D9
Public Property Get Cuotas() As CuotaManager
    On Error GoTo PagosErr
    'para q se mantenga actualizado
    Set mCuotas = Nothing
    Set mCuotas = New CuotaManager
    mCuotas.CargarCuotasByAfiliado Me
    Set Cuotas = mCuotas

    Exit Property
PagosErr:
    Call ErrorLog("Afiliado", "Pagos Property Get", Err.Description)
End Property

'##ModelId=46A91EC103DC
Public Property Set Cuotas(ByVal vNewValue As CuotaManager)
    On Error GoTo PagosErr

    '## Generated default body ...
    Set mCuotas = vNewValue

    Exit Property
PagosErr:
    Call ErrorLog("Afiliado", "Pagos Property Set", Err.Description)
End Property

'##ModelId=46A91EC103DE
Public Property Get Cobrador() As Empleado
    On Error GoTo CobradorErr

    '## Generated default body ...
    Set Cobrador = mCobrador

    Exit Property
CobradorErr:
    Call ErrorLog("Afiliado", "Cobrador Property Get", Err.Description)
End Property

'##ModelId=46A91EC103E1
Public Property Set Cobrador(ByVal vNewValue As Empleado)
    On Error GoTo CobradorErr

    '## Generated default body ...
    Set mCobrador = vNewValue

    Exit Property
CobradorErr:
    Call ErrorLog("Afiliado", "Cobrador Property Set", Err.Description)
End Property

'##ModelId=46A91EC103E3
Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("Afiliado", "Direccion Property Get", Err.Description)
End Property

'##ModelId=46A91EC103E6
Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("Afiliado", "Direccion Property Set", Err.Description)
End Property

'##ModelId=46A91EC103E8
Public Property Get PersonasACargo() As AfiliadoManager
    On Error GoTo PersonasACargoErr
    

'    If Not bPersonasACargoCargados Then
'        Set mPersonasACargo = AfiliadosLocal.GetAfiliadosACargo(Me)
'        bPersonasACargoCargados = True
'    End If
    If mPersonasACargo Is Nothing Then
        Set mPersonasACargo = AfiliadosLocal.GetAfiliadosACargo(Me)
        bPersonasACargoCargados = True
    End If
    Set PersonasACargo = mPersonasACargo

    Exit Property
PersonasACargoErr:
    Call ErrorLog("Afiliado", "PersonasACargo Property Get", Err.Description)
End Property

'##ModelId=46A91EC103EB
Public Property Set PersonasACargo(ByVal vNewValue As AfiliadoManager)
    On Error GoTo PersonasACargoErr

    '## Generated default body ...
    Set mPersonasACargo = vNewValue

    Exit Property
PersonasACargoErr:
    Call ErrorLog("Afiliado", "PersonasACargo Property Set", Err.Description)
End Property

'##ModelId=46A91EC103ED
Public Property Get Parent() As Afiliado
    On Error GoTo ParentErr

    '## Generated default body ...
    Set Parent = mParent

    Exit Property
ParentErr:
    Call ErrorLog("Afiliado", "Parent Property Get", Err.Description)
End Property

'##ModelId=46A91EC103F0
Public Property Set Parent(ByVal vNewValue As Afiliado)
    On Error GoTo ParentErr

    '## Generated default body ...
    Set mParent = vNewValue

    Exit Property
ParentErr:
    Call ErrorLog("Afiliado", "Parent Property Set", Err.Description)
End Property

Public Property Get Parentezco() As Parentezco
    On Error GoTo ParentErr

    Set Parentezco = mParentezco

    Exit Property
ParentErr:
    Call ErrorLog("Afiliado", "Parentezco Property Get", Err.Description)
End Property

Public Property Set Parentezco(ByVal vNewValue As Parentezco)
    On Error GoTo ParentErr

    Set mParentezco = vNewValue

    Exit Property
ParentErr:
    Call ErrorLog("Afiliado", "Parentezco Property Set", Err.Description)
End Property

Public Property Get Importe() As Currency
    On Error GoTo ImporteErr

    Importe = mImporte

    Exit Property
ImporteErr:
    Call ErrorLog("Afiliado", "Importe Property Get", Err.Description)
End Property

Public Property Let Importe(ByVal vNewValue As Currency)
    On Error GoTo ImporteErr

    mImporte = vNewValue

    Exit Property
ImporteErr:
    Call ErrorLog("Afiliado", "Importe Property Let", Err.Description)
End Property

Public Property Get HistoriaClinica() As HistoriaClinica
    
    If Not mHistoriaClinica.Cargada Then mHistoriaClinica.Cargar Me
    
    Set HistoriaClinica = mHistoriaClinica
    
End Property

Friend Property Set HistoriaClinica(pValue As HistoriaClinica)
    Set mHistoriaClinica = pValue
    
End Property

Public Property Get TipoAfiliado() As eTipoAfiliado
    TipoAfiliado = eTipo
End Property

Friend Property Let TipoAfiliado(pTipo As eTipoAfiliado)
     eTipo = pTipo
End Property

Public Function Edad() As Integer
On Error Resume Next
Edad = Year(Date) - Year(mFechaNacimiento)
End Function

Public Property Get IdF() As String
    On Error GoTo idErr

      If Me.TipoAfiliado = eTitular Then
                IdF = Trim(Str(Me.Id / 1000))
            Else
                Dim idAux As Long
                idAux = Int(Me.Id / 1000)
                IdF = Trim(Str(Me.Id - (idAux * 1000)))
            End If

    Exit Property
idErr:
    Call ErrorLog("Afiliado", "Id Property Get", Err.Description)
End Property

Public Property Get IdCompleto() As String
    On Error GoTo idErr

     If Me.TipoAfiliado = eTitular Then
                IdCompleto = Me.IdF
            Else
                IdCompleto = Me.Parent.IdF + " - " + Me.IdF
            End If

    Exit Property
idErr:
    Call ErrorLog("Afiliado", "Id Property Get", Err.Description)
End Property

Public Function GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "parentezco"
            GetProperty = Parentezco.Nombre
        Case "aac"
            If Me.TipoAfiliado = eTitular Then
                GetProperty = Trim(Str(Me.PersonasACargo.Count))
            Else
                GetProperty = "-"
            End If
        Case "cantatenciones"
            GetProperty = Me.Atenciones.CantidadAtenciones
        Case "estadopagos"
            If Me.TipoAfiliado = eTitular Then
                Dim cant As Integer
                cant = Cuotas.GetCuotasByEstado(eImpaga).Count
                GetProperty = IIf(cant = 0, "Al dia", "Debe " + IIf(cant = 1, "1 mes", Trim(Str(cant) + " meses")))
            Else
                GetProperty = "-"
            End If
                            
    End Select
End Function

Private Sub Class_Initialize()
    Set mHistoriaClinica = New HistoriaClinica
    Set mDireccion = New Direccion
    'Set mPersonasACargo = New AfiliadoManager
End Sub

Public Property Get NombreCompleto() As String
    NombreCompleto = mApellido + ", " + mNombre
End Property

Friend Sub InsertTitular()

Dim trans As Transaccion
Set trans = BD.Transaccion
Dim params As New ParametrosManager

'habria q preguntar si la direccion es la misma q de los afiliados a cargo, entonces no la vuelvo a guardar
Direccion.Guardar trans
If Not Vehiculo Is Nothing Then Vehiculo.Guardar trans
'aca agrego los campos q se tienen q guardar independiente de si el afiliado es titular o no
LlenarParametros params

With params
    .Add "eliminado", "false" 'cuando lo insertamos no esta eliminado
    .Add "TopeAtenciones", TopeAtenciones
    .Add "idEmpleado", Cobrador.Id
    .Add "importe", Importe
    PersonasACargo.Guardar trans, Me
       
End With

trans.AddInsertCommand tAfiliado, params

HistoriaClinica.Save trans, Me

Telefonos.GuardarTelefonos Me.Id, eOTAfiliado, trans

trans.Execute
End Sub

Friend Sub InsertACargo(trans As Transaccion, pParent As Afiliado)

Dim params As New ParametrosManager

'habria q preguntar si la direccion es la misma q de los afiliados a cargo, entonces no la vuelvo a guardar
Direccion.Guardar trans

Set Me.Parent = pParent
LlenarParametros params
With params
    
    .Add "idAfiliadoTitular", Parent.Id
    .Add "idParentezco", Parentezco.Id
    .Add "eliminado", False
End With

trans.AddInsertCommand tAfiliado, params

HistoriaClinica.Save trans, Me

Telefonos.GuardarTelefonos Me.Id, eOTAfiliado, trans

End Sub

Public Sub GuardarModificaciones()

'terminar!!!!!

Dim trans As Transaccion
Set trans = BD.Transaccion
Dim params As New ParametrosManager

'habria q preguntar si la direccion es la misma q de los afiliados a cargo, entonces no la vuelvo a guardar
Direccion.GuardarModificaciones trans
If Not Vehiculo Is Nothing Then Vehiculo.GuardarModificaciones trans
'aca agrego los campos q se tienen q guardar independiente de si el afiliado es titular o no
LlenarParametros params

With params
    If Not Vehiculo Is Nothing Then .Add "idVehiculo", Vehiculo.Id
    .Add "TopeAtenciones", TopeAtenciones
    .Add "idEmpleado", Cobrador.Id
    .Add "importe", Importe
    'Vehiculo ver
     'no va porq las modificaciones las guardo cuando se las hago a cada af a cargo
    'PersonasACargo.GuardarModificaciones trans
       
End With

trans.AddUpdateCommand tAfiliado, params

HistoriaClinica.SaveChanges trans

Telefonos.SaveChanges trans

trans.Execute
End Sub

Public Sub GuardarModificacionesACargo()

'terminar!!!!!

Dim params As New ParametrosManager
Dim trans As Transaccion
Set trans = BD.Transaccion

Direccion.GuardarModificaciones trans

LlenarParametros params

With params
    'fijarse si se puede cambiar el afiliado titular
    .Add "idParentezco", Parentezco.Id
          
End With

trans.AddUpdateCommand tAfiliado, params

HistoriaClinica.SaveChanges trans

Telefonos.SaveChanges trans

trans.Execute
End Sub

Public Sub DarDeBaja()
    Dim c As New Parametro
    c.Nombre = "id"
    c.Valor = mId
    BD.BorradoLogico tAfiliado, c
End Sub

Public Sub Restaurar()
    Dim c As New Parametro
    c.Nombre = "id"
    c.Valor = mId
    BD.RestaurarElemento tAfiliado, c
End Sub

Private Sub LlenarParametros(pParametros As ParametrosManager)
With pParametros
    .Add "Apellido", Apellido
    'Atenciones no tiene todavia
    .Add "EstadoCivil", EstadoCivil.Id
    .Add "FechaInscripcion", FechaInscripcion
    .Add "FechaNacimiento", FechaNacimiento
    .Add "id", Id, True
    .Add "InicioPrestacion", InicioPrestacion
    .Add "Nombre", Nombre
    .Add "NroDoc", NroDoc
    If Not ObraSocial Is Nothing Then .Add "idObraSocial", ObraSocial.Id
    .Add "Observaciones", Observaciones
    .Add "idOcupacion", Ocupacion.Id
    'Pagos no tiene todavia
    .Add "Sexo", Sexo
    .Add "TipoDoc", TipoDoc.Id
    .Add "idDireccion", Direccion.Id
End With
End Sub
