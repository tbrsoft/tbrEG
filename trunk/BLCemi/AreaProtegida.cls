VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AreaProtegida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FF470213"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A8047002EE
Private mId As Long

'##ModelId=46A804C000EA
Private mNombreArea As String

'##ModelId=46A817F10213
Private mFechaInscripcion As Date

'##ModelId=46A817F4035B
Private mInicioPrestacion As Date

'##ModelId=46A817F700AB
Private mTopeAtenciones As Integer

'##ModelId=46A817FC034B
Private mCobrador As Empleado

'##ModelId=46A8180100AB
Private mCuotas As CuotaManager

'##ModelId=46A818030213
Private mAtenciones As AtencionManager

'##ModelId=46A8180500AB
Private mObservaciones As String

'##ModelId=46A8183000BB
Private mApellidoResp As String

'##ModelId=46A8183201C5
Private mTipoDocResp As TipoDoc

'##ModelId=46A81833032C
Private mNroDocResp As Long

'##ModelId=46A818670109
Private mNombreResp As String

'##ModelId=46A8199300CB
Private mDireccion As Direccion

Private mImporte As Currency

Private mTelefonos As TelefonoManager

Private mAfiliados As AfiliadoExternoManager

Private bAfiliadosCargados As Boolean

'##ModelId=46A9203401CF
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("AreaProtegida", "Id Property Get", err.Description)
End Property

'##ModelId=46A9203401D0
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("AreaProtegida", "Id Property Let", err.Description)
End Property

'##ModelId=46A9203401D6
Public Property Get NombreArea() As String
    On Error GoTo NombreAreaErr

    '## Generated default body ...
    NombreArea = mNombreArea

    Exit Property
NombreAreaErr:
    Call ErrorLog("AreaProtegida", "NombreArea Property Get", err.Description)
End Property

'##ModelId=46A9203401D7
Public Property Let NombreArea(ByVal vNewValue As String)
    On Error GoTo NombreAreaErr

    '## Generated default body ...
    mNombreArea = vNewValue

    Exit Property
NombreAreaErr:
    Call ErrorLog("AreaProtegida", "NombreArea Property Let", err.Description)
End Property

'##ModelId=46A9203401DB
Public Property Get FechaInscripcion() As Date
    On Error GoTo FechaInscripcionErr

    '## Generated default body ...
    FechaInscripcion = mFechaInscripcion

    Exit Property
FechaInscripcionErr:
    Call ErrorLog("AreaProtegida", "FechaInscripcion Property Get", err.Description)
End Property

'##ModelId=46A9203401DC
Public Property Let FechaInscripcion(ByVal vNewValue As Date)
    On Error GoTo FechaInscripcionErr

    '## Generated default body ...
    mFechaInscripcion = vNewValue

    Exit Property
FechaInscripcionErr:
    Call ErrorLog("AreaProtegida", "FechaInscripcion Property Let", err.Description)
End Property

'##ModelId=46A9203401E0
Public Property Get InicioPrestacion() As Date
    On Error GoTo InicioPrestacionErr

    '## Generated default body ...
    InicioPrestacion = mInicioPrestacion

    Exit Property
InicioPrestacionErr:
    Call ErrorLog("AreaProtegida", "InicioPrestacion Property Get", err.Description)
End Property

'##ModelId=46A9203401E1
Public Property Let InicioPrestacion(ByVal vNewValue As Date)
    On Error GoTo InicioPrestacionErr

    '## Generated default body ...
    mInicioPrestacion = vNewValue

    Exit Property
InicioPrestacionErr:
    Call ErrorLog("AreaProtegida", "InicioPrestacion Property Let", err.Description)
End Property

'##ModelId=46A9203401E5
Public Property Get TopeAtenciones() As Integer
    On Error GoTo TopeAtencionesErr

    '## Generated default body ...
    TopeAtenciones = mTopeAtenciones

    Exit Property
TopeAtencionesErr:
    Call ErrorLog("AreaProtegida", "TopeAtenciones Property Get", err.Description)
End Property

'##ModelId=46A9203401E6
Public Property Let TopeAtenciones(ByVal vNewValue As Integer)
    On Error GoTo TopeAtencionesErr

    '## Generated default body ...
    mTopeAtenciones = vNewValue

    Exit Property
TopeAtencionesErr:
    Call ErrorLog("AreaProtegida", "TopeAtenciones Property Let", err.Description)
End Property

'##ModelId=46A9203401EA
Public Property Get Cobrador() As Empleado
    On Error GoTo CobradorErr

    '## Generated default body ...
    Set Cobrador = mCobrador

    Exit Property
CobradorErr:
    Call ErrorLog("AreaProtegida", "Cobrador Property Get", err.Description)
End Property

'##ModelId=46A9203401ED
Public Property Set Cobrador(ByVal vNewValue As Empleado)
    On Error GoTo CobradorErr

    '## Generated default body ...
    Set mCobrador = vNewValue

    Exit Property
CobradorErr:
    Call ErrorLog("AreaProtegida", "Cobrador Property Set", err.Description)
End Property

'##ModelId=46A9203401EF
Public Property Get Cuotas() As CuotaManager
    On Error GoTo PagosErr

    Set mCuotas = Nothing
    Set mCuotas = New CuotaManager
    mCuotas.CargarCuotasByAreaProtegida Me
    Set Cuotas = mCuotas

    Exit Property
PagosErr:
    Call ErrorLog("AreaProtegida", "Pagos Property Get", err.Description)
End Property

'##ModelId=46A9203401F2
Public Property Set Cuotas(ByVal vNewValue As CuotaManager)
    On Error GoTo PagosErr

    '## Generated default body ...
    Set mCuotas = vNewValue

    Exit Property
PagosErr:
    Call ErrorLog("AreaProtegida", "Pagos Property Set", err.Description)
End Property

'##ModelId=46A9203401F4
Public Property Get Atenciones() As AtencionManager
    On Error GoTo AtencionesErr

    'para q se mantenga actualizado lo pido cada vez q lo necesito
    Set mAtenciones = AtencionesLocal.GetByTipo(eAreaProtegida, Me.Id)

    Set Atenciones = mAtenciones

    Exit Property
AtencionesErr:
    Call ErrorLog("AreaProtegida", "Atenciones Property Get", err.Description)
End Property

'##ModelId=46A9203401F7
Public Property Set Atenciones(ByVal vNewValue As AtencionManager)
    On Error GoTo AtencionesErr

    '## Generated default body ...
    Set mAtenciones = vNewValue

    Exit Property
AtencionesErr:
    Call ErrorLog("AreaProtegida", "Atenciones Property Set", err.Description)
End Property

'##ModelId=46A9203401F9
Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("AreaProtegida", "Observaciones Property Get", err.Description)
End Property

'##ModelId=46A9203401FA
Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("AreaProtegida", "Observaciones Property Let", err.Description)
End Property

'##ModelId=46A9203401FE
Public Property Get ApellidoResp() As String
    On Error GoTo ApellidoRespErr

    '## Generated default body ...
    ApellidoResp = mApellidoResp

    Exit Property
ApellidoRespErr:
    Call ErrorLog("AreaProtegida", "ApellidoResp Property Get", err.Description)
End Property

'##ModelId=46A9203401FF
Public Property Let ApellidoResp(ByVal vNewValue As String)
    On Error GoTo ApellidoRespErr

    '## Generated default body ...
    mApellidoResp = vNewValue

    Exit Property
ApellidoRespErr:
    Call ErrorLog("AreaProtegida", "ApellidoResp Property Let", err.Description)
End Property

'##ModelId=46A920340203
Public Property Get TipoDocResp() As TipoDoc
    On Error GoTo TipoDocRespErr

    '## Generated default body ...
    Set TipoDocResp = mTipoDocResp

    Exit Property
TipoDocRespErr:
    Call ErrorLog("AreaProtegida", "TipoDocResp Property Get", err.Description)
End Property

'##ModelId=46A920340204
Public Property Set TipoDocResp(ByVal vNewValue As TipoDoc)
    On Error GoTo TipoDocRespErr

    '## Generated default body ...
    Set mTipoDocResp = vNewValue

    Exit Property
TipoDocRespErr:
    Call ErrorLog("AreaProtegida", "TipoDocResp Property Let", err.Description)
End Property

'##ModelId=46A920340208
Public Property Get NroDocResp() As Long
    On Error GoTo NroDocRespErr

    '## Generated default body ...
    NroDocResp = mNroDocResp

    Exit Property
NroDocRespErr:
    Call ErrorLog("AreaProtegida", "NroDocResp Property Get", err.Description)
End Property

'##ModelId=46A920340209
Public Property Let NroDocResp(ByVal vNewValue As Long)
    On Error GoTo NroDocRespErr

    '## Generated default body ...
    mNroDocResp = vNewValue

    Exit Property
NroDocRespErr:
    Call ErrorLog("AreaProtegida", "NroDocResp Property Let", err.Description)
End Property

'##ModelId=46A92034020D
Public Property Get NombreResp() As String
    On Error GoTo NombreRespErr

    '## Generated default body ...
    NombreResp = mNombreResp

    Exit Property
NombreRespErr:
    Call ErrorLog("AreaProtegida", "NombreResp Property Get", err.Description)
End Property

'##ModelId=46A92034020E
Public Property Let NombreResp(ByVal vNewValue As String)
    On Error GoTo NombreRespErr

    '## Generated default body ...
    mNombreResp = vNewValue

    Exit Property
NombreRespErr:
    Call ErrorLog("AreaProtegida", "NombreResp Property Let", err.Description)
End Property

'##ModelId=46A920340212
Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("AreaProtegida", "Direccion Property Get", err.Description)
End Property

'##ModelId=46A920340215
Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("AreaProtegida", "Direccion Property Set", err.Description)
End Property

Public Property Get Importe() As Currency
    On Error GoTo ImporteErr

    Importe = mImporte

    Exit Property
ImporteErr:
    Call ErrorLog("AreaProtegida", "Importe Property Get", err.Description)
End Property

Public Property Let Importe(ByVal vNewValue As Currency)
    On Error GoTo ImporteErr

    mImporte = vNewValue

    Exit Property
ImporteErr:
    Call ErrorLog("AreaProtegida", "Importe Property Let", err.Description)
End Property

Public Property Get Telefonos() As TelefonoManager
    On Error GoTo TelefonosErr
    
    If mTelefonos Is Nothing Then
        Set mTelefonos = TelefonosLocal.GetByOwner(mId, eOTAreaProtegida)
    End If
    
    Set Telefonos = mTelefonos

    Exit Property
TelefonosErr:
    Call ErrorLog("AreaProtegida", "Telefonos Property Get", err.Description)
End Property

Public Property Set Telefonos(ByVal vNewValue As TelefonoManager)
    On Error GoTo TelefonosErr
        
    Set mTelefonos = vNewValue
    If Not mTelefonos Is Nothing Then
        mTelefonos.OwnerId = mId
        mTelefonos.OwnerType = eOTAreaProtegida
    End If
    
    Exit Property
TelefonosErr:
    Call ErrorLog("AreaProtegida", "Telefonos Property Set", err.Description)
End Property

Public Property Get Afiliados() As AfiliadoExternoManager
    On Error GoTo AfiliadosErr
    
    If Not bAfiliadosCargados Then
        mAfiliados.CargarAfiliadoExternos Me
        bAfiliadosCargados = True
    End If
    Set Afiliados = mAfiliados

    Exit Property
AfiliadosErr:
    Call ErrorLog("AreaProtegida", "Afiliados Property Get", err.Description)
End Property

Public Property Set Afiliados(ByVal vNewValue As AfiliadoExternoManager)
    On Error GoTo AfiliadosErr

    Set mAfiliados = vNewValue

    Exit Property
AfiliadosErr:
    Call ErrorLog("AreaProtegida", "Afiliados Property Set", err.Description)
End Property

Public Property Get NombreCompleto() As String
    NombreCompleto = mApellidoResp + ", " + mNombreResp
End Property

Public Function GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "cantatenciones"
            GetProperty = Me.Atenciones.CantidadAtenciones
        Case "estadopagos"
            Dim cant As Integer
            cant = Cuotas.GetCuotasByEstado(eImpaga).Count
            GetProperty = IIf(cant = 0, "Al dia", "Debe " + IIf(cant = 1, "1 mes", Trim(Str(cant) + " meses")))
    End Select
End Function


Friend Sub Guardar()
    Dim params As New ParametrosManager
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Direccion.Guardar trans
    
    Id = BD.getIdDisponible(tAreaProtegida)
    With params
        .Add "id", Id
        .Add "ApellidoResp", ApellidoResp
        .Add "idEmpleado", Cobrador.Id
        .Add "idDireccion", Direccion.Id
        .Add "FechaInscripcion", FechaInscripcion
        .Add "InicioPrestacion", InicioPrestacion
        .Add "NombreArea", NombreArea
        .Add "NombreResp", NombreResp
        .Add "NroDocResp", NroDocResp
        .Add "Observaciones", Observaciones
        .Add "TipoDocResp", TipoDocResp.Id
        .Add "TopeAtenciones", TopeAtenciones
        .Add "Importe", Importe
    End With
        
    trans.AddInsertCommand tAreaProtegida, params
    Telefonos.GuardarTelefonos Me.Id, eOTAreaProtegida, trans
    trans.Execute
End Sub

Public Sub GuardarModificaciones()
Dim params As New ParametrosManager
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Direccion.GuardarModificaciones trans
        
    With params
        .Add "id", Id, True
        .Add "ApellidoResp", ApellidoResp
        .Add "idEmpleado", Cobrador.Id
        '.Add "idDireccion", Direccion.id
        .Add "FechaInscripcion", FechaInscripcion
        .Add "InicioPrestacion", InicioPrestacion
        .Add "NombreArea", NombreArea
        .Add "NombreResp", NombreResp
        .Add "NroDocResp", NroDocResp
        .Add "Observaciones", Observaciones
        .Add "TipoDocResp", TipoDocResp.Id
        .Add "TopeAtenciones", TopeAtenciones
        .Add "Importe", Importe
    End With
        
    trans.AddUpdateCommand tAreaProtegida, params
    Telefonos.SaveChanges trans
    trans.Execute
End Sub

Private Sub Class_Initialize()
    Set mDireccion = New Direccion
    Set mAfiliados = New AfiliadoExternoManager
End Sub
