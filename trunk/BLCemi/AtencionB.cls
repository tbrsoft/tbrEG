VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mId As Long
Private mNroIncidente As String
Private mNroIncidenteInterno As String
Private mDespachador As Empleado
Private mFecha As Date
'Las horas las manejo string porq parece facil manejarse con un datetimepicker.
Private mHoraLlamada As String
'hora de salida del movil de preinspeccion
Private mSalidaPreInspeccion As String
'hora de llegada del movil de preinspeccion
Private mLlegadaPreInspeccion As String
'hora de salida de la dotacion
Private mSalidaDotacion As String
'Hora de llegada al lugar de la emergencia
Private mQTH As String
'Hora de liberacion del movil
Private mVL As String

Private mEquipos As EquipoManager
Private mDireccion As Direccion
Private mCronica As String 'reseña
Private mObservaciones As String
Private mSintoma As Sintoma
Private mEstado As eEstadoAtencion

Private mInvolucrados As InvolucradoManager
Private mVehiculos As VehiculoManager

Dim mDescripcionPerjuiciosCampo As String
Dim mMaterialesCombustibles As String
Dim mDescripcionPerjuiciosVivienda As String
Dim mAccesoPor As String
Dim mReferencias As String
Dim mAseguradora As String
Dim mDescripcionMaterial As String
Dim mEquiposEspeciales As String
Dim mPoliza As String
Dim mAmbientesAfectadosVivienda As Integer
Dim mHectareasAfectadasCampo As Integer
Dim mInstalacionElectrica As InstElectrica
Dim mInstalacionGas As InstGas
Dim mColaboracionBomberos As ColaboracionManager

Dim mPoliciaACargo As String
Dim mPoliciaCantidad As Integer
Dim mPoliciaMovil As String
Dim mSEMedico As String
Dim mSECentroAsistencial As String
Dim mSEMedicoMP As String
Dim mSEAmbulancias As Boolean

'falta
'mSeguro
'mMaterialLuchaFuego as boolean
Public Property Get SEAmbulancias() As Boolean
    SEAmbulancias = mSEAmbulancias
End Property

Public Property Let SEAmbulancias(ByVal vNewValue As Boolean)
    mSEAmbulancias = vNewValue
End Property

Public Property Get SEMedicoMP() As String
    SEMedicoMP = mSEMedicoMP
End Property

Public Property Let SEMedicoMP(ByVal vNewValue As String)
    mSEMedicoMP = vNewValue
End Property

Public Property Get SECentroAsistencial() As String
    SECentroAsistencial = mSECentroAsistencial
End Property

Public Property Let SECentroAsistencial(ByVal vNewValue As String)
    mSECentroAsistencial = vNewValue
End Property

Public Property Get SEMedico() As String
    SEMedico = mSEMedico
End Property

Public Property Let SEMedico(ByVal vNewValue As String)
    mSEMedico = vNewValue
End Property

Public Property Get PoliciaACargo() As String
    PoliciaACargo = mPoliciaACargo
End Property

Public Property Let PoliciaACargo(ByVal vNewValue As String)
    mPoliciaACargo = vNewValue
End Property

Public Property Get PoliciaCantidad() As Integer
    PoliciaCantidad = mPoliciaCantidad
End Property

Public Property Let PoliciaCantidad(ByVal vNewValue As Integer)
    mPoliciaCantidad = vNewValue
End Property

Public Property Get PoliciaMovil() As String
    PoliciaMovil = mPoliciaMovil
End Property

Public Property Let PoliciaMovil(ByVal vNewValue As String)
    mPoliciaMovil = vNewValue
End Property

Public Property Get ColaboracionBomberos() As ColaboracionManager
    If mColaboracionBomberos Is Nothing Then
        Set mColaboracionBomberos = New ColaboracionManager
        mColaboracionBomberos.CargarPorSiniestro Me.Id
    End If
    
    Set ColaboracionBomberos = mColaboracionBomberos
End Property

Public Property Set ColaboracionBomberos(vNewValue As ColaboracionManager)
    Set mColaboracionBomberos = vNewValue
End Property

Public Property Get InstalacionGas() As InstGas
    Set InstalacionGas = mInstalacionGas
End Property

Public Property Set InstalacionGas(ByVal vNewValue As InstGas)
    Set mInstalacionGas = vNewValue
End Property

Public Property Get InstalacionElectrica() As InstElectrica
    Set InstalacionElectrica = mInstalacionElectrica
End Property

Public Property Set InstalacionElectrica(ByVal vNewValue As InstElectrica)
    Set mInstalacionElectrica = vNewValue
End Property

Public Property Get HectareasAfectadasCampo() As Integer
    HectareasAfectadasCampo = mHectareasAfectadasCampo
End Property

Public Property Let HectareasAfectadasCampo(ByVal vNewValue As Integer)
    mHectareasAfectadasCampo = vNewValue
End Property

Public Property Get AmbientesAfectadosVivienda() As Integer
    AmbientesAfectadosVivienda = mAmbientesAfectadosVivienda
End Property

Public Property Let AmbientesAfectadosVivienda(ByVal vNewValue As Integer)
    mAmbientesAfectadosVivienda = vNewValue
End Property

Public Property Get Poliza() As String
    Poliza = mPoliza
End Property

Public Property Let Poliza(ByVal vNewValue As String)
    mPoliza = vNewValue
End Property

Public Property Get EquiposEspeciales() As String
    EquiposEspeciales = mEquiposEspeciales
End Property

Public Property Let EquiposEspeciales(ByVal vNewValue As String)
    mEquiposEspeciales = vNewValue
End Property

Public Property Get DescripcionMaterial() As String
    DescripcionMaterial = mDescripcionMaterial
End Property

Public Property Let DescripcionMaterial(ByVal vNewValue As String)
    mDescripcionMaterial = vNewValue
End Property

Public Property Get Aseguradora() As String
    Aseguradora = mAseguradora
End Property

Public Property Let Aseguradora(ByVal vNewValue As String)
    mAseguradora = vNewValue
End Property

Public Property Get MaterialesCombustibles() As String
    MaterialesCombustibles = mMaterialesCombustibles
End Property

Public Property Let MaterialesCombustibles(ByVal vNewValue As String)
    mMaterialesCombustibles = vNewValue
End Property

Public Property Get DescripcionPerjuiciosCampo() As String
    DescripcionPerjuiciosCampo = mDescripcionPerjuiciosCampo
End Property

Public Property Let DescripcionPerjuiciosCampo(ByVal vNewValue As String)
    mDescripcionPerjuiciosCampo = vNewValue
End Property

Public Property Get DescripcionPerjuiciosVivienda() As String
    DescripcionPerjuiciosVivienda = mDescripcionPerjuiciosVivienda
End Property

Public Property Let DescripcionPerjuiciosVivienda(ByVal vNewValue As String)
    mDescripcionPerjuiciosVivienda = vNewValue
End Property

Public Property Get AccesoPor() As String
    AccesoPor = mAccesoPor
End Property

Public Property Let AccesoPor(ByVal vNewValue As String)
   mAccesoPor = vNewValue
End Property

Public Property Get Referencias() As String
    Referencias = mReferencias
End Property

Public Property Let Referencias(ByVal vNewValue As String)
    mReferencias = vNewValue
End Property

Public Property Set Vehiculos(ByVal vNewValue As VehiculoManager)
    On Error GoTo DespachadorErr
    
    Set mVehiculos = vNewValue

    Exit Property
DespachadorErr:
    Call ErrorLog("AtencionB", "Vehiculos Property Set", Err.Description)
End Property

Public Property Get Vehiculos() As VehiculoManager
    On Error GoTo AfiliadoErr

    If mVehiculos Is Nothing Then
        Set mVehiculos = New VehiculoManager
        mVehiculos.CargarPorSiniestro Me.Id
    End If
    Set Vehiculos = mVehiculos

    Exit Property
AfiliadoErr:
    Call ErrorLog("AtencionB", "Vehiculos Property Get", Err.Description)
End Property

'##ModelId=46A92069031F
Public Property Set Involucrados(ByVal vNewValue As InvolucradoManager)
    On Error GoTo DespachadorErr
    
    Set mInvolucrados = vNewValue

    Exit Property
DespachadorErr:
    Call ErrorLog("AtencionB", "Involucrados Property Set", Err.Description)
End Property

Public Property Get Involucrados() As InvolucradoManager
    On Error GoTo AfiliadoErr

    If mInvolucrados Is Nothing Then
        Set mInvolucrados = New InvolucradoManager
        mInvolucrados.CargarInvolucrados Me
    End If
    
    Set Involucrados = mInvolucrados

    Exit Property
AfiliadoErr:
    Call ErrorLog("AtencionB", "Involucrados Property Get", Err.Description)
End Property

'##ModelId=46A920690312
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("AtencionB", "Id Property Get", Err.Description)
End Property

'##ModelId=46A920690313
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("AtencionB", "Id Property Let", Err.Description)
End Property

'##ModelId=46A920690317
Public Property Get NroIncidente() As String
    On Error GoTo NroIncidenteErr

    '## Generated default body ...
    NroIncidente = mNroIncidente

    Exit Property
NroIncidenteErr:
    Call ErrorLog("AtencionB", "NroIncidente Property Get", Err.Description)
End Property

'##ModelId=46A920690318
Public Property Let NroIncidente(ByVal vNewValue As String)
    On Error GoTo NroIncidenteErr

    '## Generated default body ...
    mNroIncidente = vNewValue

    Exit Property
NroIncidenteErr:
    Call ErrorLog("AtencionB", "NroIncidente Property Let", Err.Description)
End Property

Public Property Get nroIncidenteInterno() As String
    On Error GoTo NroIncidenteInternoErr

    '## Generated default body ...
    nroIncidenteInterno = mNroIncidenteInterno

    Exit Property
NroIncidenteInternoErr:
    Call ErrorLog("AtencionB", "NroIncidenteInterno Property Get", Err.Description)
End Property

'##ModelId=46A920690318
Public Property Let nroIncidenteInterno(ByVal vNewValue As String)
    On Error GoTo NroIncidenteInternoErr

    '## Generated default body ...
    mNroIncidenteInterno = vNewValue

    Exit Property
NroIncidenteInternoErr:
    Call ErrorLog("AtencionB", "NroIncidenteInterno Property Let", Err.Description)
End Property
'##ModelId=46A92069031C
Public Property Get Despachador() As Empleado
    On Error GoTo DespachadorErr

    '## Generated default body ...
    Set Despachador = mDespachador

    Exit Property
DespachadorErr:
    Call ErrorLog("AtencionB", "Despachador Property Get", Err.Description)
End Property

'##ModelId=46A92069031F
Public Property Set Despachador(ByVal vNewValue As Empleado)
    On Error GoTo DespachadorErr

    '## Generated default body ...
    Set mDespachador = vNewValue

    Exit Property
DespachadorErr:
    Call ErrorLog("AtencionB", "Despachador Property Set", Err.Description)
End Property

'##ModelId=46A92069032B
Public Property Get Fecha() As Date
    On Error GoTo FechaErr

    '## Generated default body ...
    Fecha = mFecha

    Exit Property
FechaErr:
    Call ErrorLog("AtencionB", "Fecha Property Get", Err.Description)
End Property

'##ModelId=46A92069032C
Public Property Let Fecha(ByVal vNewValue As Date)
    On Error GoTo FechaErr

    '## Generated default body ...
    mFecha = vNewValue

    Exit Property
FechaErr:
    Call ErrorLog("AtencionB", "Fecha Property Let", Err.Description)
End Property

'##ModelId=46A920690330
Public Property Get HoraLlamada() As String
    On Error GoTo HoraLlamadaErr

    '## Generated default body ...
    HoraLlamada = mHoraLlamada

    Exit Property
HoraLlamadaErr:
    Call ErrorLog("AtencionB", "HoraLlamada Property Get", Err.Description)
End Property

'##ModelId=46A920690331
Public Property Let HoraLlamada(ByVal vNewValue As String)
    On Error GoTo HoraLlamadaErr

    '## Generated default body ...
    mHoraLlamada = vNewValue

    Exit Property
HoraLlamadaErr:
    Call ErrorLog("AtencionB", "HoraLlamada Property Let", Err.Description)
End Property

Public Property Get LlegadaPreInspeccion() As String
    On Error GoTo QTHErr

    LlegadaPreInspeccion = mLlegadaPreInspeccion

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "LlegadaPreInspeccion Property Get", Err.Description)
End Property


Public Property Let LlegadaPreInspeccion(ByVal vNewValue As String)
    On Error GoTo QTHErr

    mLlegadaPreInspeccion = vNewValue

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "LlegadaPreInspeccion Property Let", Err.Description)
End Property

Public Property Get SalidaDotacion() As String
    On Error GoTo QTHErr

    SalidaDotacion = mSalidaDotacion

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "SalidaDotacion Property Get", Err.Description)
End Property


Public Property Let SalidaDotacion(ByVal vNewValue As String)
    On Error GoTo QTHErr

    mSalidaDotacion = vNewValue

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "SalidaDotacion Property Let", Err.Description)
End Property

Public Property Get SalidaPreInspeccion() As String
    On Error GoTo QTHErr

    SalidaPreInspeccion = mSalidaPreInspeccion

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "SalidaPreInspeccion Property Get", Err.Description)
End Property


Public Property Let SalidaPreInspeccion(ByVal vNewValue As String)
    On Error GoTo QTHErr

    mSalidaPreInspeccion = vNewValue

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "SalidaPreInspeccion Property Let", Err.Description)
End Property

'##ModelId=46A920690335
Public Property Get QTH() As String
    On Error GoTo QTHErr

    '## Generated default body ...
    QTH = mQTH

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "QTH Property Get", Err.Description)
End Property

'##ModelId=46A920690336
Public Property Let QTH(ByVal vNewValue As String)
    On Error GoTo QTHErr

    '## Generated default body ...
    mQTH = vNewValue

    Exit Property
QTHErr:
    Call ErrorLog("AtencionB", "QTH Property Let", Err.Description)
End Property

'##ModelId=46A92069033A
Public Property Get VL() As String
    On Error GoTo VLErr

    '## Generated default body ...
    VL = mVL

    Exit Property
VLErr:
    Call ErrorLog("AtencionB", "VL Property Get", Err.Description)
End Property

'##ModelId=46A92069033B
Public Property Let VL(ByVal vNewValue As String)
    On Error GoTo VLErr

    '## Generated default body ...
    mVL = vNewValue

    Exit Property
VLErr:
    Call ErrorLog("AtencionB", "VL Property Let", Err.Description)
End Property

Public Property Get Equipos() As EquipoManager
    On Error GoTo EquiposErr

    If mEquipos Is Nothing Then
        Set mEquipos = New EquipoManager
        mEquipos.CargarPorAtencion Me.Id
    End If
    Set Equipos = mEquipos

    Exit Property
EquiposErr:
    Call ErrorLog("AtencionB", "Equipos Property Get", Err.Description)
End Property

Public Property Set Equipos(ByVal vNewValue As EquipoManager)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mEquipos = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("AtencionB", "Equipos Property Set", Err.Description)
End Property

'##ModelId=46A920690349
Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("AtencionB", "Direccion Property Get", Err.Description)
End Property

'##ModelId=46A92069034C
Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("AtencionB", "Direccion Property Set", Err.Description)
End Property

Public Property Get Cronica() As String
    On Error GoTo NroIncidenteErr
   
    Cronica = mCronica
    
    Exit Property
NroIncidenteErr:
    Call ErrorLog("AtencionB", "Cronica Property Get", Err.Description)
End Property

Public Property Let Cronica(ByVal vNewValue As String)
    On Error GoTo NroIncidenteErr

      mCronica = vNewValue

    Exit Property
NroIncidenteErr:
    Call ErrorLog("AtencionB", "Cronica Property Let", Err.Description)
End Property

Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("AtencionB", "Observaciones Property Get", Err.Description)
End Property

'##ModelId=46A920690318
Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("AtencionB", "Observaciones Property Let", Err.Description)
End Property

Public Property Get Sintoma() As Sintoma
    Set Sintoma = mSintoma
End Property

Public Property Set Sintoma(pValue As Sintoma)
    Set mSintoma = pValue
End Property

Public Property Get Estado() As eEstadoAtencion
Estado = mEstado
End Property

Public Property Let Estado(pValue As eEstadoAtencion)
    mEstado = pValue
End Property

Public Property Get HasEquipo() As Boolean
    HasEquipo = mEquipos.Count <> 0
End Property

Public Function GetVencimiento() As String
    Dim aux
    If mSintoma.Parent.Vencimiento <> 0 Then
        aux = Round(mSintoma.Parent.Vencimiento - Minute((Time - TimeValue(mHoraLlamada))))
        If aux < 0 Then
            GetVencimiento = "00"
        Else
            GetVencimiento = IIf(aux < 10, "0" + Trim(Str(aux)), Trim(Str(aux)))
        End If
    Else
        GetVencimiento = "-"
    End If
End Function

Public Property Get GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "codigo"
            GetProperty = mSintoma.Parent.Nombre
        Case "sintoma"
            GetProperty = mSintoma.Nombre
        Case "gas"
            If Not mInstalacionGas Is Nothing Then
                GetProperty = mInstalacionGas.Nombre
            Else
                GetProperty = "-Sin Asignar-"
            End If
        Case "electricidad"
            If Not mInstalacionElectrica Is Nothing Then
                GetProperty = mInstalacionElectrica.Nombre
            Else
                GetProperty = "-Sin Asignar-"
            End If
        Case "despachador"
            GetProperty = mDespachador.NombreCompleto
        Case "movil"
            GetProperty = IIf(Equipos.Count = 0, "<Sin Asignar>", Equipos.GetNombreMoviles)
        Case "pgdireccion"
            GetProperty = mDireccion.GetBarrioCiudadProvincia + " - Calle: " + mDireccion.Calle + " " + mDireccion.Nro
            
        'lo agrego aca para no romper la compatibilidad
        Case "transcurrido"
            Dim h As String
            Dim m As String
           ' h = Trim(Str(DateDiff("h", TimeValue(mHoraLlamada), Time)))
            m = Trim(Str(DateDiff("n", TimeValue(mHoraLlamada), Time)))
            h = Int(Val(m) / 60)
            m = m - h * 60
            GetProperty = IIf(h < 10, "0" + h, h) + ":" + IIf(m < 10, "0" + m, m)
    End Select
End Property

Public Sub Guardar()
Dim params As New ParametrosManager
Dim trans As Transaccion
Set trans = BD.Transaccion

mId = BD.getIdDisponible(tAtencion)

params.Add "Id", mId
params.Add "idDespachador", mDespachador.Id
params.Add "Fecha", mFecha
params.Add "HoraLlamada", mHoraLlamada

LlenarParametros params

params.Add "idSintoma", mSintoma.Id

Direccion.Guardar trans
params.Add "idDireccion", mDireccion.Id

trans.AddInsertCommand tAtencionB, params
Me.Equipos.Guardar trans, Me.Id

Me.Involucrados.Save trans, Me

Me.ColaboracionBomberos.Guardar Me.Id

Me.Vehiculos.GuardarPorSiniestro trans, Me.Id

trans.Execute
'para q avise a la red q hubo cambios
AtencionesLocal.AtencionAgregadaoModificada
End Sub

Public Sub GuardarModificaciones(pEmpleadoActual As Empleado)
Dim segSiniestro As New SeguimientoSiniestro

segSiniestro.ObtenerOriginal mId

'analizar bien
Dim params As New ParametrosManager
Dim trans As Transaccion
Set trans = BD.Transaccion

params.Add "Id", mId, True

'params.Add "idDespachador", mDespachador.id

'la fecha y hora no deberian ser modificabless
'params.Add "Fecha", mFecha
'params.Add "HoraLlamada", mHoraLlamada

LlenarParametros params

params.Add "idSintoma", mSintoma.Id

Direccion.GuardarModificaciones trans

trans.AddUpdateCommand tAtencionB, params
Equipos.GuardarCambios trans, Me.Id

Me.Involucrados.Save trans, Me

Me.ColaboracionBomberos.Guardar Me.Id

Me.Vehiculos.GuardarPorSiniestro trans, Me.Id

trans.Execute
'muy probablemente sea mejor usar la misma transaccion, ver si se puede
segSiniestro.RegistrarCambios pEmpleadoActual

'para q avise a la red q hubo cambios
AtencionesBLocal.AtencionBAgregadaoModificada

End Sub

Private Sub LlenarParametros(pParams As ParametrosManager)
    pParams.Add "NroIncidente", mNroIncidente
    pParams.Add "NroIncidenteInterno", mNroIncidenteInterno
    pParams.Add "SalidaPreInspeccion", mSalidaPreInspeccion
    pParams.Add "LlegadaPreinspeccion", mLlegadaPreInspeccion
    pParams.Add "SalidaDotacion", mSalidaDotacion
    pParams.Add "QTH", mQTH
    pParams.Add "VL", mVL
    pParams.Add "Cronica", mCronica
    pParams.Add "Observaciones", mObservaciones
    pParams.Add "DescripcionPerjuiciosCampo", mDescripcionPerjuiciosCampo
    pParams.Add "MaterialesCombustibles", mMaterialesCombustibles
    pParams.Add "DescripcionPerjuiciosVivienda", mDescripcionPerjuiciosVivienda
    pParams.Add "AccesoPor", mAccesoPor
    pParams.Add "Referencias", mReferencias
    pParams.Add "Aseguradora", mAseguradora
    pParams.Add "DescripcionMaterial", mDescripcionMaterial
    pParams.Add "EquiposEspeciales", mEquiposEspeciales
    pParams.Add "Poliza", mPoliza
    pParams.Add "AmbientesAfectadosVivienda", mAmbientesAfectadosVivienda
    pParams.Add "HectareasAfectadasCampo", mHectareasAfectadasCampo
    
    pParams.Add "PoliciaACargo", mPoliciaACargo
    pParams.Add "PoliciaCantidad", mPoliciaCantidad
    pParams.Add "PoliciaMovil", mPoliciaMovil
    pParams.Add "SEMedico", mSEMedico
    pParams.Add "SECentroAsistencial", mSECentroAsistencial
    pParams.Add "SEMedicoMP", mSEMedicoMP
    pParams.Add "SEAmbulancias", IIf(mSEAmbulancias, 1, 0)
    
    If Not mInstalacionElectrica Is Nothing Then
        pParams.Add "idInstalacionElectrica", mInstalacionElectrica.Id
    Else
        pParams.Add "idInstalacionElectrica", 0
    End If
    
    If Not mInstalacionGas Is Nothing Then
        pParams.Add "idInstalacionGas", mInstalacionGas.Id
    Else
        pParams.Add "idInstalacionGas", 0
    End If
    pParams.Add "Estado", mEstado
End Sub

Private Sub Class_Initialize()
    'Set mInvolucrados = New InvolucradoManager
End Sub
