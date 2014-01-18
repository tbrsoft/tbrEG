VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Atencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FEA9037A"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
'Faltan datos
Option Explicit

Public Enum eEstadoAtencion
    ePendiente = 1
    eFinalizado = 2
    eListaParaCerrar = 3
    'eCancelado=4 hara falta?
End Enum

Private mId As Long
Private mNroIncidente As String
Private mNroIncidenteInterno As String
Private mDespachador As Empleado
Private mAfiliado As Afiliado
Private mAfiliadoExterno As AfiliadoExterno
Private mFecha As Date
'Las horas las manejo string porq parece facil manejarse con un datetimepicker.
Private mHoraLlamada As String
'Hora de llegada al lugar de la emergencia
Private mQTH As String
'Hora de liberacion del movil
Private mVL As String
Private mEquipos As EquipoManager
Private mDireccion As BLCemi.Direccion
Private mDireccionOrigen As BLCemi.Direccion
Private mDireccionDestino As BLCemi.Direccion

Private mAreaProtegida As AreaProtegida
Private mObraSocial As ObraSocial
Private mServicioEmergencia As ServicioEmergencia

'ver
Private mOperador As String

Private mDiagnostico As String
Private mObservaciones As String

Private mSintoma As Sintoma
Private mTelefono As Telefono

'para no obligarlo a registrar el telefono
Private mTelefonoAuxiliar As String

Private mEstado As eEstadoAtencion

'diurno, nocturno, SADOFE, etc...
Private mTipoCodigo As TipoCodigo

Private mInfoContable As New InfoContable

'##ModelId=46A920690312
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Atencion", "Id Property Get", Err.Description)
End Property

'##ModelId=46A920690313
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Atencion", "Id Property Let", Err.Description)
End Property

'##ModelId=46A920690317
Public Property Get NroIncidente() As String
    On Error GoTo NroIncidenteErr

    '## Generated default body ...
    NroIncidente = mNroIncidente

    Exit Property
NroIncidenteErr:
    Call ErrorLog("Atencion", "NroIncidente Property Get", Err.Description)
End Property

'##ModelId=46A920690318
Public Property Let NroIncidente(ByVal vNewValue As String)
    On Error GoTo NroIncidenteErr

    '## Generated default body ...
    mNroIncidente = vNewValue

    Exit Property
NroIncidenteErr:
    Call ErrorLog("Atencion", "NroIncidente Property Let", Err.Description)
End Property

Public Property Get nroIncidenteInterno() As String
    On Error GoTo NroIncidenteInternoErr

    '## Generated default body ...
    nroIncidenteInterno = mNroIncidenteInterno

    Exit Property
NroIncidenteInternoErr:
    Call ErrorLog("Atencion", "NroIncidenteInterno Property Get", Err.Description)
End Property

'##ModelId=46A920690318
Public Property Let nroIncidenteInterno(ByVal vNewValue As String)
    On Error GoTo NroIncidenteInternoErr

    '## Generated default body ...
    mNroIncidenteInterno = vNewValue

    Exit Property
NroIncidenteInternoErr:
    Call ErrorLog("Atencion", "NroIncidenteInterno Property Let", Err.Description)
End Property
'##ModelId=46A92069031C
Public Property Get Despachador() As Empleado
    On Error GoTo DespachadorErr

    '## Generated default body ...
    Set Despachador = mDespachador

    Exit Property
DespachadorErr:
    Call ErrorLog("Atencion", "Despachador Property Get", Err.Description)
End Property

'##ModelId=46A92069031F
Public Property Set Despachador(ByVal vNewValue As Empleado)
    On Error GoTo DespachadorErr

    '## Generated default body ...
    Set mDespachador = vNewValue

    Exit Property
DespachadorErr:
    Call ErrorLog("Atencion", "Despachador Property Set", Err.Description)
End Property

'##ModelId=46A920690321
Public Property Get Afiliado() As Afiliado
    On Error GoTo AfiliadoErr

    '## Generated default body ...
    Set Afiliado = mAfiliado

    Exit Property
AfiliadoErr:
    Call ErrorLog("Atencion", "Afiliado Property Get", Err.Description)
End Property

'##ModelId=46A920690324
Public Property Set Afiliado(ByVal vNewValue As Afiliado)
    On Error GoTo AfiliadoErr

    '## Generated default body ...
    Set mAfiliado = vNewValue

    Exit Property
AfiliadoErr:
    Call ErrorLog("Atencion", "Afiliado Property Set", Err.Description)
End Property

'##ModelId=46A920690326
Public Property Get AfiliadoExterno() As AfiliadoExterno
    On Error GoTo AfiliadoExternoErr

    '## Generated default body ...
    Set AfiliadoExterno = mAfiliadoExterno

    Exit Property
AfiliadoExternoErr:
    Call ErrorLog("Atencion", "AfiliadoExterno Property Get", Err.Description)
End Property

'##ModelId=46A920690329
Public Property Set AfiliadoExterno(ByVal vNewValue As AfiliadoExterno)
    On Error GoTo AfiliadoExternoErr

    '## Generated default body ...
    Set mAfiliadoExterno = vNewValue

    Exit Property
AfiliadoExternoErr:
    Call ErrorLog("Atencion", "AfiliadoExterno Property Set", Err.Description)
End Property

'##ModelId=46A92069032B
Public Property Get Fecha() As Date
    On Error GoTo FechaErr

    '## Generated default body ...
    Fecha = mFecha

    Exit Property
FechaErr:
    Call ErrorLog("Atencion", "Fecha Property Get", Err.Description)
End Property

'##ModelId=46A92069032C
Public Property Let Fecha(ByVal vNewValue As Date)
    On Error GoTo FechaErr

    '## Generated default body ...
    mFecha = vNewValue

    Exit Property
FechaErr:
    Call ErrorLog("Atencion", "Fecha Property Let", Err.Description)
End Property

'##ModelId=46A920690330
Public Property Get HoraLlamada() As String
    On Error GoTo HoraLlamadaErr

    '## Generated default body ...
    HoraLlamada = mHoraLlamada

    Exit Property
HoraLlamadaErr:
    Call ErrorLog("Atencion", "HoraLlamada Property Get", Err.Description)
End Property

'##ModelId=46A920690331
Public Property Let HoraLlamada(ByVal vNewValue As String)
    On Error GoTo HoraLlamadaErr

    '## Generated default body ...
    mHoraLlamada = vNewValue

    Exit Property
HoraLlamadaErr:
    Call ErrorLog("Atencion", "HoraLlamada Property Let", Err.Description)
End Property

'##ModelId=46A920690335
Public Property Get QTH() As String
    On Error GoTo QTHErr

    '## Generated default body ...
    QTH = mQTH

    Exit Property
QTHErr:
    Call ErrorLog("Atencion", "QTH Property Get", Err.Description)
End Property

'##ModelId=46A920690336
Public Property Let QTH(ByVal vNewValue As String)
    On Error GoTo QTHErr

    '## Generated default body ...
    mQTH = vNewValue

    Exit Property
QTHErr:
    Call ErrorLog("Atencion", "QTH Property Let", Err.Description)
End Property

'##ModelId=46A92069033A
Public Property Get VL() As String
    On Error GoTo VLErr

    '## Generated default body ...
    VL = mVL

    Exit Property
VLErr:
    Call ErrorLog("Atencion", "VL Property Get", Err.Description)
End Property

'##ModelId=46A92069033B
Public Property Let VL(ByVal vNewValue As String)
    On Error GoTo VLErr

    '## Generated default body ...
    mVL = vNewValue

    Exit Property
VLErr:
    Call ErrorLog("Atencion", "VL Property Let", Err.Description)
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
    Call ErrorLog("Atencion", "Equipos Property Get", Err.Description)
End Property

Public Property Set Equipos(ByVal vNewValue As EquipoManager)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mEquipos = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("Atencion", "Equipos Property Set", Err.Description)
End Property

'##ModelId=46A920690349
Public Property Get Direccion() As BLCemi.Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("Atencion", "Direccion Property Get", Err.Description)
End Property

'##ModelId=46A92069034C
Public Property Set Direccion(ByVal vNewValue As BLCemi.Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("Atencion", "Direccion Property Set", Err.Description)
End Property

Public Property Get DireccionDestino() As BLCemi.Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set DireccionDestino = mDireccionDestino

    Exit Property
DireccionErr:
    Call ErrorLog("Atencion", "DireccionDestino Property Get", Err.Description)
End Property

Public Property Set DireccionDestino(ByVal vNewValue As BLCemi.Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccionDestino = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("Atencion", "DireccionDestino Property Set", Err.Description)
End Property

Public Property Get DireccionOrigen() As BLCemi.Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set DireccionOrigen = mDireccionOrigen

    Exit Property
DireccionErr:
    Call ErrorLog("Atencion", "DireccionOrigen Property Get", Err.Description)
End Property

Public Property Set DireccionOrigen(ByVal vNewValue As BLCemi.Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccionOrigen = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("Atencion", "DireccionOrigen Property Set", Err.Description)
End Property


'##ModelId=46A9819C019D
Public Property Get AreaProtegida() As AreaProtegida
    On Error GoTo AreaProtegidaErr

    '## Generated default body ...
    Set AreaProtegida = mAreaProtegida

    Exit Property
AreaProtegidaErr:
    Call ErrorLog("Atencion", "AreaProtegida Property Get", Err.Description)
End Property

'##ModelId=46A9819C01A0
Public Property Set AreaProtegida(ByVal vNewValue As AreaProtegida)
    On Error GoTo AreaProtegidaErr

    '## Generated default body ...
    Set mAreaProtegida = vNewValue

    Exit Property
AreaProtegidaErr:
    Call ErrorLog("Atencion", "AreaProtegida Property Set", Err.Description)
End Property

Public Property Get ObraSocial() As ObraSocial
    On Error GoTo ObraSocialErr
   
    Set ObraSocial = mObraSocial

    Exit Property
ObraSocialErr:
    Call ErrorLog("Atencion", "ObraSocial Property Get", Err.Description)
End Property

Public Property Set ObraSocial(ByVal vNewValue As ObraSocial)
    On Error GoTo ObraSocialErr
   
    Set mObraSocial = vNewValue

    Exit Property
ObraSocialErr:
    Call ErrorLog("Atencion", "ObraSocial Property Set", Err.Description)
End Property

Public Property Get ServicioEmergencia() As ServicioEmergencia
    On Error GoTo ServicioEmergenciaErr

    '## Generated default body ...
    Set ServicioEmergencia = mServicioEmergencia

    Exit Property
ServicioEmergenciaErr:
    Call ErrorLog("Atencion", "ServicioEmergencia Property Get", Err.Description)
End Property

Public Property Set ServicioEmergencia(ByVal vNewValue As ServicioEmergencia)
    On Error GoTo ServicioEmergenciaErr

    '## Generated default body ...
    Set mServicioEmergencia = vNewValue

    Exit Property
ServicioEmergenciaErr:
    Call ErrorLog("Atencion", "ServicioEmergencia Property Set", Err.Description)
End Property
Public Property Get Operador() As String
    On Error GoTo NroIncidenteErr
   
    Operador = mOperador
    
    Exit Property
NroIncidenteErr:
    Call ErrorLog("Atencion", "Operador Property Get", Err.Description)
End Property

Public Property Let Operador(ByVal vNewValue As String)
    On Error GoTo NroIncidenteErr

      mOperador = vNewValue

    Exit Property
NroIncidenteErr:
    Call ErrorLog("Atencion", "Operador Property Let", Err.Description)
End Property

Public Property Get Diagnostico() As String
    On Error GoTo DiagnosticoErr

    '## Generated default body ...
    Diagnostico = mDiagnostico

    Exit Property
DiagnosticoErr:
    Call ErrorLog("Atencion", "Diagnostico Property Get", Err.Description)
End Property

'##ModelId=46A920690318
Public Property Let Diagnostico(ByVal vNewValue As String)
    On Error GoTo DiagnosticoErr

    '## Generated default body ...
    '2011-03-13 en la base de datos el campo es de 255 !!!
    If Len(vNewValue) > 250 Then vNewValue = Left(vNewValue, 250)
    mDiagnostico = vNewValue

    Exit Property
DiagnosticoErr:
    Call ErrorLog("Atencion", "Diagnostico Property Let", Err.Description)
End Property

Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("Atencion", "Observaciones Property Get", Err.Description)
End Property

'##ModelId=46A920690318
Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("Atencion", "Observaciones Property Let", Err.Description)
End Property

Public Property Get Sintoma() As Sintoma
    Set Sintoma = mSintoma
End Property

Public Property Set Sintoma(pValue As Sintoma)
    Set mSintoma = pValue
End Property

Public Property Get Telefono() As Telefono
    Set Telefono = mTelefono
End Property

Public Property Set Telefono(pValue As Telefono)
    Set mTelefono = pValue
End Property

Public Property Get Estado() As eEstadoAtencion
Estado = mEstado
End Property

Public Property Let Estado(pValue As eEstadoAtencion)
    mEstado = pValue
End Property

Public Property Get TelefonoAuxilar() As String
    On Error GoTo NroIncidenteErr
   
    TelefonoAuxilar = mTelefonoAuxiliar
    
    Exit Property
NroIncidenteErr:
    Call ErrorLog("Atencion", "TelefonoAuxiliar Property Get", Err.Description)
End Property

Public Property Let TelefonoAuxilar(ByVal vNewValue As String)
    On Error GoTo NroIncidenteErr

      mTelefonoAuxiliar = vNewValue

    Exit Property
NroIncidenteErr:
    Call ErrorLog("Atencion", "TelefonoAuxiliar Property Let", Err.Description)
End Property

Public Property Get TipoCodigo() As TipoCodigo
    Set TipoCodigo = mTipoCodigo
End Property

Public Property Set TipoCodigo(pValue As TipoCodigo)
    Set mTipoCodigo = pValue
End Property

Public Property Get InfoContable() As InfoContable
    Set InfoContable = mInfoContable
End Property

Public Property Set InfoContable(pValue As InfoContable)
    Set mInfoContable = pValue
End Property

Public Property Get HasEquipo() As Boolean
    HasEquipo = mEquipos.Count <> 0
End Property

Public Property Get HasInfoContable() As Boolean
    HasInfoContable = mInfoContable.IsComplete
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
        Case "afiliado"
            Dim nom As String
            If Not mAfiliado Is Nothing Then
                nom = mAfiliado.NombreCompleto
            ElseIf Not mAfiliadoExterno Is Nothing Then
                nom = mAfiliadoExterno.NombreCompleto
            Else
                nom = "No Registrado"
            End If
            GetProperty = nom
        Case "despachador"
            GetProperty = mDespachador.NombreCompleto
        Case "tipocodigo"
            GetProperty = IIf(mTipoCodigo Is Nothing, "No aplica", mTipoCodigo.Nombre)
        Case "movil"
            GetProperty = IIf(Equipos.Count = 0, "<Sin Asignar>", Equipos.GetNombreMoviles)
        Case "pgdireccion"
            'si es un translado
            If mSintoma.Parent.Id = 100 Then
                GetProperty = mDireccionOrigen.GetBarrioCiudadProvincia + " - Calle: " + mDireccionOrigen.Calle + " " + mDireccionOrigen.Nro
            Else
                'pasa a veces que si la atencion no tenia cargada la direccion falla esto !!
                'If (mDireccion Is Nothing) Then
                '    Set mDireccion = New BLCemi.Direccion
                'End If
                GetProperty = mDireccion.GetBarrioCiudadProvincia + " - Calle: " + NN_str(mDireccion.Calle) + " " + NN_str(mDireccion.Nro)
            End If
        'lo agrego aca para no romper la compatibilidad
        Case "transcurrido"
            Dim h As String
            Dim m As String
           ' h = Trim(Str(DateDiff("h", TimeValue(mHoraLlamada), Time)))
            m = Trim(Str(DateDiff("n", TimeValue(mHoraLlamada), Time)))
            h = Int(val(m) / 60)
            m = m - h * 60
            GetProperty = IIf(h < 10, "0" + h, h) + ":" + IIf(m < 10, "0" + m, m)
        Case "pgtelefono"
            If mTelefono Is Nothing Then
                GetProperty = mTelefonoAuxiliar
            Else
                GetProperty = mTelefono.Numero
            End If
        
        Case "pgmontoservicio"
            GetProperty = Trim(Str(mInfoContable.Servicio))
        Case "pgcoseguro"
            GetProperty = Trim(Str(mInfoContable.Coseguro))
        Case "pgmontoabonado"
            GetProperty = Trim(Str(mInfoContable.MontoAbonado))
    End Select
End Property

Public Sub Guardar()
Dim params As New ParametrosManager
Dim trans As Transaccion
Set trans = BD.Transaccion

mId = BD.getIdDisponible(tAtencion) 'obtiene un id disponible para hacer registro nuevo

params.Add "Id", mId
params.Add "NroIncidente", mNroIncidente
params.Add "NroIncidenteInterno", mNroIncidenteInterno
params.Add "idDespachador", mDespachador.Id
If Not mAfiliado Is Nothing Then params.Add "idAfiliadoPropio", mAfiliado.Id
If Not mAfiliadoExterno Is Nothing Then params.Add "idAfiliadoExterno", mAfiliadoExterno.Id
params.Add "Fecha", mFecha
params.Add "HoraLlamada", mHoraLlamada
params.Add "QTH", mQTH
params.Add "VL", mVL
'params.Add "Equipos"
params.Add "idSintoma", mSintoma.Id

If mSintoma.Parent.Id = 100 Then 'traslado
    DireccionDestino.Guardar trans
    DireccionOrigen.Guardar trans
    
    params.Add "idDireccionOrigen", mDireccionOrigen.Id
    params.Add "idDireccionDestino", mDireccionDestino.Id
Else
    If Not mTelefono Is Nothing Then
        params.Add "idTelefono", mTelefono.Id
    Else
        params.Add "telefonoAuxiliar", mTelefonoAuxiliar
    End If
    
    Direccion.Guardar trans
    params.Add "idDireccion", mDireccion.Id
End If

If Not mAreaProtegida Is Nothing Then params.Add "idAreaProtegida", mAreaProtegida.Id
If Not mObraSocial Is Nothing Then params.Add "idObraSocial", mObraSocial.Id
If Not mServicioEmergencia Is Nothing Then params.Add "idServicioEmergencia", mServicioEmergencia.Id

'ver
params.Add "Operador", mOperador
params.Add "Diagnostico", mDiagnostico
params.Add "Observaciones", mObservaciones

params.Add "Estado", mEstado

If Not mTipoCodigo Is Nothing Then params.Add "idTipoCodigo", mTipoCodigo.Id

'aca agrego la informacion contable
mInfoContable.FillParameters params

trans.AddInsertCommand tAtencion, params
Me.Equipos.Guardar trans, Me.Id
trans.Execute
'para q avise a la red q hubo cambios
AtencionesLocal.AtencionAgregadaoModificada
End Sub

Public Sub GuardarModificaciones(pEmpleadoActual As Empleado)
'para registrar los cambios
Dim segAtencion As New SeguimientoAtencion

segAtencion.ObtenerOriginal mId

'analizar bien
Dim params As New ParametrosManager
Dim trans As Transaccion
Set trans = BD.Transaccion

params.Add "Id", mId, True
params.Add "NroIncidente", mNroIncidente
params.Add "NroIncidenteInterno", mNroIncidenteInterno

'params.Add "idDespachador", mDespachador.id
If Not mAfiliado Is Nothing Then params.Add "idAfiliadoPropio", mAfiliado.Id
If Not mAfiliadoExterno Is Nothing Then params.Add "idAfiliadoExterno", mAfiliadoExterno.Id

'la fecha y hora no deberian ser modificabless
'params.Add "Fecha", mFecha
'params.Add "HoraLlamada", mHoraLlamada

params.Add "QTH", mQTH
params.Add "VL", mVL
'params.Add "Equipos"
params.Add "idSintoma", mSintoma.Id

If mSintoma.Parent.Id = 100 Then 'traslado
    DireccionDestino.GuardarModificaciones trans
    DireccionOrigen.GuardarModificaciones trans
    
'    params.Add "idDireccionOrigen", mDireccionOrigen.id
'    params.Add "idDireccionDestino", mDireccionDestino.id
Else
    Direccion.GuardarModificaciones trans
    'params.Add "idDireccion", mDireccion.id
End If

If Not mTelefono Is Nothing Then
    'si ya le asigne telefono lo guardo y borro el auxiliar
    params.Add "idTelefono", mTelefono.Id
    params.Add "telefonoauxiliar", ""
End If

If Not mAreaProtegida Is Nothing Then params.Add "idAreaProtegida", mAreaProtegida.Id
If Not mObraSocial Is Nothing Then params.Add "idObraSocial", mObraSocial.Id
If Not mServicioEmergencia Is Nothing Then params.Add "idServicioEmergencia", mServicioEmergencia.Id

'ver
params.Add "Operador", mOperador
params.Add "Diagnostico", mDiagnostico
params.Add "Observaciones", mObservaciones

If Not mTipoCodigo Is Nothing Then params.Add "idTipoCodigo", mTipoCodigo.Id

mInfoContable.FillParameters params

params.Add "Estado", mEstado

trans.AddUpdateCommand tAtencion, params
Equipos.GuardarCambios trans, Me.Id


trans.Execute
'muy probablemente sea mejor usar la misma transaccion, ver si se puede
segAtencion.RegistrarCambios pEmpleadoActual

'para q avise a la red q hubo cambios
AtencionesLocal.AtencionAgregadaoModificada

End Sub

