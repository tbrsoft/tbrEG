VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeguimientoAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim rsOriginal
Dim rsModificado
Dim mId As Long
Dim fila

Private NroIncidente As String
Private idDespachador As Long
Private idAfiliadoPropio As Long
Private idAfiliadoExterno As Long
Private Fecha As Date
Private HoraLlamada As String
Private QTH As String
Private VL As String
'para comparar si son los mismos equipos
Private idsEquipos As String

Private idAreaProtegida As Long
Private idObraSocial As Long
Private idServicioEmergencia As Long
Private Operador As String
Private Diagnostico As String
Private Observaciones As String
Private idSintoma As Long
Private idTelefono As Long
Private TelefonoAuxiliar As String
Private Estado As eEstadoAtencion
'direcciones
Private mDireccion As Direccion
Private mDireccionOrigen As Direccion
Private mDireccionDestino As Direccion
Private idDireccion As Long
Private idDireccionOrigen As Long
Private idDireccionDestino As Long
Private idTipoCodigo As Long
'campos de infocontable
Private Servicio As Currency
Private copago As Currency
Private abonado As Currency
Private nroIncidenteInterno As String

Public Sub ObtenerOriginal(pId As Long)
    mId = pId
    Set rsOriginal = BD.SelectSQL("select * from atencion where id=" + Str(pId))
    
    
    NroIncidente = "" & rsOriginal!NroIncidente
    Fecha = rsOriginal!Fecha
    HoraLlamada = "" & rsOriginal!HoraLlamada
    QTH = "" & rsOriginal!QTH
    VL = "" & rsOriginal!VL
    idSintoma = rsOriginal!idSintoma
    idDireccion = ReplaceNulls(rsOriginal!idDireccion)
    Operador = "" & rsOriginal!Operador
    idDespachador = rsOriginal!idDespachador
    idAreaProtegida = ReplaceNulls(rsOriginal!idAreaProtegida)
    idObraSocial = ReplaceNulls(rsOriginal!idObraSocial)
    idServicioEmergencia = ReplaceNulls(rsOriginal!idServicioEmergencia)
    idAfiliadoPropio = ReplaceNulls(rsOriginal!idAfiliadoPropio)
    idAfiliadoExterno = ReplaceNulls(rsOriginal!idAfiliadoExterno)
    idDireccionDestino = ReplaceNulls(rsOriginal!idDireccionDestino)
    idDireccionOrigen = ReplaceNulls(rsOriginal!idDireccionOrigen)
    idTipoCodigo = ReplaceNulls(rsOriginal!idTipoCodigo)
    Estado = rsOriginal!Estado
    Diagnostico = "" & rsOriginal!Diagnostico
    Observaciones = "" & rsOriginal!Observaciones
    idTelefono = ReplaceNulls(rsOriginal!idTelefono)
    TelefonoAuxiliar = "" & rsOriginal!TelefonoAuxiliar
    
    Servicio = rsOriginal!Servicio
    copago = rsOriginal!copago
    abonado = rsOriginal!abonado
    nroIncidenteInterno = "" & rsOriginal!nroIncidenteInterno
    'controlo las direcciones
    If idDireccion <> 0 Then
        Set mDireccion = New Direccion
        mDireccion.Cargar idDireccion
    End If
     
    If idDireccionDestino <> 0 Then
        Set mDireccionDestino = New Direccion
        mDireccionDestino.Cargar idDireccionDestino
    End If
    
    If idDireccionOrigen <> 0 Then
        Set mDireccionOrigen = New Direccion
        mDireccionOrigen.Cargar idDireccionOrigen
    End If
    
    Set rsOriginal = BD.SelectSQL("select * from equipoxatencion where idatencion=" + Str(pId) + " order by idequipo")
    'guardo los equipos en una cadena, estan ordenados
    While Not rsOriginal.EOF
        idsEquipos = idsEquipos + Str(rsOriginal!idEquipo) + ","
        rsOriginal.MoveNext
    Wend
    Set rsOriginal = Nothing
End Sub

Public Sub RegistrarCambios(pEmpleado As Empleado)

    Set rsModificado = BD.SelectSQL("select * from atencion where id=" + Str(mId))
           
    Dim seg As New DetalleSegManager
           
    seg.idAtencion = mId
           
    If NroIncidente <> "" & rsModificado!NroIncidente Then seg.Nuevo "NroIncidente", NroIncidente, rsModificado!NroIncidente, pEmpleado, Date, Time$
    If Fecha <> rsModificado!Fecha Then seg.Nuevo "fecha", Str(Fecha), Str(rsModificado!Fecha), pEmpleado, Date, Time$
    If HoraLlamada <> "" & rsModificado!HoraLlamada Then seg.Nuevo "HoraLlamada", Str(HoraLlamada), rsModificado!HoraLlamada, pEmpleado, Date, Time$
    If QTH <> "" & rsModificado!QTH Then seg.Nuevo "QTH", QTH, rsModificado!QTH, pEmpleado, Date, Time$
    If VL <> "" & rsModificado!VL Then seg.Nuevo "VL", VL, rsModificado!VL, pEmpleado, Date, Time$
    If idSintoma <> rsModificado!idSintoma Then seg.Nuevo "Sintoma", Str(idSintoma), rsModificado!idSintoma, pEmpleado, Date, Time$
    If Operador <> "" & rsModificado!Operador Then seg.Nuevo "Operador", Operador, rsModificado!Operador, pEmpleado, Date, Time$
   'no se modifica If idDespachador <> rsModificado!idDespachador Then seg.Nuevo "idDespachador", Str(idDespachador), rsModificado!idDespachador, pEmpleado, Date, Time$
    If idAreaProtegida <> ReplaceNulls(rsModificado!idAreaProtegida) Then seg.Nuevo "AreaProtegida", Str(idAreaProtegida), rsModificado!idAreaProtegida, pEmpleado, Date, Time$
    If idObraSocial <> ReplaceNulls(rsModificado!idObraSocial) Then seg.Nuevo "ObraSocial", Str(idObraSocial), rsModificado!idObraSocial, pEmpleado, Date, Time$
    If idServicioEmergencia <> ReplaceNulls(rsModificado!idServicioEmergencia) Then seg.Nuevo "ServicioEmergencia", Str(idServicioEmergencia), rsModificado!idServicioEmergencia, pEmpleado, Date, Time$
    If idAfiliadoPropio <> ReplaceNulls(rsModificado!idAfiliadoPropio) Then seg.Nuevo "AfiliadoPropio", Str(idAfiliadoPropio), rsModificado!idAfiliadoPropio, pEmpleado, Date, Time$
    If idAfiliadoExterno <> ReplaceNulls(rsModificado!idAfiliadoExterno) Then seg.Nuevo "AfiliadoExterno", Str(idAfiliadoExterno), rsModificado!idAfiliadoExterno, pEmpleado, Date, Time$
    If idTipoCodigo <> ReplaceNulls(rsModificado!idTipoCodigo) Then seg.Nuevo "TipoCodigo", Str(idTipoCodigo), rsModificado!idTipoCodigo, pEmpleado, Date, Time$
    If Estado <> rsModificado!Estado Then seg.Nuevo "Estado", Str(Estado), rsModificado!Estado, pEmpleado, Date, Time$
    If Diagnostico <> "" & rsModificado!Diagnostico Then seg.Nuevo "Diagnostico", Diagnostico, rsModificado!Diagnostico, pEmpleado, Date, Time$
    If Observaciones <> "" & rsModificado!Observaciones Then seg.Nuevo "Observaciones", Observaciones, rsModificado!Observaciones, pEmpleado, Date, Time$
    If idTelefono <> ReplaceNulls(rsModificado!idTelefono) Then seg.Nuevo "Telefono", Str(idTelefono), rsModificado!idTelefono, pEmpleado, Date, Time$
    If TelefonoAuxiliar <> "" & rsModificado!TelefonoAuxiliar Then seg.Nuevo "telefonoAuxiliar", TelefonoAuxiliar, rsModificado!TelefonoAuxiliar, pEmpleado, Date, Time$
    If Servicio <> "" & rsModificado!Servicio Then seg.Nuevo "Servicio", Str(Servicio), rsModificado!Servicio, pEmpleado, Date, Time$
    If copago <> "" & rsModificado!copago Then seg.Nuevo "Copago", Str(copago), rsModificado!copago, pEmpleado, Date, Time$
    If abonado <> "" & rsModificado!abonado Then seg.Nuevo "MontoAbonado", Str(abonado), rsModificado!abonado, pEmpleado, Date, Time$
    If nroIncidenteInterno <> "" & rsModificado!nroIncidenteInterno Then seg.Nuevo "NroIncidenteInterno", nroIncidenteInterno, rsModificado!nroIncidenteInterno, pEmpleado, Date, Time$
    
    'comparo las direcciones
    Dim mDireccionNew As Direccion
    Dim mDireccionOrigenNew As Direccion
    Dim mDireccionDestinoNew As Direccion

    
    If Not mDireccion Is Nothing And ReplaceNulls(rsModificado!idDireccion) <> 0 Then
        Set mDireccionNew = New Direccion
        mDireccionNew.Cargar rsModificado!idDireccion
        If Not mDireccion.EqualsTo(mDireccionNew) Then seg.Nuevo "Direccion", mDireccion.GetBarrioCiudadProvincia + ", Calle: " + mDireccion.Calle + " " + mDireccion.Nro, mDireccionNew.GetBarrioCiudadProvincia + ", Calle: " + mDireccionNew.Calle + " " + mDireccionNew.Nro, pEmpleado, Date, Time$
    End If
    
    If Not mDireccionDestino Is Nothing And ReplaceNulls(rsModificado!idDireccionDestino) <> 0 Then
        Set mDireccionDestinoNew = New Direccion
        mDireccionDestinoNew.Cargar rsModificado!idDireccionDestino
        If Not mDireccionDestino.EqualsTo(mDireccionDestinoNew) Then seg.Nuevo "Direccion Destino", mDireccionDestino.GetBarrioCiudadProvincia + ", Calle: " + mDireccionDestino.Calle + " " + mDireccionDestino.Nro, mDireccionDestinoNew.GetBarrioCiudadProvincia + ", Calle: " + mDireccionDestinoNew.Calle + " " + mDireccionDestinoNew.Nro, pEmpleado, Date, Time$
    End If
    If Not mDireccionOrigen Is Nothing And ReplaceNulls(rsModificado!idDireccionOrigen) <> 0 Then
        Set mDireccionOrigenNew = New Direccion
        mDireccionOrigenNew.Cargar rsModificado!idDireccionOrigen
        If Not mDireccionOrigen.EqualsTo(mDireccionOrigenNew) Then seg.Nuevo "Direccion Origen", mDireccionOrigen.GetBarrioCiudadProvincia + ", Calle: " + mDireccionOrigen.Calle + " " + mDireccionOrigen.Nro, mDireccionOrigenNew.GetBarrioCiudadProvincia + ", Calle: " + mDireccionOrigenNew.Calle + " " + mDireccionOrigenNew.Nro, pEmpleado, Date, Time$
    End If
   
   'VER si se modifican los equipos
    Dim idsEquiposMod  As String
    Set rsModificado = BD.SelectSQL("select * from equipoxatencion where idatencion=" + Str(mId) + " order by idequipo")
    'guardo los equipos en una cadena, estan ordenados
    While Not rsModificado.EOF
        idsEquiposMod = idsEquiposMod + Str(rsModificado!idEquipo) + ","
        rsModificado.MoveNext
    Wend
    Set rsModificado = Nothing
    Dim eOrig() As String
    Dim eMod() As String
    eOrig = Split(Replace(idsEquipos, " ", ""), ",")
    eMod = Split(Replace(idsEquiposMod, " ", ""), ",")
    Dim i As Integer
    
    If UBound(eOrig) <> UBound(eMod) Then
       seg.Nuevo "Dotaciones", idsEquipos, idsEquiposMod, pEmpleado, Date, Time$
    Else
        For i = 0 To UBound(eOrig)
            If eOrig(i) <> eMod(i) Then
                seg.Nuevo "Dotaciones", idsEquipos, idsEquiposMod, pEmpleado, Date, Time$
                Exit For
            End If
        Next
    End If
    seg.Save
    
End Sub

Private Function ReplaceNulls(value) As Long
    ReplaceNulls = IIf(IsNull(value), 0, value)
End Function

'CAMPOS DE ATENCION
'id
'NroIncidente
'Fecha
'HoraLlamada
'QTH
'VL
'idSintoma
'idDireccion
'Operador
'idDespachador
'idAreaProtegida
'idObraSocial
'idServicioEmergencia
'idAfiliadoPropio
'idAfiliadoExterno
'idDireccionDestino
'idDireccionOrigen
'Estado
'Diagnostico
'Observaciones
'idTelefono
'telefonoAuxiliar


'CAMPOS DE ATENCIONB
'Private mId As Long
'Private mNroIncidente As String
'Private mNroIncidenteInterno As String
'Private mDespachador As Empleado
'Private mFecha As Date
'Private mHoraLlamada As String
'Private mSalidaPreInspeccion As String
'Private mLlegadaPreInspeccion As String
'Private mSalidaDotacion As String
'Private mQTH As String
'Private mVL As String
'
'Private mEquipos As EquipoManager
'Private mDireccion As Direccion
'Private mCronica As String 'reseña
'Private mObservaciones As String
'Private mSintoma As Sintoma
'Private mEstado As eEstadoAtencion
'
'Private mInvolucrados As InvolucradoManager
'Private mVehiculos As VehiculoManager
'
'Dim mDescripcionPerjuiciosCampo As String
'Dim mMaterialesCombustibles As String
'Dim mDescripcionPerjuiciosVivienda As String
'Dim mAccesoPor As String
'Dim mReferencias As String
'Dim mAseguradora As String
'Dim mDescripcionMaterial As String
'Dim mEquiposEspeciales As String
'Dim mPoliza As String
'Dim mAmbientesAfectadosVivienda As Integer
'Dim mHectareasAfectadasCampo As Integer
'Dim mInstalacionElectrica As InstElectrica
'Dim mInstalacionGas As InstGas
'Dim mColaboracionBomberos As ColaboracionManager
'
'Dim mPoliciaACargo As String
'Dim mPoliciaCantidad As Integer
'Dim mPoliciaMovil As String
'Dim mSEMedico As String
'Dim mSECentroAsistencial As String
'Dim mSEMedicoMP As String
'Dim mSEAmbulancias As Boolean
