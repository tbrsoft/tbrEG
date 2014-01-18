VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeguimientoSiniestro"
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
Private Fecha As Date
Private HoraLlamada As String
Private QTH As String
Private VL As String
Private mSalidaPreInspeccion As String
Private mLlegadaPreInspeccion As String
Private mSalidaDotacion As String

'para comparar si son los mismos equipos
Private idsEquipos As String
Private idsColaboraciones As String
Private idsInvolucrados As String
Private idsVehiculos As String

Private Observaciones As String
Private idSintoma As Long
Private Estado As eEstadoAtencion
'direcciones
Private mDireccion As Direccion
Private idDireccion As Long
Private nroIncidenteInterno As String
Private mCronica As String 'reseña
Private mSintoma As Sintoma
Private mInvolucrados As InvolucradoManager
Private mVehiculos As VehiculoManager
Private mDescripcionPerjuiciosCampo As String
Private mMaterialesCombustibles As String
Private mDescripcionPerjuiciosVivienda As String
Private mAccesoPor As String
Private mReferencias As String
Private mAseguradora As String
Private mDescripcionMaterial As String
Private mEquiposEspeciales As String
Private mPoliza As String
Private mAmbientesAfectadosVivienda As Integer
Private mHectareasAfectadasCampo As Integer
Private mIdInstalacionElectrica As Long
Private mIDInstalacionGas As Long
Private mColaboracionBomberos As ColaboracionManager
Private mPoliciaACargo As String
Private mPoliciaCantidad As Integer
Private mPoliciaMovil As String
Private mSEMedico As String
Private mSECentroAsistencial As String
Private mSEMedicoMP As String
Private mSEAmbulancias As Boolean

'warning: ver como comparar sub clases (involucrados, vehiculos y colaboraciones)

Public Sub ObtenerOriginal(pId As Long)
    mId = pId
    Set rsOriginal = BD.SelectSQL("select * from atencionb where id=" + Str(pId))
        
    NroIncidente = "" & rsOriginal!NroIncidente
    Fecha = rsOriginal!Fecha
    HoraLlamada = "" & rsOriginal!HoraLlamada
    QTH = "" & rsOriginal!QTH
    VL = "" & rsOriginal!VL
    mSalidaPreInspeccion = "" & rsOriginal!SalidaPreInspeccion
    mLlegadaPreInspeccion = "" & rsOriginal!LlegadaPreInspeccion
    mSalidaDotacion = "" & rsOriginal!SalidaDotacion

    idSintoma = rsOriginal!idSintoma
    idDireccion = ReplaceNulls(rsOriginal!idDireccion)
    idDespachador = rsOriginal!idDespachador
    Estado = rsOriginal!Estado
    Observaciones = "" & rsOriginal!Observaciones
    nroIncidenteInterno = "" & rsOriginal!nroIncidenteInterno
    mCronica = "" & rsOriginal!Cronica
    mDescripcionPerjuiciosCampo = "" & rsOriginal!DescripcionPerjuiciosCampo
    mMaterialesCombustibles = "" & rsOriginal!MaterialesCombustibles
    mDescripcionPerjuiciosVivienda = "" & rsOriginal!DescripcionPerjuiciosVivienda
    mAccesoPor = "" & rsOriginal!AccesoPor
    mReferencias = "" & rsOriginal!Referencias
    mAseguradora = "" & rsOriginal!Aseguradora
    mDescripcionMaterial = "" & rsOriginal!DescripcionMaterial
    mEquiposEspeciales = "" & rsOriginal!EquiposEspeciales
    mPoliza = "" & rsOriginal!Poliza
    mAmbientesAfectadosVivienda = ReplaceNulls(rsOriginal!AmbientesAfectadosVivienda)
    mHectareasAfectadasCampo = ReplaceNulls(rsOriginal!HectareasAfectadasCampo)
    mIdInstalacionElectrica = ReplaceNulls(rsOriginal!idInstalacionElectrica)
    mIDInstalacionGas = ReplaceNulls(rsOriginal!idInstalacionGas)
    mPoliciaACargo = "" & rsOriginal!PoliciaACargo
    mPoliciaCantidad = ReplaceNulls(rsOriginal!PoliciaCantidad)
    mPoliciaMovil = "" & rsOriginal!PoliciaMovil
    mSEMedico = "" & rsOriginal!SEMedico
    mSECentroAsistencial = "" & rsOriginal!SECentroAsistencial
    mSEMedicoMP = "" & rsOriginal!SEMedicoMP
    mSEAmbulancias = rsOriginal!SEAmbulancias

    'controlo la direccion
    If idDireccion <> 0 Then
        Set mDireccion = New Direccion
        mDireccion.Cargar idDireccion
    End If
    
    'EQUIPOS
    Set rsOriginal = BD.SelectSQL("select * from equipoxatencion where idatencion=" + Str(pId) + " order by idequipo")
    'guardo los equipos en una cadena, estan ordenados
    While Not rsOriginal.EOF
        idsEquipos = idsEquipos + Str(rsOriginal!idEquipo) + ","
        rsOriginal.MoveNext
    Wend
    
'    'COLABORACIONES, NO FUNCIONA ASI
'    Set rsOriginal = BD.SelectSQL("select * from equipoxatencion where idatencion=" + Str(pId) + " order by idequipo")
'    'guardo los equipos en una cadena, estan ordenados
'    While Not rsOriginal.EOF
'        idsColaboraciones = idsColaboraciones + Str(rsOriginal!idEquipo) + ","
'        rsOriginal.MoveNext
'    Wend
'
'    'INVOLUCRADOS
'    Set rsOriginal = BD.SelectSQL("select * from equipoxatencion where idatencion=" + Str(pId) + " order by idequipo")
'    'guardo los equipos en una cadena, estan ordenados
'    While Not rsOriginal.EOF
'        idsEquipos = idsEquipos + Str(rsOriginal!idEquipo) + ","
'        rsOriginal.MoveNext
'    Wend
'
'    'VEHICULOS
'    Set rsOriginal = BD.SelectSQL("select * from equipoxatencion where idatencion=" + Str(pId) + " order by idequipo")
'    'guardo los equipos en una cadena, estan ordenados
'    While Not rsOriginal.EOF
'        idsEquipos = idsEquipos + Str(rsOriginal!idEquipo) + ","
'        rsOriginal.MoveNext
'    Wend
    
    Set rsOriginal = Nothing
End Sub

Public Sub RegistrarCambios(pEmpleado As Empleado)

    Set rsModificado = BD.SelectSQL("select * from atencionb where id=" + Str(mId))
           
    Dim seg As New DetalleSegManager
           
    seg.idAtencion = mId
           
    If NroIncidente <> "" & rsModificado!NroIncidente Then seg.Nuevo "NroIncidente", Str(NroIncidente), rsModificado!NroIncidente, pEmpleado, Date, Time$
    If Fecha <> rsModificado!Fecha Then seg.Nuevo "fecha", Str(Fecha), Str(rsModificado!Fecha), pEmpleado, Date, Time$
    If HoraLlamada <> "" & rsModificado!HoraLlamada Then seg.Nuevo "HoraLlamada", Str(HoraLlamada), rsModificado!HoraLlamada, pEmpleado, Date, Time$
    If QTH <> "" & rsModificado!QTH Then seg.Nuevo "QTH", QTH, rsModificado!QTH, pEmpleado, Date, Time$
    If VL <> "" & rsModificado!VL Then seg.Nuevo "VL", VL, rsModificado!VL, pEmpleado, Date, Time$
    If mSalidaPreInspeccion <> "" & rsModificado!SalidaPreInspeccion Then seg.Nuevo "SalidaPreInspeccion", mSalidaPreInspeccion, rsModificado!SalidaPreInspeccion, pEmpleado, Date, Time$
    If mLlegadaPreInspeccion <> "" & rsModificado!LlegadaPreInspeccion Then seg.Nuevo "LlegadaPreInspeccion", mLlegadaPreInspeccion, rsModificado!LlegadaPreInspeccion, pEmpleado, Date, Time$
    If mSalidaDotacion <> "" & rsModificado!SalidaDotacion Then seg.Nuevo "SalidaDotacion", mSalidaDotacion, rsModificado!SalidaDotacion, pEmpleado, Date, Time$
    If idSintoma <> rsModificado!idSintoma Then seg.Nuevo "Sintoma", Str(idSintoma), rsModificado!idSintoma, pEmpleado, Date, Time$
    If Estado <> rsModificado!Estado Then seg.Nuevo "Estado", Str(Estado), rsModificado!Estado, pEmpleado, Date, Time$
    If Observaciones <> "" & rsModificado!Observaciones Then seg.Nuevo "Observaciones", Observaciones, rsModificado!Observaciones, pEmpleado, Date, Time$
    If nroIncidenteInterno <> "" & rsModificado!nroIncidenteInterno Then seg.Nuevo "NroIncidenteInterno", nroIncidenteInterno, rsModificado!nroIncidenteInterno, pEmpleado, Date, Time$
    
    If mCronica <> "" & rsModificado!Cronica Then seg.Nuevo "Cronica", mCronica, rsModificado!Cronica, pEmpleado, Date, Time$
    If mDescripcionPerjuiciosCampo <> "" & rsModificado!DescripcionPerjuiciosCampo Then seg.Nuevo "DescripcionPerjuiciosCampo", mDescripcionPerjuiciosCampo, rsModificado!DescripcionPerjuiciosCampo, pEmpleado, Date, Time$
    If mMaterialesCombustibles <> "" & rsModificado!MaterialesCombustibles Then seg.Nuevo "MaterialesCombustibles", mMaterialesCombustibles, rsModificado!MaterialesCombustibles, pEmpleado, Date, Time$
    If mDescripcionPerjuiciosVivienda <> "" & rsModificado!DescripcionPerjuiciosVivienda Then seg.Nuevo "DescripcionPerjuiciosVivienda", mDescripcionPerjuiciosVivienda, rsModificado!DescripcionPerjuiciosVivienda, pEmpleado, Date, Time$
    If mAccesoPor <> "" & rsModificado!AccesoPor Then seg.Nuevo "AccesoPor", mAccesoPor, rsModificado!AccesoPor, pEmpleado, Date, Time$
    If mReferencias <> "" & rsModificado!Referencias Then seg.Nuevo "Referencias", mReferencias, rsModificado!Referencias, pEmpleado, Date, Time$
    If mAseguradora <> "" & rsModificado!Aseguradora Then seg.Nuevo "Aseguradora", mAseguradora, rsModificado!Aseguradora, pEmpleado, Date, Time$
    If mDescripcionMaterial <> "" & rsModificado!DescripcionMaterial Then seg.Nuevo "DescripcionMaterial", mDescripcionMaterial, rsModificado!DescripcionMaterial, pEmpleado, Date, Time$
    If mEquiposEspeciales <> "" & rsModificado!EquiposEspeciales Then seg.Nuevo "EquiposEspeciales", mEquiposEspeciales, rsModificado!EquiposEspeciales, pEmpleado, Date, Time$
    If mPoliza <> "" & rsModificado!Poliza Then seg.Nuevo "Poliza", mPoliza, rsModificado!Poliza, pEmpleado, Date, Time$
    If mAmbientesAfectadosVivienda <> rsModificado!AmbientesAfectadosVivienda Then seg.Nuevo "AmbientesAfectadosVivienda", Str(mAmbientesAfectadosVivienda), rsModificado!AmbientesAfectadosVivienda, pEmpleado, Date, Time$
    If mHectareasAfectadasCampo <> rsModificado!HectareasAfectadasCampo Then seg.Nuevo "HectareasAfectadasCampo", Str(mHectareasAfectadasCampo), rsModificado!HectareasAfectadasCampo, pEmpleado, Date, Time$
    If mIdInstalacionElectrica <> rsModificado!idInstalacionElectrica Then seg.Nuevo "InstalacionElectrica", Str(mIdInstalacionElectrica), rsModificado!idInstalacionElectrica, pEmpleado, Date, Time$
    If mIDInstalacionGas <> rsModificado!idInstalacionGas Then seg.Nuevo "InstalacionGas", Str(mIDInstalacionGas), rsModificado!idInstalacionGas, pEmpleado, Date, Time$
    If mPoliciaACargo <> "" & rsModificado!PoliciaACargo Then seg.Nuevo "PoliciaACargo", mPoliciaACargo, rsModificado!PoliciaACargo, pEmpleado, Date, Time$
    If mPoliciaCantidad <> rsModificado!PoliciaCantidad Then seg.Nuevo "PoliciaCantidad", Str(mPoliciaCantidad), rsModificado!PoliciaCantidad, pEmpleado, Date, Time$
    If mPoliciaMovil <> "" & rsModificado!PoliciaMovil Then seg.Nuevo "PoliciaMovil", mPoliciaMovil, rsModificado!PoliciaMovil, pEmpleado, Date, Time$
    If mSEMedico <> "" & rsModificado!SEMedico Then seg.Nuevo "SEMedico", mSEMedico, rsModificado!SEMedico, pEmpleado, Date, Time$
    If mSECentroAsistencial <> "" & rsModificado!SECentroAsistencial Then seg.Nuevo "SECentroAsistencial", mSECentroAsistencial, rsModificado!SECentroAsistencial, pEmpleado, Date, Time$
    If mSEMedicoMP <> "" & rsModificado!SEMedicoMP Then seg.Nuevo "SEMedicoMP", mSEMedicoMP, rsModificado!SEMedicoMP, pEmpleado, Date, Time$
    If mSEAmbulancias <> rsModificado!SEAmbulancias Then seg.Nuevo "SEAmbulancias", Str(mSEAmbulancias), rsModificado!SEAmbulancias, pEmpleado, Date, Time$

'terminar con estos 3
'Private mColaboracionBomberos As ColaboracionManager
'Private mInvolucrados As InvolucradoManager
'Private mVehiculos As VehiculoManager
    
    'comparo las direcciones
    Dim mDireccionNew As Direccion
    
    If Not mDireccion Is Nothing And ReplaceNulls(rsModificado!idDireccion) <> 0 Then
        Set mDireccionNew = New Direccion
        mDireccionNew.Cargar rsModificado!idDireccion
        If Not mDireccion.EqualsTo(mDireccionNew) Then seg.Nuevo "Direccion", mDireccion.GetBarrioCiudadProvincia + ", Calle: " + mDireccion.Calle + " " + mDireccion.Nro, mDireccionNew.GetBarrioCiudadProvincia + ", Calle: " + mDireccionNew.Calle + " " + mDireccionNew.Nro, pEmpleado, Date, Time$
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
    
    CompararColecciones seg, "Dotaciones", idsEquipos, idsEquiposMod, pEmpleado
    
    seg.Save
    
End Sub

Private Sub CompararColecciones(pSeg As DetalleSegManager, pNombreCampo As String, pIdsOriginales As String, pIdsModificados As String, pEmpleado As Empleado)
    Dim eOrig() As String
    Dim eMod() As String
    eOrig = Split(Replace(pIdsOriginales, " ", ""), ",")
    eMod = Split(Replace(pIdsModificados, " ", ""), ",")
    Dim i As Integer
    
    If UBound(eOrig) <> UBound(eMod) Then
       pSeg.Nuevo pNombreCampo, pIdsOriginales, pIdsModificados, pEmpleado, Date, Time$
    Else
        For i = 0 To UBound(eOrig)
            If eOrig(i) <> eMod(i) Then
                pSeg.Nuevo pNombreCampo, pIdsOriginales, pIdsModificados, pEmpleado, Date, Time$
                Exit For
            End If
        Next
    End If
End Sub

Private Function ReplaceNulls(value) As Long
    ReplaceNulls = IIf(IsNull(value), 0, value)
End Function

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

