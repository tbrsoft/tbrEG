VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionBManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event HasChanged()

Private mCol As Collection

Private mSQL As String

Private WithEvents mNetMonitor As NetMonitor
Attribute mNetMonitor.VB_VarHelpID = -1

Friend Property Get Sql() As String
    Sql = mSQL
End Property

Friend Property Let Sql(pSql As String)
    mSQL = pSql
End Property

Friend Sub AtencionBAgregadaoModificada()
    'prueba
    mNetMonitor.MiRedLocal.EnviarEventoABM Alta, "AtencionB", 1
    'Reload 'porq aca??
End Sub

'Public Sub Remove(pId As Long)
'    On Error GoTo RemoveErr
'
'    '## Your code goes here ...
'
'    Exit Sub
'RemoveErr:
'    Call ErrorLog(MyUnhandledError, "Remove Sub")
'End Sub

'##ModelId=46A80BE40222
Public Sub AddItem(pAtencionB As AtencionB)
    On Error GoTo AddItemErr

    mCol.Add pAtencionB, "k" + Str(pAtencionB.Id)

    Exit Sub
AddItemErr:
If Err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
   ErrorLog "AtencionBManager", "AddItem", Err.Description
End If
End Sub

Public Sub Reload()
    If mSQL <> "" Then 'esto es porq AtencionBeslocal no tiene q recargarse...
        Dim rs
        Set rs = BD.SelectSQL(mSQL)
        Set mCol = New Collection
        Cargar rs, Me
        RaiseEvent HasChanged
    End If
End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mNetMonitor = NetMonitorLocal
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

Public Property Get Item(pId As Long) As AtencionB
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Function LoadById(pId As Long) As AtencionB
On Error GoTo errMaq
Dim am As New AtencionBManager

Dim rs
'siempre va a cargar una sola
am.Sql = "select * from AtencionB where id=" + Str(pId)
Set rs = BD.SelectSQL(am.Sql)

Cargar rs, am
Set LoadById = am.Item(pId)
Set am = Nothing

Exit Function
errMaq:
ErrorLog "AtencionBManager", "LoadById ", Err.Description

End Function

Public Function GetByEstado(pEstado As eEstadoAtencion) As AtencionBManager
On Error GoTo errMaq
Dim am As New AtencionBManager

Dim rs

am.Sql = "select * from AtencionB where estado=" + Str(pEstado) 'esto es para el metodo reload
Set rs = BD.SelectSQL(am.Sql)

Cargar rs, am
Set GetByEstado = am
Exit Function
errMaq:
ErrorLog "AtencionBManager", "GetByEstado ", Err.Description

End Function

Public Function GetByDate(pDate As Date) As AtencionBManager
On Error GoTo errMaq
Dim am As New AtencionBManager

Dim rs

am.Sql = "select * from AtencionB where fecha=" + Str(pDate) 'esto es para el metodo reload

Set rs = BD.SelectSQL(am.Sql)

Cargar rs, am
Set GetByDate = am
Exit Function
errMaq:
ErrorLog "AtencionBManager", "GetByDate ", Err.Description
End Function

'Public Function GetByTipo(pTipo As eDestinoAtencionB, pId As Long) As AtencionBManager
''las cargo cada vez para q esten actualizadas, revisar
''TERMINAR
'On Error GoTo errMaq
'Dim am As New AtencionBManager
'
'Dim campoID As String
'campoID = Choose(pTipo, "idAfiliadoPropio", "idAreaProtegida", "idObraSocial", "idServicioEmergencia", "idAfiliadoExterno")
'
'Dim rs
'am.Sql = "select * from AtencionB where " + campoID + "=" + Str(pId)
'Set rs = BD.Selectsql(am.Sql)
'
'Cargar rs, am
'Set GetByTipo = am
'Exit Function
'errMaq:
'ErrorLog "Error AtencionBManager, sub GetByTipo " + Err.Description
'
'End Function

Private Sub Cargar(rs, pAtenciones As AtencionBManager)
Dim at As AtencionB

While Not rs.EOF
        
    Set at = Add(rs!Id, rs!NroIncidente, EmpleadosLocal.Item(rs!idDespachador), rs!Fecha, rs!HoraLlamada, "" & rs!SalidaPreInspeccion, "" & rs!LlegadaPreInspeccion, "" & rs!SalidaDotacion, "" & rs!QTH, "" & rs!VL, "" & rs!Observaciones, SintomasLocal.Item(rs!idSintoma), rs!Estado)
    pAtenciones.AddItem at
    'lo pongo aca para no modificar la funcion add
    at.nroIncidenteInterno = rs!nroIncidenteInterno & ""
    at.PoliciaACargo = "" & rs!PoliciaACargo
    at.PoliciaMovil = "" & rs!PoliciaMovil
    at.SEMedico = "" & rs!SEMedico
    at.SECentroAsistencial = "" & rs!SECentroAsistencial
    at.SEMedicoMP = "" & rs!SEMedicoMP
    at.SEAmbulancias = IIf(rs!SEAmbulancias = 1, True, False)
    at.AccesoPor = "" & rs!AccesoPor
    at.Aseguradora = "" & rs!Aseguradora
    at.Cronica = "" & rs!Cronica
    at.DescripcionMaterial = "" & rs!DescripcionMaterial
    at.DescripcionPerjuiciosCampo = "" & rs!DescripcionPerjuiciosCampo
    at.DescripcionPerjuiciosVivienda = "" & rs!DescripcionPerjuiciosVivienda
    at.EquiposEspeciales = "" & rs!EquiposEspeciales
    at.MaterialesCombustibles = "" & rs!MaterialesCombustibles
    at.Poliza = "" & rs!Poliza
    at.Referencias = "" & rs!Referencias
      
    If Not IsNull(rs!PoliciaCantidad) Then
        at.PoliciaCantidad = rs!PoliciaCantidad
    End If
    If Not IsNull(rs!AmbientesAfectadosVivienda) Then
        at.AmbientesAfectadosVivienda = rs!AmbientesAfectadosVivienda
    End If
    If Not IsNull(rs!HectareasAfectadasCampo) Then
         at.HectareasAfectadasCampo = rs!HectareasAfectadasCampo
    End If
    
    If Not IsNull(rs!idInstalacionElectrica) Then
         Set at.InstalacionElectrica = InstElectricasLocal.Item(rs!idInstalacionElectrica)
    End If
    
    If Not IsNull(rs!idInstalacionGas) Then
         Set at.InstalacionGas = InstalacionesGasLocal.Item(rs!idInstalacionGas)
    End If
    
    Set at.Direccion = New Direccion
    at.Direccion.Cargar rs!idDireccion
           
'    If Not IsNull(rs!idAfiliadoPropio) Then
'        Set at.Afiliado = AfiliadosLocal.Item(rs!idAfiliadoPropio)
'    End If
        
    rs.MoveNext
    
Wend
End Sub

'aca cargo solamente los campos comunes a todas las AtencionBes
Private Function Add(pId As Long, pNroIncidente As String, pDespachador As Empleado, pFecha As Date, pHora As String, pSalidaPreInspeccion As String, pLlegadaPreInspeccion As String, pSalidaDotacion As String, pQTH As String, pVL As String, pObservaciones As String, pSintoma As Sintoma, pEstado As eEstadoAtencion) As AtencionB
     Dim at As AtencionB
     Set at = New AtencionB
     
     at.Id = pId
     at.NroIncidente = pNroIncidente
     Set at.Despachador = pDespachador
     at.Observaciones = pObservaciones
     
     at.Fecha = pFecha
     at.HoraLlamada = pHora
     at.SalidaPreInspeccion = pSalidaPreInspeccion
     at.SalidaDotacion = pSalidaDotacion
     at.LlegadaPreInspeccion = pLlegadaPreInspeccion
     at.QTH = pQTH
     at.VL = pVL
     Set at.Sintoma = pSintoma
     at.Estado = pEstado
     'Set at.Telefono = pTelefono
     
     AddItem at
     Set Add = at
    
End Function

Public Function Filter(pFechaDesde As String, pFechaHasta As String, Optional pMes As Integer = 0, Optional pYear As Integer = 0, Optional pDespachador As Empleado, Optional pCodigoE As CodigoEmergencia) As AtencionBManager
On Error GoTo errMaq

Dim Sql As String
Dim criterios As String

'filtro por fecha
If pFechaDesde <> "" And pFechaHasta <> "" Then
    criterios = " AND fecha BETWEEN #" + pFechaDesde + "# AND #" + pFechaHasta + "#"
ElseIf pMes <> 0 Then
    criterios = " AND month(fecha)=" + Str(pMes)
ElseIf pYear <> 0 Then
    criterios = " AND year(fecha)=" + Str(pYear)
End If

'filtro por despachador
If Not pDespachador Is Nothing Then
    criterios = criterios + " AND idDespachador=" + Str(pDespachador.Id)
End If

'filtro por codigo de emergencia,no lo tengo directamente, me fijo en el sintoma...
If Not pCodigoE Is Nothing Then
    criterios = criterios + " AND idsintoma BETWEEN " + Str(pCodigoE.Id * 100) + " AND " + Str((pCodigoE.Id * 100) + 100)
End If

'controlo para ubicar bien el where y el and de los criterios segun sea necesario
If criterios <> "" Then
    criterios = Right(criterios, Len(criterios) - 4)
    criterios = " WHERE " + criterios
End If

'armo la consulta
mSQL = "SELECT * FROM AtencionB " + criterios

Dim am As New AtencionBManager

Dim rs
am.Sql = mSQL
Set rs = BD.SelectSQL(mSQL)

Cargar rs, am
Set Filter = am

Exit Function
errMaq:
ErrorLog "AtencionBManager", "filter ", Err.Description
End Function

Public Function GetAsignadas() As AtencionBManager
Dim at As AtencionB
Dim ats As AtencionBManager
Set ats = New AtencionBManager

For Each at In Me
    If at.Equipos.Count <> 0 And at.Estado = ePendiente Then ats.AddItem at
Next
Set GetAsignadas = ats
End Function

Public Function GetSinAsignar() As AtencionBManager
Dim at As AtencionB
Dim ats As AtencionBManager
Set ats = New AtencionBManager

For Each at In Me
    If at.Equipos.Count = 0 And at.Estado = ePendiente Then ats.AddItem at
Next
Set GetSinAsignar = ats
End Function

'Public Function GetListasParaCerrar()
'Dim at As AtencionB
'Dim ats As AtencionBManager
'Set ats = New AtencionBManager
'
'For Each at In Me
'    If at.Estado = eListaParaCerrar Then ats.AddItem at
'Next
'Set GetListasParaCerrar = ats
'End Function

Private Sub mNetMonitor_AtencionesChanged()
    'este es el unico evento q me interesa aca
    Me.Reload
End Sub

Public Function CantidadAtenciones() As Integer
    Dim a As AtencionB
    Dim cont As Integer
    
    For Each a In Me
        If DateDiff("m", a.Fecha, Date) = 0 Then cont = cont + 1
    Next
    CantidadAtenciones = cont
End Function

