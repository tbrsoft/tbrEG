VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AtencionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A803D50196"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit
Public Event HasChanged()

Public Enum eDestinoAtencion
    eAfiliadoPropio = 1
    eAreaProtegida = 2
    eObraSocial = 3
    eServicioEmergencia = 4
    eAfiliadoExterno = 5
End Enum

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

Friend Sub AtencionAgregadaoModificada()
    'prueba
    mNetMonitor.MiRedLocal.EnviarEventoABM Alta, "atencion", 1
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
Public Sub AddItem(pAtencion As Atencion)
    On Error GoTo AddItemErr

    mCol.Add pAtencion, "k" + Str(pAtencion.Id)

    Exit Sub
AddItemErr:
If Err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "AtencionManager", "AddItem", Err.Description
End If
End Sub

Public Sub Reload()
    If mSQL <> "" Then 'esto es porq atencioneslocal no tiene q recargarse...
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

Public Property Get Item(pId As Long) As Atencion
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Function LoadById(pId As Long) As Atencion
On Error GoTo errMaq
Dim am As New AtencionManager

Dim rs
'siempre va a cargar una sola
am.Sql = "select * from atencion where id=" + Str(pId)
Set rs = BD.SelectSQL(am.Sql)

Cargar rs, am
Set LoadById = am.Item(pId)
Set am = Nothing

Exit Function
errMaq:
ErrorLog "AtencionManager", "LoadById", Err.Description

End Function

Public Function GetByEstado(pEstado As eEstadoAtencion) As AtencionManager
    
    TERR.Anotar "atMng001"
    On Error GoTo errMaq
    Dim am As New AtencionManager
    
    Dim rs
    
    am.Sql = "select * from atencion where estado=" + Str(pEstado) 'esto es para el metodo reload
    Set rs = BD.SelectSQL(am.Sql)
    
    Cargar rs, am
    
    Set GetByEstado = am
    Exit Function
errMaq:
    ErrorLog "AtencionManager", "GetByEstado ", Err.Description
    TERR.AppendLog "at.mng.GbSt", TERR.ErrToTXT(Err)

End Function

Public Function GetByDate(pDate As Date) As AtencionManager
On Error GoTo errMaq
Dim am As New AtencionManager

Dim rs

am.Sql = "select * from atencion where fecha=" + Str(pDate) 'esto es para el metodo reload

Set rs = BD.SelectSQL(am.Sql)

Cargar rs, am
Set GetByDate = am
Exit Function
errMaq:
ErrorLog "AtencionManager", "GetByDate ", Err.Description
End Function

'estan mal puestos los nombres, deberia ser cargar por tipo no get
Public Function GetByTipo(pTipo As eDestinoAtencion, pId As Long) As AtencionManager
'las cargo cada vez para q esten actualizadas, revisar
'TERMINAR
On Error GoTo errMaq
Dim am As New AtencionManager

Dim campoID As String
campoID = Choose(pTipo, "idAfiliadoPropio", "idAreaProtegida", "idObraSocial", "idServicioEmergencia", "idAfiliadoExterno")

Dim rs
am.Sql = "select * from atencion where " + campoID + "=" + Str(pId)
Set rs = BD.SelectSQL(am.Sql)

Cargar rs, am
Set GetByTipo = am
Exit Function
errMaq:
ErrorLog "AtencionManager", "GetByTipo ", Err.Description

End Function

Private Sub Cargar(rs, pAtenciones As AtencionManager)
    On Local Error GoTo ErrCRG
    
    TERR.Anotar "at.mng.crg001"
    Dim at As Atencion
    
    While Not rs.EOF
        TERR.Anotar "at.mng.crg002", NN_num(rs!Id), NN_num(rs!NroIncidente)
        TERR.Anotar "at.mng.crg003", NN_num(rs!idDespachador), NN_str(rs!Fecha), NN_str(rs!HoraLlamada)
        TERR.Anotar "at.mng.crg004", NN_str(rs!QTH), NN_str(rs!VL), NN_str(rs!Diagnostico), NN_str(rs!Observaciones)
        TERR.Anotar "at.mng.crg005", NN_num(rs!idSintoma), NN_str(rs!Estado)
            
        Set at = Add(NN_num(rs!Id), NN_num(rs!NroIncidente), _
            EmpleadosLocal.Item(NN_num(rs!idDespachador)), _
            rs!Fecha, rs!HoraLlamada, _
            NN_str(rs!QTH), _
            NN_str(rs!VL), _
            NN_str(rs!Diagnostico), _
            NN_str(rs!Observaciones), _
            SintomasLocal.Item(NN_num(rs!idSintoma)), _
            NN_str(rs!Estado))
        
        TERR.Anotar "at.mng.crg007"
        pAtenciones.AddItem at
        'lo pongo aca para no modificar la funcion add
        at.nroIncidenteInterno = rs!nroIncidenteInterno & ""
        TERR.Anotar "at.mng.crg006"
        
        If at.Sintoma.Parent.Id = 100 Then 'es un traslado
            TERR.Anotar "at.mng.crg008"
            Set at.DireccionDestino = New Direccion
            Set at.DireccionOrigen = New Direccion
            at.DireccionDestino.Cargar rs!idDireccionDestino
            at.DireccionOrigen.Cargar rs!idDireccionOrigen
        Else
            TERR.Anotar "at.mng.crg009"
            Set at.Telefono = TelefonosLocal.Item(NN_num(rs!idTelefono))
            'si no tiene telefono cargo el auxiliar q esta si o si
            If at.Telefono Is Nothing Then at.TelefonoAuxilar = NN_num(rs!TelefonoAuxiliar)
            Set at.Direccion = New Direccion
            Dim idDir As Long
            idDir = NN_num(rs!idDireccion)
            TERR.Anotar "at.mng.crg009b", idDir
            If idDir = 0 Then
                TERR.AppendLog "Dir No valida!"
            Else
                at.Direccion.Cargar NN_num(rs!idDireccion)
            End If
            TERR.Anotar "at.mng.crg009c"
            at.Operador = NN_str(rs!Operador)
        End If
        
        TERR.Anotar "at.mng.crg010"
        If Not IsNull(rs!idAfiliadoPropio) Then
            TERR.Anotar "at.mng.crg011"
            Set at.Afiliado = AfiliadosLocal.Item(rs!idAfiliadoPropio)
        ElseIf Not IsNull(rs!idAreaProtegida) Then
            TERR.Anotar "at.mng.crg012"
            Set at.AreaProtegida = AreasProtegidasLocal.Item(rs!idAreaProtegida)
            If Not IsNull(rs!idAfiliadoExterno) Then
                TERR.Anotar "at.mng.crg013"
                Set at.AfiliadoExterno = at.AreaProtegida.Afiliados.Item(rs!idAfiliadoExterno)
            End If
        ElseIf Not IsNull(rs!idObraSocial) Then
            TERR.Anotar "at.mng.crg014"
            Set at.ObraSocial = ObrasSocialesLocal.Item(rs!idObraSocial)
            If Not IsNull(rs!idAfiliadoExterno) Then
                TERR.Anotar "at.mng.crg015"
                Set at.AfiliadoExterno = at.ObraSocial.Afiliados.Item(rs!idAfiliadoExterno)
            End If
        ElseIf Not IsNull(rs!idServicioEmergencia) Then
            TERR.Anotar "at.mng.crg016"
            Set at.ServicioEmergencia = ServiciosEmergenciaLocal.Item(rs!idServicioEmergencia)
            TERR.Anotar "at.mng.crg017"
            If Not IsNull(rs!idAfiliadoExterno) Then
                TERR.Anotar "at.mng.crg018"
                Set at.AfiliadoExterno = at.ServicioEmergencia.Afiliados.Item(rs!idAfiliadoExterno)
            End If
        End If
        
        TERR.Anotar "at.mng.crg019"
        If Not IsNull(rs!idTipoCodigo) Then
            TERR.Anotar "at.mng.crg020"
            Set at.TipoCodigo = TiposCodigoLocal.Item(rs!idTipoCodigo)
        End If
        TERR.Anotar "at.mng.crg021"
        
        'aca cargo la infocontable
        Dim infoC As New InfoContable
        infoC.Constructor rs!Servicio, rs!copago, rs!abonado, rs!condIva, rs!Id, rs!MontoIVA
        Set at.InfoContable = infoC
        TERR.Anotar "at.mng.crg022"
        rs.MoveNext
        
    Wend
    
    TERR.Anotar "at.mng.crg023"
    
    Exit Sub
ErrCRG:
    TERR.AppendLog "at.mng.FAIL", TERR.ErrToTXT(Err)

End Sub

'aca cargo solamente los campos comunes a todas las atenciones
Private Function Add(pId As Long, pNroIncidente As String, pDespachador As Empleado, pFecha As Date, pHora As String, pQTH As String, pVL As String, pDiagnostico As String, pObservaciones As String, pSintoma As Sintoma, pEstado As eEstadoAtencion) As Atencion
     Dim at As Atencion
     Set at = New Atencion
     
     at.Id = pId
     at.NroIncidente = pNroIncidente
     Set at.Despachador = pDespachador
     at.Diagnostico = pDiagnostico
     at.Observaciones = pObservaciones
     
     at.Fecha = pFecha
     at.HoraLlamada = pHora
     at.QTH = pQTH
     at.VL = pVL
     Set at.Sintoma = pSintoma
     at.Estado = pEstado
     'Set at.Telefono = pTelefono
     
     AddItem at
     Set Add = at
    
End Function

Public Function Filter(pAfiliado As Afiliado, pServicioEmergencia As ServicioEmergencia, pObraSocial As ObraSocial, pAreaProtegida As AreaProtegida, pFechaDesde As String, pFechaHasta As String, Optional pMes As Integer = 0, Optional pYear As Integer = 0, Optional pDespachador As Empleado, Optional pCodigoE As CodigoEmergencia) As AtencionManager
On Error GoTo errMaq

Dim idDestino As Long
Dim campoID As String
Dim Sql As String
Dim criterios As String
'aca armo para filtrar por destino de la atencion
If Not pAfiliado Is Nothing Then
    idDestino = pAfiliado.Id
    campoID = "idAfiliadoPropio"
ElseIf Not pServicioEmergencia Is Nothing Then
    idDestino = pServicioEmergencia.Id
    campoID = "idServicioEmergencia"
ElseIf Not pObraSocial Is Nothing Then
    idDestino = pObraSocial.Id
    campoID = "idObraSocial"
ElseIf Not pAreaProtegida Is Nothing Then
    idDestino = pAreaProtegida.Id
    campoID = "idAreaProtegida"
End If

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
If criterios <> "" And campoID = "" Then
    criterios = Right(criterios, Len(criterios) - 4)
    criterios = " WHERE " + criterios
End If

'armo la consulta
mSQL = "SELECT * FROM ATENCION " + IIf(campoID <> "", "WHERE " + campoID + " = " + Str(idDestino), "") + criterios

Dim am As New AtencionManager

Dim rs
am.Sql = mSQL
Set rs = BD.SelectSQL(mSQL)

Cargar rs, am
Set Filter = am

Exit Function
errMaq:
ErrorLog "AtencionManager", "filter ", Err.Description
End Function

Public Function GetAsignadas() As AtencionManager
Dim at As Atencion
Dim ats As AtencionManager
Set ats = New AtencionManager

For Each at In Me
    If at.Equipos.Count <> 0 And at.Estado = ePendiente Then ats.AddItem at
Next
Set GetAsignadas = ats
End Function

Public Function GetSinAsignar() As AtencionManager
Dim at As Atencion
Dim ats As AtencionManager
Set ats = New AtencionManager

For Each at In Me
    If at.Equipos.Count = 0 And at.Estado = ePendiente Then ats.AddItem at
Next
Set GetSinAsignar = ats
End Function

'Public Function GetListasParaCerrar()
'Dim at As Atencion
'Dim ats As AtencionManager
'Set ats = New AtencionManager
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
    Dim a As Atencion
    Dim cont As Integer
    
    For Each a In Me
        If DateDiff("m", a.Fecha, Date) = 0 Then cont = cont + 1
    Next
    CantidadAtenciones = cont
End Function
