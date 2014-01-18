VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuotaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A803E20196"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit
Public Event HasChanged()

Private mCol As Collection
Private mSQL As String

Private WithEvents mNetMonitor As NetMonitor
Attribute mNetMonitor.VB_VarHelpID = -1

Friend Property Let Sql(pSql As String)
    mSQL = pSql
End Property

'##ModelId=46A80BDA01D4
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("CuotaManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A80BE60157
Public Sub AddItem(pCuota As Cuota)
    On Error GoTo AddItemErr

    mCol.Add pCuota, "k" + Str(pCuota.Id)

    Exit Sub
AddItemErr:
    ErrorLog "CuotaManager", "AddItem", err.Description
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

Public Sub Reload()
    If mSQL <> "" Then
        Dim rs
        Set rs = BD.SelectSQL(mSQL)
        Set mCol = Nothing
        Set mCol = New Collection
        Cargar rs, Me
    End If
    RaiseEvent HasChanged
End Sub

Friend Sub CargarCuotasByAfiliado(pAfiliado As Afiliado)
On Error GoTo errMaq
Dim rs
mSQL = "select * from cuota where idAfiliado=" + Str(pAfiliado.Id)
Set rs = BD.SelectSQL(mSQL)
Cargar rs, Me
Exit Sub
errMaq:
ErrorLog "CuotasManager", "cargarTodos ", err.Description
End Sub

Friend Sub CargarCuotasByAreaProtegida(pAreaProtegida As AreaProtegida)
On Error GoTo errMaq
Dim c As Cuota
Dim rs
mSQL = "select * from cuota where idAreaProtegida=" + Str(pAreaProtegida.Id)
Set rs = BD.SelectSQL(mSQL)
Cargar rs, Me
Exit Sub
errMaq:
ErrorLog "CuotasManager", "cargarTodos ", err.Description
End Sub

Friend Sub CargarCuotasByEstado(pEstado As eEstadoCuota)
On Error GoTo errMaq
Dim rs
If pEstado = eAnulado Or pEstado = ePedirRecibo Then
    mSQL = "select * from cuotasanuladas where estado=" + Str(pEstado)
Else
    mSQL = "select * from cuota where estado=" + Str(pEstado)
End If

Set rs = BD.SelectSQL(mSQL)
Cargar rs, Me

Exit Sub
errMaq:
ErrorLog "CuotasManager", "cargarTodos ", err.Description
End Sub

Private Sub Cargar(rs, cm As CuotaManager)
Dim c As Cuota

While Not rs.EOF
        
    Set c = cm.Add(rs!Id, Nothing, Nothing, rs!NroRecibo, rs!Mes, rs!mYear, rs!Monto, EmpleadosLocal.Item(rs!idEmisor), EmpleadosLocal.Item(rs!idCobrador), rs!Estado, rs!FechaEmision)
    If c.Estado = ePagado Then c.FechaPago = rs!FechaPago
    If Not IsNull(rs!idAfiliado) Then Set c.Afiliado = AfiliadosLocal.Item(rs!idAfiliado)
    If Not IsNull(rs!idAreaProtegida) Then Set c.AreaProtegida = AreasProtegidasLocal.Item(rs!idAreaProtegida)
    If c.Estado = eAnulado Or c.Estado = ePedirRecibo Then
        'solo si son cuotas anuladas o para pedir recibo
        If Not IsNull(rs!idanuladopor) Then Set c.AnuladoPor = EmpleadosLocal.Item(rs!idanuladopor)
        If Not IsNull(rs!iddevolucionregistradapor) Then Set c.DevolucionRegistradaPor = EmpleadosLocal.Item(rs!iddevolucionregistradapor)
    ElseIf c.Estado = ePagado Then If Not IsNull(rs!FechaPago) Then c.FechaPago = rs!FechaPago
    End If
    rs.MoveNext
    
Wend

End Sub

Public Function GetCuotasByEstado(pEstado As eEstadoCuota) As CuotaManager
    Dim c As Cuota
    Dim cm As New CuotaManager
    
    For Each c In Me
        If c.Estado = eImpaga Then cm.AddItem c
    Next
    Set GetCuotasByEstado = cm
    
End Function

Public Function GetCuotasByCobrador(pCobrador As Empleado) As CuotaManager
    Dim c As Cuota
    Dim cm As New CuotaManager
    
    For Each c In Me
        If c.Cobrador.Id = pCobrador.Id Then cm.AddItem c
    Next
    Set GetCuotasByCobrador = cm
    
End Function

Friend Function Add(pId As Long, pAfiliado As Afiliado, pAreaProtegida As AreaProtegida, pNroRecibo As Long, pMes As Integer, pYear As Integer, pMonto As Currency, pEmisor As Empleado, pCobrador As Empleado, pEstado As eEstadoCuota, pFechaEmision As Date) As Cuota
    Dim c As New Cuota
    Set c.Afiliado = pAfiliado
    Set c.AreaProtegida = pAreaProtegida
    c.aYear = pYear
    
    Set c.Emisor = pEmisor
    Set c.Cobrador = pCobrador
    c.Estado = pEstado
    c.FechaEmision = pFechaEmision
    c.Id = pId
    c.Mes = pMes
    c.Monto = pMonto
    c.NroRecibo = pNroRecibo
    
    Set Add = c
    AddItem c
    
End Function

Public Function Nuevo(pAfiliado As Afiliado, pAreaProtegida As AreaProtegida, pNroRecibo As Long, pMes As Integer, pYear As Integer, pMonto As Currency, pEmisor As Empleado) As Cuota
    If ItemByPeriodo(pMes, pYear) Is Nothing Then
        Dim c As New Cuota
        Set c.Afiliado = pAfiliado
        Set c.AreaProtegida = pAreaProtegida
        c.aYear = pYear
        If Not c.Afiliado Is Nothing Then
            Set c.Cobrador = c.Afiliado.Cobrador
        ElseIf Not c.AreaProtegida Is Nothing Then
            Set c.Cobrador = c.AreaProtegida.Cobrador
        Else
            ErrorLog "CuotaManager", "Nuevo", "No se establecio un Afiliado o un Area Protegida."
        End If
        Set c.Emisor = pEmisor
        c.Estado = eImpaga
        c.FechaEmision = Date
        'c.FechaPago null hasta el momento
        'c.id
        c.Mes = pMes
        c.Monto = pMonto
        c.NroRecibo = pNroRecibo
        c.Guardar
        Set Nuevo = c
        AddItem c
    End If
End Function

Public Property Get ItemByPeriodo(pMes As Integer, pYear As Integer) As Cuota
    Dim c As Cuota
    For Each c In Me
        If c.Mes = pMes And c.aYear = pYear Then
            Set ItemByPeriodo = c
            Exit For
        End If
    Next
End Property

Public Function GetUltimoNroRecibo() As Long
'    Dim rs
'    Dim aux As Long
'    Dim aux2 As Long
'    Set rs = BD.Selectsql("Select max(nrorecibo) from Cuota")
'    aux = CLng(IIf(IsNull(rs.fields(0).Value), 0, rs.fields(0).Value))
'    Set rs = BD.Selectsql("Select max(nrorecibo) from CuotasAnuladas")
'    aux2 = CLng(IIf(IsNull(rs.fields(0).Value), 0, rs.fields(0).Value))
'    Set rs = Nothing
'    GetUltimoNroRecibo = IIf(aux < aux2, aux2, aux)
    GetUltimoNroRecibo = BD.getIdDisponible(tRecibo)
End Function

'son para configurar y consultar el numero de recibo sin q se incremente el contador
Public Function GetNumeroDeReciboActual() As Long
On Error GoTo errman
Dim rs
Set rs = BD.SelectSQL("select id from IDRecibo")
GetNumeroDeReciboActual = rs!Id
Exit Function
errman:
GetNumeroDeReciboActual = 0
End Function

Public Sub SetNumeroDeReciboActual(pRecibo As Long)
BD.SelectSQL ("update IDRecibo set id=" + Str(pRecibo))
End Sub

Private Sub mNetMonitor_CuotasChanged()
    Reload
End Sub

Public Sub CuotasModificadasoAgregadas()
    mNetMonitor.MiRedLocal.EnviarEventoABM Alta, "cuota", 1
End Sub
