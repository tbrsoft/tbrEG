VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cuota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FEAB005D"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Enum eEstadoCuota
    eImpaga = 1
    ePagado = 2
    eAnulado = 3 'cuando el cobrador ya me devolvio el recibo
    ePedirRecibo = 4 'para marcar q el cobrador me debe devolver el recibo
End Enum

Private mId As Long
Private mNroRecibo As Long
Private mMonth As Integer
Private mYear As Integer
Private mMonto As Currency
Private mFechaPago As Date
Private mFechaEmision As Date
Private mFechaAnulado As Date
Private mAfiliado As Afiliado
Private mAreaProtegida As AreaProtegida
Private mCobrador As Empleado
Private mEmisor As Empleado
Private mCobroRegistradoPor As Empleado
Private mAnuladoPor As Empleado
Private mDevolucionRegistradaPor As Empleado

Private mEstado As eEstadoCuota

Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Cuota", "Id Property Get", err.Description)
End Property

Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Cuota", "Id Property Let", err.Description)
End Property

Public Property Get NroRecibo() As Long
    On Error GoTo NroReciboErr

    '## Generated default body ...
    NroRecibo = mNroRecibo

    Exit Property
NroReciboErr:
    Call ErrorLog("Cuota", "NroRecibo Property Get", err.Description)
End Property

Friend Property Let NroRecibo(ByVal vNewValue As Long)
    On Error GoTo NroReciboErr

    '## Generated default body ...
    mNroRecibo = vNewValue

    Exit Property
NroReciboErr:
    Call ErrorLog("Cuota", "NroRecibo Property Let", err.Description)
End Property
'##ModelId=46A977B20330
Public Property Get Monto() As Currency
    On Error GoTo MontoErr

    '## Generated default body ...
    Monto = mMonto

    Exit Property
MontoErr:
    Call ErrorLog("Cuota", "Monto Property Get", err.Description)
End Property

'##ModelId=46A977B20331
Public Property Let Monto(ByVal vNewValue As Currency)
    On Error GoTo MontoErr

    '## Generated default body ...
    mMonto = vNewValue

    Exit Property
MontoErr:
    Call ErrorLog("Cuota", "Monto Property Let", err.Description)
End Property

'##ModelId=46A977B20335
Public Property Get FechaPago() As Date
    On Error GoTo FechaPagoErr

    '## Generated default body ...
    FechaPago = mFechaPago

    Exit Property
FechaPagoErr:
    Call ErrorLog("Cuota", "FechaPago Property Get", err.Description)
End Property

'##ModelId=46A977B20336
Public Property Let FechaPago(ByVal vNewValue As Date)
    On Error GoTo FechaPagoErr

    '## Generated default body ...
    mFechaPago = vNewValue

    Exit Property
FechaPagoErr:
    Call ErrorLog("Cuota", "FechaPago Property Let", err.Description)
End Property

Public Property Get FechaAnulado() As Date
    On Error GoTo FechaAnuladoErr

    FechaAnulado = mFechaAnulado

    Exit Property
FechaAnuladoErr:
    Call ErrorLog("Cuota", "FechaAnulado Property Get", err.Description)
End Property

Public Property Let FechaAnulado(ByVal vNewValue As Date)
    On Error GoTo FechaAnuladoErr

    mFechaAnulado = vNewValue

    Exit Property
FechaAnuladoErr:
    Call ErrorLog("Cuota", "FechaAnulado Property Let", err.Description)
End Property

Public Property Get Mes() As Integer
    On Error GoTo MesErr

    '## Generated default body ...
    Mes = mMonth

    Exit Property
MesErr:
    Call ErrorLog("Cuota", "Mes Property Get", err.Description)
End Property

Public Property Let Mes(ByVal vNewValue As Integer)
    On Error GoTo MesErr

    '## Generated default body ...
    mMonth = vNewValue

    Exit Property
MesErr:
    Call ErrorLog("Cuota", "Mes Property Let", err.Description)
End Property

Public Property Get aYear() As Integer
    On Error GoTo aYearErr

    '## Generated default body ...
    aYear = mYear

    Exit Property
aYearErr:
    Call ErrorLog("Cuota", "aYear Property Get", err.Description)
End Property

Public Property Let aYear(ByVal vNewValue As Integer)
    On Error GoTo aYearErr

    '## Generated default body ...
    mYear = vNewValue

    Exit Property
aYearErr:
    Call ErrorLog("Cuota", "aYear Property Let", err.Description)
End Property

'##ModelId=46A977B20335
Public Property Get FechaEmision() As Date
    On Error GoTo FechaEmisionErr

    '## Generated default body ...
    FechaEmision = mFechaEmision

    Exit Property
FechaEmisionErr:
    Call ErrorLog("Cuota", "FechaEmision Property Get", err.Description)
End Property

'##ModelId=46A977B20336
Public Property Let FechaEmision(ByVal vNewValue As Date)
    On Error GoTo FechaEmisionErr

    '## Generated default body ...
    mFechaEmision = vNewValue

    Exit Property
FechaEmisionErr:
    Call ErrorLog("Cuota", "FechaEmision Property Let", err.Description)
End Property

Public Property Get Afiliado() As Afiliado
    On Error GoTo AfiliadoErr

    Set Afiliado = mAfiliado

    Exit Property
AfiliadoErr:
    Call ErrorLog("Cuota", "Afiliado Property Get", err.Description)
End Property

Public Property Set Afiliado(ByVal vNewValue As Afiliado)
    On Error GoTo AfiliadoErr

    Set mAfiliado = vNewValue

    Exit Property
AfiliadoErr:
    Call ErrorLog("Cuota", "Afiliado Property Set", err.Description)
End Property

Public Property Get AreaProtegida() As AreaProtegida
    On Error GoTo AreaProtegidaErr

    Set AreaProtegida = mAreaProtegida

    Exit Property
AreaProtegidaErr:
    Call ErrorLog("Cuota", "AreaProtegida Property Get", err.Description)
End Property

Public Property Set AreaProtegida(ByVal vNewValue As AreaProtegida)
    On Error GoTo AreaProtegidaErr

    Set mAreaProtegida = vNewValue

    Exit Property
AreaProtegidaErr:
    Call ErrorLog("Cuota", "AreaProtegida Property Set", err.Description)
End Property

Public Property Get Cobrador() As Empleado
    On Error GoTo CobradorErr

    Set Cobrador = mCobrador

    Exit Property
CobradorErr:
    Call ErrorLog("Cuota", "Cobrador Property Get", err.Description)
End Property

Public Property Set Cobrador(ByVal vNewValue As Empleado)
    On Error GoTo CobradorErr

    Set mCobrador = vNewValue

    Exit Property
CobradorErr:
    Call ErrorLog("Cuota", "Cobrador Property Set", err.Description)
End Property

Public Property Get CobroRegistradoPor() As Empleado
    On Error GoTo CobradorErr

    Set CobroRegistradoPor = mCobroRegistradoPor

    Exit Property
CobradorErr:
    Call ErrorLog("Cuota", "Cobrador Property Get", err.Description)
End Property

Public Property Set DevolucionRegistradaPor(ByVal vNewValue As Empleado)
    On Error GoTo CobradorErr

    Set mDevolucionRegistradaPor = vNewValue

    Exit Property
CobradorErr:
    Call ErrorLog("Cuota", "Cobrador Property Set", err.Description)
End Property

Public Property Get DevolucionRegistradaPor() As Empleado
    On Error GoTo CobradorErr

    Set DevolucionRegistradaPor = mDevolucionRegistradaPor

    Exit Property
CobradorErr:
    Call ErrorLog("Cuota", "Cobrador Property Get", err.Description)
End Property

Public Property Set CobroRegistradoPor(ByVal vNewValue As Empleado)
    On Error GoTo CobradorErr

    Set mCobroRegistradoPor = vNewValue

    Exit Property
CobradorErr:
    Call ErrorLog("Cuota", "Cobrador Property Set", err.Description)
End Property
Public Property Get Emisor() As Empleado
    On Error GoTo EmisorErr

    Set Emisor = mEmisor

    Exit Property
EmisorErr:
    Call ErrorLog("Cuota", "Emisor Property Get", err.Description)
End Property

Public Property Set Emisor(ByVal vNewValue As Empleado)
    On Error GoTo EmisorErr

    Set mEmisor = vNewValue

    Exit Property
EmisorErr:
    Call ErrorLog("Cuota", "Emisor Property Set", err.Description)
End Property

Public Property Get AnuladoPor() As Empleado
    On Error GoTo AnuladoPorErr

    Set AnuladoPor = mAnuladoPor

    Exit Property
AnuladoPorErr:
    Call ErrorLog("Cuota", "AnuladoPor Property Get", err.Description)
End Property

Public Property Set AnuladoPor(ByVal vNewValue As Empleado)
    On Error GoTo AnuladoPorErr

    Set mAnuladoPor = vNewValue

    Exit Property
AnuladoPorErr:
    Call ErrorLog("Cuota", "AnuladoPor Property Set", err.Description)
End Property
Public Property Get Estado() As eEstadoCuota
    On Error GoTo EstadoErr

    Estado = mEstado

    Exit Property
EstadoErr:
    Call ErrorLog("Cuota", "Estado Property Get", err.Description)
End Property

Public Property Let Estado(ByVal vNewValue As eEstadoCuota)
    On Error GoTo EstadoErr

    mEstado = vNewValue

    Exit Property
EstadoErr:
    Call ErrorLog("Cuota", "Estado Property Set", err.Description)
End Property

Public Function EstadoToString() As String
    On Error GoTo EstadoErr

    EstadoToString = Choose(mEstado, "Pendiente", "Pagado", "Anulado", "PedirRecibo")
   
    Exit Function
EstadoErr:
    Call ErrorLog("Cuota", "EstadoToString Function", err.Description)
End Function

Public Function GetProperty(Nombre As String) As String
    Dim aux As String
    aux = LCase(Nombre)
    Dim APoAf As Object 'lo uso para cuando tienen las mismas propiedades
    If Not Afiliado Is Nothing Then
        Set APoAf = Afiliado
    ElseIf Not AreaProtegida Is Nothing Then
        Set APoAf = AreaProtegida
    End If
    
    Select Case aux
        Case "nombre"
            If Not Afiliado Is Nothing Then
                GetProperty = Afiliado.NombreCompleto
            ElseIf Not AreaProtegida Is Nothing Then
                GetProperty = AreaProtegida.NombreArea
            End If
        Case "nro"
            If Not Afiliado Is Nothing Then
                GetProperty = Afiliado.IdCompleto
            ElseIf Not AreaProtegida Is Nothing Then
                GetProperty = AreaProtegida.Id
            End If
        Case "cobrador"
            GetProperty = Cobrador.NombreCompleto
        Case "ciudad"
            GetProperty = APoAf.Direccion.Ciudad.Nombre
        Case "barrio"
            GetProperty = APoAf.Direccion.Barrio.Nombre
        Case "calle"
            GetProperty = APoAf.Direccion.Calle + " " + APoAf.Direccion.Nro
        Case "$monto"
            GetProperty = "$" + Trim(Str(Me.Monto))
    End Select
End Function

Friend Sub Guardar()
    
    Dim params As New ParametrosManager
    Me.Id = BD.getIdDisponible(tCuota)
    params.Add "id", Me.Id
    If Not mAfiliado Is Nothing Then params.Add "idAfiliado", Me.Afiliado.Id
    If Not mAreaProtegida Is Nothing Then params.Add "idAreaProtegida", Me.AreaProtegida.Id

    params.Add "mYear", Me.aYear
    params.Add "idCobrador", Me.Cobrador.Id
    params.Add "idEmisor", Me.Emisor.Id
    params.Add "Estado", Me.Estado
    params.Add "FechaEmision", Me.FechaEmision
    
    params.Add "Mes", Me.Mes
    params.Add "Monto", Me.Monto
    params.Add "NroRecibo", Me.NroRecibo

    BD.Insertar tCuota, params
    'aca no aviso a la red porq se va a ejecutar muchas veces juntas,
    'aviso desde el frm generadorde cuotas
End Sub

Public Sub Anular(pAnuladoPor As Empleado, pEstado As eEstadoCuota)
    
    Set AnuladoPor = pAnuladoPor
    FechaAnulado = Date
    
    Dim params As New ParametrosManager
   
    params.Add "id", Me.Id
    If Not mAfiliado Is Nothing Then params.Add "idAfiliado", Me.Afiliado.Id
    If Not mAreaProtegida Is Nothing Then params.Add "idAreaProtegida", Me.AreaProtegida.Id

    params.Add "mYear", Me.aYear
    params.Add "idCobrador", Me.Cobrador.Id
    params.Add "idEmisor", Me.Emisor.Id
    params.Add "Estado", pEstado
    params.Add "FechaEmision", Me.FechaEmision
    params.Add "Mes", Me.Mes
    params.Add "Monto", Me.Monto
    params.Add "NroRecibo", Me.NroRecibo
    params.Add "idAnuladoPor", AnuladoPor.Id
    params.Add "FechaAnulado", FechaAnulado
    
    BD.Insertar tCuotasAnuladas, params
    'para q avise a la red
    Dim cm As New CuotaManager
    cm.CuotasModificadasoAgregadas
    Set cm = Nothing
End Sub

Public Sub RegistrarDevolucion(pDevolucionRegistradaPor As Empleado)
    Set DevolucionRegistradaPor = pDevolucionRegistradaPor
    Estado = eAnulado
    
    Dim params As New ParametrosManager
    params.Add "id", Id, True
    params.Add "Estado", Estado
    params.Add "idDevolucionRegistradapor", pDevolucionRegistradaPor.Id
    
    BD.Update tCuotasAnuladas, params
    'para q avise a la red
    Dim cm As New CuotaManager
    cm.CuotasModificadasoAgregadas
    Set cm = Nothing
End Sub

Public Sub RegistrarCobro(pCobroRegistradoPor As Empleado, Optional pNroRecibo As Long)
'el nrorecibo es para cuando registro el cobro despues de anular el recibo anterior
    Set CobroRegistradoPor = pCobroRegistradoPor
    Estado = ePagado
    FechaPago = Date
    
    Dim params As New ParametrosManager
    params.Add "id", Id, True
    params.Add "Estado", Estado
    params.Add "fechapago", FechaPago
    params.Add "idCobradoPor", CobroRegistradoPor.Id
    If pNroRecibo <> 0 Then
        NroRecibo = pNroRecibo
        params.Add "NroRecibo", NroRecibo
    End If
    BD.Update tCuota, params
    'para q avise a la red
    Dim cm As New CuotaManager
    cm.CuotasModificadasoAgregadas
    Set cm = Nothing
End Sub
