VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiquidacionEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mTipoEmpresa As eDestinoCodigo
Private mIdEmpresa As Long
Private mFecha As Date
Private mTotal As Currency
Private mDetalle As InfoContableManager
Private mEmitida As Boolean
Private mId As Long
Private mMes As Integer
Private mYear As Integer

Public Property Get Mes() As Integer
Mes = mMes
End Property

Public Property Let Mes(ByVal vNewValue As Integer)
mMes = vNewValue
End Property

Public Property Get Year() As Integer
Year = mYear
End Property

Public Property Let Year(ByVal vNewValue As Integer)
mYear = vNewValue
End Property

Public Property Get Id() As Long
    Id = mId
End Property

Friend Property Let Id(ByVal vNewValue As Long)
    mId = vNewValue
End Property

Public Property Get idEmpresa() As Long
    idEmpresa = mIdEmpresa
End Property

Friend Property Let idEmpresa(ByVal vNewValue As Long)
    mIdEmpresa = vNewValue
End Property

Public Property Get TipoEmpresa() As eDestinoCodigo
    TipoEmpresa = mTipoEmpresa
End Property

Public Property Let TipoEmpresa(ByVal vNewValue As eDestinoCodigo)
    mTipoEmpresa = vNewValue
End Property

Public Property Get Fecha() As Date
Fecha = mFecha
End Property

Public Property Let Fecha(ByVal vNewValue As Date)
mFecha = vNewValue
End Property

Public Property Get Total() As Currency
    Total = mTotal
End Property

'ver bien
Friend Property Let Total(pValue As Currency)
    mTotal = pValue
End Property

Public Property Get Detalle() As InfoContableManager
    If mDetalle Is Nothing Then
        Set mDetalle = New InfoContableManager
        mDetalle.LoadByLiquidacion (Me.Id)
    End If
    Set Detalle = mDetalle
End Property

Public Property Set Detalle(pValue As InfoContableManager)
    Set mDetalle = pValue
End Property

'para marcar si la imprimio o no
Public Property Get Emitida() As Boolean
    Emitida = mEmitida
End Property

Public Property Let Emitida(ByVal vNewValue As Boolean)
    mEmitida = vNewValue
End Property

Public Function GetProperty(pNombre As String) As String
    Dim pName As String
    pName = LCase(pNombre)
    Select Case pName
        Case "empresa"
            If mTipoEmpresa = eDCServicioEmergencia Then
                If Not ServiciosEmergenciaLocal.Item(mIdEmpresa) Is Nothing Then
                    GetProperty = ServiciosEmergenciaLocal.Item(mIdEmpresa).Nombre
                End If
            Else
                If Not ObrasSocialesLocal.Item(mIdEmpresa) Is Nothing Then
                    GetProperty = ObrasSocialesLocal.Item(mIdEmpresa).Nombre
                End If
            End If
        Case "tipotostring"
            GetProperty = IIf(mTipoEmpresa = eDCObraSocial, "Obra Social", "Servicio Emergencia")
        Case "saldo"
            GetProperty = Detalle.GetSaldo
    End Select
End Function


Public Sub Save()
    On Error GoTo errman
    Dim params As New ParametrosManager
    mId = BD.getIdDisponible(tLiquidacionEmpresa)
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    params.Add "id", mId, True
    params.Add "fecha", mFecha
    params.Add "idEmpresa", mIdEmpresa
    params.Add "tipoEmpresa", mTipoEmpresa
    params.Add "mes", mMes
    params.Add "aYear", mYear
    params.Add "total", mTotal
    mDetalle.Save mId, trans
    trans.AddInsertCommand tLiquidacionEmpresa, params
    trans.Execute
    Exit Sub
errman:
    ErrorLog "LiquidacionEmpresa", "Save", err.Description
End Sub
