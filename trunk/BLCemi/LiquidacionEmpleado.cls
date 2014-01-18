VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiquidacionEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'es practicamente igual a guardia pero esta se puede usar para pagos mensuales,
'aparte es un resumen de cada mes

Private mId As Long
Private mFecha As Date
Private mMonto As Currency
Private mAdelanto As Currency
Private mPlus As Currency
Private mObservaciones As String
Private mDetalle As GuardiaManager
Private mMes As Integer
Private mYear As Integer
'para uso interno
Private mIdEmpleado As Long

Friend Property Get IdEmpleado() As Long
IdEmpleado = mIdEmpleado
End Property

Friend Property Let IdEmpleado(ByVal vNewValue As Long)
mIdEmpleado = vNewValue
End Property

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

Public Property Let Id(ByVal vNewValue As Long)
mId = vNewValue
End Property

Public Property Get Fecha() As Date
Fecha = mFecha
End Property

Public Property Let Fecha(ByVal vNewValue As Date)
mFecha = vNewValue
End Property

Public Property Get Adelanto() As Currency
Adelanto = mAdelanto
End Property

Public Property Let Adelanto(ByVal vNewValue As Currency)
mAdelanto = vNewValue
End Property

Public Property Get Monto() As Currency
Monto = mMonto
End Property

Public Property Let Monto(ByVal vNewValue As Currency)
mMonto = vNewValue
End Property

Public Property Get Plus() As Variant
Plus = mPlus
End Property

Public Property Let Plus(ByVal vNewValue As Variant)
mPlus = vNewValue
End Property

Public Property Get Observaciones() As String
    Observaciones = mObservaciones
End Property

Public Property Let Observaciones(ByVal vNewValue As String)
    mObservaciones = vNewValue
End Property

Public Property Get Detalle() As GuardiaManager
    If mDetalle Is Nothing Then
        Set mDetalle = New GuardiaManager
        mDetalle.LoadByLiquidacion mId
    End If
    Set Detalle = mDetalle
End Property

Public Property Set Detalle(ByVal vNewValue As GuardiaManager)
    Set mDetalle = vNewValue
End Property

Public Function GetProperty(Nombre As String) As String
    Dim aux As String
    aux = LCase(Nombre)
    Select Case aux
        Case "nombreempleado"
            GetProperty = EmpleadosLocal.Item(mIdEmpleado).NombreCompleto
    End Select
End Function

Public Function GetEmpleado() As Empleado
    Set GetEmpleado = EmpleadosLocal.Item(mIdEmpleado)
End Function

Public Function GetSaldo() As Currency
GetSaldo = mMonto - mAdelanto + mPlus
End Function

Friend Sub Save(pIdEmpleado As Long)
    On Error GoTo errman
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
    mId = BD.getIdDisponible(tLiquidacionEmpleado)
    params.Add "id", mId, True
    params.Add "monto", mMonto
    params.Add "fecha", mFecha
    params.Add "adelanto", mAdelanto
    params.Add "plus", mPlus
    params.Add "idEmpleado", pIdEmpleado
    mIdEmpleado = pIdEmpleado
    params.Add "mes", mMes
    params.Add "aYear", mYear
    params.Add "Observaciones", mObservaciones
    trans.AddInsertCommand tLiquidacionEmpleado, params
    mDetalle.Liquidar trans, mId
    trans.Execute
    Exit Sub
errman:
    ErrorLog "LiquidacionEmpleado", "Liquidar", err.Description
End Sub


