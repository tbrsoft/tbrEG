VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Guardia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mId As Long
Private mFecha As Date
Private mMonto As Currency
Private mAdelanto As Currency
Private mPlus As Currency
Private mObservaciones As String

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

Friend Sub Save(pIdEmpleado As Long)
    Dim params As New ParametrosManager
    mId = BD.getIdDisponible(tGuardia)
    params.Add "id", mId, True
    params.Add "monto", mMonto
    params.Add "fecha", mFecha
    params.Add "adelanto", mAdelanto
    params.Add "plus", mPlus
    params.Add "idEmpleado", pIdEmpleado
    params.Add "Observaciones", mObservaciones
    BD.Insertar tGuardia, params
End Sub

'setea la guardia como liquidada
Friend Sub Liquidar(pIdLiquidacion As Long, trans As Transaccion)
    Dim params As New ParametrosManager
    params.Add "id", mId, True
    params.Add "idLiquidacion", pIdLiquidacion
    trans.AddUpdateCommand tGuardia, params
End Sub

Public Function GetSaldo() As Currency
GetSaldo = mMonto - mAdelanto + mPlus
End Function


