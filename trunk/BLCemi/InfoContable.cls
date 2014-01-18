VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'esta clase es distinta a las demas, se guarda con la atencion,
'pero se puede cargar y guardar sola
Public Enum eCondicionIVA
    eCINoInformado = 0
    eCIGravado = 1
    eCINoGravado = 2
    eCIExento = 3
End Enum

Private mCoseguro As Currency
Private mAbonado As Currency
Private mServicio As Currency 'este se registra en la OS o SE, esta por si lo cambia
'Private mRendido As Boolean
Private mCondIva As eCondicionIVA
Private mIdAtencion As Long
'para no tener q recargar cada vez en getproperty
Private mAtencionAux As Atencion
'lo necesito por si modifica la informacion en la empresa
Private mMontoIva As Currency

Public Property Get Coseguro() As Currency
    Coseguro = mCoseguro
End Property

Public Property Let Coseguro(ByVal vNewValue As Currency)
    mCoseguro = vNewValue
End Property

Public Property Get Servicio() As Currency
    Servicio = mServicio
End Property

Public Property Let Servicio(ByVal vNewValue As Currency)
    mServicio = vNewValue
End Property

Public Property Get MontoAbonado() As Currency
    MontoAbonado = mAbonado
End Property

Public Property Let MontoAbonado(ByVal vNewValue As Currency)
    mAbonado = vNewValue
End Property

'Public Property Get Rendido() As Boolean
'    Rendido = mRendido
'End Property
'
'Public Property Let Rendido(ByVal vNewValue As Boolean)
'    mRendido = vNewValue
'End Property

Public Property Get CondicionIVA() As eCondicionIVA
    CondicionIVA = mCondIva
End Property

Public Property Let CondicionIVA(ByVal vNewValue As eCondicionIVA)
    mCondIva = vNewValue
End Property

Public Property Get MontoIVA() As Currency
    MontoIVA = mMontoIva
End Property

Friend Property Let MontoIVA(pValue As Currency)
    mMontoIva = pValue
End Property

Friend Sub Constructor(pServicio As Currency, pCoseguro As Currency, pAbonado As Currency, pCondIva As eCondicionIVA, pIdAtencion As Long, pMontoIva As Currency)
    mAbonado = pAbonado
    mCondIva = pCondIva
    mCoseguro = pCoseguro
    mIdAtencion = pIdAtencion
    mServicio = pServicio
    mMontoIva = pMontoIva
End Sub

Public Property Get IsComplete() As Boolean
IsComplete = (mAbonado <> -1) And (mCondIva <> eCINoInformado) And (mCoseguro <> -1) _
                And (mIdAtencion <> -1) And (mServicio <> -1)
End Property

'le paso los parametros para que agregue esta informacion adicional
Friend Sub FillParameters(params As ParametrosManager)
    params.Add "copago", mCoseguro
    params.Add "abonado", mAbonado
    params.Add "condIva", mCondIva
    params.Add "servicio", mServicio
    params.Add "montoiva", mMontoIva
End Sub

Friend Sub Save(pIdLiquidacion As Long, trans As Transaccion)
    Dim params As New ParametrosManager
    params.Add "id", mIdAtencion, True
    params.Add "idLiquidacion", pIdLiquidacion
    FillParameters params
    trans.AddUpdateCommand tAtencion, params
End Sub

Public Function GetSaldo() As Currency
    GetSaldo = mServicio - mAbonado + mMontoIva
End Function

Public Property Get idAtencion() As Long
    idAtencion = mIdAtencion
End Property

Friend Property Let idAtencion(ByVal vNewValue As Long)
    mIdAtencion = vNewValue
End Property

Public Function GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "fecha"
            GetProperty = GetAtencion().Fecha
                    
        End Select
End Function

Private Function GetAtencion() As Atencion
If mAtencionAux Is Nothing Then
    Dim ats As New AtencionManager
    Set mAtencionAux = ats.LoadById(mIdAtencion)
End If
Set GetAtencion = mAtencionAux
End Function


'la inicializo vacia
Private Sub Class_Initialize()
    mAbonado = -1
    mCondIva = eCINoInformado
    mCoseguro = -1
    mIdAtencion = -1
    mServicio = -1
    mMontoIva = -1
End Sub

'calcula el iva segun los datos de la empresa, no se si esta bien ubicado aca
Public Function GetIva(pInfoContableEmp As InfoContableEmp) As Currency
    If mCondIva = eCIGravado Then
        If pInfoContableEmp.Formula = eServicioMenosCopagoPorIVA Then
            mMontoIva = (mServicio - mCoseguro) * (pInfoContableEmp.IVA / 100)
        Else
            mMontoIva = mServicio * (pInfoContableEmp.IVA / 100)
        End If
    Else
        'warning revisar si esta bien el monto de iva
        mMontoIva = 0
    End If
    GetIva = mMontoIva
End Function

