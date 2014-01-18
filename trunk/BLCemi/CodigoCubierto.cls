VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodigoCubierto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum eDestinoCodigo
    eDCObraSocial = 1
    eDCServicioEmergencia = 2
End Enum

Private mCodigo As CodigoEmergencia
Private mCoseguro As Currency
Private mTipoCodigo As TipoCodigo
Private mServicio As Currency

Public Property Get Codigo() As CodigoEmergencia
    Set Codigo = mCodigo
End Property

Public Property Set Codigo(pValue As CodigoEmergencia)
    Set mCodigo = pValue
End Property

Public Property Get Servicio() As Currency
    Servicio = mServicio
End Property

Public Property Let Servicio(pValue As Currency)
    mServicio = pValue
End Property

Public Property Get Coseguro() As Currency
    Coseguro = mCoseguro
End Property

Public Property Let Coseguro(pValue As Currency)
    mCoseguro = pValue
End Property

Public Property Get TipoCodigo() As TipoCodigo
    Set TipoCodigo = mTipoCodigo
End Property

Public Property Set TipoCodigo(pValue As TipoCodigo)
    Set mTipoCodigo = pValue
End Property

Public Property Get NombreCompuesto() As String
    NombreCompuesto = IIf(mCodigo.Id < 10, "0" + Trim(Str(mCodigo.Id)), Trim(Str(mCodigo.Id))) + " - " + mCodigo.Nombre
End Property

Public Property Get GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "codigo"
            GetProperty = mCodigo.Nombre
        Case "tipo"
            If mTipoCodigo Is Nothing Then
                GetProperty = "<Sin Datos>"
            Else
                GetProperty = mTipoCodigo.Nombre
            End If
    End Select
    
End Property

Friend Sub Save(pTrans As Transaccion, pTipo As eDestinoCodigo, pIdEmpresa As Long)

    Dim params As ParametrosManager
    Set params = New ParametrosManager
    params.Add "idCodigo", mCodigo.Id
    params.Add "idEmpresa", pIdEmpresa
    params.Add "coseguro", mCoseguro
    params.Add "servicio", mServicio
    params.Add "tipoEmpresa", pTipo
    If Not mTipoCodigo Is Nothing Then params.Add "idtipocodigo", mTipoCodigo.Id
    pTrans.AddInsertCommand tCodigoXEmpresa, params
    
End Sub

Public Function Clone() As CodigoCubierto
    Dim cc As New CodigoCubierto
    Set cc.Codigo = Me.Codigo
    cc.Coseguro = Me.Coseguro
    cc.Servicio = Me.Servicio
    Set cc.TipoCodigo = Me.TipoCodigo
    Set Clone = cc
End Function
