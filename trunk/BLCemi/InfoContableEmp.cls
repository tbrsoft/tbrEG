VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoContableEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'esta clase es para empresas
Public Enum eFormula
    eServicioMenosCopagoPorIVA = 1
    eServicioPorIVA = 2
End Enum

Private mIva As Double
Private mFormula As eFormula

Public Property Get IVA() As Double
    IVA = mIva
End Property

Public Property Let IVA(ByVal vNewValue As Double)
    mIva = vNewValue
End Property

Public Property Get Formula() As eFormula
    Formula = mFormula
End Property

Public Property Let Formula(ByVal vNewValue As eFormula)
    mFormula = vNewValue
End Property

Friend Sub Constructor(pIva As Double, pFormula As eFormula)
    mIva = pIva
    mFormula = pFormula
End Sub

'le paso los parametros para que agregue esta informacion adicional
Friend Sub FillParameters(params As ParametrosManager)
    params.Add "formula", mFormula
    params.Add "Iva", mIva
End Sub
