VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private vUtilizarTipos As Boolean
Private vHabilitarVencimiento As Boolean
Private vCosegurosPorCodigo As Boolean
Private vExigirCodigos As Boolean

Public Property Get ExigirCodigos() As Boolean
    ExigirCodigos = vExigirCodigos
End Property

Public Property Let ExigirCodigos(pValor As Boolean)
     vExigirCodigos = pValor
End Property

Public Property Get UtilizarTipos() As Boolean
    UtilizarTipos = vUtilizarTipos
End Property

Public Property Let UtilizarTipos(pValor As Boolean)
     vUtilizarTipos = pValor
End Property

Public Property Get HabilitarVencimiento() As Boolean
    HabilitarVencimiento = vHabilitarVencimiento
End Property

Public Property Let HabilitarVencimiento(pValor As Boolean)
     vHabilitarVencimiento = pValor
End Property

Public Property Get CosegurosPorCodigo() As Boolean
    CosegurosPorCodigo = vCosegurosPorCodigo
End Property

Public Property Let CosegurosPorCodigo(pValor As Boolean)
     vCosegurosPorCodigo = pValor
End Property

Private Sub Class_Initialize()
    Load
End Sub

Friend Sub Load()
    vUtilizarTipos = IIf(GetSetting("TbrEmergencyGroup", "Codigo", "UtilizarTipos", "True") = "True", True, False)
    vHabilitarVencimiento = IIf(GetSetting("TbrEmergencyGroup", "Codigo", "HabilitarVencimiento", "True") = "True", True, False)
    vCosegurosPorCodigo = IIf(GetSetting("TbrEmergencyGroup", "Codigo", "CosegurosPorCodigo", "True") = "True", True, False)
    vExigirCodigos = IIf(GetSetting("TbrEmergencyGroup", "Codigo", "ExigirCodigos", "True") = "True", True, False)

End Sub

Private Sub Class_Terminate()
    Save
End Sub

Friend Sub Save()
    SaveSetting "TbrEmergencyGroup", "Codigo", "UtilizarTipos", IIf(vUtilizarTipos, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Codigo", "HabilitarVencimiento", IIf(vHabilitarVencimiento, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Codigo", "CosegurosPorCodigo", IIf(vCosegurosPorCodigo, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Codigo", "ExigirCodigos", IIf(vExigirCodigos, "True", "False")
End Sub

