VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCamposRequeridos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private vExigirDNIAE As Boolean
Private vExigirDNIAP As Boolean
Private vExigirDNIRespArea As Boolean
Private vUsarTopeAtencAP As Boolean
Private vUsarTopeAtencArea As Boolean
Private vUsarTopeAtencAE As Boolean

Public Property Get ExigirDNIAE() As Boolean
    ExigirDNIAE = vExigirDNIAE
End Property

Public Property Let ExigirDNIAE(pValor As Boolean)
     vExigirDNIAE = pValor
End Property

Public Property Get ExigirDNIAP() As Boolean
    ExigirDNIAP = vExigirDNIAP
End Property

Public Property Let ExigirDNIAP(pValor As Boolean)
     vExigirDNIAP = pValor
End Property

Public Property Get ExigirDNIRespArea() As Boolean
    ExigirDNIRespArea = vExigirDNIRespArea
End Property

Public Property Let ExigirDNIRespArea(pValor As Boolean)
     vExigirDNIRespArea = pValor
End Property

Public Property Get UsarTopeAtencAP() As Boolean
    UsarTopeAtencAP = vUsarTopeAtencAP
End Property

Public Property Let UsarTopeAtencAP(pValor As Boolean)
     vUsarTopeAtencAP = pValor
End Property

Public Property Get UsarTopeAtencAE() As Boolean
    UsarTopeAtencAE = vUsarTopeAtencAE
End Property

Public Property Let UsarTopeAtencAE(pValor As Boolean)
     vUsarTopeAtencAE = pValor
End Property

Public Property Get UsarTopeAtencArea() As Boolean
    UsarTopeAtencArea = vUsarTopeAtencArea
End Property

Public Property Let UsarTopeAtencArea(pValor As Boolean)
     vUsarTopeAtencArea = pValor
End Property

Private Sub Class_Initialize()
    Load
End Sub

Friend Sub Load()
    vExigirDNIAE = IIf(GetSetting("TbrEmergencyGroup", "Requeridos", "ExigirDNIAE", "True") = "True", True, False)
    vUsarTopeAtencArea = IIf(GetSetting("TbrEmergencyGroup", "Requeridos", "UsarTopeAtencArea", "True") = "True", True, False)
    vUsarTopeAtencAP = IIf(GetSetting("TbrEmergencyGroup", "Requeridos", "UsarTopeAtencAP", "True") = "True", True, False)
    vUsarTopeAtencAE = IIf(GetSetting("TbrEmergencyGroup", "Requeridos", "UsarTopeAtencAE", "True") = "True", True, False)
    vExigirDNIAP = IIf(GetSetting("TbrEmergencyGroup", "Requeridos", "ExigirDNIAP", "True") = "True", True, False)
    vExigirDNIRespArea = IIf(GetSetting("TbrEmergencyGroup", "Requeridos", "ExigirDNIRespArea", "True") = "True", True, False)
End Sub

Private Sub Class_Terminate()
    Save
End Sub

Friend Sub Save()
    SaveSetting "TbrEmergencyGroup", "Requeridos", "ExigirDNIAE", IIf(vExigirDNIAE, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Requeridos", "UsarTopeAtencArea", IIf(vUsarTopeAtencArea, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Requeridos", "UsarTopeAtencAP", IIf(vUsarTopeAtencAP, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Requeridos", "UsarTopeAtencAE", IIf(vUsarTopeAtencAE, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Requeridos", "ExigirDNIAP", IIf(vExigirDNIAP, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Requeridos", "ExigirDNIRespArea", IIf(vExigirDNIRespArea, "True", "False")
End Sub
