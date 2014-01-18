VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDefaults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim vProvincia As String
Dim vCiudad As String
Dim vBarrio As String
Dim vPais As String

Public Property Get Pais() As String
    Pais = vPais
End Property

Public Property Let Pais(pValue As String)
    vPais = pValue
End Property

Public Property Get Provincia() As String
    Provincia = vProvincia
End Property

Public Property Let Provincia(pValue As String)
    vProvincia = pValue
End Property

Public Property Get Ciudad() As String
    Ciudad = vCiudad
End Property

Public Property Let Ciudad(pValue As String)
    vCiudad = pValue
End Property

Public Property Get Barrio() As String
    Barrio = vBarrio
End Property

Public Property Let Barrio(pValue As String)
    vBarrio = pValue
End Property

Private Sub Class_Initialize()
    Load
End Sub

Private Sub Class_Terminate()
    Save
End Sub

Friend Sub Save()
    SaveSetting "TbrEmergencyGroup", "Defaults", "Pais", vPais
    SaveSetting "TbrEmergencyGroup", "Defaults", "Provincia", vProvincia
    SaveSetting "TbrEmergencyGroup", "Defaults", "Ciudad", vCiudad
    SaveSetting "TbrEmergencyGroup", "Defaults", "Barrio", vBarrio
End Sub

Friend Sub Load()
    vPais = GetSetting("TbrEmergencyGroup", "Defaults", "Pais", "argentina")
    vProvincia = GetSetting("TbrEmergencyGroup", "Defaults", "Provincia", "cordoba")
    vCiudad = GetSetting("TbrEmergencyGroup", "Defaults", "Ciudad", "unquillo")
    vBarrio = GetSetting("TbrEmergencyGroup", "Defaults", "Barrio", "centro")
End Sub
