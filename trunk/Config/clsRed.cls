VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private vModoServer As Boolean
Private vPuertoTCPLocal As Long
Private vPuertoTCPRemoto As Long
Private vDirIPRemota As String
Private mName As String

Public Property Get ModoServer() As Boolean
    ModoServer = vModoServer
End Property

Public Property Let ModoServer(pValor As Boolean)
     vModoServer = pValor
End Property

Public Property Get PuertoTCPLocal() As Long
    PuertoTCPLocal = vPuertoTCPLocal
End Property

Public Property Let PuertoTCPLocal(pValor As Long)
     vPuertoTCPLocal = pValor
End Property

Public Property Get PuertoTCPRemoto() As Long
    PuertoTCPRemoto = vPuertoTCPRemoto
End Property

Public Property Let PuertoTCPRemoto(pValor As Long)
     vPuertoTCPRemoto = pValor
End Property

Public Property Get DirIPRemota() As String
    DirIPRemota = vDirIPRemota
End Property

Public Property Let DirIPRemota(pValor As String)
     vDirIPRemota = pValor
End Property

Public Property Get Nombre() As String
    Nombre = mName
End Property

Public Property Let Nombre(pNombre As String)
    mName = pNombre
End Property

Friend Sub Load()
    vDirIPRemota = GetSetting("TbrEmergencyGroup", "Red", "DirIpRemota", "192.168.0.1")
    vModoServer = IIf(GetSetting("TbrEmergencyGroup", "Red", "ModoServer", "True") = "True", True, False)
    mName = GetSetting("TbrEmergencyGroup", "Red", "Nombre", "Cliente")
End Sub

Private Sub Class_Initialize()
    Load
End Sub

Friend Sub Save()
    SaveSetting "TbrEmergencyGroup", "Red", "DirIpRemota", vDirIPRemota
    SaveSetting "TbrEmergencyGroup", "Red", "ModoServer", IIf(vModoServer, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Red", "Nombre", mName
End Sub

Private Sub Class_Terminate()
    Save
End Sub
