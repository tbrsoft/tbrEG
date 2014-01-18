VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfiguracionInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

'variables locales para almacenar los valores de las propiedades
Private mvarPathDB As String 'copia local
Private mvarIPServidor As String 'copia local
Private mvarPais As String 'copia local
Private mvarProvincia As String 'copia local
Private mvarCiudad As String 'copia local
Private mvarBarrio As String 'copia local
Private mvarModoFuncionamiento As Integer 'copia local

Public Property Let ModoFuncionamiento(ByVal vData As Integer)
    mvarModoFuncionamiento = vData
End Property

Public Property Get ModoFuncionamiento() As Integer
    ModoFuncionamiento = mvarModoFuncionamiento
End Property

Public Sub Cargar(pPath As String)
    PathDB = Leer_Ini(pPath, "Config", "PathDB", APh + "..\BD\database.mdb") 'TODO este path sirve instalado pero no aqui
    IPServidor = Leer_Ini(pPath, "Config", "IPServidor", "127.0.0.1")
    ModoFuncionamiento = Val(Leer_Ini(pPath, "Config", "ModoFuncionamiento", "1"))
    Pais = Leer_Ini(pPath, "Config", "Pais", "Argentina")
    Provincia = Leer_Ini(pPath, "Config", "Provincia", "Cordoba")
    Ciudad = Leer_Ini(pPath, "Config", "Ciudad", "Capital")
    Barrio = Leer_Ini(pPath, "Config", "Barrio", "Centro")
End Sub

Public Function Save(pPath As String) As Integer
    PathDB = GetSetting("TbrEmergencyGroup", "DBLayer", "PathDB", APh + "..\BD\database.mdb")
    IPServidor = GetIPAddress
    
    Dim Server As String
    Dim SharePath As String
    Dim success As Long
    
    SharePath = Left(PathDB, Len(PathDB) - InStrRev(PathDB, "\"))
    
    Server = "\\" & Environ$("COMPUTERNAME")
    
    success = ShareAdd(Server, SharePath, "DBEmergencyGroup", "", "")
    
    Save = success
    
    Grabar_Ini pPath, "Config", "pathDB", Server + "\" + "DBEmergencyGroup" + "\database.mdb" 'asumo q es siempre database
    Grabar_Ini pPath, "Config", "IPServidor", IPServidor
    Grabar_Ini pPath, "Config", "ModoFuncionamiento", ModoFuncionamiento
    Grabar_Ini pPath, "Config", "Pais", Pais
    Grabar_Ini pPath, "Config", "Provincia", Provincia
    Grabar_Ini pPath, "Config", "Ciudad", Ciudad
    Grabar_Ini pPath, "Config", "Barrio", Barrio
End Function

Public Property Let Barrio(ByVal vData As String)
    mvarBarrio = vData
End Property

Public Property Get Barrio() As String
    Barrio = mvarBarrio
End Property

Public Property Let Pais(ByVal vData As String)
    mvarPais = vData
End Property

Public Property Get Pais() As String
    Pais = mvarPais
End Property

Public Property Let Ciudad(ByVal vData As String)
    mvarCiudad = vData
End Property

Public Property Get Ciudad() As String
    Ciudad = mvarCiudad
End Property

Public Property Let Provincia(ByVal vData As String)
    mvarProvincia = vData
End Property

Public Property Get Provincia() As String
    Provincia = mvarProvincia
End Property

Public Property Let IPServidor(ByVal vData As String)
    mvarIPServidor = vData
End Property

Public Property Get IPServidor() As String
    IPServidor = mvarIPServidor
End Property

Public Property Let PathDB(ByVal vData As String)
    mvarPathDB = vData
End Property

Public Property Get PathDB() As String
    PathDB = mvarPathDB
End Property
