VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBDLayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim vPathDB As String
Dim FSO As New Scripting.FileSystemObject

Public Property Get PathDB() As String
    'CORRECION ESPECIFICA PARA SANTAFESION MARZO 2010, borrar en el futuro
    If FSO.FileExists(vPathDB) = False Then vPathDB = App.Path + "\..\BD\database.mdb"
    PathDB = vPathDB
End Property

Public Property Let PathDB(pValue As String)
    vPathDB = pValue
End Property

Private Sub Class_Initialize()
    Load
End Sub

Private Sub Class_Terminate()
    Save
End Sub

Friend Sub Save()
    SaveSetting "TbrEmergencyGroup", "DBLayer", "PathDB", vPathDB
End Sub

Friend Sub Load()
    vPathDB = GetSetting("TbrEmergencyGroup", "DBLayer", "PathDB", App.Path + "\..\BD\database.mdb")
    'TODO arreglar esto por algo mas global del sistema!!!
    'CORRECION ESPECIFICA PARA SANTAFESION MARZO 2010, borrar en el futuro
    If FSO.FileExists(vPathDB) = False Then vPathDB = App.Path + "\..\BD\database.mdb"
    
End Sub
