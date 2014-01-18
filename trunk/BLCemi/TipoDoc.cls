VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TipoDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"445E82BB00DA"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"entity"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'##ModelId=445E82C7037A
Private mId As Integer

'##ModelId=445E82CF002E
Private mNombre As String

'##ModelId=448F2462027C
Public Property Get Id() As Integer
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("TipoDoc", "id Property Get", Err.Description)
End Property

'##ModelId=448F2462027D
Public Property Let Id(ByVal vNewValue As Integer)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("TipoDoc", "id Property Let", Err.Description)
End Property

'##ModelId=448F24620281
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("TipoDoc", "nombre Property Get", Err.Description)
End Property

'##ModelId=448F24620282
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("TipoDoc", "nombre Property Let", Err.Description)
End Property
