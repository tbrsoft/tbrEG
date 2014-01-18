VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'##ModelId=46A80493034B
Private mId As Long

'##ModelId=46A804A30271
Private mNombre As String
Private mPrimerOrden As String
Private mSegundoOrden As String
Private mTercerOrden As String

Private mProvincias As ProvinciaManager

'##ModelId=46A977E701F6
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Pais", "Id Property Get", Err.Description)
End Property

'##ModelId=46A977E701F7
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Pais", "Id Property Let", Err.Description)
End Property

'##ModelId=46A977E701FB
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Pais", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A977E701FC
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Pais", "Nombre Property Let", Err.Description)
End Property

Public Property Get PrimerOrden() As String
    PrimerOrden = mPrimerOrden
End Property

Public Property Let PrimerOrden(ByVal vNewValue As String)
    mPrimerOrden = vNewValue
End Property

Public Property Get SegundoOrden() As String
    SegundoOrden = mSegundoOrden
End Property

Public Property Let SegundoOrden(ByVal vNewValue As String)
    mSegundoOrden = vNewValue
End Property

Public Property Get TercerOrden() As String
    TercerOrden = mTercerOrden
End Property

Public Property Let TercerOrden(ByVal vNewValue As String)
    mTercerOrden = vNewValue
End Property

Public Property Get Provincias() As ProvinciaManager
    Set Provincias = ProvinciasLocal.GetByPais(Me)
End Property
