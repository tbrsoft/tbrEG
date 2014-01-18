VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EstadoCivil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mId As Integer

Private mNombre As String

Public Property Get Id() As Integer
    On Error GoTo idErr

    Id = mId

    Exit Property
idErr:
    Call ErrorLog("EstadoCivil", "id Property Get", Err.Description)
End Property

Public Property Let Id(ByVal vNewValue As Integer)
    On Error GoTo idErr

    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("EstadoCivil", "id Property Let", Err.Description)
End Property

Public Property Get Nombre() As String
    On Error GoTo nombreErr

    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("EstadoCivil", "nombre Property Get", Err.Description)
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("EstadoCivil", "nombre Property Let", Err.Description)
End Property

