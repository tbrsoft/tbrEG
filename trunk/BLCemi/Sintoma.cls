VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sintoma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FF4600AB"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A8046F009C
Private mId As Long

'##ModelId=46A81627034B
Private mNombre As String

Private mParent As CodigoEmergencia

'##ModelId=46A9787F02B2
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Sintoma", "Id Property Get", err.Description)
End Property

'##ModelId=46A9787F02B3
Public Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Sintoma", "Id Property Let", err.Description)
End Property

'##ModelId=46A9787F02B7
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Sintoma", "Nombre Property Get", err.Description)
End Property

'##ModelId=46A9787F02B8
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Sintoma", "Nombre Property Let", err.Description)
End Property

Public Property Get NombreCompuesto() As String
Attribute NombreCompuesto.VB_Description = "Devuelve el codigo concatenado con el nombre."
    NombreCompuesto = IIf(mId < 10, "0" + Trim(Str(mId)), Trim(Str(mId))) + " - " + mNombre
End Property

Public Property Get Parent() As CodigoEmergencia
    Set Parent = mParent
End Property

Friend Property Set Parent(pCodEmergencia As CodigoEmergencia)
    Set mParent = pCodEmergencia
End Property

Public Function GetProperty(pNombre As String) As String
    Dim pName As String
    pName = LCase(pNombre)
    Select Case pName
        Case "codigo"
            GetProperty = Me.Parent.NombreCompuesto
    End Select
End Function

