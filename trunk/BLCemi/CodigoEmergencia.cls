VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodigoEmergencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FF44037A"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A8047F0251
Private mId As Long

'##ModelId=46A804B702AF
Private mNombre As String

'##ModelId=46A816DD00AB
Private mSintomas As SintomaManager

Private mVencimiento As Integer

Private mBold As Boolean

Private mColorFuente As Long

Public Property Get ColorFuente() As Long
    ColorFuente = mColorFuente
End Property

Public Property Let ColorFuente(pColor As Long)
    mColorFuente = pColor
End Property

Public Property Get Bold() As Boolean
    
    Bold = mBold

End Property

Public Property Let Bold(ByVal vNewValue As Boolean)
    
    mBold = vNewValue

End Property

Public Property Get Vencimiento() As Integer
    
    Vencimiento = mVencimiento

End Property

Public Property Let Vencimiento(ByVal vNewValue As Integer)
    
    mVencimiento = vNewValue

End Property

'##ModelId=46A9762F013C
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("CodigoEmergencia", "Id Property Get", err.Description)
End Property

'##ModelId=46A9762F013D
Public Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("CodigoEmergencia", "Id Property Let", err.Description)
End Property

'##ModelId=46A9762F0141
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("CodigoEmergencia", "Nombre Property Get", err.Description)
End Property

'##ModelId=46A9762F0142
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("CodigoEmergencia", "Nombre Property Let", err.Description)
End Property

'##ModelId=46A9762F0146
Public Property Get Sintomas() As SintomaManager
    On Error GoTo SintomasErr
    
    If mSintomas Is Nothing Then
       Set mSintomas = SintomasLocal.GetByCodEmergencia(Me)
    End If
    
    Set Sintomas = mSintomas

    Exit Property
SintomasErr:
    Call ErrorLog("CodigoEmergencia", "Sintomas Property Get", err.Description)
End Property

'##ModelId=46A9762F0149
Public Property Set Sintomas(ByVal vNewValue As SintomaManager)
    On Error GoTo SintomasErr

    '## Generated default body ...
    Set mSintomas = vNewValue

    Exit Property
SintomasErr:
    Call ErrorLog("CodigoEmergencia", "Sintomas Property Set", err.Description)
End Property

Public Property Get NombreCompuesto() As String
    NombreCompuesto = IIf(mId < 10, "0" + Trim(Str(mId)), Trim(Str(mId))) + " - " + mNombre
End Property

Public Property Get GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "pgbold"
            GetProperty = IIf(mBold, "Si", "No")
        Case "pgcolorfuente"
        'despues se podra asignar un color personalizado
            GetProperty = IIf(mColorFuente <> 0, "Si", "No")
    End Select
End Property

Public Sub SaveChanges()
    Dim params As New ParametrosManager
    params.Add "id", mId, True
    params.Add "vencimiento", mVencimiento
    params.Add "nombre", mNombre
    params.Add "bold", IIf(mBold, "True", "False")
    params.Add "colorfuente", mColorFuente
    BD.Update tCodigoEmergencia, params
End Sub


Public Function Clone() As CodigoEmergencia
Attribute Clone.VB_Description = "No clona los sintomas!"
    Dim ce As New CodigoEmergencia
    ce.Id = Me.Id
    ce.Nombre = Me.Nombre
    ce.Vencimiento = Me.Vencimiento
    ce.Bold = Me.Bold
    ce.ColorFuente = Me.ColorFuente
    Set Clone = ce
End Function
