VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A81906000F"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A8191F002E
Private mId As Long

'##ModelId=46A801290399
Private mCalle As String

'##ModelId=46A80169030D
Private mNro As String

'##ModelId=46A801760242
Private mBarrio As Barrio

'##ModelId=46A81940002E
Private mCiudad As Ciudad

'##ModelId=46A81950005D
Private mProvincia As Provincia

Private mDepto As String

Private mPiso As String
Private mTorre As String
Private mEntrecalles As String

'##ModelId=46A9767702D0
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Direccion", "Id Property Get", Err.Description)
End Property

'##ModelId=46A9767702D1
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Direccion", "Id Property Let", Err.Description)
End Property

'##ModelId=46A9767702D5
Public Property Get Calle() As String
    On Error GoTo CalleErr

    '## Generated default body ...
    Calle = mCalle

    Exit Property
CalleErr:
    Call ErrorLog("Direccion", "Calle Property Get", Err.Description)
End Property

'##ModelId=46A9767702D6
Public Property Let Calle(ByVal vNewValue As String)
    On Error GoTo CalleErr

    '## Generated default body ...
    mCalle = vNewValue

    Exit Property
CalleErr:
    Call ErrorLog("Direccion", "Calle Property Let", Err.Description)
End Property

'##ModelId=46A9767702DA
Public Property Get Nro() As String
    On Error GoTo NroErr

    '## Generated default body ...
    Nro = mNro

    Exit Property
NroErr:
    Call ErrorLog("Direccion", "Nro Property Get", Err.Description)
End Property

'##ModelId=46A9767702DB
Public Property Let Nro(ByVal vNewValue As String)
    On Error GoTo NroErr

    '## Generated default body ...
    mNro = vNewValue

    Exit Property
NroErr:
    Call ErrorLog("Direccion", "Nro Property Let", Err.Description)
End Property

'##ModelId=46A9767702DF
Public Property Get Barrio() As Barrio
    On Error GoTo BarrioErr

    '## Generated default body ...
    Set Barrio = mBarrio

    Exit Property
BarrioErr:
    Call ErrorLog("Direccion", "Barrio Property Get", Err.Description)
End Property

'##ModelId=46A9767702E2
Public Property Set Barrio(ByVal vNewValue As Barrio)
    On Error GoTo BarrioErr

    '## Generated default body ...
    Set mBarrio = vNewValue

    Exit Property
BarrioErr:
    Call ErrorLog("Direccion", "Barrio Property Set", Err.Description)
End Property

'##ModelId=46A9767702E4
Public Property Get Ciudad() As Ciudad
    On Error GoTo CiudadErr

    '## Generated default body ...
    Set Ciudad = mBarrio.Parent

    Exit Property
CiudadErr:
    Call ErrorLog("Direccion", "Ciudad Property Get", Err.Description)
End Property

'##ModelId=46A9767702E9
Public Property Get Provincia() As Provincia
    On Error GoTo ProvinciaErr

    If Not mBarrio Is Nothing Then
        Set Provincia = mBarrio.Parent.Parent
    End If
    
    Exit Property
ProvinciaErr:
    Call ErrorLog("Direccion", "Provincia Property Get", Err.Description)
End Property

Public Property Get Pais() As Pais
    On Error GoTo ProvinciaErr

    If Not mBarrio Is Nothing Then
        Set Pais = mBarrio.Parent.Parent.Parent
    End If
    
    Exit Property
ProvinciaErr:
    Call ErrorLog("Direccion", "Provincia Property Get", Err.Description)
End Property

Public Property Get Depto() As String
    Depto = mDepto
End Property

Public Property Let Depto(ByVal vNewValue As String)
    mDepto = vNewValue
End Property

Public Property Get Piso() As String
    Piso = mPiso
End Property

Public Property Let Piso(ByVal vNewValue As String)
    mPiso = vNewValue
End Property

Public Property Get Torre() As String
    Torre = mTorre
End Property

Public Property Let Torre(ByVal vNewValue As String)
    mTorre = vNewValue
End Property

Public Property Get Entrecalles() As String
    Entrecalles = mEntrecalles
End Property

Public Property Let Entrecalles(ByVal vNewValue As String)
    mEntrecalles = vNewValue
End Property

Public Function GetBarrioCiudadProvincia() As String
    'If (Me.Barrio Is Nothing) Then Set Me.Barrio = New BLCemi.Barrio
    'If (Me.Ciudad Is Nothing) Then Set Me.Ciudad = New BLCemi.Ciudad
    'If (Me.Provincia Is Nothing) Then Set Me.Provincia = New BLCemi.Provincia
    GetBarrioCiudadProvincia = Me.Barrio.Nombre + ", " + Me.Ciudad.Nombre + ", " + Me.Provincia.Nombre + "."
End Function

Public Function Clone() As Direccion
    Dim d As New Direccion
    Set d.Barrio = Me.Barrio
    d.Calle = Me.Calle
    d.Id = Me.Id
    d.Nro = Me.Nro
    d.Piso = Me.Piso
    d.Depto = Me.Depto
    d.Entrecalles = Me.Entrecalles
    d.Torre = Me.Torre
    Set Clone = d
    Set d = Nothing
End Function

Public Function EqualsTo(pDireccion As Direccion) As Boolean
    EqualsTo = (pDireccion.Barrio.Id = Me.Barrio.Id And pDireccion.Calle = Me.Calle And pDireccion.Id = Me.Id And pDireccion.Nro = Me.Nro And pDireccion.Piso = Me.Piso And pDireccion.Depto = Me.Depto And pDireccion.Torre = Me.Torre And pDireccion.Entrecalles = Me.Entrecalles)
End Function

Friend Sub Cargar(pId As Long)
    'esta clase y vehiculo son las unicas q cargo de esta forma
    Dim rs
    Set rs = BD.SelectSQL("select * from direccion where id=" + Str(pId))
    Set Me.Barrio = BarriosLocal.Item(rs!idBarrio)
    Me.Calle = rs!Calle
    Me.Nro = rs!Nro
    Me.Id = rs!Id
    Me.Depto = "" & rs!Depto
    Me.Piso = "" & rs!Piso
    Me.Torre = "" & rs!Torre
    Me.Entrecalles = "" & rs!Entrecalles

End Sub

Friend Sub Guardar(pTrans As Transaccion)
    Dim params As New ParametrosManager
    Me.Id = BD.getIdDisponible(tDireccion)
    params.Add "id", Me.Id
    params.Add "calle", Me.Calle
    params.Add "nro", Me.Nro
    params.Add "idBarrio", Me.Barrio.Id
    params.Add "piso", Me.Piso
    params.Add "depto", Me.Depto
    params.Add "torre", Me.Torre
    params.Add "entrecalles", Me.Entrecalles
        
    pTrans.AddInsertCommand tDireccion, params
End Sub

Friend Sub GuardarModificaciones(pTrans As Transaccion)
    Dim params As New ParametrosManager
    params.Add "id", Me.Id, True
    params.Add "calle", Me.Calle
    params.Add "nro", Me.Nro
    params.Add "idBarrio", Me.Barrio.Id
    params.Add "piso", Me.Piso
    params.Add "depto", Me.Depto
    params.Add "torre", Me.Torre
    params.Add "entrecalles", Me.Entrecalles
    
    pTrans.AddUpdateCommand tDireccion, params
End Sub



