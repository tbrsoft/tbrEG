VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Vehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FEA202BF"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A80446008C
Private mId As Long

'##ModelId=46A8051C0157
Private mMarca As String

'##ModelId=46A8052703B9
Private mModelo As String

'##ModelId=46A8053600EA
Private mPatente As String

Private mTipo As String

Private mColor As String

Private mPerjuicios As String

Public Property Get Perjuicios() As String
    On Error GoTo MarcaErr

    '## Generated default body ...
    Perjuicios = mPerjuicios

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Perjuicios Property Get", Err.Description)
End Property

Public Property Let Perjuicios(ByVal vNewValue As String)
    On Error GoTo MarcaErr

    '## Generated default body ...
    mPerjuicios = vNewValue

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Perjuicios Property Let", Err.Description)
End Property

Public Property Get Color() As String
    On Error GoTo MarcaErr

    '## Generated default body ...
    Color = mColor

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Color Property Get", Err.Description)
End Property

Public Property Let Color(ByVal vNewValue As String)
    On Error GoTo MarcaErr

    '## Generated default body ...
    mColor = vNewValue

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Color Property Let", Err.Description)
End Property

Public Property Get tipo() As String
    On Error GoTo MarcaErr

    '## Generated default body ...
    tipo = mTipo

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Tipo Property Get", Err.Description)
End Property

Public Property Let tipo(ByVal vNewValue As String)
    On Error GoTo MarcaErr

    '## Generated default body ...
    mTipo = vNewValue

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Tipo Property Let", Err.Description)
End Property

'##ModelId=46A978B603C9
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Vehiculo", "Id Property Get", Err.Description)
End Property

'##ModelId=46A978B603CA
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Vehiculo", "Id Property Let", Err.Description)
End Property

'##ModelId=46A978B603CE
Public Property Get Marca() As String
    On Error GoTo MarcaErr

    '## Generated default body ...
    Marca = mMarca

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Marca Property Get", Err.Description)
End Property

'##ModelId=46A978B603CF
Public Property Let Marca(ByVal vNewValue As String)
    On Error GoTo MarcaErr

    '## Generated default body ...
    mMarca = vNewValue

    Exit Property
MarcaErr:
    Call ErrorLog("Vehiculo", "Marca Property Let", Err.Description)
End Property

'##ModelId=46A978B603D3
Public Property Get Modelo() As String
    On Error GoTo ModeloErr

    '## Generated default body ...
    Modelo = mModelo

    Exit Property
ModeloErr:
    Call ErrorLog("Vehiculo", "Modelo Property Get", Err.Description)
End Property

'##ModelId=46A978B603D4
Public Property Let Modelo(ByVal vNewValue As String)
    On Error GoTo ModeloErr

    '## Generated default body ...
    mModelo = vNewValue

    Exit Property
ModeloErr:
    Call ErrorLog("Vehiculo", "Modelo Property Let", Err.Description)
End Property

'##ModelId=46A978B603D8
Public Property Get Patente() As String
    On Error GoTo PatenteErr

    '## Generated default body ...
    Patente = mPatente

    Exit Property
PatenteErr:
    Call ErrorLog("Vehiculo", "Patente Property Get", Err.Description)
End Property

'##ModelId=46A978B603D9
Public Property Let Patente(ByVal vNewValue As String)
    On Error GoTo PatenteErr

    '## Generated default body ...
    mPatente = vNewValue

    Exit Property
PatenteErr:
    Call ErrorLog("Vehiculo", "Patente Property Let", Err.Description)
End Property

Friend Sub Cargar(pId As Long)
'esta clase es la unica q cargo de esta forma
Dim rs
Set rs = BD.SelectSQL("select * from vehiculo where id=" + Str(pId))
Me.Marca = rs!Marca
Me.Modelo = rs!Modelo
Me.Id = rs!Id
Me.Patente = rs!Patente
Me.Color = "" & rs!Color
Me.Perjuicios = "" & rs!Perjuicios
Me.tipo = "" & rs!tipo
End Sub

Friend Sub Guardar(pTrans As Transaccion)
    Dim params As New ParametrosManager
    Me.Id = BD.getIdDisponible(tVehiculo)
    params.Add "id", Me.Id
    params.Add "marca", Me.Marca
    params.Add "patente", Me.Patente
    params.Add "modelo", Me.Modelo
'    params.Add "tipo", Me.Tipo
'    params.Add "perjuicios", Me.Perjuicios
'    params.Add "color", Me.Color
    
    pTrans.AddInsertCommand tVehiculo, params
End Sub

Friend Sub GuardarModificaciones(pTrans As Transaccion)
    If Me.Id = 0 Then 'esto es por si agrega un nuevo vehiculo en modificar afiliado
        Guardar pTrans
    Else
        Dim params As New ParametrosManager
        params.Add "id", Me.Id, True
        params.Add "marca", Me.Marca
        params.Add "patente", Me.Patente
        params.Add "modelo", Me.Modelo
        params.Add "tipo", Me.tipo
        params.Add "perjuicios", Me.Perjuicios
        params.Add "color", Me.Color
    
        pTrans.AddUpdateCommand tVehiculo, params
    End If
End Sub

Friend Sub GuardarPorSiniestro(pTrans As Transaccion, pIdSiniestro As Long)
    Dim params As New ParametrosManager
    
    params.Add "marca", Me.Marca
    params.Add "patente", Me.Patente
    params.Add "modelo", Me.Modelo
    params.Add "tipo", Me.tipo
    params.Add "perjuicios", Me.Perjuicios
    params.Add "color", Me.Color
    params.Add "idSiniestro", pIdSiniestro
    If Me.Id = 0 Then
        Me.Id = BD.getIdDisponible(tVehiculo)
        params.Add "id", Me.Id
        pTrans.AddInsertCommand tVehiculo, params
    Else
        params.Add "id", Me.Id, True
        pTrans.AddUpdateCommand tVehiculo, params
    End If
End Sub
