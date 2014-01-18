VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServicioEmergencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FEA50399"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A804480232
Private mId As Long

'##ModelId=46A804CD037A
Private mNombre As String

'##ModelId=46A8111102EE
Private mTelefonos As TelefonoManager

'##ModelId=46A819BC0148
Private mDireccion As Direccion

Private mAfiliados As AfiliadoExternoManager

Private mCodigos As CodigoCubiertoManager
Private mInfoContable As InfoContableEmp

Private bAfiliadosCargados As Boolean

'##ModelId=46A978000330
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("ServicioEmergencia", "Id Property Get", Err.Description)
End Property

'##ModelId=46A978000331
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("ServicioEmergencia", "Id Property Let", Err.Description)
End Property

'##ModelId=46A978000335
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("ServicioEmergencia", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A978000336
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("ServicioEmergencia", "Nombre Property Let", Err.Description)
End Property

'##ModelId=46A97800033A
Public Property Get Telefonos() As TelefonoManager
    On Error GoTo TelefonosErr

    If mTelefonos Is Nothing Then
        Set mTelefonos = TelefonosLocal.GetByOwner(Me.Id, eOTServicioEmergencia)
    End If
    
    Set Telefonos = mTelefonos

    Exit Property
TelefonosErr:
    Call ErrorLog("ServicioEmergencia", "Telefonos Property Get", Err.Description)
End Property

'##ModelId=46A97800033D
Public Property Set Telefonos(ByVal vNewValue As TelefonoManager)
    On Error GoTo TelefonosErr
    
    Set mTelefonos = vNewValue
    If Not mTelefonos Is Nothing Then
        mTelefonos.OwnerId = mId
        mTelefonos.OwnerType = eOTServicioEmergencia
    End If
    
    Exit Property
TelefonosErr:
    Call ErrorLog("ServicioEmergencia", "Telefonos Property Set", Err.Description)
End Property

'##ModelId=46A97800033F
Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("ServicioEmergencia", "Direccion Property Get", Err.Description)
End Property

'##ModelId=46A978000342
Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("ServicioEmergencia", "Direccion Property Set", Err.Description)
End Property

'##ModelId=46A9778400CF
Public Property Get Afiliados() As AfiliadoExternoManager
    On Error GoTo AfiliadosErr

    If Not bAfiliadosCargados Then
        mAfiliados.CargarAfiliadoExternos Me
        bAfiliadosCargados = True
    End If
    Set Afiliados = mAfiliados

    Exit Property
AfiliadosErr:
    Call ErrorLog("ServicioEmergencia", "Afiliados Property Get", Err.Description)
End Property

'##ModelId=46A9778400D2
Public Property Set Afiliados(ByVal vNewValue As AfiliadoExternoManager)
    On Error GoTo AfiliadosErr

    '## Generated default body ...
    Set mAfiliados = vNewValue

    Exit Property
AfiliadosErr:
    Call ErrorLog("ServicioEmergencia", "Afiliados Property Set", Err.Description)
End Property
Public Property Get CodigosCubiertos() As CodigoCubiertoManager
    On Error GoTo DireccionErr

        If mCodigos Is Nothing Then
            Set mCodigos = New CodigoCubiertoManager
            mCodigos.Cargar eDCServicioEmergencia, mId
        End If
        Set CodigosCubiertos = mCodigos

    Exit Property
DireccionErr:
    Call ErrorLog("ServicioEmergencia", "CodigosCubiertos Property Get", Err.Description)
End Property

Public Property Set CodigosCubiertos(ByVal vNewValue As CodigoCubiertoManager)
    On Error GoTo DireccionErr
    
    Set mCodigos = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("ServicioEmergencia", "CodigosCubiertos Property Set", Err.Description)
End Property

Public Property Get InfoContable() As InfoContableEmp
On Error GoTo DireccionErr

        Set InfoContable = mInfoContable

    Exit Property
DireccionErr:
    Call ErrorLog("ServicioEmergencia", "InfoContable Property Get", Err.Description)
End Property

Public Property Set InfoContable(ByVal vNewValue As InfoContableEmp)
    On Error GoTo DireccionErr
    
    Set mInfoContable = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("ServicioEmergencia", "CodigosCubiertos Property Set", Err.Description)
End Property
Public Function GetProperty(pNombre As String) As String
    Dim pName As String
    pName = LCase(pNombre)
    Select Case pName
        Case "codigos"
            GetProperty = Me.CodigosCubiertos.GetCadenaCodigos
    End Select
End Function

Friend Sub Guardar()
       
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
    Me.Id = BD.getIdDisponible(tServicioEmergencia)
    params.Add "id", Me.Id
    params.Add "Nombre", Me.Nombre
    
    Me.Direccion.Guardar trans 'va aca asi se le asigna un id
    
    params.Add "idDireccion", Me.Direccion.Id
    mInfoContable.FillParameters params
    trans.AddInsertCommand tServicioEmergencia, params
    
    Me.Telefonos.GuardarTelefonos Me.Id, eOTServicioEmergencia, trans
    Me.CodigosCubiertos.Guardar eDCServicioEmergencia, trans, Me.Id
    
    trans.Execute
    
End Sub

Public Sub GuardarModificaciones()
       
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
    
    params.Add "id", Me.Id, True
    params.Add "Nombre", Me.Nombre
    
    Me.Direccion.GuardarModificaciones trans 'va aca asi se le asigna un id
    mInfoContable.FillParameters params
    trans.AddUpdateCommand tServicioEmergencia, params
    
    Me.Telefonos.SaveChanges trans
    Me.CodigosCubiertos.SaveChanges eDCServicioEmergencia, trans, Me.Id
    trans.Execute
    
End Sub

Private Sub Class_Initialize()
    Set mDireccion = New Direccion
    Set mAfiliados = New AfiliadoExternoManager
End Sub
