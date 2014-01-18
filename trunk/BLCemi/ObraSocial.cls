VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObraSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FEA40251"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A8044C0177
Private mId As Long

'##ModelId=46A804CB00DA
Private mNombre As String

'##ModelId=46A80D6E0138
Private mTelefonos As TelefonoManager

'##ModelId=46A80D79007D
Private mServicioEmergencia As ServicioEmergencia

'##ModelId=46A80F48033C
Private mObservaciones As String

'##ModelId=46A80F5C004E
Private mAfiliados As AfiliadoExternoManager

'##ModelId=46A8114200BB
Private mCoseguro As Currency

'##ModelId=46A819D801D4
Private mDireccion As Direccion

Private mCodigos As CodigoCubiertoManager
Private mInfoContable As InfoContableEmp

Private bTelefonosCargados As Boolean
Private bAfiliadosCargados As Boolean

'##ModelId=46A9778400B1
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("ObraSocial", "Id Property Get", Err.Description)
End Property

'##ModelId=46A9778400B2
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("ObraSocial", "Id Property Let", Err.Description)
End Property

'##ModelId=46A9778400B6
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("ObraSocial", "Nombre Property Get", Err.Description)
End Property

'##ModelId=46A9778400B7
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("ObraSocial", "Nombre Property Let", Err.Description)
End Property

'##ModelId=46A9778400BB
Public Property Get Telefonos() As TelefonoManager
    On Error GoTo TelefonosErr

    If mTelefonos Is Nothing Then
       Set mTelefonos = TelefonosLocal.GetByOwner(Me.Id, eOTObraSocial)
    End If
    
    Set Telefonos = mTelefonos

    Exit Property
TelefonosErr:
    Call ErrorLog("ObraSocial", "Telefonos Property Get", Err.Description)
End Property

'##ModelId=46A9778400BE
Public Property Set Telefonos(ByVal vNewValue As TelefonoManager)
    On Error GoTo TelefonosErr
    
    Set mTelefonos = vNewValue
    If Not mTelefonos Is Nothing Then
        mTelefonos.OwnerId = mId
        mTelefonos.OwnerType = eOTObraSocial
    End If
    
    Exit Property
TelefonosErr:
    Call ErrorLog("ObraSocial", "Telefonos Property Set", Err.Description)
End Property

'##ModelId=46A9778400C0
Public Property Get ServicioEmergencia() As ServicioEmergencia
    On Error GoTo ServicioEmergenciaErr

    '## Generated default body ...
    Set ServicioEmergencia = mServicioEmergencia

    Exit Property
ServicioEmergenciaErr:
    Call ErrorLog("ObraSocial", "ServicioEmergencia Property Get", Err.Description)
End Property

'##ModelId=46A9778400C3
Public Property Set ServicioEmergencia(ByVal vNewValue As ServicioEmergencia)
    On Error GoTo ServicioEmergenciaErr

    '## Generated default body ...
    Set mServicioEmergencia = vNewValue

    Exit Property
ServicioEmergenciaErr:
    Call ErrorLog("ObraSocial", "ServicioEmergencia Property Set", Err.Description)
End Property

'##ModelId=46A9778400CA
Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("ObraSocial", "Observaciones Property Get", Err.Description)
End Property

'##ModelId=46A9778400CB
Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("ObraSocial", "Observaciones Property Let", Err.Description)
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
    Call ErrorLog("ObraSocial", "Afiliados Property Get", Err.Description)
End Property

'##ModelId=46A9778400D2
Public Property Set Afiliados(ByVal vNewValue As AfiliadoExternoManager)
    On Error GoTo AfiliadosErr

    bAfiliadosCargados = True
    Set mAfiliados = vNewValue

    Exit Property
AfiliadosErr:
    Call ErrorLog("ObraSocial", "Afiliados Property Set", Err.Description)
End Property

'##ModelId=46A9778400D4
Public Property Get Coseguro() As Currency
    On Error GoTo CoseguroErr

    '## Generated default body ...
    Coseguro = mCoseguro

    Exit Property
CoseguroErr:
    Call ErrorLog("ObraSocial", "Coseguro Property Get", Err.Description)
End Property

'##ModelId=46A9778400D5
Public Property Let Coseguro(ByVal vNewValue As Currency)
    On Error GoTo CoseguroErr

    '## Generated default body ...
    mCoseguro = vNewValue

    Exit Property
CoseguroErr:
    Call ErrorLog("ObraSocial", "Coseguro Property Let", Err.Description)
End Property

'##ModelId=46A9778400D9
Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("ObraSocial", "Direccion Property Get", Err.Description)
End Property

'##ModelId=46A9778400DC
Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("ObraSocial", "Direccion Property Set", Err.Description)
End Property

Public Property Get CodigosCubiertos() As CodigoCubiertoManager
    On Error GoTo DireccionErr

        If mCodigos Is Nothing Then
            Set mCodigos = New CodigoCubiertoManager
            mCodigos.Cargar eDCObraSocial, mId
        End If
        Set CodigosCubiertos = mCodigos

    Exit Property
DireccionErr:
    Call ErrorLog("ObraSocial", "CodigosCubiertos Property Get", Err.Description)
End Property

Public Property Set CodigosCubiertos(ByVal vNewValue As CodigoCubiertoManager)
    On Error GoTo DireccionErr
    
    Set mCodigos = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("ObraSocial", "CodigosCubiertos Property Set", Err.Description)
End Property

Public Property Get InfoContable() As InfoContableEmp
On Error GoTo DireccionErr

        Set InfoContable = mInfoContable

    Exit Property
DireccionErr:
    Call ErrorLog("ObraSocial", "InfoContable Property Get", Err.Description)
End Property

Public Property Set InfoContable(ByVal vNewValue As InfoContableEmp)
    On Error GoTo DireccionErr
    
    Set mInfoContable = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("ObraSocial", "CodigosCubiertos Property Set", Err.Description)
End Property

Friend Function Insert()
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
    
    Id = BD.getIdDisponible(tObraSocial)
    Direccion.Guardar trans 'aca vuelve con un id valido
    
    params.Add "Coseguro", Coseguro
    params.Add "idDireccion", Direccion.Id
    params.Add "id", Id
    params.Add "Nombre", Nombre
    params.Add "Observaciones", Observaciones
    params.Add "idServicioEmergencia", ServicioEmergencia.Id
    mInfoContable.FillParameters params
        
    BD.Transaccion.AddInsertCommand tObraSocial, params
    
    Me.Telefonos.GuardarTelefonos Me.Id, eOTObraSocial, trans
    Me.CodigosCubiertos.Guardar eDCObraSocial, trans, Me.Id
    BD.Transaccion.Execute
End Function

Public Function GuardarModificaciones()
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
    
    Direccion.GuardarModificaciones trans
    
    params.Add "Coseguro", Coseguro
    params.Add "id", Id, True
    params.Add "Nombre", Nombre
    params.Add "Observaciones", Observaciones
    params.Add "idServicioEmergencia", ServicioEmergencia.Id
    mInfoContable.FillParameters params
    BD.Transaccion.AddUpdateCommand tObraSocial, params
    
    Me.Telefonos.SaveChanges trans
    Me.CodigosCubiertos.SaveChanges eDCObraSocial, trans, Me.Id
    'ver que esta pasando aca que me dice q los cabois harian claves repetidas
    BD.Transaccion.Execute
End Function


Public Function GetProperty(pNombre As String) As String
    Dim pName As String
    pName = LCase(pNombre)
    Select Case pName
        Case "servicio"
            'andres marzo 2010
            If (Me.ServicioEmergencia Is Nothing) Then
                GetProperty = "No hay servicio asignado!"
            Else
                GetProperty = Me.ServicioEmergencia.Nombre
            End If
        Case "codigos"
            GetProperty = Me.CodigosCubiertos.GetCadenaCodigos
    End Select
End Function

Private Sub Class_Initialize()
    Set mDireccion = New Direccion
    'Set mCodigos = New CodigoCubiertoManager
    Set mAfiliados = New AfiliadoExternoManager
End Sub
