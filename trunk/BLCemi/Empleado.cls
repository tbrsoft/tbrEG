VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Empleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A7FF770186"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

'##ModelId=46A8044E0000
Private mId As Long

'##ModelId=46A8044F029F
Private mNombre As String

'##ModelId=46A8045101B5
Private mApellido As String

'##ModelId=46A8045301C5
Private mTipoDoc As TipoDoc

'##ModelId=46A804550138
Private mNroDoc As Long

'##ModelId=46A8045B00BB
Private mFechaNacimiento As Date

'##ModelId=46A8046101F4
Private mTelefonos As TelefonoManager

'##ModelId=46A81056002E
Private mCargos As CargoManager

'##ModelId=46A8197F002E
Private mDireccion As Direccion

'##ModelId=46A8E4A000DA
Private mLogin As String

'##ModelId=46A8E4A70157
Private mPass As String

'##ModelId=46A8E4F9004E
Private mPermisos As PermisoManager

Private mMP As String

Private mGuardias As GuardiaManager

'##ModelId=46A976A503DA
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Empleado", "Id Property Get", err.Description)
End Property

'##ModelId=46A976A503DB
Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Empleado", "Id Property Let", err.Description)
End Property

'##ModelId=46A976A503DF
Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Empleado", "Nombre Property Get", err.Description)
End Property

'##ModelId=46A976A503E0
Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Empleado", "Nombre Property Let", err.Description)
End Property

'##ModelId=46A976A503E4
Public Property Get Apellido() As String
    On Error GoTo ApellidoErr

    '## Generated default body ...
    Apellido = mApellido

    Exit Property
ApellidoErr:
    Call ErrorLog("Empleado", "Apellido Property Get", err.Description)
End Property

'##ModelId=46A976A503E5
Public Property Let Apellido(ByVal vNewValue As String)
    On Error GoTo ApellidoErr

    '## Generated default body ...
    mApellido = vNewValue

    Exit Property
ApellidoErr:
    Call ErrorLog("Empleado", "Apellido Property Let", err.Description)
End Property

'##ModelId=46A976A503E9
Public Property Get TipoDoc() As TipoDoc
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set TipoDoc = mTipoDoc

    Exit Property
TipoDocErr:
    Call ErrorLog("Empleado", "TipoDoc Property Get", err.Description)
End Property

'##ModelId=46A976A503EA
Public Property Set TipoDoc(ByVal vNewValue As TipoDoc)
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set mTipoDoc = vNewValue

    Exit Property
TipoDocErr:
    Call ErrorLog("Empleado", "TipoDoc Property Let", err.Description)
End Property

'##ModelId=46A976A503EE
Public Property Get NroDoc() As Long
    On Error GoTo NroDocErr

    '## Generated default body ...
    NroDoc = mNroDoc

    Exit Property
NroDocErr:
    Call ErrorLog("Empleado", "NroDoc Property Get", err.Description)
End Property

'##ModelId=46A976A503EF
Public Property Let NroDoc(ByVal vNewValue As Long)
    On Error GoTo NroDocErr

    '## Generated default body ...
    mNroDoc = vNewValue

    Exit Property
NroDocErr:
    Call ErrorLog("Empleado", "NroDoc Property Let", err.Description)
End Property

'##ModelId=46A976A503F3
Public Property Get FechaNacimiento() As Date
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    FechaNacimiento = mFechaNacimiento

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("Empleado", "FechaNacimiento Property Get", err.Description)
End Property

'##ModelId=46A976A503F4
Public Property Let FechaNacimiento(ByVal vNewValue As Date)
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    mFechaNacimiento = vNewValue

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("Empleado", "FechaNacimiento Property Let", err.Description)
End Property

'##ModelId=46A976A503F8
Public Property Get Telefonos() As TelefonoManager
    On Error GoTo TelefonosErr
    
    If mTelefonos Is Nothing Then
        Set mTelefonos = TelefonosLocal.GetByOwner(mId, eOTEmpleado)
    End If
    
    Set Telefonos = mTelefonos

    Exit Property
TelefonosErr:
    Call ErrorLog("Empleado", "Telefonos Property Get", err.Description)
End Property

'##ModelId=46A976A503FB
Public Property Set Telefonos(ByVal vNewValue As TelefonoManager)
    On Error GoTo TelefonosErr
        
    Set mTelefonos = vNewValue
    If Not mTelefonos Is Nothing Then
        mTelefonos.OwnerId = mId
        mTelefonos.OwnerType = eOTEmpleado
    End If
    
    Exit Property
TelefonosErr:
    Call ErrorLog("Empleado", "Telefonos Property Set", err.Description)
End Property

'##ModelId=46A976A503FD
Public Property Get Cargos() As CargoManager
    On Error GoTo CargosErr

    '## Generated default body ...
    Set Cargos = mCargos

    Exit Property
CargosErr:
    Call ErrorLog("Empleado", "Cargos Property Get", err.Description)
End Property

'##ModelId=46A976A50400
Public Property Set Cargos(ByVal vNewValue As CargoManager)
    On Error GoTo CargosErr

    '## Generated default body ...
    Set mCargos = vNewValue

    Exit Property
CargosErr:
    Call ErrorLog("Empleado", "Cargos Property Set", err.Description)
End Property

'##ModelId=46A976A50402
Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("Empleado", "Direccion Property Get", err.Description)
End Property

'##ModelId=46A976A50405
Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("Empleado", "Direccion Property Set", err.Description)
End Property

'##ModelId=46A976A50407
Public Property Get Login() As String
    On Error GoTo LoginErr

    '## Generated default body ...
    Login = mLogin

    Exit Property
LoginErr:
    Call ErrorLog("Empleado", "Login Property Get", err.Description)
End Property

'##ModelId=46A976A50408
Public Property Let Login(ByVal vNewValue As String)
    On Error GoTo LoginErr

    '## Generated default body ...
    mLogin = vNewValue

    Exit Property
LoginErr:
    Call ErrorLog("Empleado", "Login Property Let", err.Description)
End Property

'##ModelId=46A976A5040C
Public Property Get Pass() As String
    On Error GoTo PassErr

    '## Generated default body ...
    Pass = mPass

    Exit Property
PassErr:
    Call ErrorLog("Empleado", "Pass Property Get", err.Description)
End Property

'##ModelId=46A976A5040D
Public Property Let Pass(ByVal vNewValue As String)
    On Error GoTo PassErr

    '## Generated default body ...
    mPass = vNewValue

    Exit Property
PassErr:
    Call ErrorLog("Empleado", "Pass Property Let", err.Description)
End Property

'##ModelId=46A976A50411
Public Property Get Permisos() As PermisoManager
    On Error GoTo PermisosErr
    
    Set Permisos = mPermisos

    Exit Property
PermisosErr:
    Call ErrorLog("Empleado", "Permisos Property Get", err.Description)
End Property

'##ModelId=46A976A50414
Public Property Set Permisos(ByVal vNewValue As PermisoManager)
    On Error GoTo PermisosErr

    '## Generated default body ...
    Set mPermisos = vNewValue

    Exit Property
PermisosErr:
    Call ErrorLog("Empleado", "Permisos Property Set", err.Description)
End Property

Public Property Get MP() As String
    MP = mMP
End Property

Public Property Let MP(pValue As String)
    mMP = pValue
End Property

Public Property Get Guardias() As GuardiaManager
    If mGuardias Is Nothing Then
        Set mGuardias = New GuardiaManager
        mGuardias.LoadNoRendidasByEmpleado Me.Id
    End If
    Set Guardias = mGuardias
End Property

Public Property Set Guardias(pGuardias As GuardiaManager)
    Set mGuardias = pGuardias
End Property

Public Property Get NombreCompleto() As String
    NombreCompleto = mApellido + ", " + mNombre
End Property

Public Function CargosToString() As String
    CargosToString = Cargos.CargosToString
End Function

Private Sub Class_Initialize()
    Set mCargos = New CargoManager
    Set mDireccion = New Direccion
    Set mPermisos = New PermisoManager
End Sub

Friend Sub DarDeBaja()
    Dim c As New Parametro
    c.Nombre = "id"
    c.Valor = mId
    BD.BorradoLogico tEmpleado, c
End Sub

Friend Sub Restaurar()
    Dim c As New Parametro
    c.Nombre = "id"
    c.Valor = mId
    BD.RestaurarElemento tEmpleado, c
End Sub

Friend Sub Guardar()
   
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
    Me.Id = BD.getIdDisponible(tEmpleado)
    params.Add "id", Me.Id
    params.Add "Apellido", Me.Apellido
    params.Add "FechaNacimiento", Me.FechaNacimiento
    params.Add "Login", Me.Login
    params.Add "Nombre", Me.Nombre
    params.Add "NroDoc", Me.NroDoc
    params.Add "Pass", Me.Pass
    params.Add "TipoDoc", Me.TipoDoc.Id
    params.Add "Permisos", Me.Permisos.CadenaDePermisos
    params.Add "EsSuperUsuario", IIf(Me.Permisos.EsSuperUsuario, 1, 0)
    params.Add "mp", Me.MP
    params.Add "eliminado", False 'cuando lo creo no esta eliminado
    Me.Direccion.Guardar trans 'va aca asi se le asigna un id
    
    params.Add "idDireccion", Me.Direccion.Id
    
    trans.AddInsertCommand tEmpleado, params
    
    Me.Telefonos.GuardarTelefonos Me.Id, eOTEmpleado, trans
    Me.Cargos.Guardar trans, Me.Id
 
    trans.Execute
End Sub

Public Sub GuardarModificaciones()
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    Dim params As New ParametrosManager

    params.Add "id", Me.Id, True
    
    params.Add "Apellido", Me.Apellido
    params.Add "FechaNacimiento", Me.FechaNacimiento
    params.Add "Login", Me.Login
    params.Add "Nombre", Me.Nombre
    params.Add "NroDoc", Me.NroDoc
    params.Add "Pass", Me.Pass
    params.Add "TipoDoc", Me.TipoDoc.Id
    params.Add "Permisos", Me.Permisos.CadenaDePermisos
    params.Add "mp", Me.MP
    'el id de direccion no va a cambiar
    trans.AddUpdateCommand tEmpleado, params
        
    Me.Telefonos.SaveChanges trans
    Me.Cargos.SaveChanges trans, Me.Id
    
    Me.Direccion.GuardarModificaciones trans
    
    trans.Execute
End Sub

