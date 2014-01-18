VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Involucrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum eRelacion
    eAfectado = 1
    eSolicitante = 2
End Enum

Private mId As Long
Private mNombre As String
Private mApellido As String
Private mTipoDoc As TipoDoc
Private mNroDoc As Long
Private mFechaNacimiento As Date
Private mSexo As Integer
Private mObservaciones As String
Private mDireccion As Direccion
'para saber si el año es edad o es fecha nac
Private mEsEdad As Boolean
Private mRelacion As eRelacion
Private mDescripcionRelacion As String
Private mTelefono As String

Public Property Get Telefono() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Telefono = mTelefono

    Exit Property
nombreErr:
    Call ErrorLog("Telefono", "Telefono Property Get", err.Description)
End Property

Public Property Let Telefono(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mTelefono = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Telefono", "Telefono Property Let", err.Description)
End Property

Public Property Get DescripcionRelacion() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    DescripcionRelacion = mDescripcionRelacion

    Exit Property
nombreErr:
    Call ErrorLog("Telefono", "DescripcionRelacion Property Get", err.Description)
End Property

Public Property Let DescripcionRelacion(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mDescripcionRelacion = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Telefono", "DescripcionRelacion Property Let", err.Description)
End Property

Public Property Get TipoRelacion() As eRelacion
    On Error GoTo idErr

    TipoRelacion = mRelacion

    Exit Property
idErr:
    Call ErrorLog("Telefono", "Id Property Get", err.Description)
End Property

Public Property Let TipoRelacion(ByVal vNewValue As eRelacion)
    On Error GoTo idErr

    mRelacion = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Telefono", "Id Property Let", err.Description)
End Property

Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Telefono", "Id Property Get", err.Description)
End Property

Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Telefono", "Id Property Let", err.Description)
End Property

Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("Telefono", "Nombre Property Get", err.Description)
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("Telefono", "Nombre Property Let", err.Description)
End Property

Public Property Get Apellido() As String
    On Error GoTo ApellidoErr

    '## Generated default body ...
    Apellido = mApellido

    Exit Property
ApellidoErr:
    Call ErrorLog("Telefono", "Apellido Property Get", err.Description)
End Property

Public Property Let Apellido(ByVal vNewValue As String)
    On Error GoTo ApellidoErr

    '## Generated default body ...
    mApellido = vNewValue

    Exit Property
ApellidoErr:
    Call ErrorLog("Telefono", "Apellido Property Let", err.Description)
End Property

Public Property Get TipoDoc() As TipoDoc
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set TipoDoc = mTipoDoc

    Exit Property
TipoDocErr:
    Call ErrorLog("Telefono", "TipoDoc Property Get", err.Description)
End Property

Public Property Set TipoDoc(ByVal vNewValue As TipoDoc)
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set mTipoDoc = vNewValue

    Exit Property
TipoDocErr:
    Call ErrorLog("Telefono", "TipoDoc Property Let", err.Description)
End Property

Public Property Get NroDoc() As Long
    On Error GoTo NroDocErr

    '## Generated default body ...
    NroDoc = mNroDoc

    Exit Property
NroDocErr:
    Call ErrorLog("Telefono", "NroDoc Property Get", err.Description)
End Property

Public Property Let NroDoc(ByVal vNewValue As Long)
    On Error GoTo NroDocErr

    '## Generated default body ...
    mNroDoc = vNewValue

    Exit Property
NroDocErr:
    Call ErrorLog("Telefono", "NroDoc Property Let", err.Description)
End Property

Public Property Get FechaNacimiento() As Date
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    FechaNacimiento = mFechaNacimiento

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("Telefono", "FechaNacimiento Property Get", err.Description)
End Property

Public Property Let FechaNacimiento(ByVal vNewValue As Date)
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    mFechaNacimiento = vNewValue

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("Telefono", "FechaNacimiento Property Let", err.Description)
End Property

Public Property Get Sexo() As Integer
    On Error GoTo SexoErr

    '## Generated default body ...
    Sexo = mSexo

    Exit Property
SexoErr:
    Call ErrorLog("Telefono", "Sexo Property Get", err.Description)
End Property

Public Property Let Sexo(ByVal vNewValue As Integer)
    On Error GoTo SexoErr

    '## Generated default body ...
    mSexo = vNewValue

    Exit Property
SexoErr:
    Call ErrorLog("Telefono", "Sexo Property Let", err.Description)
End Property

Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("Telefono", "Observaciones Property Get", err.Description)
End Property

Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("Telefono", "Observaciones Property Let", err.Description)
End Property

Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("Telefono", "Direccion Property Get", err.Description)
End Property

Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("Telefono", "Direccion Property Set", err.Description)
End Property

'Public Property Get Telefonos() As TelefonoManager
'    On Error GoTo TelefonosErr
'
'    If mTelefonos Is Nothing Then
'        Set mTelefonos = TelefonosLocal.GetByOwner(Me.Id, eOTAfiliadoExterno)
'    End If
'
'    Set Telefonos = mTelefonos
'
'    Exit Property
'TelefonosErr:
'    Call ErrorLog("Telefono", "Telefonos Property Get", err.Description )
'End Property
'
'Public Property Set Telefonos(ByVal vNewValue As TelefonoManager)
'    On Error GoTo TelefonosErr
'
'    Set mTelefonos = vNewValue
'    If Not mTelefonos Is Nothing Then
'        mTelefonos.OwnerId = Me.Id
'        mTelefonos.OwnerType = eOTAfiliadoExterno
'    End If
'
'    Exit Property
'TelefonosErr:
'    Call ErrorLog("Telefono", "Telefonos Property Set", err.Description )
'End Property

Public Function Edad() As Integer
On Error Resume Next
Edad = Year(Date) - Year(mFechaNacimiento)
End Function

Public Property Get EsEdad() As Boolean
EsEdad = mEsEdad
End Property

Public Property Let EsEdad(pValue As Boolean)
mEsEdad = pValue
End Property

Public Function GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "pgdireccion"
            GetProperty = mDireccion.GetBarrioCiudadProvincia + " - Calle: " + mDireccion.Calle + " " + mDireccion.Nro
    End Select
End Function

Private Sub Class_Initialize()
    Set mDireccion = New Direccion
End Sub

Public Property Get NombreCompleto() As String
    NombreCompleto = mApellido + ", " + mNombre
End Property

Friend Sub Save(trans As Transaccion, pIdSiniestro As Long)
    
    Dim params As New ParametrosManager

    If Me.Id = 0 Then
        
        Direccion.Guardar trans
        Me.Id = BD.getIdDisponible(tInvolucrado)
        params.Add "idSiniestro", pIdSiniestro
        LlenarParametros params
        trans.AddInsertCommand tInvolucrado, params
        
    Else
        
        Direccion.GuardarModificaciones trans
        LlenarParametros params
        trans.AddUpdateCommand tInvolucrado, params

    End If
    
End Sub

Private Sub LlenarParametros(pParametros As ParametrosManager)
With pParametros
    .Add "Apellido", Apellido
    .Add "FechaNacimiento", FechaNacimiento
    .Add "id", Id, True
    .Add "Nombre", Nombre
    .Add "NroDoc", NroDoc
    .Add "Observaciones", Observaciones
    .Add "Sexo", Sexo
    .Add "TipoDoc", TipoDoc.Id
    .Add "idDireccion", Direccion.Id
    .Add "esEdad", IIf(EsEdad, 1, 0)
    .Add "idTipoRelacion", IIf(mRelacion = eAfectado, 1, 2)
    .Add "DescripcionRelacion", mDescripcionRelacion
    .Add "Telefono", mTelefono
End With
End Sub


