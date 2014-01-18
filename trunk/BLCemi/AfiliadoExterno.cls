VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AfiliadoExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mId As Long
Private mNroAfiliado As String
Private mNombre As String
Private mApellido As String
Private mTipoDoc As TipoDoc
Private mNroDoc As Long
Private mFechaNacimiento As Date
Private mObraSocial As ObraSocial
Private mSexo As Integer
Private mTopeAtenciones As Integer
Private mObservaciones As String
Private mDireccion As Direccion
Private mHistoriaClinica As HistoriaClinica
Private mTelefonos As TelefonoManager
Private mAtenciones As AtencionManager

Private mParent As Object
'para saber si el año es edad o es fecha nac
Private mEsEdad As Boolean

Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("AfiliadoExternor", "Id Property Get", Err.Description)
End Property

Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("AfiliadoExternor", "Id Property Let", Err.Description)
End Property

Public Property Get Nombre() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("AfiliadoExternor", "Nombre Property Get", Err.Description)
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("AfiliadoExternor", "Nombre Property Let", Err.Description)
End Property

Public Property Get NroAfiliado() As String
    On Error GoTo nombreErr

    '## Generated default body ...
    NroAfiliado = mNroAfiliado

    Exit Property
nombreErr:
    Call ErrorLog("AfiliadoExternor", "nroAfiliado Property Get", Err.Description)
End Property

Public Property Let NroAfiliado(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNroAfiliado = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("AfiliadoExternor", "NroAfiliado Property Let", Err.Description)
End Property

Public Property Get Apellido() As String
    On Error GoTo ApellidoErr

    '## Generated default body ...
    Apellido = mApellido

    Exit Property
ApellidoErr:
    Call ErrorLog("AfiliadoExternor", "Apellido Property Get", Err.Description)
End Property

Public Property Let Apellido(ByVal vNewValue As String)
    On Error GoTo ApellidoErr

    '## Generated default body ...
    mApellido = vNewValue

    Exit Property
ApellidoErr:
    Call ErrorLog("AfiliadoExternor", "Apellido Property Let", Err.Description)
End Property

Public Property Get TipoDoc() As TipoDoc
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set TipoDoc = mTipoDoc

    Exit Property
TipoDocErr:
    Call ErrorLog("AfiliadoExternor", "TipoDoc Property Get", Err.Description)
End Property

Public Property Set TipoDoc(ByVal vNewValue As TipoDoc)
    On Error GoTo TipoDocErr

    '## Generated default body ...
    Set mTipoDoc = vNewValue

    Exit Property
TipoDocErr:
    Call ErrorLog("AfiliadoExternor", "TipoDoc Property Let", Err.Description)
End Property

Public Property Get NroDoc() As Long
    On Error GoTo NroDocErr

    '## Generated default body ...
    NroDoc = mNroDoc

    Exit Property
NroDocErr:
    Call ErrorLog("AfiliadoExternor", "NroDoc Property Get", Err.Description)
End Property

Public Property Let NroDoc(ByVal vNewValue As Long)
    On Error GoTo NroDocErr

    '## Generated default body ...
    mNroDoc = vNewValue

    Exit Property
NroDocErr:
    Call ErrorLog("AfiliadoExternor", "NroDoc Property Let", Err.Description)
End Property

Public Property Get FechaNacimiento() As Date
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    FechaNacimiento = mFechaNacimiento

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("AfiliadoExternor", "FechaNacimiento Property Get", Err.Description)
End Property

Public Property Let FechaNacimiento(ByVal vNewValue As Date)
    On Error GoTo FechaNacimientoErr

    '## Generated default body ...
    mFechaNacimiento = vNewValue

    Exit Property
FechaNacimientoErr:
    Call ErrorLog("AfiliadoExternor", "FechaNacimiento Property Let", Err.Description)
End Property

Public Property Get ObraSocial() As ObraSocial
    On Error GoTo ObraSocialErr

    '## Generated default body ...
    Set ObraSocial = mObraSocial

    Exit Property
ObraSocialErr:
    Call ErrorLog("AfiliadoExternor", "ObraSocial Property Get", Err.Description)
End Property

Public Property Set ObraSocial(ByVal vNewValue As ObraSocial)
    On Error GoTo ObraSocialErr

    '## Generated default body ...
    Set mObraSocial = vNewValue

    Exit Property
ObraSocialErr:
    Call ErrorLog("AfiliadoExternor", "ObraSocial Property Set", Err.Description)
End Property

Public Property Get Sexo() As Integer
    On Error GoTo SexoErr

    '## Generated default body ...
    Sexo = mSexo

    Exit Property
SexoErr:
    Call ErrorLog("AfiliadoExternor", "Sexo Property Get", Err.Description)
End Property

Public Property Let Sexo(ByVal vNewValue As Integer)
    On Error GoTo SexoErr

    '## Generated default body ...
    mSexo = vNewValue

    Exit Property
SexoErr:
    Call ErrorLog("AfiliadoExternor", "Sexo Property Let", Err.Description)
End Property

Public Property Get TopeAtenciones() As Integer
    On Error GoTo TopeAtencionesErr

    '## Generated default body ...
    TopeAtenciones = mTopeAtenciones

    Exit Property
TopeAtencionesErr:
    Call ErrorLog("AfiliadoExternor", "TopeAtenciones Property Get", Err.Description)
End Property

Public Property Let TopeAtenciones(ByVal vNewValue As Integer)
    On Error GoTo TopeAtencionesErr

    '## Generated default body ...
    mTopeAtenciones = vNewValue

    Exit Property
TopeAtencionesErr:
    Call ErrorLog("AfiliadoExternor", "TopeAtenciones Property Let", Err.Description)
End Property

Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("AfiliadoExternor", "Observaciones Property Get", Err.Description)
End Property

Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("AfiliadoExternor", "Observaciones Property Let", Err.Description)
End Property

Public Property Get Atenciones() As AtencionManager
    On Error GoTo AtencionesErr

    Set mAtenciones = AtencionesLocal.GetByTipo(eAfiliadoExterno, Me.Id)
    Set Atenciones = mAtenciones

    Exit Property
AtencionesErr:
    Call ErrorLog("AfiliadoExternor", "Atenciones Property Get", Err.Description)
End Property

Public Property Set Atenciones(ByVal vNewValue As AtencionManager)
    On Error GoTo AtencionesErr

    '## Generated default body ...
    Set mAtenciones = vNewValue

    Exit Property
AtencionesErr:
    Call ErrorLog("AfiliadoExternor", "Atenciones Property Set", Err.Description)
End Property

Public Property Get Direccion() As Direccion
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set Direccion = mDireccion

    Exit Property
DireccionErr:
    Call ErrorLog("AfiliadoExternor", "Direccion Property Get", Err.Description)
End Property

Public Property Set Direccion(ByVal vNewValue As Direccion)
    On Error GoTo DireccionErr

    '## Generated default body ...
    Set mDireccion = vNewValue

    Exit Property
DireccionErr:
    Call ErrorLog("AfiliadoExternor", "Direccion Property Set", Err.Description)
End Property

Public Property Get Telefonos() As TelefonoManager
    On Error GoTo TelefonosErr
    
    If mTelefonos Is Nothing Then
        Set mTelefonos = TelefonosLocal.GetByOwner(Me.Id, eOTAfiliadoExterno)
    End If
    
    Set Telefonos = mTelefonos
    
    Exit Property
TelefonosErr:
    Call ErrorLog("AfiliadoExternor", "Telefonos Property Get", Err.Description)
End Property

Public Property Set Telefonos(ByVal vNewValue As TelefonoManager)
    On Error GoTo TelefonosErr

    Set mTelefonos = vNewValue
    If Not mTelefonos Is Nothing Then
        mTelefonos.OwnerId = Me.Id
        mTelefonos.OwnerType = eOTAfiliadoExterno
    End If
    
    Exit Property
TelefonosErr:
    Call ErrorLog("AfiliadoExternor", "Telefonos Property Set", Err.Description)
End Property


Public Property Get Parent() As Object
    Set Parent = mParent
End Property

Friend Property Set Parent(pValue As Object)
    Set mParent = pValue
End Property

Public Property Get HistoriaClinica() As HistoriaClinica
    
    If Not mHistoriaClinica.Cargada Then mHistoriaClinica.Cargar Me
    
    Set HistoriaClinica = mHistoriaClinica
    
End Property

Friend Property Set HistoriaClinica(pValue As HistoriaClinica)
    Set mHistoriaClinica = pValue
End Property

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
        
    End Select
End Function

Private Sub Class_Initialize()
    Set mHistoriaClinica = New HistoriaClinica
    Set mDireccion = New Direccion
End Sub

Public Property Get NombreCompleto() As String
    NombreCompleto = mApellido + ", " + mNombre
End Property

Friend Sub Insert(pParent As Object)

'terminar!!!!!
Set mParent = pParent

Dim trans As Transaccion
Set trans = BD.Transaccion
Dim params As New ParametrosManager

Direccion.Guardar trans

Me.Id = BD.getIdDisponible(tAfiliadoExterno)

LlenarParametros params

trans.AddInsertCommand tAfiliadoExterno, params

HistoriaClinica.Save trans, Me

InsertTablaIntermedia trans
Telefonos.GuardarTelefonos Me.Id, eOTAfiliadoExterno, trans

trans.Execute
End Sub

Public Sub GuardarModificaciones()

'terminar!!!!!

Dim trans As Transaccion
Set trans = BD.Transaccion
Dim params As New ParametrosManager

Direccion.GuardarModificaciones trans

LlenarParametros params

trans.AddUpdateCommand tAfiliadoExterno, params

HistoriaClinica.SaveChanges trans

Telefonos.SaveChanges trans

trans.Execute
End Sub

Private Sub InsertTablaIntermedia(pTrans As Transaccion)
    Dim pTabla As eTablas
    Dim campoID As String

    'aca le mando el campo y la tabla y esta funcion me los devuelve seteados
    GetParentType campoID, pTabla

    Dim params As ParametrosManager
    Set params = New ParametrosManager
    params.Add "idAfiliadoExterno", Me.Id
    params.Add campoID, mParent.Id
    pTrans.AddInsertCommand pTabla, params
End Sub

Private Sub GetParentType(ByRef outCampoID As String, ByRef outTabla As eTablas)
        
    If TypeOf mParent Is AreaProtegida Then
        outTabla = tAfiliadoExternoXAreaProtegida
        outCampoID = "idAreaProtegida"
    ElseIf TypeOf mParent Is ServicioEmergencia Then
        outTabla = tAfiliadoExternoXServicioEmergencia
        outCampoID = "idServicioEmergencia"
    ElseIf TypeOf mParent Is ObraSocial Then
        outTabla = tAfiliadoExternoXObraSocial
        outCampoID = "idObraSocial"
    End If

End Sub

Private Sub LlenarParametros(pParametros As ParametrosManager)
With pParametros
    .Add "Apellido", Apellido
    'Atenciones no tiene todavia
    .Add "FechaNacimiento", FechaNacimiento
    .Add "id", Id, True
    .Add "Nombre", Nombre
    .Add "NroDoc", NroDoc
    If Not ObraSocial Is Nothing Then .Add "idObraSocial", ObraSocial.Id
    .Add "Observaciones", Observaciones
    'Pagos no tiene todavia
    .Add "Sexo", Sexo
    .Add "TipoDoc", TipoDoc.Id
    .Add "TopeAtenciones", TopeAtenciones
    .Add "NroAfiliado", NroAfiliado
    .Add "idDireccion", Direccion.Id
    .Add "esEdad", IIf(EsEdad, 1, 0)
End With
End Sub

