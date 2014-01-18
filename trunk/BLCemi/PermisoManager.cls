VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PermisoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A8E11B0213"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Enum ePermisos
    ConsultarAfiliados = 1
    AltaAfiliado = 2
    ModificacionAfiliado = 3
    BajaAfiliado = 4
    
    ConsultarAreaProtegida = 5
    AltaAreaProtegida = 6
    BajaAreaProtegida = 7
    ModificacionAreaProtegida = 8

    ConsultarEmpleado = 9
    AltaEmpleado = 10
    BajaEmpleado = 11
    ModificacionEmpleado = 12
    AsignarPermisosEmpleado = 13

    ConsultarObraSocial = 14
    AltaObraSocial = 15
    BajaObraSocial = 16
    ModificacionObraSocial = 17

    ConsultarPago = 18
    AltaPago = 19
    RegistrarDevolucionRecibosAnulados = 20

    ConsultarServicioEmergencia = 21
    AltaServicioEmergencia = 22
    BajaServicioEmergencia = 23
    ModificacionServicioEmergencia = 24

    ConsultarAtencion = 25
    AltaAtencion = 26
    BajaAtencion = 27
    ModificacionAtencion = 28

    ConsultarEquipo = 29
    AltaEquipo = 30
    ModificacionEquipo = 31
    
    ConsultarMovil = 32
    AltaMovil = 33
    BajaMovil = 34
    ModificacionMovil = 35

    'ConsultarCargo = 36
    AltaCargo = 36
    BajaCargo = 37
    ModificacionCargo = 38

    'ConsultarMedicamento = 40
    AltaMedicamento = 39
    BajaMedicamento = 40
    ModificacionMedicamento = 41

    'ConsultarEnfermedad = 44
    AltaEnfermedad = 42
    BajaEnfermedad = 43
    ModificacionEnfermedad = 44

    'ConsultarAlergia = 48
    AltaAlergia = 45
    BajaAlergia = 46
    ModificacionAlergia = 47

    'ConsultarTipoTelefono = 52
    AltaTipoTelefono = 48
    BajaTipoTelefono = 49
    ModificacionTipoTelefono = 50

    'ConsultarOcupacion = 56
    AltaOcupacion = 51
    BajaOcupacion = 52
    ModificacionOcupacion = 53

    'ConsultarParentezco = 60
    AltaParentezco = 54
    BajaParentezco = 55
    ModificacionParentezco = 56

    'ConsultarLugar = 64
    AltaLugar = 57
    BajaLugar = 58
    ModificacionLugar = 59

    'ConsultarProvincia = 68

    'ConsultarCiudad
    AltaCiudadBarrio = 60
    BajaCiudadBarrio = 61
    ModificacionCiudadBarrio = 62
    
    EmitirListadoPagos = 70
    
    'configuracion
    ConfigurarApariencia = 75
    ConfigurarRed = 76
    ConfigurarBaseDatos = 77
    ConfigurarComportamiento = 78
    
    SetearNumeroRecibo = 79
    
    'agregado 09/01/2009
    BajaEquipo = 80
    
    AltaTipoCodigo = 88
    BajaTipoCodigo = 81
    ModificarTipoCodigo = 82
    
    ConfigurarCodigo = 83
    
    RegistrarGuardia = 84
    LiquidacionEmpleado = 85
    LiquidacionEmpresa = 86
    
    ConfigurarDefaults = 87
    'el 88 esta usuado mas arriba
    
    ConsultarLiquidacionEmpleado = 89
    ConsultarLiquidacionEmpresas = 90
    
    VerInformacionContableAtencion = 91
End Enum

Private mPermisos As String
Private mIsSuperUser As Boolean

'Devuelve true si tiene permiso para la funcion pasada por parametro
'##ModelId=46A8E1440138a
Public Function Can(pTask As ePermisos) As Boolean
    On Error GoTo CanErr
    
    If mIsSuperUser Then
        Can = True
    Else
        Can = InStr(1, mPermisos, "-" + Trim(Str(pTask)) + "-") <> 0
    End If
    
    Exit Function
CanErr:
    Call ErrorLog("PermisoManager", "Can Function", err.Description)
End Function

'##ModelId=46A8E2AE0222
Public Sub Grant(pTask As ePermisos)
    On Error GoTo GrantErr

    If Not Can(pTask) Then mPermisos = mPermisos + Trim(Str(pTask)) + "-"

    Exit Sub
GrantErr:
    Call ErrorLog("PermisoManager", "Grant Sub", err.Description)
End Sub

'##ModelId=46A8E2BE01F4
Public Sub Revoke(pTask As ePermisos)
    On Error GoTo RevokeErr

    If Can(pTask) Then mPermisos = Replace(mPermisos, "-" + Trim(Str(pTask)) + "-", "-")

    Exit Sub
RevokeErr:
    Call ErrorLog("PermisoManager", "Revoke Sub", err.Description)
End Sub

''Carga los permisos del empleado pasado por id.
''##ModelId=46A8E39601A5
'Friend Sub Load(pId As Long)
'    On Error GoTo LoadErr
''hacer
'    mPermisos = "-"
'
'    Exit Sub
'LoadErr:
'    Call ErrorLog("PermisoManager", "Load Sub")
'End Sub

Friend Property Get CadenaDePermisos() As String
    CadenaDePermisos = mPermisos
End Property

Friend Property Let CadenaDePermisos(pValue As String)
    mPermisos = pValue
End Property

Private Sub Class_Initialize()
    mPermisos = "-"
End Sub

Public Property Get EsSuperUsuario() As Boolean
    EsSuperUsuario = mIsSuperUser
End Property

Public Property Let EsSuperUsuario(pValue As Boolean)
    mIsSuperUser = pValue
End Property

