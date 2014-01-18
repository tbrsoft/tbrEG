VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClaseGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub Init_BlCemi(Optional ptLog As String)
    
    Dim ptLog2 As String
    ptLog2 = ""
    If ptLog <> "" Then
        TERR.FileLog = ptLog + "regBLCEMI.log"
        ptLog2 = ptLog + "regDBLayer.log"
    End If
    
    TERR.Set_ADN "blCemi_v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    TERR.AppendSinHist "Init_blCemi"
    
    If ptLog2 <> "" Then
        BD.Init ptLog2
    Else
        BD.Init
    End If
End Sub

'-----------------------funciones----------------------
'esto es lo mas chongo que hay
Public Function ExecuteSQL(mSQL As String) As Object
    Set ExecuteSQL = BD.SelectSQL(mSQL)
End Function

Public Sub PrintToErrorLog(pClass As String, pFunctionOrPropertyName As String, pError As String, Optional DefinePath As String = "")
    ErrorLog pClass, pFunctionOrPropertyName, pError, DefinePath
End Sub

Public Sub MostrarEstadoRed()
    NetMonitorLocal.MiRedLocal.VerEstado
End Sub

Public Function VerificarRutaBaseDatos(pPath As String) As Boolean
     VerificarRutaBaseDatos = BD.VerificarBaseDatos(pPath)
End Function

Public Function GetDatabaseVersion() As String
    GetDatabaseVersion = BD.GetDBVersion
End Function
'es para cuando no se puede conectar
Public Function ConectarBaseDatos() As Boolean
    ConectarBaseDatos = BD.ReConectar
End Function

Public Sub ResetearColecciones()
    ResetearBL
End Sub

Public Function GetCantidadRegistros(pTabla As String) As Long
    On Error Resume Next
    Dim rs
    Set rs = BD.SelectSQL("select count(id) from " + pTabla)
    GetCantidadRegistros = rs.Fields(0).value
End Function

'-------------propiedades-----------------------------

Public Property Get DBMonitorGBL() As DBMonitor
    Set DBMonitorGBL = DBMonitorLocal
End Property

Public Property Get NetMonitorGBL() As NetMonitor
    Set NetMonitorGBL = NetMonitorLocal
End Property

Public Property Get ErrHandlerGBL() As ErrHandler
    Set ErrHandlerGBL = ErrHandlerLocal
End Property

'---------------colecciones--------------------------

Public Property Get LugaresGBL() As LugarManager
    Set LugaresGBL = LugaresLocal
End Property

Public Property Get OcupacionesGBL() As OcupacionManager
    Set OcupacionesGBL = OcupacionesLocal
End Property

Public Property Get ObrasSocialesGBL() As ObraSocialManager
    Set ObrasSocialesGBL = ObrasSocialesLocal
End Property

Public Property Get EmpleadosGBL() As EmpleadoManager
    Set EmpleadosGBL = EmpleadosLocal
End Property

Public Property Get CodigoEmergenciaGBL() As CodigoEmergenciaManager
    Set CodigoEmergenciaGBL = CodigoEmergenciaLocal
End Property

Public Property Get CargosGBL() As CargoManager
    Set CargosGBL = CargosLocal
End Property

Public Property Get TiposDocumentoGBL() As TipoDocManager
    Set TiposDocumentoGBL = TiposDocumentoLocal
End Property

Public Property Get TiposTelefonoGBL() As TipoTelManager
    Set TiposTelefonoGBL = TiposTelefonoLocal
End Property

Public Property Get ProvinciasGBL() As ProvinciaManager
    Set ProvinciasGBL = ProvinciasLocal
End Property

Public Property Get ParentezcosGBL() As ParentezcoManager
    Set ParentezcosGBL = ParentezcoLocal
End Property

Public Property Get AfiliadosGBL() As AfiliadoManager
    Set AfiliadosGBL = AfiliadosLocal
End Property

Public Property Get ServiciosEmergenciaGBL() As ServicioEmergenciaManager
    Set ServiciosEmergenciaGBL = ServiciosEmergenciaLocal
End Property

Public Property Get EnfermedadesGBL() As EnfermedadManager
    Set EnfermedadesGBL = EnfermedadesLocal
End Property

Public Property Get AlergiasGBL() As AlergiaManager
    Set AlergiasGBL = AlergiasLocal
End Property

Public Property Get MedicamentosGBL() As MedicamentoManager
    Set MedicamentosGBL = MedicamentosLocal
End Property

Public Property Get MovilesGBL() As MovilManager
    Set MovilesGBL = MovilesLocal
End Property

Public Property Get AreasProtegidasGBL() As AreaProtegidaManager
    Set AreasProtegidasGBL = AreasProtegidasLocal
End Property

Public Property Get EstadosCivilesGBL() As EstadoCivilManager
    Set EstadosCivilesGBL = EstadosCivilesLocal
End Property

Public Property Get AtencionesGBL() As AtencionManager
    Set AtencionesGBL = AtencionesLocal
End Property

Public Property Get AtencionesBGBL() As AtencionBManager
    Set AtencionesBGBL = AtencionesBLocal
End Property

Public Property Get CuotasByEstadoGBL(pEstado As eEstadoCuota) As CuotaManager
    Set CuotasByEstadoGBL = CuotasByEstadoLocal(pEstado)
End Property

Public Property Get SintomasGBL() As SintomaManager
    Set SintomasGBL = SintomasLocal
End Property

Public Property Get EquiposGBL() As EquipoManager
    Set EquiposGBL = EquiposLocal
End Property

Public Property Get TelefonosGBL() As TelefonoManager
    Set TelefonosGBL = TelefonosLocal
End Property

Public Property Get TiposCodigoGBL() As TipoCodigoManager
    Set TiposCodigoGBL = TiposCodigoLocal
End Property

Public Property Get InstElectricasGBL() As InstElectricaManager
    Set InstElectricasGBL = InstElectricasLocal
End Property

Public Property Get InstalacionesGasGBL() As InstGasManager
    Set InstalacionesGasGBL = InstalacionesGasLocal
End Property

Public Property Get CuerposDeBomberosGBL()
    Set CuerposDeBomberosGBL = CuerposDeBomberosLocal
End Property

Public Property Get PaisesGBL() As PaisManager
    Set PaisesGBL = PaisesLocal
End Property
