Attribute VB_Name = "Modulo"
Public TERR As New tbrErrores.clsTbrERR
Public CCFFGG As New tbrconfig.GlobalCls
Public BD As New DataBaseLayer.BaseDeDatos

Private mNetMonitor As NetMonitor
Private mDBMonitor As DBMonitor
Private mErrHandler As ErrHandler

'colecciones necesarias en BL
Private mCargos As CargoManager
Private mEstadosCiviles As EstadoCivilManager
Private mTiposDoc As TipoDocManager
Private mTiposTelefono As TipoTelManager
Private mBarrios As BarrioManager
Private mCiudades As CiudadManager
Private mProvincias As ProvinciaManager
Private mPaises As PaisManager
Private mParentezcos As ParentezcoManager
Private mServiciosEmergencia As ServicioEmergenciaManager
Private mCodigosEmergencia As CodigoEmergenciaManager
Private mEnfermedades As EnfermedadManager
Private mAlergias As AlergiaManager
Private mMedicamentos As MedicamentoManager
Private mLugares As LugarManager
Private mOcupaciones As OcupacionManager
Private mObrasSociales As ObraSocialManager
Private mMoviles As MovilManager
Private mEmpleados As EmpleadoManager
Private mSintomas As SintomaManager
Private mTelefonos As TelefonoManager

Private mAfiliados As AfiliadoManager
Private mAreasProtegidas As AreaProtegidaManager

Private mAtenciones As AtencionManager
Private mAtencionesB As AtencionBManager

Private mCuotas As CuotaManager
Private mEquipos As EquipoManager

Private mTiposCodigo As TipoCodigoManager

Private mInstElectricas As InstElectricaManager
Private mInstGas As InstGasManager

Private mCuerposBomberos As CuerpoBomberosManager

Public Sub ResetearBL()
    Set mAfiliados = Nothing
    Set mCargos = Nothing
    Set mTiposTelefono = Nothing
    Set mBarrios = Nothing
    Set mCiudades = Nothing
    Set mProvincias = Nothing
    Set mParentezcos = Nothing
    Set mServiciosEmergencia = Nothing
    Set mCodigosEmergencia = Nothing
    Set mEnfermedades = Nothing
    Set mAlergias = Nothing
    Set mMedicamentos = Nothing
    Set mLugares = Nothing
    Set mOcupaciones = Nothing
    Set mObrasSociales = Nothing
    Set mMoviles = Nothing
    Set mEmpleados = Nothing
    Set mSintomas = Nothing
    Set mTelefonos = Nothing
    
    Set mAreasProtegidas = Nothing
    
    Set mAtenciones = Nothing
    Set mAtencionesB = Nothing
    
    Set mCuotas = Nothing
    Set mEquipos = Nothing
    Set mTiposCodigo = Nothing
End Sub

Public Property Get CargosLocal() As CargoManager
    If mCargos Is Nothing Then
       Set mCargos = New CargoManager
       mCargos.CargarTodos
    End If
    Set CargosLocal = mCargos
End Property

Public Property Get TiposDocumentoLocal() As TipoDocManager
    If mTiposDoc Is Nothing Then
       Set mTiposDoc = New TipoDocManager
       mTiposDoc.cargarTipoDoc
    End If
    Set TiposDocumentoLocal = mTiposDoc
End Property

Public Property Get EstadosCivilesLocal() As EstadoCivilManager
    If mEstadosCiviles Is Nothing Then
       Set mEstadosCiviles = New EstadoCivilManager
       mEstadosCiviles.cargarEstadoCivil
    End If
    Set EstadosCivilesLocal = mEstadosCiviles
End Property

Public Property Get TiposTelefonoLocal() As TipoTelManager
    If mTiposTelefono Is Nothing Then
       Set mTiposTelefono = New TipoTelManager
       mTiposTelefono.CargarTodos
    End If
    Set TiposTelefonoLocal = mTiposTelefono
End Property

Public Property Get BarriosLocal() As BarrioManager
    If mBarrios Is Nothing Then
       Set mBarrios = New BarrioManager
       mBarrios.CargarTodos
    End If
    Set BarriosLocal = mBarrios
End Property

Public Property Get CiudadesLocal() As CiudadManager
    If mCiudades Is Nothing Then
       Set mCiudades = New CiudadManager
       mCiudades.CargarTodos
    End If
    Set CiudadesLocal = mCiudades
End Property

Public Property Get ProvinciasLocal() As ProvinciaManager
    If mProvincias Is Nothing Then
       Set mProvincias = New ProvinciaManager
       mProvincias.CargarTodos
    End If
    Set ProvinciasLocal = mProvincias
End Property

Public Property Get PaisesLocal() As PaisManager
    If mPaises Is Nothing Then
       Set mPaises = New PaisManager
       mPaises.CargarTodos
    End If
    Set PaisesLocal = mPaises
End Property

Public Property Get ParentezcoLocal() As ParentezcoManager
    If mParentezcos Is Nothing Then
       Set mParentezcos = New ParentezcoManager
       mParentezcos.CargarTodos
    End If
    Set ParentezcoLocal = mParentezcos
End Property

Public Property Get AfiliadosLocal() As AfiliadoManager
    If mAfiliados Is Nothing Then
        Set mAfiliados = New AfiliadoManager
        mAfiliados.CargarTodos
    End If
    
    Set AfiliadosLocal = mAfiliados
End Property

Public Property Get ServiciosEmergenciaLocal() As ServicioEmergenciaManager
    If mServiciosEmergencia Is Nothing Then
        Set mServiciosEmergencia = New ServicioEmergenciaManager
        mServiciosEmergencia.CargarTodos
    End If
    
    Set ServiciosEmergenciaLocal = mServiciosEmergencia
End Property

Public Property Get CodigoEmergenciaLocal() As CodigoEmergenciaManager
    If mCodigosEmergencia Is Nothing Then
       Set mCodigosEmergencia = New CodigoEmergenciaManager
       mCodigosEmergencia.CargarTodos
    End If
    Set CodigoEmergenciaLocal = mCodigosEmergencia
End Property

Public Property Get EnfermedadesLocal() As EnfermedadManager
    If mEnfermedades Is Nothing Then
       Set mEnfermedades = New EnfermedadManager
       mEnfermedades.CargarTodos
    End If
    Set EnfermedadesLocal = mEnfermedades
End Property

Public Property Get AlergiasLocal() As AlergiaManager
    If mAlergias Is Nothing Then
       Set mAlergias = New AlergiaManager
       mAlergias.CargarTodos
    End If
    Set AlergiasLocal = mAlergias
End Property

Public Property Get MedicamentosLocal() As MedicamentoManager
    If mMedicamentos Is Nothing Then
       Set mMedicamentos = New MedicamentoManager
       mMedicamentos.CargarTodos
    End If
    Set MedicamentosLocal = mMedicamentos
End Property

Public Property Get AreasProtegidasLocal() As AreaProtegidaManager
    If mAreasProtegidas Is Nothing Then
       Set mAreasProtegidas = New AreaProtegidaManager
       mAreasProtegidas.CargarTodos
    End If
    Set AreasProtegidasLocal = mAreasProtegidas
End Property

Public Property Get LugaresLocal() As LugarManager
    If mLugares Is Nothing Then
       Set mLugares = New LugarManager
       mLugares.CargarTodos
    End If
    Set LugaresLocal = mLugares
End Property

Public Property Get OcupacionesLocal() As OcupacionManager
    If mOcupaciones Is Nothing Then
       Set mOcupaciones = New OcupacionManager
       mOcupaciones.CargarTodos
    End If
    Set OcupacionesLocal = mOcupaciones
End Property

Public Property Get ObrasSocialesLocal() As ObraSocialManager
    If mObrasSociales Is Nothing Then
       Set mObrasSociales = New ObraSocialManager
       mObrasSociales.CargarTodos
    End If
    Set ObrasSocialesLocal = mObrasSociales
End Property

Public Property Get MovilesLocal() As MovilManager
    If mMoviles Is Nothing Then
       Set mMoviles = New MovilManager
       mMoviles.CargarTodos
    End If
    Set MovilesLocal = mMoviles
End Property

Public Property Get EmpleadosLocal() As EmpleadoManager
    If mEmpleados Is Nothing Then
       Set mEmpleados = New EmpleadoManager
       mEmpleados.CargarEmpleados
    End If
    Set EmpleadosLocal = mEmpleados
End Property

Public Property Get AtencionesLocal() As AtencionManager
    If mAtenciones Is Nothing Then
        Set mAtenciones = New AtencionManager
    End If
    Set AtencionesLocal = mAtenciones
End Property

Public Property Get AtencionesBLocal() As AtencionBManager
    If mAtencionesB Is Nothing Then
        Set mAtencionesB = New AtencionBManager
    End If
    Set AtencionesBLocal = mAtencionesB
End Property

Public Property Get CuotasByEstadoLocal(eEstado As eEstadoCuota) As CuotaManager
    'If mCuotasImpagas Is Nothing Then
    'siempre las refresco, esta mal pero me parece la unica forma de mantenerme actualizado
        Set mCuotas = Nothing
        Set mCuotas = New CuotaManager
        mCuotas.CargarCuotasByEstado (eEstado)
    'End If
    Set CuotasByEstadoLocal = mCuotas
End Property

Public Property Get SintomasLocal() As SintomaManager
    If mSintomas Is Nothing Then
        Set mSintomas = New SintomaManager
        mSintomas.CargarSintomas
    End If
    Set SintomasLocal = mSintomas
End Property

Public Property Get TelefonosLocal() As TelefonoManager
    If mTelefonos Is Nothing Then
        Set mTelefonos = New TelefonoManager
        mTelefonos.CargarTodos
    End If
    Set TelefonosLocal = mTelefonos
End Property

Public Property Get EquiposLocal() As EquipoManager
    If mEquipos Is Nothing Then
        Set mEquipos = New EquipoManager
        mEquipos.CargarTodos
    End If
    Set EquiposLocal = mEquipos
End Property

Public Property Get TiposCodigoLocal() As TipoCodigoManager
    If mTiposCodigo Is Nothing Then
       Set mTiposCodigo = New TipoCodigoManager
       mTiposCodigo.CargarTodos
    End If
    Set TiposCodigoLocal = mTiposCodigo
End Property

Public Property Get InstElectricasLocal() As InstElectricaManager
    If mInstElectricas Is Nothing Then
       Set mInstElectricas = New InstElectricaManager
       mInstElectricas.CargarTodos
    End If
    Set InstElectricasLocal = mInstElectricas
End Property

Public Property Get InstalacionesGasLocal() As InstGasManager
    If mInstGas Is Nothing Then
       Set mInstGas = New InstGasManager
       mInstGas.CargarTodos
    End If
    Set InstalacionesGasLocal = mInstGas
End Property

Public Property Get CuerposDeBomberosLocal() As CuerpoBomberosManager
    If mCuerposBomberos Is Nothing Then
       Set mCuerposBomberos = New CuerpoBomberosManager
       mCuerposBomberos.CargarTodos
    End If
    Set CuerposDeBomberosLocal = mCuerposBomberos
End Property

Public Property Get NetMonitorLocal() As NetMonitor
    If mNetMonitor Is Nothing Then Set mNetMonitor = New NetMonitor
    Set NetMonitorLocal = mNetMonitor
End Property

Public Property Get DBMonitorLocal() As DBMonitor
    If mDBMonitor Is Nothing Then Set mDBMonitor = New DBMonitor
    Set DBMonitorLocal = mDBMonitor
End Property

Public Property Get ErrHandlerLocal() As ErrHandler
    If mErrHandler Is Nothing Then Set mErrHandler = New ErrHandler
    Set ErrHandlerLocal = mErrHandler
End Property

'evitar valores no nulos
Public Function NN_str(val, Optional default As String = "") As String
    If IsNull(val) Then
        NN_str = default
    Else
        NN_str = CStr(val)
    End If
End Function

Public Function NN_num(val, Optional default As Long = 0) As Long
    If IsNull(val) Then
        NN_num = default
    Else
        If IsNumeric(val) Then
            NN_num = CLng(val)
        Else
            NN_num = default
        End If
    End If
End Function

