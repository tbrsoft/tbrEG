VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BaseDeDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eTablas
    tAfiliado = 1
    tAlergia = 2
    tAreaProtegida = 3
    tAtencion = 4
    tBarrio = 5
    tCargo = 6
    tCiudad = 7
    tCodigoEmergencia = 8
    tDireccion = 9
    tEmpleado = 10
    tEnfermedad = 11
    tLugar = 12
    tMovil = 13
    tObraSocial = 14
    tOcupacion = 15
    tCuota = 16
    tParentezco = 17
    tProvincia = 18
    tServicioEmergencia = 19
    tSintoma = 20
    tTelefono = 21
    tTipoTelefono = 22
    tVehiculo = 23
    tMedicamento = 24
    'tablas secundarias
    tTelefonoXAfiliado = 25
    tTelefonoXEmpleado = 26
    tTelefonoXAreaProtegida = 27
    tTelefonoXServicioEmergencia = 28
    tCargoXEmpleado = 29
    
    tTelefonoXObraSocial = 30
    tCodigoXEmpresa = 31
    
    tAlergiaXAfiliado = 32
    tEnfermedadXAfiliado = 33
    tMedicamentoXAfiliado = 34
    
    tAfiliadoExterno = 35
    
    tAlergiaXAfiliadoExterno = 36
    tEnfermedadXAfiliadoExterno = 37
    tMedicamentoXAfiliadoExterno = 38
    
    tAfiliadoExternoXObraSocial = 39
    tAfiliadoExternoXServicioEmergencia = 40
    tAfiliadoExternoXAreaProtegida = 41
    
    tCuotasAnuladas = 42
    tEquipo = 43
    tEmpleadoXEquipo = 44
    tEquipoXAtencion = 45
    
    tRecibo = 46
    
    tDetalleSeguimiento = 47
    tTipoCodigo = 48
    
    tLiquidacionEmpresa = 49
    tGuardia = 50
    tLiquidacionEmpleado = 51
    
'---------Bomberos-----------------
    
    tAtencionB = 52
    tInvolucrado = 53
    tInstElectrica = 54
    tInstGas = 55
    
    tPais = 56
    
    tCuerpoBomberos = 57
    tResponsableCuerpo = 58
    tUnidadCuerpo = 59
    tColaboracionCuerpo = 60
End Enum

Dim cn As Connection
Private mTransaction As Transaccion
Private mDBPath As String
Private WithEvents mConfig As clsConfiguracion
Attribute mConfig.VB_VarHelpID = -1
Private mDBEvents As New DBEvents

Public Property Get DBEventsManager() As DBEvents
    Set DBEventsManager = mDBEvents
End Property

Public Property Get Transaccion() As Transaccion
    Set Transaccion = mTransaction
End Property

Friend Property Get Conexion() As Connection
    Set Conexion = cn
End Property

Private Sub mConfig_ConfigChanged()
'si cambio la configuracion me fijo si sigue el mismo path...
If mDBPath <> mConfig.DBLayer.PathDB Then
    mDBPath = mConfig.DBLayer.PathDB
    Conectar
End If
End Sub

Private Sub Class_Initialize()
    'se creo el sub Init por andres 2010 03 17
    'da ocote que este aqui
End Sub

Public Sub Init(Optional ptLog As String)
    
    If ptLog <> "" Then TERR.FileLog = ptLog
    TERR.Set_ADN "dblayer_v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    TERR.AppendSinHist "Init_DBLayer"
    
    Set mConfig = CCFFGG.Configuracion 'aqui se crea el objeto config global
    mDBPath = CCFFGG.Configuracion.DBLayer.PathDB
    TERR.AppendSinHist ".:\|/:." + mDBPath
    Conectar
    Set mTransaction = New Transaccion
    mTransaction.Inicializar Me
    
    'aca reviso la version y hago el upgrade de ser necesario
    UpgradeDatabase
End Sub

Public Sub Conectar()
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cn = Nothing
    Set cn = New Connection
    Dim cs As String
    cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mDBPath + ";Persist Security Info=False;Jet OLEDB:Database Password=zapato"
    cn.ConnectionString = cs
    On Error GoTo errman
    cn.Open
    Exit Sub
    
errman:
    TERR.AppendLog "erConn", TERR.ErrToTXT(Err)
    
   Select Case Err.Number
   Case -2147467259:
        mDBEvents.FireErrorEvent eDBCantFindDB, Err.Description
   Case Else
        mDBEvents.FireErrorEvent eUnknown, Err.Description
   End Select
End Sub

'para intentar reconectar si no se encuentra la bd, no levanta eventos

Public Function ReConectar() As Boolean
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cn = Nothing
    Set cn = New Connection
    Dim cs As String
    cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mDBPath + ";Persist Security Info=False;Jet OLEDB:Database Password=zapato"
    cn.ConnectionString = cs
    On Error GoTo errman
    cn.Open
    ReConectar = True
    Exit Function
errman:
   ReConectar = False
End Function

Public Function VerificarBaseDatos(pPath As String) As Boolean
    On Error GoTo errman
    Dim cnn As New Connection
    Dim cs As String
    cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pPath + ";Persist Security Info=False;Jet OLEDB:Database Password=zapato"
    cnn.ConnectionString = cs
    cnn.Open
    Dim rs As Recordset
    'aca se podria usar "version"
    Set rs = cnn.Execute("select * from Afiliado where 1=0")
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
    VerificarBaseDatos = True
    Exit Function
errman:
    VerificarBaseDatos = False
End Function

Public Function GetDBVersion() As String
    Dim rs As Recordset
    Set rs = cn.Execute("select * from version")
    GetDBVersion = Trim(Str(rs!principal)) + "." + Trim(Str(rs!secundario)) + "." + Trim(Str(rs!Revision))
    Set rs = Nothing
End Function

'Esta funcion modifica la base de datos para adaptarla a la version correspondiente
Public Function UpgradeDatabase()
    'VERSION ACTUAL 2.0.0
    If GetDBVersion = "1.0.0" Then
        'agrego la columna que falta
        cn.Execute "ALTER table Atencion add NroIncidenteInterno text(10) "
        'actualizo la version
        cn.Execute "Update version set principal=1, secundario=0, revision=1"
    End If
    'es sumativo, asi q si era 1.0.0 aca llega 1.0.1
    If GetDBVersion = "1.0.1" Then
        cn.Execute "ALTER table CodigoEmergencia add bold text(6)"
        cn.Execute "ALTER table CodigoEmergencia add colorfuente long"
        'le asigno valores por defecto
        cn.Execute "UPDATE CodigoEmergencia set bold='false',colorfuente=0"
        'actualizo la version
        cn.Execute "Update version set principal=1, secundario=0, revision=2"
    End If
    
    If GetDBVersion = "1.0.2" Then
        'agrego la tabla atencionb
        Dim mSql As String
        mSql = "CREATE TABLE AtencionB (id long, NroIncidente text(20), Fecha datetime, HoraLlamada text(50), SalidaPreInspeccion text(50), LlegadaPreinspeccion text(50), SalidaDotacion text(50), QTH text(50), VL text(50), idSintoma long,idDireccion long, idDespachador long, Estado long, Observaciones  text(255), Cronica MEMO, NroIncidenteInterno text(50)"
        'mas campos...
        mSql = mSql + ", DescripcionPerjuiciosCampo MEMO, MaterialesCombustibles Memo, DescripcionPerjuiciosVivienda MEMO, AccesoPor Text(255), Referencias MEMO, Aseguradora text(255), DescripcionMaterial MEMO, EquiposEspeciales MEMO, Poliza text(25), AmbientesAfectadosVivienda long, HectareasAfectadasCampo long, idInstalacionElectrica long, idInstalacionGas long"
        'y mas campos...
        mSql = mSql + ", PoliciaACargo text(50),PoliciaCantidad long, PoliciaMovil Text(50), SEMedico Text(50), SECentroAsistencial Text(255), SEMedicoMP Text(25), SEAmbulancias long"
        'y la pk
        mSql = mSql + ",constraint pk primary key (id))"
        
        cn.Execute mSql
        
        CreateTableID tAtencionB, 0
        
        'agrego tabla inst electrica
        cn.Execute "Create Table InstElectrica(id long, nombre Text(25),constraint pk primary key (id))"
        'agrego algunos valores...
        cn.Execute "Insert Into InstElectrica(id, nombre) values (1, 'Sin Instalacion')"
        cn.Execute "Insert Into InstElectrica(id, nombre) values (2, '12v')"
        cn.Execute "Insert Into InstElectrica(id, nombre) values (3, '110v')"
        cn.Execute "Insert Into InstElectrica(id, nombre) values (4, '220v')"
        cn.Execute "Insert Into InstElectrica(id, nombre) values (5, '380v')"
        
        CreateTableID tInstElectrica, 5
        
        'agrego tabla instGas
        cn.Execute "Create Table InstGas(id long, nombre Text(25),constraint pk primary key (id))"
        'inserto algunos valores...
        cn.Execute "Insert Into InstGas(id, nombre) values (1, 'Sin Instalacion')"
        cn.Execute "Insert Into InstGas(id, nombre) values (2, 'Cilindros 45kg')"
        cn.Execute "Insert Into InstGas(id, nombre) values (3, 'Garrafas 10kg')"
        
        CreateTableID tInstGas, 3
        
        'agrego tabla involucrado
        cn.Execute "Create Table Involucrado(id long, nombre Text(50),apellido Text(50),observaciones Text(255),DescripcionRelacion Memo,Telefono Text(25),idTipoRelacion long, TipoDoc long,NroDoc long, sexo long, idDireccion long, EsEdad long,FechaNacimiento DateTime, idSiniestro long, constraint pk primary key (id))"
        
        CreateTableID tInvolucrado, 0
        
        'le agrego los campos faltantes a vehiculo
        cn.Execute "ALTER table Vehiculo add Tipo text(30)"
        cn.Execute "ALTER table Vehiculo add Color text(20)"
        cn.Execute "ALTER table Vehiculo add Perjuicios MEMO"
        cn.Execute "ALTER table Vehiculo add idSiniestro long"
              
        'agrego la tabla pais
        cn.Execute "Create Table Pais(id long, nombre Text(35), PrimerOrden Text(35),SegundoOrden Text(35), TercerOrden Text(35), constraint pk primary key (id))"
        CreateTableID tPais, 1
        
        'agrego Argentina a la tabla
        cn.Execute "Insert Into Pais(id, nombre, PrimerOrden, SegundoOrden, TercerOrden) values (1, 'Argentina', 'Provincia','Ciudad','Barrio')"
        
        'ahora se pueden aregar items a la tabla provincias
        CreateTableID tProvincia, 23
        
        'agrego la columna idPais a la provincia
        cn.Execute "ALTER table Provincia add idPais long"
        
        'todas las provincias ya cargadas son de argentina
        cn.Execute "Update Provincia set idPais=1"
                       
        'agrego las tablas auxiliares de colaboracion
        cn.Execute "Create Table CuerpoBomberos(id long, nombre Text(50),constraint pk primary key (id))"
        CreateTableID tCuerpoBomberos, 0
        cn.Execute "Create Table ResponsableCuerpo(id long,idCuerpo long, nombre Text(50),constraint pk primary key (id))"
        CreateTableID tResponsableCuerpo, 0
        cn.Execute "Create Table UnidadCuerpo(id long,idCuerpo long, nombre Text(50),constraint pk primary key (id))"
        CreateTableID tUnidadCuerpo, 0
        cn.Execute "Create Table ColaboracionCuerpo(id long,idCuerpo long,idSiniestro long,idUnidad long,idResponsable long,CantidadEfectivos long, Observaciones MEMO,constraint pk primary key (id))"
        CreateTableID tColaboracionCuerpo, 0

        'actualizo la version
        cn.Execute "Update version set principal=2, secundario=0, revision=0"
    End If
End Function

Public Sub CreateTableID(pTable As eTablas, idInicial As Long)
    cn.Execute "Create Table ID" + GetTabla(pTable) + "(id long,constraint pk primary key (id))"
    cn.Execute "Insert Into ID" + GetTabla(pTable) + "(id) values (" + Str(idInicial) + ")"
End Sub

Public Function SelectSQL(sql As String) As Recordset
    On Error GoTo e:
    Dim rs As Recordset
    WriteToLog App.path + "\db.log", Str(Now) + ": SelectSql= " + sql
    Set rs = cn.Execute(sql)
    Set SelectSQL = rs
    Set rs = Nothing
    Exit Function
e:
    Debug.Print Err.Description
    Select Case Err.Number
    Case 3704:
        mDBEvents.FireErrorEvent eDBConnectionClosed, "La conexion esta cerrada."
    Case Else:
        mDBEvents.FireErrorEvent eUnknown, Err.Description
    End Select
End Function

Public Function SelectTabla(pTabla As eTablas) As Recordset
    On Error GoTo e:
    Dim rs As Recordset
    TERR.Anotar "SQL:", "selectTabla - sql= " + "SELECT * FROM " + GetTabla(pTabla)
    Set rs = cn.Execute("SELECT * FROM " + GetTabla(pTabla))
    Set SelectTabla = rs
    Exit Function
e:
    Debug.Print Err.Description
    Select Case Err.Number
    Case 3704:
        mDBEvents.FireErrorEvent eDBConnectionClosed, "La conexion esta cerrada."
    Case Else:
        mDBEvents.FireErrorEvent eUnknown, Err.Description
    End Select
End Function

Public Function Insertar(pTabla As eTablas, pParametros As ParametrosManager) As Boolean
    Dim auxNombres As String
    Dim auxValues As String
    Dim sql As String
   
    Dim cmd As Command
    Set cmd = New Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = cn
    
    Dim p As Parametro
    Dim par As Parameter
        
    For Each p In pParametros
        auxNombres = auxNombres + ", " + p.Nombre
        auxValues = auxValues + ", @" + p.Nombre
    Next
       
    sql = "INSERT INTO " + GetTabla(pTabla) + " (" + Right(auxNombres, Len(auxNombres) - 2) + ") VALUES (" + Right(auxValues, Len(auxValues) - 2) + ")"
    cmd.CommandText = sql
    
    For Each p In pParametros
        Set par = cmd.CreateParameter("@" + p.Nombre, , , , p.Valor)
        cmd.Parameters("@" + p.Nombre).Value = p.Valor
    Next
    
    On Error GoTo e:
    WriteToLog App.path + "\db.log", "insertar - sql= " + sql
    cmd.Execute
    Insertar = True
    Exit Function
e:
    Insertar = False
End Function

Public Function Update(pTabla As eTablas, pParametros As ParametrosManager) As Boolean
    Dim auxNombres As String
    Dim auxWhere As String
    Dim sql As String
   
    Dim cmd As New Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = cn
    
    Dim p As Parametro
    Dim par As Parameter
        
    For Each p In pParametros
        If p.EsPK Then
            auxWhere = auxWhere + " AND " + p.Nombre + "= @" + p.Nombre
        Else
            auxNombres = auxNombres + ", " + p.Nombre + "= @" + p.Nombre
        End If
    Next
       
    sql = "UPDATE " + GetTabla(pTabla) + " SET " + Right(auxNombres, Len(auxNombres) - 2) + " WHERE " + Right(auxWhere, Len(auxWhere) - 5)
    cmd.CommandText = sql
    
    For Each p In pParametros
        Set par = cmd.CreateParameter("@" + p.Nombre, , , , p.Valor)
        cmd.Parameters("@" + p.Nombre).Value = p.Valor
    Next
    On Error GoTo e
    WriteToLog App.path + "\db.log", "update - sql= " + sql
    cmd.Execute
    Update = True
    Exit Function
e:
    Update = False
End Function

Public Sub BorradoLogico(pTabla As eTablas, campoId As Parametro)
    SelectSQL ("update " + GetTabla(pTabla) + " set eliminado='true' where " + campoId.Nombre + "=" + Str(campoId.Valor))
End Sub

Public Sub RestaurarElemento(pTabla As eTablas, campoId As Parametro)
    SelectSQL ("update " + GetTabla(pTabla) + " set eliminado='false' where " + campoId.Nombre + "=" + Str(campoId.Valor))
End Sub

Public Function getIdDisponible(pTabla As eTablas) As Long
    'utilizo el sistema de una tabla especial para guardar el ultimo id utilizado
    'soporta concurrencia...
    
    Dim rs As Recordset
    cn.BeginTrans
    cn.Execute "Update ID" + GetTabla(pTabla) + " set id=id+1"
    Set rs = cn.Execute("select id from ID" + GetTabla(pTabla))
    cn.CommitTrans
    getIdDisponible = rs!id
End Function

Private Sub WriteToLog(path As String, contents As String)
   ' On Error GoTo e
    Dim fso ' As FileSystemObject
    Dim f, ts
    'Dim ts As TextStream
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(path) Then
       fso.CreateTextFile path
    End If
   
    Set f = fso.GetFile(path)
    Set ts = f.OpenAsTextStream(8, 0) '8=ForAppending
    
    ts.WriteLine contents
    'ts.write cadena
    ts.Close
'    Exit Sub
'e:
    
End Sub
