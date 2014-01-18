Attribute VB_Name = "ModuloConfiguracionListados"
'Función api que recupera un valor-dato de un archivo Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

'Función api que Escribe un valor - dato en un archivo Ini
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpString As String, _
    ByVal lpFileName As String) As Long
    
    
Public Enum eListado
    eListadoAtencionesPendientes = 1
    eListadoAtencionesGeneral = 2
    eListadoAtencionesBPendientes = 3
    eListadoAtencionesBGeneral = 4
End Enum
'Lee un dato -----------------------------
'Recibe la ruta del archivo, la clave a leer y el valor por defecto en caso de que la Key no exista
Public Function Leer_Ini(Path_INI As String, Group As String, Key As String, Default As Variant) As String

Dim bufer As String * 256
Dim Len_Value As Long

        Len_Value = GetPrivateProfileString(Group, _
                                         Key, _
                                         Default, _
                                         bufer, _
                                         Len(bufer), _
                                         Path_INI)
        
        Leer_Ini = Left$(bufer, Len_Value)

End Function

'Escribe un dato en el INI _
-----------------------------
'Recibe la ruta del archivo, La clave a escribir y el valor a añadir en dicha clave

Public Function Grabar_Ini(Path_INI As String, Group As String, Key As String, Valor As Variant) As String
    WritePrivateProfileString Group, Key, Valor, Path_INI
End Function

'no hace falta porq si da error cuando cargo no se la asigno al listado y listo
Public Function GetEncabezadosDefault(listado As eListado) As ControlesPOO.LVCEncabezadoManager
    Dim encs As New ControlesPOO.LVCEncabezadoManager
    Select Case listado
        Case eListadoAtencionesPendientes
            encs.Add "Codigo", "codigo", 8
            encs.Add "Sintoma", "sintoma", 19
            encs.Add "Afiliado", "afiliado", 18
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Vence", "GetVencimiento", 8
            encs.Add "Despachador", "despachador", 18
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
            encs.Add "Transcurrido", "transcurrido", 10
        Case eListadoAtencionesGeneral
            encs.Add "Codigo", "codigo", 8
            encs.Add "Sintoma", "sintoma", 19
            encs.Add "Afiliado", "afiliado", 18
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Despachador", "despachador", 18
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
        Case eListadoAtencionesBGeneral
            encs.Add "Tipo", "codigo", 8
            encs.Add "Codigo", "sintoma", 19
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Despachador", "despachador", 18
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
        Case eListadoAtencionesBPendientes
            encs.Add "Tipo", "codigo", 8
            encs.Add "Codigo", "sintoma", 19
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Despachador", "despachador", 18
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
            encs.Add "Transcurrido", "transcurrido", 10
    End Select
    Set GetEncabezadosDefault = encs
End Function

Public Function GetEncabezadosDisponibles(listado As eListado) As ControlesPOO.LVCEncabezadoManager
    Dim encs As New ControlesPOO.LVCEncabezadoManager
    Select Case listado
        Case eListadoAtencionesPendientes
            encs.Add "Codigo", "codigo", 8
            encs.Add "Sintoma", "sintoma", 19
            encs.Add "Afiliado", "afiliado", 18
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Vence", "GetVencimiento", 8
            encs.Add "Despachador", "despachador", 18
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
            encs.Add "Transcurrido", "transcurrido", 20
            
            encs.Add "Nro. Incidente", "NroIncidente", 20
            encs.Add "Nro. Incidente Interno", "NroIncidenteInterno", 20
            encs.Add "Operador", "Operador", 30
            encs.Add "Diagnostico", "Diagnostico", 30
            encs.Add "Observaciones", "Observaciones", 30
            encs.Add "Telefono", "pgtelefono", 15
            encs.Add "Monto Servicio", "pgmontoservicio", 20
            encs.Add "Coseguro", "pgcoseguro", 20
            encs.Add "Monto Abonado", "pgmontoabonado", 20
            'ver si faltan campos
        Case eListadoAtencionesGeneral
            encs.Add "Codigo", "codigo", 8
            encs.Add "Sintoma", "sintoma", 19
            encs.Add "Afiliado", "afiliado", 18
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Despachador", "despachador", 18
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
            
            encs.Add "Nro. Incidente", "NroIncidente", 20
            encs.Add "Nro. Incidente Interno", "NroIncidenteInterno", 20
            encs.Add "Operador", "Operador", 30
            encs.Add "Diagnostico", "Diagnostico", 30
            encs.Add "Observaciones", "Observaciones", 30
            encs.Add "Telefono", "pgtelefono", 15
            encs.Add "Monto Servicio", "pgmontoservicio", 20
            encs.Add "Coseguro", "pgcoseguro", 20
            encs.Add "Monto Abonado", "pgmontoabonado", 20
            
        Case eListadoAtencionesBGeneral
            encs.Add "Tipo", "codigo", 8
            encs.Add "Codigo", "sintoma", 19
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Despachador", "despachador", 18
            encs.Add "Sal. PreInsp.", "salidapreinspeccion", 7
            encs.Add "Lleg. PreInsp.", "llegadapreinspeccion", 7
            encs.Add "Salida Dot.", "salidadotacion", 7
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
            encs.Add "Nro. Incidente", "NroIncidente", 20
            encs.Add "Nro. Incidente Interno", "NroIncidenteInterno", 20
            encs.Add "Operador", "Operador", 30
            encs.Add "Observaciones", "Observaciones", 30
        Case eListadoAtencionesBPendientes
            encs.Add "Tipo", "codigo", 8
            encs.Add "Codigo", "sintoma", 19
            encs.Add "Hora Llamado", "horallamada", 11
            encs.Add "Despachador", "despachador", 18
            encs.Add "Sal. PreInsp.", "salidapreinspeccion", 7
            encs.Add "Lleg. PreInsp.", "llegadapreinspeccion", 7
            encs.Add "Salida Dot.", "salidadotacion", 7
            encs.Add "QTH", "qth", 7
            encs.Add "VL", "vl", 7
            encs.Add "Movil", "movil", 10
            encs.Add "Direccion", "pgdireccion", 40
            encs.Add "Nro. Incidente", "NroIncidente", 20
            encs.Add "Nro. Incidente Interno", "NroIncidenteInterno", 20
            encs.Add "Operador", "Operador", 30
            encs.Add "Observaciones", "Observaciones", 30
            encs.Add "Transcurrido", "transcurrido", 20
    End Select
    Set GetEncabezadosDisponibles = encs
End Function

Public Function GetEncabezados(pPath As String) As ControlesPOO.LVCEncabezadoManager
    On Error GoTo errman
    Dim encs As New ControlesPOO.LVCEncabezadoManager
    Dim Nombre As String
    Dim miembro As String
    Dim ancho As Integer
    Dim I As Integer
    I = 0
    Do
        Nombre = Leer_Ini(pPath, "ListViewC", "NEncabezado" + Trim(Str(I)), "")
        miembro = Leer_Ini(pPath, "ListViewC", "MEncabezado" + Trim(Str(I)), "")
        ancho = CInt(Leer_Ini(pPath, "ListViewC", "AEncabezado" + Trim(Str(I)), 0))
        If Nombre <> "" Then
            encs.Add Nombre, miembro, ancho
        Else
            Exit Do
        End If
        I = I + 1
    Loop
    Set GetEncabezados = encs
    Exit Function
errman:
    Set GetEncabezados = Nothing
End Function

Public Sub SaveEncabezados(pPath As String, encs As ControlesPOO.LVCEncabezadoManager)
Dim I As Integer
Dim enc As ControlesPOO.LVCEncabezado
For Each enc In encs
    Grabar_Ini pPath, "ListViewC", "NEncabezado" + Trim(Str(I)), enc.Nombre
    Grabar_Ini pPath, "ListViewC", "MEncabezado" + Trim(Str(I)), enc.miembro
    Grabar_Ini pPath, "ListViewC", "AEncabezado" + Trim(Str(I)), Trim(Str(enc.ancho))
    I = I + 1
Next
End Sub
