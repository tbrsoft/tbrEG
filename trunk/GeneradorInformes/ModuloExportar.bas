Attribute VB_Name = "ModuloAux"
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

Public Function GetEncabezados(pPath As String) As ControlesPOO.LVCEncabezadoManager
    On Error GoTo errman
    Dim encs As New ControlesPOO.LVCEncabezadoManager
    Dim Nombre As String
    Dim miembro As String
    Dim ancho As Integer
    Dim i As Integer
    i = 0
    Do
        Nombre = Leer_Ini(pPath, "ListViewC", "NEncabezado" + Trim(Str(i)), "")
        miembro = Leer_Ini(pPath, "ListViewC", "MEncabezado" + Trim(Str(i)), "")
        ancho = CInt(Leer_Ini(pPath, "ListViewC", "AEncabezado" + Trim(Str(i)), 0))
        If Nombre <> "" Then
            encs.Add Nombre, miembro, ancho
        Else
            Exit Do
        End If
        i = i + 1
    Loop
    Set GetEncabezados = encs
    Exit Function
errman:
    Set GetEncabezados = Nothing
End Function

Public Function GetParametros(pPath As String) As LParameterManager
    On Error GoTo errman
    Dim params As New LParameterManager
    Dim Nombre As String
    Dim Tipo As String
    Dim Descripcion As String
    Dim i As Integer
    i = 0
    Do
        Nombre = Leer_Ini(pPath, "Parametros", "Nombre" + Trim(Str(i)), "")
        Tipo = Leer_Ini(pPath, "Parametros", "Tipo" + Trim(Str(i)), "")
        Descripcion = Leer_Ini(pPath, "Parametros", "Descripcion" + Trim(Str(i)), "")

        If Nombre <> "" Then
            params.Add Nombre, Tipo, Descripcion
        Else
            Exit Do
        End If
        i = i + 1
    Loop
    Set GetParametros = params
    Exit Function
errman:
    Set GetParametros = Nothing
End Function

Public Sub SaveEncabezados(pPath As String, encs As ControlesPOO.LVCEncabezadoManager)
Dim i As Integer
Dim enc As ControlesPOO.LVCEncabezado
For Each enc In encs
    Grabar_Ini pPath, "ListViewC", "NEncabezado" + Trim(Str(i)), enc.Nombre
    Grabar_Ini pPath, "ListViewC", "MEncabezado" + Trim(Str(i)), enc.miembro
    Grabar_Ini pPath, "ListViewC", "AEncabezado" + Trim(Str(i)), Trim(Str(enc.ancho))
    i = i + 1
Next
End Sub

Public Sub SaveParametros(pPath As String, params As LParameterManager)
Dim i As Integer
Dim par As LParameter
If Not params Is Nothing Then
For Each par In params
    Grabar_Ini pPath, "Parametros", "Nombre" + Trim(Str(i)), par.Nombre
    Grabar_Ini pPath, "Parametros", "Tipo" + Trim(Str(i)), par.Tipo
    Grabar_Ini pPath, "Parametros", "Descripcion" + Trim(Str(i)), par.Descripcion
    i = i + 1
Next
End If
End Sub

Public Function ExecuteSQL(pSQl As String) As Recordset
    Dim con As New Connection
    Dim cs As String
    cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + GetSetting("TbrEmergencyGroup", "DBLayer", "PathDB", "") + ";Persist Security Info=False;Jet OLEDB:Database Password=zapato"
    con.Open cs
    Dim rs As Recordset
    Set rs = con.Execute(pSQl)
    Set ExecuteSQL = rs
End Function
