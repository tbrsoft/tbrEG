Attribute VB_Name = "Modulo"
Public tErr As New tbrErrores.clsTbrERR

Public Enum eApplication
    eWord = 1
    eExcel = 2
    eOpenOffice = 3
End Enum

Private Enum eTipoLlamado
    ePropertyGet = 1
    eMethod = 2
    eGetPropertyImplementado = 3
End Enum

Public Sub Main()
    APh = App.path
    If Right(APh, 1) <> "\" Then APh = APh + "\"
    
    tErr.FileLog = APh + "reg_tlv.log"
    tErr.Set_ADN "tlv_v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    tErr.AppendSinHist "init_sys_tlv"
End Sub

Public Function GetApplication(eApp As eApplication) As Object
    On Error GoTo e
    If eApp = eExcel Then
        Set GetApplication = CreateObject("excel.application")
    ElseIf eApp = eWord Then
        Set GetApplication = CreateObject("word.application")
    ElseIf eApp = eOpenOffice Then
        Set GetApplication = CreateObject("com.sun.star.ServiceManager")
    End If
    Exit Function
e:
End Function

Public Function getValue(var As Object, pMiembro As String) As String
    On Error GoTo e
    Dim mTipoLlamado As eTipoLlamado
    mTipoLlamado = detectarMetodoLlamado(var, pMiembro)
    Select Case mTipoLlamado
        Case eGetPropertyImplementado
            getValue = var.GetProperty(pMiembro)
        Case eMethod
            getValue = CallByName(var, pMiembro, VbMethod)
        Case ePropertyGet
            getValue = CallByName(var, pMiembro, VbGet)
    End Select
    Exit Function
e: 'teoricamente nunca deberia llegar aca
    getValue = "0" 'andres agrega esto 2011/03/17 por que "transcurrido no se encuentra" ???
    'Err.Raise 2010, , pMiembro + " No se encuentra"
    tErr.AppendLog "das014", "No se encuentra la columna: " + pMiembro
End Function

Private Function detectarMetodoLlamado(var As Object, pMiembro As String) As eTipoLlamado
'detecta de que tipo es el miembro del encabezado, si es property get, un metodo o la funcion getPROPERTY
On Error Resume Next
    'limpio si hay errores
    Err.Clear
    a = CallByName(var, pMiembro, VbGet)
    If Err.Number = 0 Then
        detectarMetodoLlamado = ePropertyGet
        Exit Function
    End If

    Err.Clear
    a = CallByName(var, pMiembro, VbMethod)
    If Err.Number = 0 Then
        detectarMetodoLlamado = eMethod
        Exit Function
    End If

    Err.Clear
    a = var.GetProperty(pMiembro)

    If Err.Number = 0 Then
        detectarMetodoLlamado = eGetPropertyImplementado
        Exit Function
    End If

    'si encuentra uno de los tres metodos no deberia llegar nunca a este punto
    'y si llego significa q no encontro la propiedad
    On Error GoTo 0
    Err.Clear

    Err.Raise 438 'no se encontro la propiedad o el metodo

End Function

Public Sub EscribirArchivo(path As String, cadena As String)
   ' On Error GoTo e
    Dim fso 'As FileSystemObject
    Dim f, ts
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(path) Then
       fso.CreateTextFile path
    End If
    
    Set f = fso.GetFile(path)
    Set ts = f.OpenAsTextStream(2, 0)
     
    ts.Write cadena
    ts.Close
'    Exit Sub
'e:
    
End Sub
