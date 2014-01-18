Attribute VB_Name = "Modulo"
'necesario para la ayuda
Public hWndAyudaHTML As Long

Public Enum eTipoDatos
    eString = 1
    eInteger = 2
    eMoneda = 3
    eLong = 4
    ePatenteAutomovil = 5
    eDireccionIP = 6
End Enum

'Public Enum eTeclasPermitidas
'    eNumeros = 1
'    eComa = 2
'    eSignoMenos = 4
'    eEspacio = 8
'    eCustom = 0
'End Enum

Public Enum eTipoAMB
    etALTA = 1
    etBAJA = 2
    etMODIFICACION = 3
    etCONSULTA = 4
End Enum
'enumeracion para saber si los formularios de consulta son con o sin retorno

Public Enum eTipoFormulario
    etConRetorno = 1
    etSinRetorno = 2
End Enum

'enumeracion para saber si estan instalados excel o word
Public Enum eApplications
    eWord = 1
    eExcel = 2
    eOpenOffice = 3
End Enum

Public UsuarioActual As blcemi.Empleado

Public Enum eModo
    eNoRegistrada = 0
    eVersionRegistrada = 1
    eModoDemo = 2
End Enum

Public Enum eModoFuncionamiento
    eMFEmergencia = 1
    eMFBomberos = 2
End Enum

Public modo As eModo 'licencia aceptada
Public modoSoftware As eModoFuncionamiento

Public APh As String
Public GBL As New blcemi.ClaseGlobal
Public TERR As New tbrErrores.clsTbrERR
Public CCFFGG As New tbrConfig.GlobalCls

Sub Main()

    APh = App.Path
    If Right(APh, 1) <> "\" Then APh = APh + "\"
    
    TERR.FileLog = APh + "reg_tbrEMG.log"
    TERR.Set_ADN "tEG_v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    TERR.AppendSinHist "init_sys"
    
    'asegurarse que la referencia a las cosas de blcemi este cargada correctamente
    'NO ES UN ERROR DE VERDAD
    GBL.PrintToErrorLog "ForceLOG_Ini." + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision), _
                    "begin", _
                    CStr(Now), APh
    
    'inicializar blCemi 'andres2010 03 17
    GBL.Init_BlCemi APh  'indico carpeta de los logs
    
    
    If CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento = 0 Then
        frmConfigInicial.Show vbModal
    End If

    modoSoftware = IIf(CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento = 1, eMFEmergencia, eMFBomberos)

    frmSplash.Show
    frmSplash.Mensaje = "Cargando ccffgg.configuracion..."
    CCFFGG.Configuracion.Load
    CCFFGG.Configuracion.Save
    TERR.AppendSinHist "cfg_ready"
    frmSplash.Mensaje = "Verificando..."
    VerificarLicencia
    TERR.AppendSinHist "CLIs_ready"
    frmSplash.Mensaje = "Cargando Formulario Principal..."
    
    Load MDI
    TERR.AppendSinHist "MDI_ready"
    
    MDI.MostrarIniciarSesion
    TERR.AppendSinHist "SSion_ready"
    
    'Set UsuarioActual = EmpleadosGBL.Item(1)
End Sub

Public Sub InicializarDireccion(ctl As ctlDireccion)
    On Local Error GoTo ErrInitDIR
    TERR.Anotar "abaq"
    
    'que pasa si no se cargo al inicio ????
    ' obligarlo ????
    TERR.Anotar "abar", CCFFGG.Configuracion.Defaults.Pais
    TERR.Anotar "abas", CCFFGG.Configuracion.Defaults.Provincia
    TERR.Anotar "abat", CCFFGG.Configuracion.Defaults.Ciudad
    TERR.Anotar "abau", CCFFGG.Configuracion.Defaults.Barrio
    
    Dim P As blcemi.Pais
    Set P = GBL.PaisesGBL.ItemByName(CCFFGG.Configuracion.Defaults.Pais)
    TERR.Anotar "abau2", P.Nombre
    
    Dim P2 As blcemi.Provincia
    Set P2 = GBL.ProvinciasGBL.ItemByName(CCFFGG.Configuracion.Defaults.Provincia)
    TERR.Anotar "abau3", P2.Nombre
    
    Dim P3 As blcemi.Ciudad
    Set P3 = P2.Ciudades.ItemByName(CCFFGG.Configuracion.Defaults.Ciudad)
    TERR.Anotar "abau4", P3.Nombre
    
    Dim P4 As blcemi.Barrio
    Set P4 = P3.Barrios.ItemByName(CCFFGG.Configuracion.Defaults.Barrio)
    TERR.Anotar "abau5", P4.Nombre
    
    ctl.Inicializar P, P.Provincias, P2, P3, P4
    
    TERR.Anotar "abau7"
    
    Exit Sub
ErrInitDIR:
    TERR.AppendLog "ErrInitDIR", TERR.ErrToTXT(Err)
    Resume Next
    
End Sub

Public Sub VerificarLicencia(Optional MostrarAviso As Boolean = False)

'    modo = eVersionRegistrada
'    Exit Sub

   On Error GoTo errman
   
    TERR.Anotar "arft"
    
    'VERIFICACION DE LICENCIA
    Dim TD As New tbrDATA.clsTODO
    TERR.Anotar "arft2"
    'VER EL ID final de este equipo ya
    TD.SetLog APh + "Reg23.log"
    TD.SetSF "tbrEG_v2" 'se agrego el 20/3/2010 antes no habia nada !?!?!?!?
    TD.DoNow APh + "s.a"
    
    TERR.Anotar "arfu"
    
    Dim D As String 'identificador leido ahora de este equipo
    D = TD.GetRF
    
    TERR.Anotar "arfv"
    
    Dim FF As String, Nr As Long
    Dim FS As New Scripting.FileSystemObject
    'borro la licencia que hay si la hay ...
    If FS.FileExists(APh + "lic.insertada") Then
        TERR.Anotar "arfw"
        Nr = TD.GetNR(APh + "lic.insertada", FF) 'obtengo el numero del archivo de la licencia y en FF devuelve el ID final del equipo para el que fue hecho
        'esta funcion da 100 o mas si los cambios en la pc no son muchos
        TERR.Anotar "arfx"
        If TD.GetPtosDiff(UCase(Trim(FF)), UCase(Trim(D))) < 100 Then
        'If UCase(Trim(D)) <> UCase(Trim(FF)) Then
            'TODO matar las cadenas pateticas expuestas!
            MsgBox "La licencia fue hecha para otro equipo!", vbExclamation
            'WriteToLog APh + "errorLog.txt", Str(Now) + "Intento de insertar licencia no valida..."
            TERR.Anotar "arfy", UCase(Trim(FF)), UCase(Trim(D)) ' primero el de la licencias, segundo el de la pc donde estamos
            modo = eNoRegistrada
            TERR.AppendLog "NL9011029837"
            
            Exit Sub
        End If
        TERR.Anotar "arfz", Nr

        Select Case Nr
            Case "19", "20" 'numeros de licencia ''''EN LA DE MARTIN ERA 10 y 22 !!!! SE CAMBIO EL 20 de marzo de 2010
                modo = eVersionRegistrada
                'TODO da ocote esta cadena de texto pàra crackear
                If MostrarAviso Then MsgBox "La licencia es correcta, ya puede comenzar a utilizar el software.", vbInformation
        End Select
    Else
        TERR.Anotar "arga"
        modo = eModoDemo
    End If
    TERR.Anotar "argb"
   Exit Sub
   
errman:
    TERR.AppendLog "cfCIL", TERR.ErrToTXT(Err)
    GBL.PrintToErrorLog "Modulo", "verificarLicencia", Str(Err.Number) + " - " + Err.Description
    'MsgBox "Ocurrio un error en la ejecucion y el programa no puede continuar."
    'End 'porq si no tiene licencia no lo dejo seguir
    
    modo = eModoDemo
    
End Sub

Public Sub SoloNumeros(ByRef KeyAscii As Integer, Optional AcceptSeparator = True)
    'completar
    Select Case KeyAscii
        Case 48 To 57
        Case 8 'backspace
        Case Asc(".")
        Case Else
            Beep
            KeyAscii = 0
    End Select

    If Not AcceptSeparator And KeyAscii = Asc(".") Then KeyAscii = 0
End Sub

Public Sub GuardarReporte(Contenido As String, NombreArchivo As String, Optional Ruta As String)
    Dim cd As New CommonDialog
    cd.InitDir = IIf(Ruta = "", APh, Ruta)
    cd.DefaultExt = "html"
    cd.FileName = NombreArchivo
    cd.ShowSave
    If cd.FileName <> "" Then
         Dim fso ' As FileSystemObject
         Dim f, ts
         'Dim ts As TextStream
         Dim s As String
         Set fso = CreateObject("Scripting.FileSystemObject")
         
         If Not fso.FileExists(cd.FileName) Then
            fso.CreateTextFile cd.FileName
         End If
        
         Set f = fso.GetFile(cd.FileName)
         Set ts = f.OpenAsTextStream(8, 0) '8=ForAppending
         
         ts.Write Contenido
         'ts.write cadena
         ts.Close
    End If
End Sub

'Public Sub ValidateKey(ByRef KeyAscii As Integer, pTeclas As eTeclasPermitidas, Optional pCustom As String)
''    eNumeros = 1
''    eComa = 2
''    eSignoMenos = 4
''    eEspacio = 8
''    eCustom = 0
'Select Case pTeclas
'    Case 0
'    Case 1
'    Case 2
'    case
'
'End Select
'End Sub

Public Function CCurrency(texto As String) As Currency
    CCurrency = CCur(Replace(texto, ".", ","))
End Function


Public Function TextBoxValidado(txt As TextBox, pTipo As eTipoDatos) As Boolean
'completar
Select Case pTipo
    Case eTipoDatos.eInteger
        TextBoxValidado = (IsNumeric(Trim(txt.Text)))
    Case eTipoDatos.eLong
        TextBoxValidado = (IsNumeric(Trim(txt.Text)))
    Case eTipoDatos.eMoneda
        TextBoxValidado = (IsNumeric(Trim(txt.Text)))
    Case eTipoDatos.eString
        TextBoxValidado = (Trim(txt.Text) <> "")
    Case eTipoDatos.ePatenteAutomovil
        'completar, la parte izq no tiene q tener numeros
        If Len(txt.Text) = 6 Then
            If IsNumeric(Right(txt.Text, 3)) Then
                TextBoxValidado = True
            Else
                TextBoxValidado = False
            End If
        Else
            TextBoxValidado = False
        End If
    Case eTipoDatos.eDireccionIP
        Dim aux() As String
        aux = Split(Trim(txt.Text), ".")
        If UBound(aux) <> 3 Then
            TextBoxValidado = False
        Else
            Dim band As Boolean
            band = True
            For I = 0 To 3
                If Not (IsNumeric(aux(I)) And Val(aux(I)) <= 255 And Val(aux(I)) >= 0) Then
                    band = False
                    Exit For
                End If
            Next
            TextBoxValidado = band
        End If
        
End Select
End Function

Public Sub DistribuirBotones(tBar As ToolBar)
    Dim b As Button
    Dim b2 As Button
    Dim aux As Double
    aux = 0
    For Each b In tBar.Buttons
        If b.Style = tbrDefault And b.Visible Then
            aux = aux + tBar.ButtonWidth + 60 ' el 50 es por el interlineado entre los controles
        ElseIf b.Style = tbrSeparator Then
          aux = aux + 150
        ElseIf b.Style = tbrPlaceholder Then
            Set b2 = b
        End If
    Next
    
    b2.Width = tBar.Width - aux
    
End Sub

Public Function ApplicationInstalled(eApp As eApplications) As Boolean
On Error GoTo e
    Dim aux As Object
    If eApp = eExcel Then
        Set aux = CreateObject("excel.application")
        aux.Quit 'para q no quede la referencia colgada
        Set aux = Nothing
    ElseIf eApp = eWord Then
        Set aux = CreateObject("word.application")
        aux.Quit 'para q no quede la referencia colgada
        Set aux = Nothing
    ElseIf eApp = eOpenOffice Then
        Set aux = CreateObject("com.sun.star.ServiceManager")
        Set aux = Nothing
    End If
    ApplicationInstalled = True
    Exit Function
e:
ApplicationInstalled = False
End Function

Public Sub BloquearTextBoxes(pBloquear As Boolean, pControles As Object)
Dim c As Control
For Each c In pControles
    If TypeOf c Is TextBox Then
        c.Locked = pBloquear
    End If
Next
End Sub

Public Function GetFont(pFontName As String, Optional pFontSize As Integer, Optional pBold As Boolean = False, Optional pItalic As Boolean = False, Optional pUnderline As Boolean = False) As StdFont
    Dim mFont As New StdFont
    Dim founded As Boolean
    
    For I = 0 To Screen.FontCount
        If Screen.Fonts(I) = pFontName Then
            mFont.Name = Screen.Fonts(I)
            founded = True
            Exit For
        End If
    Next
    
    If founded Then
        mFont.Bold = pBold
        mFont.Italic = pItalic
        mFont.Underline = pUnderline
        mFont.size = pFontSize
        
        Set GetFont = mFont
    Else
        Set GetFont = Nothing
    End If
    
End Function

'imprime un recibo que puede tener variAs blcemi.Cuotas, para pago en casa central
Public Sub ImprimirReciboAfiliado(pCuotas As blcemi.CuotaManager, pEndDoc As Boolean)
    Dim formRecibo As New tbrcamposimpresion.Formulario
    Dim mHoja As tbrcamposimpresion.Hoja
    Dim mTabla As tbrcamposimpresion.Tabla
    Dim aParent As blcemi.Afiliado
    Dim a As blcemi.Afiliado
    
    formRecibo.Abrir (APh + "formularios\reciboafiliados.form")
    Set mHoja = formRecibo.Hojas(1)
    
    Dim total As Currency
    Dim meses As String
    Dim ayear As String
    
    Dim c As blcemi.Cuota
    For Each c In pCuotas
        If aParent Is Nothing Then Set aParent = c.Afiliado
        total = total + c.Monto
        meses = meses + ", " + MonthName(c.mes, True)
        ayear = Trim(Str(c.ayear)) 'ver si son de años distintos
    Next
    meses = Right(meses, Len(meses) - 2)
    
    'va por triplicado
    Set mTabla = mHoja.Tablas.ItemByName("tabla1")
    LlenarTablaAfiliados mTabla, aParent
    Set mTabla = mHoja.Tablas.ItemByName("tabla2")
    LlenarTablaAfiliados mTabla, aParent
    Set mTabla = mHoja.Tablas.ItemByName("tabla3")
    LlenarTablaAfiliados mTabla, aParent
    
    mHoja.Campos.ItemByName("importe1").Valor = total
    mHoja.Campos.ItemByName("importe2").Valor = total
    mHoja.Campos.ItemByName("importe3").Valor = total
    
      
    LlenarDatos "1", mHoja, aParent, meses, ayear
    LlenarDatos "2", mHoja, aParent, meses, ayear
    LlenarDatos "3", mHoja, aParent, meses, ayear
    
    Dim frm As New frmprueba
    'frm.Imprimir formRecibo
    formRecibo.Imprimir Printer, pEndDoc 'frmprueba
        
End Sub

'para cuando se generan todos juntos
Public Sub ImprimirReciboAfiliado2(pCuota As blcemi.Cuota, pEndDoc As Boolean)
    Dim formRecibo As New tbrcamposimpresion.Formulario
    Dim mHoja As tbrcamposimpresion.Hoja
    Dim mTabla As tbrcamposimpresion.Tabla
    Dim aParent As blcemi.Afiliado
    Dim a As blcemi.Afiliado
    
    formRecibo.Abrir (APh + "formularios\reciboafiliados.form")
    Set mHoja = formRecibo.Hojas(1)
    
    Dim total As Currency
    Dim meses As String
    Dim ayear As String
    
    Set aParent = pCuota.Afiliado
    total = pCuota.Monto
    meses = MonthName(pCuota.mes, True)
    ayear = Trim(Str(pCuota.ayear))
    
    'va por triplicado
    Set mTabla = mHoja.Tablas.ItemByName("tabla1")
    LlenarTablaAfiliados mTabla, aParent
    Set mTabla = mHoja.Tablas.ItemByName("tabla2")
    LlenarTablaAfiliados mTabla, aParent
    Set mTabla = mHoja.Tablas.ItemByName("tabla3")
    LlenarTablaAfiliados mTabla, aParent
    
    mHoja.Campos.ItemByName("importe1").Valor = total
    mHoja.Campos.ItemByName("importe2").Valor = total
    mHoja.Campos.ItemByName("importe3").Valor = total
    
      
    LlenarDatos "1", mHoja, aParent, meses, ayear
    LlenarDatos "2", mHoja, aParent, meses, ayear
    LlenarDatos "3", mHoja, aParent, meses, ayear
    
    Dim frm As New frmprueba
   ' frm.Imprimir formRecibo
     formRecibo.Imprimir Printer 'frmprueba
        
End Sub
Private Sub LlenarDatos(indice As String, pHoja As tbrcamposimpresion.Hoja, pAfiliado As blcemi.Afiliado, pMeses As String, pYear As String)
    pHoja.Campos.ItemByName("afiliado" + indice).Valor = pAfiliado.NombreCompleto
    pHoja.Campos.ItemByName("dni" + indice).Valor = pAfiliado.TipoDoc.Nombre + ": " + Trim(Str(pAfiliado.NroDoc))
    pHoja.Campos.ItemByName("mes" + indice).Valor = pMeses
    pHoja.Campos.ItemByName("year" + indice).Valor = pYear
    
    Dim fecha As String
    fecha = IIf(Day(Date) < 10, "0" + Trim(Str(Day(Date))), Str(Day(Date)))
    fecha = fecha + " " + IIf(Month(Date) < 10, "0" + Trim(Str(Month(Date))), Str(Month(Date)))
    fecha = fecha + " " + Right(Str(Year(Date)), 2)
    pHoja.Campos.ItemByName("fecha" + indice).Valor = fecha
    'direccion!!!!!!
End Sub


Private Sub LlenarTablaAfiliados(pTabla As tbrcamposimpresion.Tabla, pAfiliado As blcemi.Afiliado)
Dim I As Integer
Dim a As blcemi.Afiliado
I = 0
For Each a In pAfiliado.PersonasACargo
    I = I + 1
    pTabla.Celda(I, 1) = a.NombreCompleto
Next

End Sub

'-----------------area protegida---------------

Public Sub ImprimirReciboAreaProtegida(pCuotas As blcemi.CuotaManager, pEndDoc As Boolean)
    Dim formRecibo As New tbrcamposimpresion.Formulario
    Dim mHoja As tbrcamposimpresion.Hoja
    Dim mArea As blcemi.AreaProtegida
        
    formRecibo.Abrir (APh + "formularios\reciboarea.form")
    Set mHoja = formRecibo.Hojas(1)
    
    Dim total As Currency
    Dim meses As String
    Dim ayear As String
    
    Dim c As blcemi.Cuota
    For Each c In pCuotas
        If mArea Is Nothing Then Set mArea = c.AreaProtegida
        total = total + c.Monto
        meses = meses + ", " + MonthName(c.mes, True)
        ayear = Trim(Str(c.ayear)) 'ver si son de años distintos
    Next
    meses = Right(meses, Len(meses) - 2)
            
    mHoja.Campos.ItemByName("importe1").Valor = total
    mHoja.Campos.ItemByName("importe2").Valor = total
    mHoja.Campos.ItemByName("importe3").Valor = total
    
   
    LlenarDatosAP "1", mHoja, mArea, meses, ayear
    LlenarDatosAP "2", mHoja, mArea, meses, ayear
    LlenarDatosAP "3", mHoja, mArea, meses, ayear
    
     Dim frm As New frmprueba
    'frm.Imprimir formRecibo
    formRecibo.Imprimir Printer, pEndDoc 'frmprueba
       
    
End Sub

'para cuando genero todos los recibos juntos
Public Sub ImprimirReciboAreaProtegida2(pCuota As blcemi.Cuota, pEndDoc As Boolean)
    Dim formRecibo As New tbrcamposimpresion.Formulario
    Dim mHoja As tbrcamposimpresion.Hoja
    Dim mArea As blcemi.AreaProtegida
    Dim mes As String
    
    On Local Error GoTo errman
    formRecibo.Abrir (APh + "formularios\reciboarea.form")
    Set mHoja = formRecibo.Hojas(1)
    
    Set mArea = pCuota.AreaProtegida
        
    mes = MonthName(pCuota.mes, True)
            
    mHoja.Campos.ItemByName("importe1").Valor = pCuota.Monto
    mHoja.Campos.ItemByName("importe2").Valor = pCuota.Monto
    mHoja.Campos.ItemByName("importe3").Valor = pCuota.Monto
       
    LlenarDatosAP "1", mHoja, mArea, mes, Trim(Str(pCuota.ayear)) 'original
    LlenarDatosAP "2", mHoja, mArea, mes, Trim(Str(pCuota.ayear)) 'triplicado
    LlenarDatosAP "3", mHoja, mArea, mes, Trim(Str(pCuota.ayear)) 'duplicado
    
     Dim frm As New frmprueba
    'frm.Imprimir formRecibo
    formRecibo.Imprimir Printer, pEndDoc 'frmprueba
    Exit Sub
errman:
    GBL.PrintToErrorLog "Modulo", "ImprimirReciboAreaProtegida", Err.Description
End Sub

Private Sub LlenarDatosAP(indice As String, pHoja As tbrcamposimpresion.Hoja, pArea As blcemi.AreaProtegida, pMeses As String, pYear As String)
    pHoja.Campos.ItemByName("area" + indice).Valor = pArea.NombreArea
    pHoja.Campos.ItemByName("dni" + indice).Valor = pArea.TipoDocResp.Nombre + ": " + Trim(Str(pArea.NroDocResp))
    pHoja.Campos.ItemByName("responsable" + indice).Valor = pArea.NombreCompleto
    pHoja.Campos.ItemByName("direccion" + indice).Valor = pArea.Direccion.Calle + " " + pArea.Direccion.Nro
    pHoja.Campos.ItemByName("ciudad" + indice).Valor = pArea.Direccion.Barrio.Nombre + " - " + pArea.Direccion.Ciudad.Nombre
    
    pHoja.Campos.ItemByName("mes" + indice).Valor = pMeses
    pHoja.Campos.ItemByName("year" + indice).Valor = pYear
   
    
    Dim fecha As String
    fecha = IIf(Day(Date) < 10, "0" + Trim(Str(Day(Date))), Str(Day(Date)))
    fecha = fecha + " " + IIf(Month(Date) < 10, "0" + Trim(Str(Month(Date))), Str(Month(Date)))
    fecha = fecha + " " + Right(Str(Year(Date)), 2)
    pHoja.Campos.ItemByName("fecha" + indice).Valor = fecha
    
End Sub
