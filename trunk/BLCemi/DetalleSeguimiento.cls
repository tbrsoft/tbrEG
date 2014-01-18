VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DetalleSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mId As Long
Dim mValorNuevo As String
Dim mValorAnterior As String
Dim mEmpleado As Empleado
Dim mIdAtencion As Long
Dim mCampo As String
Dim mFecha As Date
Dim mHora As String

Public Property Let Id(pId As Long)
    mId = pId
End Property
Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let ValorNuevo(pValor As String)
    mValorNuevo = pValor
End Property
Public Property Get ValorNuevo() As String
    ValorNuevo = mValorNuevo
End Property

Public Property Let ValorAnterior(pValor As String)
    mValorAnterior = pValor
End Property
Public Property Get ValorAnterior() As String
    ValorAnterior = mValorAnterior
End Property

Public Property Set Empleado(pEmpleado As Empleado)
    Set mEmpleado = pEmpleado
End Property
Public Property Get Empleado() As Empleado
    Set Empleado = mEmpleado
End Property

Friend Property Let idAtencion(pIdAtencion As Long)
    mIdAtencion = pIdAtencion
End Property

Friend Property Get idAtencion() As Long
    idAtencion = mIdAtencion
End Property

Public Property Let campo(pCampo As String)
    mCampo = pCampo
End Property
Public Property Get campo() As String
    campo = mCampo
End Property

Public Property Let Fecha(pFecha As Date)
    mFecha = pFecha
End Property
Public Property Get Fecha() As Date
    Fecha = mFecha
End Property

Public Property Let Hora(pHora As String)
    mHora = pHora
End Property
Public Property Get Hora() As String
    Hora = mHora
End Property

Public Function GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "pgvaloranterior"
            GetProperty = GetValorAnterior(mCampo)
        Case "pgvalornuevo"
            GetProperty = GetValorNuevo(mCampo)
        Case "pgempleado"
            GetProperty = mEmpleado.NombreCompleto
    End Select
End Function

Private Function GetValorAnterior(campo As String) As String

Dim aux As String
aux = LCase(campo)

Select Case aux
        Case "nroincidenteinterno", "nroincidente", "fecha", "diagnostico", "observaciones", "telefonoauxiliar", "hora llamada", "qth", "vl", "operador"
            GetValorAnterior = mValorAnterior
        'es lo mismo, los separo por las dudas...
        Case "salidapreinspeccion", "policiaacargo", "llegadapreinspeccion", "salidadotacion", "cronica", "descripcionperjuicioscampo", "materialescombustibles", "descripcionperjuiciosvivienda", "accesopor", "referencias", "aseguradora", "descripcionmaterial", "equiposespeciales", "poliza", "ambientesafectadosvivienda", "hectareasafectadascampo", "instalacionelectrica", "instalaciongas", "policiacantidad", "policiamovil", "semedico", "secentroasistencial", "semedicomp"
            GetValorAnterior = mValorAnterior
        Case "sintoma"
            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = SintomasLocal.Item(CLng(mValorAnterior)).NombreCompuesto
            End If
        Case "areaprotegida"

            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = AreasProtegidasLocal.Item(CLng(mValorAnterior)).NombreArea
            End If
        Case "obrasocial"
            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = ObrasSocialesLocal.Item(CLng(mValorAnterior)).Nombre
            End If
        Case "servicioemergencia"
            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = ServiciosEmergenciaLocal.Item(CLng(mValorAnterior)).Nombre
            End If
        Case "afiliadopropio"
            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = AfiliadosLocal.Item(CLng(mValorAnterior)).NombreCompleto
            End If
        Case "afiliadoexterno"
            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                Dim at As Atencion
                Set at = AtencionesLocal.LoadById(mIdAtencion)
                GetValorAnterior = at.AfiliadoExterno.NombreCompleto
            End If
        Case "estado"
             GetValorAnterior = IIf(CLng(mValorAnterior) = 1, "Pendiente", "Finalizado")
        Case "telefono"
            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = TelefonosLocal.Item(CLng(mValorAnterior)).Numero
            End If
        Case "tipocodigo"
            If CLng(mValorAnterior) = 0 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = TiposCodigoLocal.Item(CLng(mValorAnterior)).Nombre
            End If
        Case "servicio", "copago", "montoabonado"
            If CCur(mValorAnterior) = -1 Then
                GetValorAnterior = "<Sin valor>"
            Else
                GetValorAnterior = mValorAnterior
            End If
       
       'prueba
        Case "direccion"
        GetValorAnterior = mValorAnterior
        Case "direccion origen"
        GetValorAnterior = mValorAnterior
        Case "direccion destino"
        GetValorAnterior = mValorAnterior
        Case "dotaciones"
        GetValorAnterior = mValorAnterior
        'no   Case  "idDespachador"
        
         'warning: faltan
        'Private mSEAmbulancias As Boolean

        'Private mColaboracionBomberos As ColaboracionManager
        'Private mInvolucrados As InvolucradoManager
        'Private mVehiculos As VehiculoManager
  End Select
End Function
Private Function GetValorNuevo(campo As String) As String

Dim aux As String
aux = LCase(campo)

Select Case aux
        Case "nroincidenteinterno", "nroincidente", "fecha", "diagnostico", "observaciones", "telefonoauxiliar", "hora llamada", "qth", "vl", "operador"
            GetValorNuevo = mValorNuevo
        'es lo mismo q arriba, los separo por las dudas...
        Case "salidapreinspeccion", "policiaacargo", "llegadapreinspeccion", "salidadotacion", "cronica", "descripcionperjuicioscampo", "materialescombustibles", "descripcionperjuiciosvivienda", "accesopor", "referencias", "aseguradora", "descripcionmaterial", "equiposespeciales", "poliza", "ambientesafectadosvivienda", "hectareasafectadascampo", "instalacionelectrica", "instalaciongas", "policiacantidad", "policiamovil", "semedico", "secentroasistencial", "semedicomp"
            GetValorNuevo = mValorNuevo
        Case "sintoma"
            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = SintomasLocal.Item(CLng(mValorNuevo)).NombreCompuesto
            End If
        Case "areaprotegida"

            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = AreasProtegidasLocal.Item(CLng(mValorNuevo)).NombreArea
            End If
        Case "obrasocial"
            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = ObrasSocialesLocal.Item(CLng(mValorNuevo)).Nombre
            End If
        Case "servicioemergencia"
            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = ServiciosEmergenciaLocal.Item(CLng(mValorNuevo)).Nombre
            End If
        Case "afiliadopropio"
            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = AfiliadosLocal.Item(CLng(mValorNuevo)).NombreCompleto
            End If
        
        Case "afiliadoexterno"
            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                Dim at As Atencion
                Set at = AtencionesLocal.LoadById(mIdAtencion)
                GetValorNuevo = at.AfiliadoExterno.NombreCompleto
            End If
        Case "estado"
            GetValorNuevo = IIf(CLng(mValorNuevo) = 1, "Pendiente", "Finalizado")
        Case "telefono"
            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = TelefonosLocal.Item(CLng(mValorNuevo)).Numero
            End If
        Case "tipocodigo"
            If CLng(mValorNuevo) = 0 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = TiposCodigoLocal.Item(CLng(mValorNuevo)).Nombre
            End If
        Case "servicio", "copago", "montoabonado"
            If CCur(mValorNuevo) = -1 Then
                GetValorNuevo = "<Sin valor>"
            Else
                GetValorNuevo = mValorNuevo
            End If
            
       'prueba
        Case "direccion"
            GetValorNuevo = mValorNuevo
        Case "direccion origen"
            GetValorNuevo = mValorNuevo
        Case "direccion destino"
            GetValorNuevo = mValorNuevo
        Case "dotaciones"
            GetValorNuevo = mValorNuevo
        
        'warning: faltan
        'Private mSEAmbulancias As Boolean

        'Private mColaboracionBomberos As ColaboracionManager
        'Private mInvolucrados As InvolucradoManager
        'Private mVehiculos As VehiculoManager

  End Select
End Function

Friend Sub Save(pTrans As DataBaseLayer.Transaccion)
    Dim params As New ParametrosManager
  
    params.Add "ValorNuevo", mValorNuevo
    params.Add "ValorAnterior", mValorAnterior
    params.Add "idEmpleado", mEmpleado.Id
    params.Add "Campo", mCampo
    params.Add "Fecha", mFecha
    params.Add "Hora", mHora
    params.Add "idAtencion", mIdAtencion
    
    pTrans.AddInsertCommand tDetalleSeguimiento, params
   
End Sub

