VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InfoContableManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

'##ModelId=46A8DFD60242
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("InfoContableManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pInfoContable As InfoContable)
    On Error GoTo AddItemErr
    'uso el idatencion porq hay una sola infocontable por liquidacion
    mCol.Add pInfoContable, "k" + Str(pInfoContable.idAtencion)

    Exit Sub
AddItemErr:
    ErrorLog "InfoContableManager", "AddItem", err.Description
End Sub

Friend Function Add(pServicio As Currency, pCoseguro As Currency, pAbonado As Currency, pCondIva As eCondicionIVA, pIdAtencion As Long, pMontoIva As Currency) As InfoContable
Dim c As New InfoContable

c.Constructor pServicio, pCoseguro, pAbonado, pCondIva, pIdAtencion, pMontoIva

AddItem c
Set Add = c
Set c = Nothing
End Function

'Friend Sub CargarTodos()
''muy probablemente esto deberia ir en la capa de datos...
'On Error GoTo errMaq
'
'Dim rs
'Set rs = BD.SelectTabla(tParentezco)
'While Not rs.EOF
'
'    Add rs!id, rs!Nombre
'    rs.MoveNext
'
'Wend
'Exit Sub
'errMaq:
'ErrorLog "Error EnfermedadesManager, sub cargarTodos " + Err.Description
'End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

'Public Property Get Item(pId As Long) As Parentezco
'On Error Resume Next
'Set Item = mCol.Item("k" + Str(pId))
'End Property
'
'Public Property Get ItemByName(pNombre As String) As Parentezco
'Dim b As Parentezco
'For Each b In Me
'    If LCase(b.Nombre) = LCase(pNombre) Then
'        Set ItemByName = b
'        Exit For
'    End If
'Next
'End Property

'Public Function Nuevo(pNombre As String) As Parentezco
'    If ItemByName(pNombre) Is Nothing Then
'        Dim c As New Parentezco
'        c.id = BD.getIdDisponible(tParentezco)
'        c.Nombre = pNombre
'        c.Guardar
'        Me.AddItem c
'        Set Nuevo = c
'    End If
'End Function

Public Sub LoadByLiquidacion(pIdLiquidacion As Long)
    On Error GoTo errMaq

    Dim rs
    Dim where As String
    where = " idLiquidacion= " + Str(pIdLiquidacion)
    Set rs = BD.SelectSQL("select copago, abonado, condiva, id, servicio,montoiva FROM atencion where " + where)
    
    CargarRS rs
    
    Exit Sub
errMaq:
    ErrorLog "InfoContableManager", "LoadByLiquidacion", err.Description
End Sub

Public Sub LoadNoRendidosByEmpresa(pIdEmpresa As Long, pTipo As eDestinoCodigo)
    On Error GoTo errMaq

    Dim rs
    Dim where As String
    where = IIf(pTipo = eDCObraSocial, "idObraSocial", "idServicioEmergencia") + "=" + Str(pIdEmpresa) + " and idLiquidacion=0"
    Set rs = BD.SelectSQL("select copago, abonado, condiva, id, servicio, montoIva  FROM atencion where " + where)
    
    CargarRS rs
    
    Exit Sub
errMaq:
    ErrorLog "InfoContableManager", "LoadNoRendidosByEmpresa", err.Description
End Sub

Private Sub CargarRS(rs)
    On Error GoTo errMaq
    
    While Not rs.EOF
            
        Add rs!Servicio, rs!copago, rs!abonado, rs!condIva, rs!Id, rs!MontoIVA
        rs.MoveNext
        
    Wend
    Exit Sub
errMaq:
    ErrorLog "InfoContableManager", "CargarRs", err.Description

End Sub

Friend Sub Save(pIdLiquidacion As Long, trans As Transaccion)
    'es mas bien un update
    Dim infoC As InfoContable
    For Each infoC In Me
        infoC.Save pIdLiquidacion, trans
    Next
End Sub

Public Function GetSaldo() As Currency
    Dim ic As InfoContable
    Dim saldoAux As Currency
    For Each ic In Me
        saldoAux = saldoAux + ic.GetSaldo
    Next
    GetSaldo = saldoAux
End Function
