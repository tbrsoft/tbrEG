VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VehiculoManager"
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
    Call ErrorLog("VehiculoManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pVehiculo As Vehiculo)
    On Error GoTo AddItemErr

    mCol.Add pVehiculo ', "k" + Str(pVehiculo.Id)

    Exit Sub
AddItemErr:
If Err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "VehiculoManager", "AddItem", Err.Description
End If
End Sub

Friend Function Add(pId As Long, pTipo As String, pMarca As String, pModelo, pPatente As String, pPerjuicios As String, pColor As String) As Vehiculo
    Dim c As New Vehiculo
    c.Id = pId
    c.Marca = pMarca
    c.Modelo = pModelo
    c.Patente = pPatente
    c.Color = pColor
    c.Perjuicios = pPerjuicios
    c.tipo = pTipo
    
    AddItem c
    Set Add = c
    Set c = Nothing
End Function

Public Function Nuevo(pTipo As String, pMarca As String, pModelo, pPatente As String, pPerjuicios As String, pColor As String) As Vehiculo
    Dim c As New Vehiculo
    c.Marca = pMarca
    c.Modelo = pModelo
    c.Patente = pPatente
    c.Color = pColor
    c.Perjuicios = pPerjuicios
    c.tipo = pTipo
        
    AddItem c
    Set Nuevo = c
    Set c = Nothing
End Function

Friend Sub GuardarPorSiniestro(trans As Transaccion, pIdSiniestro As Long)
    Dim v As Vehiculo
    For Each v In Me
        v.GuardarPorSiniestro trans, pIdSiniestro
    Next
End Sub

Friend Sub CargarPorSiniestro(pIdSiniestro As Long)
    'muy probablemente esto deberia ir en la capa de datos...
    On Error GoTo errMaq
    TERR.Anotar "abba", pIdSiniestro
    Dim rs
    Set rs = BD.SelectSQL("Select * from Vehiculo where idSiniestro=" + Str(pIdSiniestro))
    
    While Not rs.EOF
        TERR.Anotar "abbc", NN_num(rs!Id), NN_str(rs!tipo), NN_str(rs!Marca), NN_str(rs!Modelo)
        TERR.Anotar "abbd", NN_str(rs!Patente), NN_str(rs!Perjuicios), NN_str(rs!Color)
        
        Add NN_num(rs!Id), _
            NN_str(rs!tipo), NN_str(rs!Marca), NN_str(rs!Modelo), _
            NN_str(rs!Patente), NN_str(rs!Perjuicios), NN_str(rs!Color)
            
        rs.MoveNext
    
    Wend
    
    Exit Sub
errMaq:
    ErrorLog "VehiculoManager", "CargarPorSiniestro", Err.Description
End Sub

'Friend Sub CargarTodos()
''muy probablemente esto deberia ir en la capa de datos...
'On Error GoTo errMaq
'
'Dim rs
'Set rs = BD.SelectTabla(tVehiculo)
'While Not rs.EOF
'
'    Add rs!Id, rs!Nombre, rs!Patente, rs!Estado
'    rs.MoveNext
'
'Wend
'Exit Sub
'errMaq:
'ErrorLog "Error VehiculoManager, sub cargarTodos " + Err.Description
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

Public Property Get ItemById(pId As Long) As Vehiculo
    On Error Resume Next
    Dim v As Vehiculo
    For Each v In Me
        If v.Id = pId Then
            Set ItemById = v
            Exit For
        End If
    Next
End Property
