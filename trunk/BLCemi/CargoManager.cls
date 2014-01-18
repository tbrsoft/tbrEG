VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CargoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A8DFB402BF"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Enum eCargosFijos
    eParamedico = 1
    eDespachador = 2
    eAdministracion = 3
    eCobrador = 4
    eMedico = 5
    eChofer = 6
End Enum

Private mCol As Collection

Private mCargosOriginal As Collection

'##ModelId=46A8DFD60242
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    mCol.Remove "k" + Str(pId)

    Exit Sub
RemoveErr:
    Call ErrorLog("CargoManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pCargo As Cargo)
    On Error GoTo AddItemErr

    mCol.Add pCargo, "k" + Str(pCargo.Id)

    Exit Sub
AddItemErr:
If Err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "CargoManager", "AddItem", Err.Description
End If
End Sub

Public Function Nuevo(pNombre As String) As Cargo
    Dim c As New Cargo
    c.Id = BD.getIdDisponible(tCargo)
    c.Nombre = pNombre
    c.Guardar '??????? no se si va aca
    Me.AddItem c
    Set Nuevo = c
End Function

Friend Function Add(pId As Long, pNombre As String) As Cargo
Dim c As New Cargo
c.Id = pId
c.Nombre = pNombre

mCol.Add c, "k" + Str(c.Id)
Set Add = c
Set c = Nothing
End Function

Friend Sub CargarPorEmpleado(pIdEmpleado As Long)
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * from CargoXEmpleado where idEmpleado=" + Str(pIdEmpleado))
While Not rs.EOF
        
    AddItem CargosLocal.Item(rs!idCargo)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "CargosManager", "cargarPorEmpleado ", Err.Description

End Sub

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

'agrego los cargos fijos
Select Case CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento
    Case 0:
        ErrorLog "CargosManager", "cargarTodos", "Modofuncionamiento no esta seteado."
    Case 1:
        Add 1, "Paramedico"
        Add 2, "Despachador"
        Add 3, "Administracion"
        Add 4, "Cobrador"
        Add 5, "Medico"
        Add 6, "Chofer"
    Case 2:
        Add 1, "Bombero"
        Add 2, "Despachador"
        Add 3, "Administracion"
        Add 4, "Cobrador"
        Add 5, "Medico" 'ver
        Add 6, "Chofer"
End Select

Dim rs
Set rs = BD.SelectTabla(tCargo)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "CargosManager", " cargarTodos ", Err.Description
End Sub

Friend Sub Guardar(pTrans As Transaccion, IdEmpleado As Long)

Dim c As Cargo
Dim params As ParametrosManager
For Each c In Me
    Set params = New ParametrosManager
    params.Add "idCargo", c.Id
    params.Add "idEmpleado", IdEmpleado
    pTrans.AddInsertCommand tCargoXEmpleado, params
Next

End Sub

Public Sub BeginEdit()
    Clone
End Sub

Public Sub CancelChanges()
    If Not mCargosOriginal Is Nothing Then
        Set mCol = mCargosOriginal
        Set mCargosOriginal = Nothing
    End If
End Sub

Friend Sub SaveChanges(pTrans As Transaccion, IdEmpleado As Long)
    'borro todos los cargos de este empleado
    Dim params As New ParametrosManager
    params.Add "idEmpleado", IdEmpleado, True
    pTrans.AddDeleteCommand tCargoXEmpleado, params
    'los vuelvo a guardar
    Me.Guardar pTrans, IdEmpleado
    
End Sub

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

Public Property Get Item(pId As Long) As Cargo
    'si no existe devuelve nothing
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pName As String) As Cargo
    'si no existe devuelve nothing
    On Error Resume Next
    Dim c As Cargo
    For Each c In Me
        If LCase(c.Nombre) = LCase(pName) Then
            Set ItemByName = c
            Exit For
        End If
    Next
End Property

Public Function Exists(pId As Long) As Boolean
    Dim t As Cargo
    Set t = Me.Item(pId)
    Exists = (Not t Is Nothing)
End Function

Private Sub Clone()
    Dim cm As Collection
    Set cm = New Collection
    Dim c As Cargo
    For Each c In Me
        cm.Add c.Clone, "k" + Str(c.Id)
    Next
    Set mCargosOriginal = cm
End Sub

Public Function CargosToString() As String
    Dim c As Cargo
    Dim aux As String
    For Each c In Me
        aux = aux + ", " + c.Nombre
    Next
    If aux <> "" Then
        CargosToString = Right(aux, Len(aux) - 2)
    Else
        CargosToString = ""
    End If
    
End Function

