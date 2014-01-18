VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MovilManager"
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
    Call ErrorLog("MovilManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pMovil As Movil)
    On Error GoTo AddItemErr

    mCol.Add pMovil, "k" + Str(pMovil.Id)

    Exit Sub
AddItemErr:
If err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "MovilManager", "AddItem", err.Description
End If
End Sub

Friend Function Add(pId As Long, pNombre As String, pPatente As String, pEstado As eEstadoMovil) As Movil
    Dim c As New Movil
    c.Id = pId
    c.Nombre = pNombre
    c.Estado = pEstado
    c.Patente = pPatente
    AddItem c
    Set Add = c
    Set c = Nothing
End Function

Public Function Nuevo(pNombre As String, pPatente As String, pEstado As eEstadoMovil) As Movil
    Dim c As New Movil
    c.Nombre = pNombre
    c.Estado = pEstado
    c.Patente = pPatente
    c.Insert
    AddItem c
    Set Nuevo = c
    Set c = Nothing
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tMovil)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre, rs!Patente, rs!Estado
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "MovilManager", "cargarTodos ", err.Description
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

Public Property Get Item(pId As Long) As Movil
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

