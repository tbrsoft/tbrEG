VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnfermedadManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A803990196"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private mCol As Collection

Private mEnfermedadOriginal As Collection

'##ModelId=46A80BD603B9
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    mCol.Remove "k" + Str(pId)

    Exit Sub
RemoveErr:
    Call ErrorLog("EnfermedadManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A80BE00213
Public Sub AddItem(pEnfermedad As Enfermedad)
    On Error GoTo AddItemErr

    mCol.Add pEnfermedad, "k" + Str(pEnfermedad.Id)

    Exit Sub
AddItemErr:
If Err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "EnfermedadManager", "AddItem", Err.Description
End If
End Sub

Public Function Nuevo(pNombre As String) As Enfermedad
    If ItemByName(pNombre) Is Nothing Then
        Dim c As New Enfermedad
        c.Id = BD.getIdDisponible(tEnfermedad)
        c.Nombre = pNombre
        c.Guardar
        Me.AddItem c
        Set Nuevo = c
    End If
End Function

Friend Function Add(pId As Long, pNombre As String) As Enfermedad
Dim e As New Enfermedad
e.Id = pId
e.Nombre = pNombre

mCol.Add e, "k" + Str(e.Id)
Set Add = e
Set e = Nothing
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tEnfermedad)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "EnfermedadesManager", "cargarTodos ", Err.Description
End Sub

Friend Sub CargarXAfiliado(pAfiliado As Object)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim tabla As String
If TypeOf pAfiliado Is Afiliado Then
    tabla = "EnfermedadXAfiliado"
ElseIf TypeOf pAfiliado Is AfiliadoExterno Then
    tabla = "EnfermedadXAfiliadoExterno"
End If

Dim rs
Set rs = BD.SelectSQL("SELECT idEnfermedad FROM " + tabla + " where idAfiliado=" + Str(pAfiliado.Id))
While Not rs.EOF
        
    AddItem EnfermedadesLocal.Item(rs!idEnfermedad)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "EnfermedadesManager", "cargarXAfiliado ", Err.Description
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

Friend Sub Guardar(pTrans As Transaccion, idAfiliado As Long, pTabla As eTablas)

Dim c As Enfermedad
Dim params As ParametrosManager
For Each c In Me
    Set params = New ParametrosManager
    params.Add "idEnfermedad", c.Id
    params.Add "idAfiliado", idAfiliado
    pTrans.AddInsertCommand pTabla, params
Next

End Sub

Friend Sub BeginEdit()
    Clone
End Sub

Friend Sub CancelChanges()
    If Not mEnfermedadOriginal Is Nothing Then
        Set mCol = mEnfermedadOriginal
        Set mEnfermedadOriginal = Nothing
    End If
End Sub

Friend Sub SaveChanges(pTrans As Transaccion, idAfiliado As Long, pTabla As eTablas)
    'borro todos las enfermedades de este afiliado
    Dim params As New ParametrosManager
    params.Add "idAfiliado", idAfiliado, True
    pTrans.AddDeleteCommand pTabla, params
    'los vuelvo a guardar
    Me.Guardar pTrans, idAfiliado, pTabla
    
End Sub

Public Property Get Item(pId As Long) As Enfermedad
    'si no existe devuelve nothing
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pName As String) As Enfermedad
    On Error Resume Next
    Dim a As Enfermedad
    For Each a In Me
        If LCase(Trim(a.Nombre)) = LCase(Trim(pName)) Then
            Set ItemByName = a
            Exit For
        End If
    Next
End Property

Public Function Exists(pId As Long) As Boolean
    Dim t As Enfermedad
    Set t = Me.Item(pId)
    Exists = (Not t Is Nothing)
End Function

Private Sub Clone()
    Dim cm As Collection
    Set cm = New Collection
    Dim c As Enfermedad
    For Each c In Me
        cm.Add c, "k" + Str(c.Id)
    Next
    Set mEnfermedadOriginal = cm
End Sub

