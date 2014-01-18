VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlergiaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A8038302EE"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private mCol As Collection

Private mAlergiasOriginal As Collection

'##ModelId=46A80BDC004E
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    mCol.Remove "k" + Str(pId)

    Exit Sub
RemoveErr:
    Call ErrorLog("AlergiaManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A80BE8001F
Public Sub AddItem(pAlergia As Alergia)
    On Error GoTo AddItemErr

    mCol.Add pAlergia, "k" + Str(pAlergia.Id)

    Exit Sub
AddItemErr:
If Err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "AlergiaManager", "AddItem", Err.Description
End If
End Sub

Public Property Get Item(pId As Long) As Alergia
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pName As String) As Alergia
    On Error Resume Next
    Dim a As Alergia
    For Each a In Me
        If LCase(Trim(a.Nombre)) = LCase(Trim(pName)) Then
            Set ItemByName = a
            Exit For
        End If
    Next
End Property

Public Function Nuevo(pNombre As String) As Alergia
    If ItemByName(pNombre) Is Nothing Then
        Dim c As New Alergia
        c.Id = BD.getIdDisponible(tAlergia)
        c.Nombre = pNombre
        c.Guardar
        Me.AddItem c
        Set Nuevo = c
    End If
End Function

Friend Function Add(pId As Long, pNombre As String) As Alergia
Dim a As New Alergia
a.Id = pId
a.Nombre = pNombre

mCol.Add a, "k" + Str(a.Id)
Set Add = a
Set a = Nothing
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tAlergia)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "AlergiasManager", "cargarTodos ", Err.Description
End Sub

Friend Sub CargarXAfiliado(pAfiliado As Object)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq
Dim tabla As String
If TypeOf pAfiliado Is Afiliado Then
    tabla = "AlergiaXAfiliado"
ElseIf TypeOf pAfiliado Is AfiliadoExterno Then
    tabla = "AlergiaXAfiliadoExterno"
End If
Dim rs
Set rs = BD.SelectSQL("SELECT idAlergia FROM " + tabla + " where idAfiliado=" + Str(pAfiliado.Id))
While Not rs.EOF
        
    AddItem AlergiasLocal.Item(rs!idAlergia)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "AlegriaManager", "cargarXAfiliado ", Err.Description
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

Dim c As Alergia
Dim params As ParametrosManager
For Each c In Me
    Set params = New ParametrosManager
    params.Add "idAlergia", c.Id
    params.Add "idAfiliado", idAfiliado
    pTrans.AddInsertCommand pTabla, params
Next

End Sub

Friend Sub BeginEdit()
    Clone
End Sub

Friend Sub CancelChanges()
    If Not mAlergiasOriginal Is Nothing Then
        Set mCol = mAlergiasOriginal
        Set mAlergiasOriginal = Nothing
    End If
End Sub

Friend Sub SaveChanges(pTrans As Transaccion, idAfiliado As Long, pTabla As eTablas)
    'borro todos los cargos de este empleado
    Dim params As New ParametrosManager
    params.Add "idAfiliado", idAfiliado, True
    pTrans.AddDeleteCommand pTabla, params
    'los vuelvo a guardar
    Me.Guardar pTrans, idAfiliado, pTabla
    
End Sub

Public Function Exists(pId As Long) As Boolean
    Dim t As Alergia
    Set t = Me.Item(pId)
    Exists = (Not t Is Nothing)
End Function

Private Sub Clone()
    Dim cm As Collection
    Set cm = New Collection
    Dim c As Alergia
    For Each c In Me
        cm.Add c, "k" + Str(c.Id)
    Next
    Set mAlergiasOriginal = cm
End Sub

