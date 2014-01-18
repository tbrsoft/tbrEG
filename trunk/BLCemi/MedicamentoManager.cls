VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MedicamentoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Private mMedicamentosOriginal As Collection

'##ModelId=46A80BDC004E
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr
    
    mCol.Remove "k" + Str(pId)
    
    Exit Sub
RemoveErr:
    Call ErrorLog("MedicamentoManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A80BE8001F
Public Sub AddItem(pMedicamento As Medicamento)
    On Error GoTo AddItemErr

    mCol.Add pMedicamento, "k" + Str(pMedicamento.Id)

    Exit Sub
AddItemErr:
If Err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "MedicamentoManager", "AddItem", Err.Description
End If
End Sub

Public Function Nuevo(pNombre As String) As Medicamento
    If ItemByName(pNombre) Is Nothing Then
        Dim c As New Medicamento
        c.Id = BD.getIdDisponible(tMedicamento)
        c.Nombre = pNombre
        c.Guardar
        Me.AddItem c
        Set Nuevo = c
    End If
End Function

Public Property Get Item(pId As Long) As Medicamento
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pName As String) As Medicamento
    On Error Resume Next
    Dim a As Medicamento
    For Each a In Me
        If LCase(Trim(a.Nombre)) = LCase(Trim(pName)) Then
            Set ItemByName = a
            Exit For
        End If
    Next
End Property


Friend Function Add(pId As Long, pNombre As String) As Medicamento
Dim a As New Medicamento
a.Id = pId
a.Nombre = pNombre

AddItem a
Set Add = a
Set a = Nothing
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tMedicamento)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "MedicamentosManager", "cargarTodos ", Err.Description
End Sub

Friend Sub CargarXAfiliado(pAfiliado As Object)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq
Dim tabla As String
If TypeOf pAfiliado Is Afiliado Then
    tabla = "MedicamentoXAfiliado"
ElseIf TypeOf pAfiliado Is AfiliadoExterno Then
    tabla = "MedicamentoXAfiliadoExterno"
End If
Dim rs
Set rs = BD.SelectSQL("SELECT idMedicamento FROM " + tabla + " where idAfiliado=" + Str(pAfiliado.Id))
While Not rs.EOF
        
    AddItem MedicamentosLocal.Item(rs!idMedicamento)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "MedicamentosManager", "cargarXAfiliado ", Err.Description
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

Dim c As Medicamento
Dim params As ParametrosManager
For Each c In Me
    Set params = New ParametrosManager
    params.Add "idMedicamento", c.Id
    params.Add "idAfiliado", idAfiliado
    pTrans.AddInsertCommand pTabla, params
Next

End Sub

Friend Sub BeginEdit()
    Clone
End Sub

Friend Sub CancelChanges()
    If Not mMedicamentosOriginal Is Nothing Then
        Set mCol = mMedicamentosOriginal
        Set mMedicamentosOriginal = Nothing
    End If
End Sub

Friend Sub SaveChanges(pTrans As Transaccion, idAfiliado As Long, pTabla As eTablas)
    'borro todos los Medicamentos de este Afiliado
    Dim params As New ParametrosManager
    params.Add "idAfiliado", idAfiliado, True
    pTrans.AddDeleteCommand pTabla, params
    'los vuelvo a guardar
    Me.Guardar pTrans, idAfiliado, pTabla
    
End Sub

Public Function Exists(pId As Long) As Boolean
    Dim t As Medicamento
    Set t = Me.Item(pId)
    Exists = (Not t Is Nothing)
End Function

Private Sub Clone()
    Dim cm As Collection
    Set cm = New Collection
    Dim c As Medicamento
    For Each c In Me
        cm.Add c, "k" + Str(c.Id)
    Next
    Set mMedicamentosOriginal = cm
End Sub

