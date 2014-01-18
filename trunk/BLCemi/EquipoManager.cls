VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EquipoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event ItemRemoved(pEquipo As Equipo)
Public Event ItemAdded(pEquipo As Equipo)

Private mCol As Collection
Private mColEliminados As Collection


Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    mCol.Remove "k" + Str(pId)

    Exit Sub
RemoveErr:
    Call ErrorLog("EquipoManager", "Remove Sub", err.Description)
End Sub

Public Sub AddItem(pEquipo As Equipo)
    On Error GoTo AddItemErr
        mCol.Add pEquipo, "k" + Str(pEquipo.Id)
        RaiseEvent ItemAdded(pEquipo)
    Exit Sub
AddItemErr:
 If err.Number <> 457 Then 'si quiere insertar dos veces el mismo elemento lo ignoro, no se si es lo mejor...
    ErrorLog "EquipoManager", "AddItem", err.Description
End If
End Sub

Public Function Nuevo(pMovil As Movil, pDotacion As EmpleadoManager) As Equipo
    Dim e As New Equipo
   
    Set e.Movil = pMovil
    Set e.Dotacion = pDotacion
    e.Insert
    mCol.Add e, "k" + Str(e.Id)
    Set Nuevo = e
    Set e = Nothing
End Function

Friend Function Add(pId As Long, pMovil As Movil, pEliminado As Boolean) As Equipo
    Dim e As New Equipo
    e.Id = pId
    Set e.Movil = pMovil
      
    If Not pEliminado Then
        mCol.Add e, "k" + Str(e.Id)
    Else
        mColEliminados.Add e, "k" + Str(e.Id)
    End If
    Set Add = e
    Set e = Nothing
End Function

Friend Sub CargarTodos()
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tEquipo)
While Not rs.EOF
        
    Add rs!Id, MovilesLocal.Item(rs!idMovil), CBool(rs!eliminado)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "EquiposManager", "cargarTodos ", err.Description
End Sub

Friend Sub CargarPorAtencion(idAtencion As Long)
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * from EquipoXAtencion where idAtencion=" + Str(idAtencion))
While Not rs.EOF
    'ver como implementar lo del equipo responsable
    AddItem EquiposLocal.Item(rs!idEquipo)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "EquiposManager", "cargarPorAtencion ", err.Description

End Sub

Friend Sub Guardar(pTrans As Transaccion, idAtencion As Long)

Dim e As Equipo
Dim params As ParametrosManager
For Each e In Me
    Set params = New ParametrosManager
    params.Add "idEquipo", e.Id
    params.Add "idAtencion", idAtencion
    'ver equipo responsable
    pTrans.AddInsertCommand tEquipoXAtencion, params
Next

End Sub

Friend Sub GuardarCambios(pTrans As Transaccion, idAtencion As Long)
    'borro todos los equipos de esta atencion
    Dim params As New ParametrosManager
    params.Add "idAtencion", idAtencion, True
    pTrans.AddDeleteCommand tEquipoXAtencion, params
    'los vuelvo a guardar
    Me.Guardar pTrans, idAtencion

End Sub

Public Function GetEliminados() As EquipoManager
    Dim em As New EquipoManager
    Dim e As Equipo
    For Each e In mColEliminados
        em.AddItem e
    Next
    Set GetEliminados = em
End Function

Public Sub DarItemDeBaja(pId As Long)
   'esto es porq la unica coleccion q tiene los elementos eliminados es equiposLocal
   EquiposLocal.DarItemDeBajaLocal pId
   RaiseEvent ItemRemoved(EquiposLocal.Item(pId))
End Sub

Public Sub DarItemDeBajaLocal(pId As Long)
    Dim e As Equipo
    Set e = Me.Item(pId)
    e.DarDeBaja
    mColEliminados.Add e, "k" + Str(e.Id)
    mCol.Remove "k" + Str(e.Id)
End Sub

Public Sub RestaurarItem(pId As Long)
    'esto es porq la unica coleccion q tiene los elementos eliminados es empleadosLocal
    EquiposLocal.RestaurarItemLocal pId
End Sub

Friend Sub RestaurarItemLocal(pId As Long)
    Dim e As Equipo
    Set e = Me.Item(pId) 'en item tambien devuelvo los eliminados
    e.Restaurar
    mColEliminados.Remove "k" + Str(e.Id)
    AddItem e
    
End Sub

Public Function GetNombreMoviles() As String
    Dim aux As String
    If Me.Count = 0 Then
        GetNombreMoviles = ""
    Else
        Dim e As Equipo
        For Each e In Me
            aux = aux + ", " + e.Movil.Nombre
        Next
        GetNombreMoviles = Right(aux, Len(aux) - 2)
    End If
End Function

Private Sub Class_Initialize()
    Set mCol = New Collection
    Set mColEliminados = New Collection
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

Public Property Get Item(pId As Long) As Equipo
On Error GoTo errman    'si salta un error probablemente este eliminado
    Set Item = mCol.Item("k" + Str(pId))
    Exit Property
errman:
    On Error Resume Next 'si salta un error aca, me esta pidiendo un empleado q no tengo
    Set Item = mColEliminados.Item("k" + Str(pId))
End Property
