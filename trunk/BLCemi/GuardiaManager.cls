VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GuardiaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("GuardiaManager", "Remove Sub", err.Description)
End Sub

Public Sub AddItem(pGuardia As Guardia)
    On Error GoTo AddItemErr

    mCol.Add pGuardia, "k" + Str(pGuardia.Id)
    
    Exit Sub
AddItemErr:
    ErrorLog "GuardiaManager", "AddItem", err.Description
End Sub

Public Function Nueva(pIdEmpleado As Long, pMonto As Currency, pAdelanto As Currency, pFecha As Date, pPlus As Currency, pObservaciones As String) As Guardia
    Dim c As New Guardia
    c.Adelanto = pAdelanto
    c.Fecha = pFecha
    c.Plus = pPlus
    c.Monto = pMonto
    c.Observaciones = pObservaciones
    c.Save pIdEmpleado
    AddItem c
    Set Nueva = c
    Set c = Nothing
End Function

Friend Function Add(pId As Long, pMonto As Currency, pAdelanto As Currency, pFecha As Date, pPlus As Currency, pObservaciones As String) As Guardia
    Dim c As New Guardia
    c.Id = pId
    c.Adelanto = pAdelanto
    c.Fecha = pFecha
    c.Plus = pPlus
    c.Monto = pMonto
    c.Observaciones = pObservaciones
    AddItem c
    Set Add = c
    Set c = Nothing
End Function

Public Sub LoadNoRendidasByEmpleado(pIdEmpleado As Long)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * From Guardia where idLiquidacion=0 AND idEmpleado=" + Str(pIdEmpleado))
CargarRS rs

Exit Sub
errMaq:
ErrorLog "GuardiaManager", "cargarTodos ", err.Description
End Sub

Public Sub LoadByLiquidacion(pIdLiquidacion As Long)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * From Guardia where idliquidacion=" + Str(pIdLiquidacion))
CargarRS rs
Exit Sub
errMaq:
ErrorLog "GuardiaManager", "cargarTodos ", err.Description
End Sub

Friend Sub Liquidar(trans As Transaccion, pIdLiquidacion As Long)
    Dim g As Guardia
    For Each g In Me
        g.Liquidar pIdLiquidacion, trans
    Next
End Sub

Private Sub CargarRS(rs)

While Not rs.EOF
    Add rs!Id, rs!Monto, rs!Adelanto, rs!Fecha, rs!Plus, "" & rs!Observaciones
    rs.MoveNext
Wend

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

Public Property Get Item(pId As Long) As Guardia
On Error Resume Next
Set Item = mCol.Item("k" + Str(pId))
End Property



