VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiqEmpleadoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...
    MsgBox "implementar"

    Exit Sub
RemoveErr:
    Call ErrorLog("LiqEmpleadoManager", "Remove Sub", err.Description)
End Sub

Public Sub AddItem(pLiquidacionEmpleado As LiquidacionEmpleado)
    On Error GoTo AddItemErr

    mCol.Add pLiquidacionEmpleado, "k" + Str(pLiquidacionEmpleado.Id)
    
    Exit Sub
AddItemErr:
    ErrorLog "LiqEmpleadoManager", "AddItem", err.Description
End Sub

Public Function Nueva(pIdEmpleado As Long, pMonto As Currency, pAdelanto As Currency, pFecha As Date, pPlus As Currency, pObservaciones As String, pDetalle As GuardiaManager, pMes As Integer, pYear As Integer) As LiquidacionEmpleado
    Dim c As New LiquidacionEmpleado
    c.Adelanto = pAdelanto
    c.Fecha = pFecha
    c.Plus = pPlus
    c.Monto = pMonto
    c.Mes = pMes
    c.Year = pYear
    c.Observaciones = pObservaciones
    Set c.Detalle = pDetalle
    c.Save pIdEmpleado
    AddItem c
    Set Nueva = c
    Set c = Nothing
End Function

Friend Function Add(pId As Long, pMonto As Currency, pAdelanto As Currency, pFecha As Date, pPlus As Currency, pObservaciones As String, pMes As Integer, pYear As Integer) As LiquidacionEmpleado
    Dim c As New LiquidacionEmpleado
    c.Id = pId
    c.Adelanto = pAdelanto
    c.Fecha = pFecha
    c.Plus = pPlus
    c.Monto = pMonto
    c.Mes = pMes
    c.Year = pYear
    c.Observaciones = pObservaciones
    AddItem c
    Set Add = c
    Set c = Nothing
End Function

Public Sub LoadByEmpleado(pIdEmpleado As Long)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * From LiquidacionEmpleado where idEmpleado=" + Str(pIdEmpleado))
CargarRS rs

Exit Sub
errMaq:
ErrorLog "LiquidacionEmpleadoManager", "LoadByEmpleado", err.Description
End Sub

Public Sub LoadByPeriodo(pPeriodo As MonthYear)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * From LiquidacionEmpleado where mes=" + Str(pPeriodo.Month) + " and ayear=" + Str(pPeriodo.Year))
CargarRS rs
Exit Sub
errMaq:
ErrorLog "LiquidacionEmpleadoManager", "LoadByPeriodo ", err.Description
End Sub

Private Sub CargarRS(rs)

While Not rs.EOF
    Add(rs!Id, rs!Monto, rs!Adelanto, rs!Fecha, rs!Plus, rs!Observaciones, rs!Mes, rs!aYear).IdEmpleado = rs!IdEmpleado
    
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

Public Property Get Item(pId As Long) As LiquidacionEmpleado
On Error Resume Next
Set Item = mCol.Item("k" + Str(pId))
End Property




