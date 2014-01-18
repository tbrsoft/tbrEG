VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiqEmpresaManager"
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
    Call ErrorLog("LiqEmpresaManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pLiquidacion As LiquidacionEmpresa)
    On Error GoTo AddItemErr

    mCol.Add pLiquidacion, "k" + Str(pLiquidacion.Id)
    
    Exit Sub
AddItemErr:
    ErrorLog "LiqEmpresaManager", "AddItem", err.Description
End Sub

Friend Function Add(pId As Long, pFecha As Date, pTotal As Currency, pTipoEmpresa As eDestinoCodigo, pEmitida As Boolean, pMes As Integer, pYear As Integer) As LiquidacionEmpresa
Dim c As New LiquidacionEmpresa
c.Id = pId
'c.Detalle
c.Emitida = pEmitida
c.Fecha = pFecha
c.TipoEmpresa = pTipoEmpresa
c.Total = pTotal
c.Mes = pMes
c.Year = pYear

AddItem c
Set Add = c
Set c = Nothing
End Function

Public Sub Load()
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tLiquidacionEmpresa)
CargarRS rs

Exit Sub
errMaq:
ErrorLog "LiqEmpresaManager", "Load", err.Description
End Sub

Public Sub LoadByPeriodo(pPeriodo As MonthYear)
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * From LiquidacionEmpresa where mes=" + Str(pPeriodo.Month) + " and ayear=" + Str(pPeriodo.Year))
CargarRS rs
Exit Sub
errMaq:
ErrorLog "LiquidacionEmpresaManager", "LoadByPeriodo ", err.Description
End Sub

Private Sub CargarRS(rs)
    While Not rs.EOF
    
        Add(rs!Id, rs!Fecha, rs!Total, IIf(rs!TipoEmpresa = 1, eDCObraSocial, eDCServicioEmergencia), rs!Emitido, rs!Mes, rs!aYear).idEmpresa = rs!idEmpresa
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

Public Function Nuevo(pDetalle As InfoContableManager, pEmitida As Boolean, pFecha As Date, pTipoEmpresa As eDestinoCodigo, pIdEmpresa As Long, pMes As Integer, pYear As Integer) As LiquidacionEmpresa
    Dim c As New LiquidacionEmpresa
    c.Id = BD.getIdDisponible(tLiquidacionEmpresa)
    Set c.Detalle = pDetalle
    c.Emitida = pEmitida
    c.Fecha = pFecha
    c.idEmpresa = pIdEmpresa
    c.TipoEmpresa = pTipoEmpresa
    c.Mes = pMes
    c.Year = pYear
    c.Total = pDetalle.GetSaldo
    c.Save
    Me.AddItem c
    Set Nuevo = c
End Function


