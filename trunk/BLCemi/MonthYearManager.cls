VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MonthYearManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Friend Function Add(pMes As Integer, pYear As Integer) As MonthYear
Dim l As New MonthYear
l.Month = pMes
l.Year = pYear

mCol.Add l
Set Add = l
Set l = Nothing
End Function

Public Sub LoadFromLiqEmpleados()
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("select distinct mes,ayear from LiquidacionEmpleado")
While Not rs.EOF
        
    Add rs!Mes, rs!aYear
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "MonthYearManager", "LoadFromLiqEmpleados ", err.Description
End Sub

Public Sub LoadFromLiqEmpresa()
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("select distinct mes,ayear from LiquidacionEmpresa")
While Not rs.EOF
        
    Add rs!Mes, rs!aYear
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "MonthYearManager", "LoadFromLiqEmpresa", err.Description
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


