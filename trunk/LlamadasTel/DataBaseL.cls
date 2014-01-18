VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataBaseL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum eTables
    etTelephoneCall = 1
    etUser = 2
    etContact = 3
    etTelephone = 4
End Enum

Dim cn As Connection
Private mDBPath As String

'Friend Property Get Conexion() As Connection
'    Set Conexion = cn
'End Property

Private Sub Class_Initialize()
    mDBPath = App.Path + "\database.mdb"
    Conectar
End Sub

Public Sub Conectar()
    
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If
    Set cn = Nothing
    Set cn = New Connection
    Dim cs As String
    cs = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + mDBPath + ";Persist Security Info=False"
    cn.ConnectionString = cs
    On Error GoTo errman
    cn.Open
    Exit Sub
errman:
   'hacer algo
End Sub

Public Function SelectSQL(sql As String) As Recordset
    On Error GoTo e:
    Dim rs As Recordset
    Set rs = cn.Execute(sql)
    Set SelectSQL = rs
    Exit Function
e:
    Debug.Print Err.Description
    
End Function

Public Function SelectTable(pTable As eTables) As Recordset
    On Error GoTo e:
    Dim rs As Recordset
    Set rs = cn.Execute("SELECT * FROM " + GetTable(pTable))
    Set SelectTable = rs
    Exit Function
e:
    Debug.Print Err.Description
    
End Function

'Public Function Insertar(pTabla As eTablas, pParametros As ParametrosManager) As Boolean
'    Dim auxNombres As String
'    Dim auxValues As String
'    Dim sql As String
'
'    Dim cmd As Command
'    Set cmd = New Command
'
'    cmd.CommandType = adCmdText
'    Set cmd.ActiveConnection = cn
'
'    Dim p As Parametro
'    Dim par As Parameter
'
'    For Each p In pParametros
'        auxNombres = auxNombres + ", " + p.Nombre
'        auxValues = auxValues + ", @" + p.Nombre
'    Next
'
'    sql = "INSERT INTO " + GetTabla(pTabla) + " (" + Right(auxNombres, Len(auxNombres) - 2) + ") VALUES (" + Right(auxValues, Len(auxValues) - 2) + ")"
'    cmd.CommandText = sql
'
'    For Each p In pParametros
'        Set par = cmd.CreateParameter("@" + p.Nombre, , , , p.Valor)
'        cmd.Parameters("@" + p.Nombre).Value = p.Valor
'    Next
'
'    On Error GoTo e:
'    cmd.Execute
'    Insertar = True
'    Exit Function
'e:
'    Insertar = False
'End Function
'
'Public Function Update(pTabla As eTablas, pParametros As ParametrosManager) As Boolean
'    Dim auxNombres As String
'    Dim auxWhere As String
'    Dim sql As String
'
'    Dim cmd As New Command
'
'    cmd.CommandType = adCmdText
'    Set cmd.ActiveConnection = cn
'
'    Dim p As Parametro
'    Dim par As Parameter
'
'    For Each p In pParametros
'        If p.EsPK Then
'            auxWhere = auxWhere + " AND " + p.Nombre + "= @" + p.Nombre
'        Else
'            auxNombres = auxNombres + ", " + p.Nombre + "= @" + p.Nombre
'        End If
'    Next
'
'    sql = "UPDATE " + GetTabla(pTabla) + " SET " + Right(auxNombres, Len(auxNombres) - 2) + " WHERE " + Right(auxWhere, Len(auxWhere) - 5)
'    cmd.CommandText = sql
'
'    For Each p In pParametros
'        Set par = cmd.CreateParameter("@" + p.Nombre, , , , p.Valor)
'        cmd.Parameters("@" + p.Nombre).Value = p.Valor
'    Next
'    On Error GoTo e
'    cmd.Execute
'    Update = True
'    Exit Function
'e:
'    Update = False
'End Function

Public Function GetAvailableId(pTable As eTables) As Long
'utilizo el sistema de una tabla especial para guardar el ultimo id utilizado
'soporta concurrencia...

Dim rs As Recordset
cn.BeginTrans
cn.Execute "Update ID" + GetTable(pTable) + " set id=id+1"
Set rs = cn.Execute("select id from ID" + GetTable(pTable))
cn.CommitTrans
GetAvailableId = rs!Id

End Function

Private Function GetTable(tTable As eTables) As String
   GetTable = Choose(tTable, "TelephoneCall", "Users", "Contact", "Telephone")
End Function

