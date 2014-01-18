VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Transaccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'una transaccion tiene una coleccion de comandos
Private mCol As Collection
Private parent As BaseDeDatos

Public Sub AddInsertCommand(pTabla As eTablas, pParametros As ParametrosManager)
    Dim auxNombres As String
    Dim auxValues As String
    Dim sql As String
   
    Dim cmd As New Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = parent.Conexion
    
    Dim p As Parametro
    Dim par As Parameter
        
    For Each p In pParametros
        auxNombres = auxNombres + ", " + p.Nombre
        auxValues = auxValues + ", @" + p.Nombre
    Next
       
    sql = "INSERT INTO " + GetTabla(pTabla) + " (" + Right(auxNombres, Len(auxNombres) - 2) + ") VALUES (" + Right(auxValues, Len(auxValues) - 2) + ")"
    cmd.CommandText = sql
    
    For Each p In pParametros
        Set par = cmd.CreateParameter("@" + p.Nombre, , , , p.Valor)
        cmd.Parameters("@" + p.Nombre).Value = p.Valor
    Next
    
    mCol.Add cmd
End Sub

Public Sub AddUpdateCommand(pTabla As eTablas, pParametros As ParametrosManager)
    Dim auxNombres As String
    Dim auxWhere As String
    Dim sql As String
   
    Dim cmd As New Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = parent.Conexion
    
    Dim p As Parametro
    Dim par As Parameter
        
    For Each p In pParametros
        If p.EsPK Then
            'si nunca pasa por aca es un error
            auxWhere = auxWhere + " AND " + p.Nombre + "= @" + p.Nombre
        Else
            auxNombres = auxNombres + ", " + p.Nombre + "= @" + p.Nombre
        End If
    Next
       
    sql = "UPDATE " + GetTabla(pTabla) + " SET " + Right(auxNombres, Len(auxNombres) - 2) + " WHERE " + Right(auxWhere, Len(auxWhere) - 5)
    cmd.CommandText = sql
    
    For Each p In pParametros
        Set par = cmd.CreateParameter("@" + p.Nombre, , , , p.Valor)
        cmd.Parameters("@" + p.Nombre).Value = p.Valor
    Next
    
    mCol.Add cmd
End Sub

'PROBAR, NO SE SI FUNCIONA!!!
Public Sub AddDeleteCommand(pTabla As eTablas, pParametros As ParametrosManager)
    Dim auxWhere As String
    Dim sql As String
   
    Dim cmd As New Command
    
    cmd.CommandType = adCmdText
    Set cmd.ActiveConnection = parent.Conexion
    
    Dim p As Parametro
    Dim par As Parameter
        
    For Each p In pParametros
        If p.EsPK Then
            'si nunca pasa por aca es un error
            auxWhere = auxWhere + " AND " + p.Nombre + "= @" + p.Nombre
        End If
    Next
       
    sql = "DELETE FROM " + GetTabla(pTabla) + " WHERE " + Right(auxWhere, Len(auxWhere) - 4)
    cmd.CommandText = sql
    
    For Each p In pParametros
        Set par = cmd.CreateParameter("@" + p.Nombre, , , , p.Valor)
        cmd.Parameters("@" + p.Nombre).Value = p.Valor
    Next
    
    mCol.Add cmd
End Sub

Public Function Execute() As Boolean
On Error GoTo e

parent.Conexion.BeginTrans
WriteToLog App.path + "\trans.log", Str(Now) + " Se inicio una transaccion"
Dim cmd As Command

For Each cmd In mCol
    'imprimirComando cmd
    WriteToLog App.path + "\trans.log", cmd.CommandText
    cmd.Execute
    
Next
'si fue todo bien...
parent.Conexion.CommitTrans
WriteToLog App.path + "\trans.log", Str(Now) + " Se termino la transaccion"

'despues de ejecutar la transaccion, la reinicio para q quede lista para otra cosa

Set mCol = Nothing
Set mCol = New Collection

Execute = True
Exit Function
e:
'si hay un error cancelo la transaccion
parent.Conexion.RollbackTrans

parent.DBEventsManager.FireErrorEvent eUnknown, "Transaccion: " + Err.Description

'si fallo tambien la reinicio para q quede lista para otra cosa

Set mCol = Nothing
Set mCol = New Collection

Execute = False
End Function

Friend Sub Inicializar(pParent As BaseDeDatos)
    Set parent = pParent
End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Private Sub imprimirComando(cmd As Command)
    On Error Resume Next
    aux = cmd.CommandText
    Dim p As Parameter
    For Each p In cmd.Parameters
        aux2 = aux2 + p.Name + " = " + p.Value + ", "
    Next
    Debug.Print aux + " - " + aux2
End Sub

Private Sub WriteToLog(path As String, contents As String)
   ' On Error GoTo e
    Dim fso ' As FileSystemObject
    Dim f, ts
    'Dim ts As TextStream
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(path) Then
       fso.CreateTextFile path
    End If
   
    Set f = fso.GetFile(path)
    Set ts = f.OpenAsTextStream(8, 0) '8=ForAppending
    
    ts.WriteLine contents
    'ts.write cadena
    ts.Close
'    Exit Sub
'e:
    
End Sub
