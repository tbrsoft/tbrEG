VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Archivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function LeerArchivo(path As String) As String
    Dim fso 'As FileSystemObject
    Dim f
        
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If path = "" Then
        LeerArchivo = ""
    Else
        Set f = fso.GetFile(path)
    End If
    Set ts = f.OpenAsTextStream(1)
    
    s = ts.ReadAll
    
    ts.Close
    LeerArchivo = s
End Function

Public Sub EscribirArchivo(path As String, cadena As String)
   ' On Error GoTo e
    Dim fso As FileSystemObject
    Dim f, ts
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(path) Then
       fso.CreateTextFile path
    End If
    
    Set f = fso.GetFile(path)
    Set ts = f.OpenAsTextStream(2, 0)
     
    ts.Write cadena
    ts.Close
'    Exit Sub
'e:
    
End Sub
