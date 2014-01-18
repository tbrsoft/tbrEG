VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Log"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public Function LeerArchivo(path As String) As String
'    Dim fso 'As FileSystemObject
'    Dim f
'
'    Dim s As String
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    If path = "" Then
'        LeerArchivo = ""
'    Else
'        Set f = fso.GetFile(path)
'    End If
'    Set ts = f.OpenAsTextStream(1)
'
'    s = ts.ReadAll
'
'    ts.Close
'    LeerArchivo = s
'End Function

Public Sub WriteToLog(Path As String, contents As String)
    
    On Error GoTo E44
    
    Dim FSO As New Scripting.FileSystemObject
    Dim TS As TextStream
       
    Set TS = FSO.OpenTextFile(Path, ForAppending, True)
        TS.WriteLine contents
    TS.Close
    Exit Sub
    
E44:
    MsgBox "Error 60112!"
    
End Sub
