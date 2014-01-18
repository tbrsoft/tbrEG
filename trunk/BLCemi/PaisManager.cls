VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PaisManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

'##ModelId=46A9178B006D
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("PaisManager", "Remove Sub", Err.Description)
End Sub

''##ModelId=46A91790006D
'Public Sub AddItem()
'    On Error GoTo AddItemErr
'
'    '## Your code goes here ...
'
'    Exit Sub
'AddItemErr:
'    Call ErrorLog(MyUnhandledError, "AddItem Sub")
'End Sub

Friend Function Add(pId As Long, pNombre As String, pPrimerOrden As String, pSegundoOrden As String, pTercerOrden As String) As Pais
    Dim p As New Pais
    p.Id = pId
    p.Nombre = pNombre
    p.PrimerOrden = pPrimerOrden
    p.SegundoOrden = pSegundoOrden
    p.TercerOrden = pTercerOrden
    mCol.Add p, "k" + Str(p.Id)
    Set Add = p
    Set p = Nothing
End Function

Friend Sub CargarTodos()
    'muy probablemente esto deberia ir en la capa de datos...
    'TODO prestar atencion a los comentarios
    On Error GoTo errMaq
    
    Dim rs
    Set rs = BD.SelectTabla(tPais)
    While Not rs.EOF
            
        Add rs!Id, rs!Nombre, rs!PrimerOrden, rs!SegundoOrden, rs!TercerOrden
        rs.MoveNext
        
    Wend
    Exit Sub
errMaq:
    ErrorLog "PaisManager", "cargarTodos ", Err.Description
End Sub

Public Property Get Item(pId As Long) As Pais
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pNombre As String) As Pais
Dim b As Pais
For Each b In Me
    If LCase(b.Nombre) = LCase(pNombre) Then
        Set ItemByName = b
        Exit For
    End If
Next
End Property

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
