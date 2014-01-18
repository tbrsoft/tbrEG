VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnidadCuerpoManager"
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
    Call ErrorLog("UnidadCuerpoManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pUnidadCuerpo As UnidadCuerpo)
    On Error GoTo AddItemErr

    mCol.Add pUnidadCuerpo ', "k" + Str(pUnidadCuerpo.Id)

    Exit Sub
AddItemErr:
    ErrorLog "UnidadCuerpoManager", "AddItem", Err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String) As UnidadCuerpo
Dim c As New UnidadCuerpo
c.Id = pId
c.Nombre = pNombre
AddItem c
Set Add = c
Set c = Nothing
End Function

Friend Sub CargarPorCuerpo(pIdCuerpo As Long)
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectSQL("Select * from UnidadCuerpo where idCuerpo=" + Str(pIdCuerpo))
While Not rs.EOF
        
    Add rs!Id, rs!Nombre
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "UnidadCuerpoManager", "cargarPorCuerpo ", Err.Description
End Sub

'falta cargar uno solo

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

Public Property Get ItemById(pId As Long) As UnidadCuerpo
    On Error Resume Next
    Dim u As UnidadCuerpo
    For Each u In Me
        If u.Id = pId Then
            Set ItemById = u
            Exit For
        End If
    Next
End Property

Public Property Get ItemByName(pNombre As String) As UnidadCuerpo
Dim b As UnidadCuerpo
For Each b In Me
    If LCase(b.Nombre) = LCase(pNombre) Then
        Set ItemByName = b
        Exit For
    End If
Next
End Property

'Public Function Nuevo(pNombre As String) As UnidadCuerpo
'    Dim c As New UnidadCuerpo
'    c.Id = BD.getIdDisponible(tUnidadCuerpo)
'    c.Nombre = pNombre
'    c.Guardar
'    Me.AddItem c
'    Set Nuevo = c
'End Function

Public Sub Guardar(pIdCuerpo As Long)

    Dim uc As UnidadCuerpo
    For Each uc In Me
        uc.Guardar pIdCuerpo
    Next

End Sub

Public Function UnidadesToString() As String
    Dim c As UnidadCuerpo
    Dim aux As String
    For Each c In Me
        aux = aux + ", " + c.Nombre
    Next
    If aux <> "" Then
        UnidadesToString = Right(aux, Len(aux) - 2)
    Else
        UnidadesToString = ""
    End If
    
End Function
