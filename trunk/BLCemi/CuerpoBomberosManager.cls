VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuerpoBomberosManager"
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
    Call ErrorLog("CuerpoBomberosManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pCuerpoBomberos As CuerpoBomberos)
    On Error GoTo AddItemErr

    mCol.Add pCuerpoBomberos, "k" + Str(pCuerpoBomberos.Id)

    Exit Sub
AddItemErr:
    ErrorLog "CuerpoBomberosManager", "AddItem", err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String) As CuerpoBomberos
Dim c As New CuerpoBomberos
c.Id = pId
c.Nombre = pNombre
AddItem c
Set Add = c
Set c = Nothing
End Function

Friend Sub CargarTodos()
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tCuerpoBomberos)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "CuerpoBomberosManager", "cargarTodos ", err.Description
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

Public Property Get Item(pId As Long) As CuerpoBomberos
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

'Public Property Get ItemByName(pNombre As String) As CuerpoBomberos
'Dim b As CuerpoBomberos
'For Each b In Me
'    If LCase(b.Nombre) = LCase(pNombre) Then
'        Set ItemByName = b
'        Exit For
'    End If
'Next
'End Property
'
Public Function Nuevo(pNombre As String, pResponsables As ResponsableCuerpoManager, pUnidades As UnidadCuerpoManager) As CuerpoBomberos
    Dim c As New CuerpoBomberos
    c.Id = BD.getIdDisponible(tCuerpoBomberos)
    c.Nombre = pNombre
    Set c.Responsables = pResponsables
    Set c.Unidades = pUnidades
    c.Guardar
    Me.AddItem c
    Set Nuevo = c
End Function

