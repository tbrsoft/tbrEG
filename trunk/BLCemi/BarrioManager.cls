VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarrioManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A9176C03C8"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private mCol As Collection

'##ModelId=46A9178901E4
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("BarrioManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A9178E0222
Public Sub AddItem(pBarrio As Barrio)
    On Error GoTo AddItemErr

    mCol.Add pBarrio, "k" + Str(pBarrio.Id)

    Exit Sub
AddItemErr:
    ErrorLog "BarrioManager", "AddItem", Err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String, pParent As Ciudad)
Dim b As New Barrio
b.Id = pId
b.Nombre = pNombre
Set b.Parent = pParent

mCol.Add b, "k" + Str(b.Id)
End Function

Public Function NuevoBarrio(pNombre As String, pParent As Ciudad) As Barrio
Dim b As New Barrio
b.Id = BD.getIdDisponible(tBarrio)
b.Nombre = pNombre
Set b.Parent = pParent
b.Guardar
'lo tengo q agregar a esta coleccion y a barrioslocal tambien, por como estan echas las cosas,
'sino, tengo q refrescar desde afuera, cosa q no recomendamos en absoluto

BarriosLocal.AddItem b
mCol.Add b, "k" + Str(b.Id)
Set NuevoBarrio = b
End Function

Friend Function GetByCiudad(pCiudad As Ciudad)
Dim b As Barrio
Dim thisBarrios As New BarrioManager
For Each b In Me
    If b.Parent.Id = pCiudad.Id Then thisBarrios.AddItem b
Next
Set GetByCiudad = thisBarrios
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tBarrio)
While Not rs.EOF

    Add rs!Id, rs!Nombre, CiudadesLocal.Item(rs!idCiudad)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "BarrioManager", "cargarPorBarrio ", Err.Description
End Sub

Public Property Get Item(pId As Long) As Barrio
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pNombre As String) As Barrio
Dim b As Barrio
For Each b In Me
    'ver como hacer menos restrictivo, o sea, q si uno esta en mayusc y el otro no, lo devuelva igual
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

