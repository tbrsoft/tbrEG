VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProvinciaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A9177600BB"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private mCol As Collection

'##ModelId=46A9178B006D
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("ProvinciaManager", "Remove Sub", Err.Description)
End Sub

Public Sub AddItem(pProvincia As Provincia)
    On Error GoTo AddItemErr

    mCol.Add pProvincia, "k" + Str(pProvincia.Id)

    Exit Sub
AddItemErr:
    ErrorLog "ProvinciaManager", "AddItem", Err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String, pParent As Pais)
Dim c As New Provincia
c.Id = pId
c.Nombre = pNombre
Set c.Parent = pParent

mCol.Add c, "k" + Str(c.Id)
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tProvincia)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre, PaisesLocal.Item(rs!idPais)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "ProvinciasManager", "cargarTodos ", Err.Description
End Sub

Public Property Get Item(pId As Long) As Provincia
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pNombre As String) As Provincia
Dim b As Provincia
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

Public Function NuevaProvincia(pNombre As String, pParent As Pais) As Provincia
Dim c As New Provincia
c.Id = BD.getIdDisponible(tProvincia)
c.Nombre = pNombre
Set c.Parent = pParent
c.Guardar
'lo tengo q agregar a esta coleccion y a ciudadeslocal tambien, por como estan echas las cosas,
'sino, tengo q refrescar desde afuera, cosa q no recomendamos en absoluto

ProvinciasLocal.AddItem c
mCol.Add c, "k" + Str(c.Id)
Set NuevaProvincia = c
End Function

Friend Function GetByPais(pPais As Pais) As ProvinciaManager
Dim c As Provincia
Dim thisProvincias As New ProvinciaManager
For Each c In Me
    If c.Parent.Id = pPais.Id Then thisProvincias.AddItem c
Next
Set GetByPais = thisProvincias
End Function


