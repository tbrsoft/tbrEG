VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CiudadManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A9175C0242"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private mCol As Collection

'##ModelId=46A91787009C
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("CiudadManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A9178D008C
Public Sub AddItem(pCiudad As Ciudad)
    On Error GoTo AddItemErr

    mCol.Add pCiudad, "k" + Str(pCiudad.Id)

    Exit Sub
AddItemErr:
    ErrorLog "CiudadManager", "AddItem", Err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String, pParent As Provincia)
Dim c As New Ciudad
c.Id = pId
c.Nombre = pNombre
Set c.Parent = pParent

mCol.Add c, "k" + Str(c.Id)
End Function

Public Function NuevaCiudad(pNombre As String, pParent As Provincia) As Ciudad
Dim c As New Ciudad
c.Id = BD.getIdDisponible(tCiudad)
c.Nombre = pNombre
Set c.Parent = pParent
c.Guardar
'lo tengo q agregar a esta coleccion y a ciudadeslocal tambien, por como estan echas las cosas,
'sino, tengo q refrescar desde afuera, cosa q no recomendamos en absoluto

CiudadesLocal.AddItem c
mCol.Add c, "k" + Str(c.Id)
Set NuevaCiudad = c
End Function

Friend Function GetByProvincia(pProvincia As Provincia) As CiudadManager
    Dim c As Ciudad
    Dim thisCiudades As New CiudadManager
    For Each c In Me
        If c.Parent.Id = pProvincia.Id Then thisCiudades.AddItem c
    Next
    Set GetByProvincia = thisCiudades
End Function

Friend Sub CargarTodos()
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tCiudad)
While Not rs.EOF

    Add rs!Id, rs!Nombre, ProvinciasLocal.Item(rs!idprovincia)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "CiudadManager", "CargarTodos ", Err.Description

End Sub

Public Property Get Item(pId As Long) As Ciudad
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByName(pNombre As String) As Ciudad
    TERR.Anotar "abau3-a", pNombre
    Dim b As Ciudad
    For Each b In Me
        TERR.Anotar "abau3-b", b.Nombre
        If LCase(b.Nombre) = LCase(pNombre) Then
            Set ItemByName = b
            Exit For
        End If
    Next
    TERR.Anotar "abau3-c"
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

