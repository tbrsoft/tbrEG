VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServicioEmergenciaManager"
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
    Call ErrorLog("ServicioEmergenciaManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pServicioEmergencia As ServicioEmergencia)
    On Error GoTo AddItemErr

    mCol.Add pServicioEmergencia, "k" + Str(pServicioEmergencia.Id)

    Exit Sub
AddItemErr:
    ErrorLog "ServicioEmergenciaManager", "AddItem", Err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String, pIdDireccion As Long, pInfoContable As InfoContableEmp) As ServicioEmergencia

Dim c As New ServicioEmergencia
c.Id = pId
c.Nombre = pNombre
c.Direccion.Cargar pIdDireccion
Set c.InfoContable = pInfoContable

mCol.Add c, "k" + Str(c.Id)
Set Add = c
Set c = Nothing
End Function

Public Function Nuevo(pNombre As String, pDireccion As Direccion, pTelefonos As TelefonoManager, pCodigosCubiertos As CodigoCubiertoManager, pInfoContable As InfoContableEmp) As ServicioEmergencia

Dim c As New ServicioEmergencia
'c.id = pId se lo asigno dentro de guardar
c.Nombre = pNombre
Set c.Direccion = pDireccion
Set c.Telefonos = pTelefonos
Set c.CodigosCubiertos = pCodigosCubiertos
Set c.InfoContable = pInfoContable
c.Guardar

AddItem c
Set c = Nothing
End Function


Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim ic As InfoContableEmp
Dim rs
Set rs = BD.SelectTabla(tServicioEmergencia)
While Not rs.EOF
    Set ic = New InfoContableEmp
    ic.Constructor rs!IVA, rs!Formula
    Add rs!Id, rs!Nombre, rs!idDireccion, ic
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "ServicioEmergenciaManager", "cargarTodos ", Err.Description
End Sub

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

Public Property Get Item(pId As Long) As ServicioEmergencia
'devuelve nothing si no lo encuentra
On Error Resume Next
Set Item = mCol.Item("k" + Str(pId))
End Property

