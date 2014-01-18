VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObraSocialManager"
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
    Call ErrorLog("ObraSocialManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pObraSocial As ObraSocial)
    On Error GoTo AddItemErr

    mCol.Add pObraSocial, "k" + Str(pObraSocial.Id)
    
    Exit Sub
AddItemErr:
    ErrorLog "ObraSocialManager", "AddItem", Err.Description
End Sub

Friend Function Add(pId As Long, pCoseguro As Currency, pNombre As String, pObservaciones As String, pServicioEmergencia As ServicioEmergencia, idDireccion As Long, pInfoContable As InfoContableEmp) As ObraSocial
Dim os As New ObraSocial
'os.Afiliados lo cargo on demand
'os.CodigosCubiertos.CargarPorObraSocial pId
os.Coseguro = pCoseguro
os.Direccion.Cargar (idDireccion)
os.Id = pId
os.Nombre = pNombre
os.Observaciones = pObservaciones
Set os.ServicioEmergencia = pServicioEmergencia
'Set os.Telefonos = pTelefonos
Set os.InfoContable = pInfoContable
mCol.Add os, "k" + Str(os.Id)
Set Add = os
Set os = Nothing
End Function

Public Function Nuevo(pCodigosCubiertos As CodigoCubiertoManager, pCoseguro As Currency, pDireccion As Direccion, pNombre As String, pObservaciones As String, pServicioEmergencia As ServicioEmergencia, pTelefonos As TelefonoManager, pInfoContable As InfoContableEmp) As ObraSocial

Dim os As New ObraSocial
'os.Afiliados no tiene todavia
Set os.CodigosCubiertos = pCodigosCubiertos
os.Coseguro = pCoseguro
Set os.Direccion = pDireccion
'os.id se lo asigno en la funcion insert
os.Nombre = pNombre
os.Observaciones = pObservaciones
Set os.ServicioEmergencia = pServicioEmergencia
Set os.Telefonos = pTelefonos
Set os.InfoContable = pInfoContable
os.Insert

Me.AddItem os
Set Nuevo = os

End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
'implementar
On Error GoTo errMaq
Dim ic As InfoContableEmp
Dim rs
Set rs = BD.SelectTabla(tObraSocial)
While Not rs.EOF
    Set ic = New InfoContableEmp
    ic.Constructor rs!IVA, rs!Formula
    Add rs!Id, rs!Coseguro, rs!Nombre, "" & rs!Observaciones, ServiciosEmergenciaLocal.Item(rs!idServicioEmergencia), rs!idDireccion, ic
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "ObraSocialManager", "cargarTodos ", Err.Description
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

Public Property Get Item(pId As Long) As ObraSocial
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property
