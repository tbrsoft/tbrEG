VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InvolucradoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    Exit Sub
RemoveErr:
    Call ErrorLog("InvolucradoManager", "Remove Sub", err.Description)
End Sub

Public Sub AddItem(pInvolucrado As Involucrado)
    On Error GoTo AddItemErr

    If pInvolucrado.Id <> 0 Then
        mCol.Add pInvolucrado, "k" + Str(pInvolucrado.Id)
    Else
        mCol.Add pInvolucrado
    End If
    
    Exit Sub
AddItemErr:
    ErrorLog "InvolucradoManager", "AddItem", err.Description
End Sub

Public Function Nuevo(pApellido As String, pDireccion As Direccion, pFechaNacimiento As Date, pNombre As String, pNrodoc As Long, pObservaciones As String, pSexo As Integer, pTipoDoc As TipoDoc, pEsEdad As Boolean, pTipoRelacion As eRelacion, pDescripcionRelacion As String, pTelefono As String) As Involucrado
Dim a As New Involucrado

a.Apellido = pApellido
Set a.Direccion = pDireccion
a.FechaNacimiento = pFechaNacimiento
a.Nombre = pNombre
a.NroDoc = pNrodoc
a.Observaciones = pObservaciones
a.Sexo = pSexo
Set a.TipoDoc = pTipoDoc
a.EsEdad = pEsEdad
a.DescripcionRelacion = pDescripcionRelacion
a.TipoRelacion = pTipoRelacion
a.Telefono = pTelefono
'a.Insert

AddItem a

Set Nuevo = a
End Function

Friend Function Add(pId As Long, pApellido As String, pNombre As String, pTipoDoc As TipoDoc, pNrodoc As Long, pIdDireccion As Long, pFechaNacimiento As Date, pObservaciones As String, pSexo As Integer, pEsEdad As Boolean, pTipoRelacion As eRelacion, pDescripcionRelacion As String, pTelefono As String) As Involucrado
Dim a As New Involucrado

a.Apellido = pApellido
a.Direccion.Cargar pIdDireccion
a.FechaNacimiento = pFechaNacimiento
a.Id = pId
a.Nombre = pNombre
a.NroDoc = pNrodoc
a.Observaciones = pObservaciones
a.Sexo = pSexo
Set a.TipoDoc = pTipoDoc
a.EsEdad = pEsEdad
a.DescripcionRelacion = pDescripcionRelacion
a.TipoRelacion = pTipoRelacion
a.Telefono = pTelefono

AddItem a
Set Add = a
End Function

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

Public Property Get Item(pId As Long) As Involucrado
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByIndex(pIndex As Long) As Involucrado
    On Error Resume Next
    Set ItemByIndex = mCol.Item(pIndex)
End Property

Public Function GetByTipo(pTipoRelacion As eRelacion) As InvolucradoManager
    
    Dim Involucrados As New InvolucradoManager
    Dim inv As Involucrado
    For Each inv In Me
        If inv.TipoRelacion = pTipoRelacion Then Involucrados.AddItem inv
    Next
    Set GetByTipo = Involucrados
    
End Function

Friend Sub CargarInvolucrados(pAtencionB As AtencionB)

Dim rs

'levanto solo los involucrados q hacen falta
Set rs = BD.SelectSQL("SELECT * FROM Involucrado WHERE idSiniestro=" + Str(pAtencionB.Id))

While Not rs.EOF
   
    Add rs!Id, rs!Apellido, rs!Nombre, TiposDocumentoLocal.Item(rs!TipoDoc), rs!NroDoc, rs!idDireccion, rs!FechaNacimiento, "" & rs!Observaciones, rs!Sexo, IIf(rs!EsEdad = 1, True, False), IIf(rs!idTipoRelacion = 1, eRelacion.eAfectado, eRelacion.eSolicitante), "" & rs!DescripcionRelacion, rs!Telefono
    rs.MoveNext
Wend

Exit Sub
errMaq:
ErrorLog "InvolucradosManager", "cargarInvolucrados ", err.Description

End Sub

Public Function LoadById(pId As Long) As Involucrado
    On Error GoTo errman
    Dim ae As Involucrado
    Dim rs
    Set rs = BD.SelectSQL("SELECT * FROM Involucrado WHERE id=" + Str(pId))
           
    Set ae = Add(rs!Id, rs!Apellido, rs!Nombre, TiposDocumentoLocal.Item(rs!TipoDoc), rs!NroDoc, rs!idDireccion, rs!FechaNacimiento, "" & rs!Observaciones, rs!Sexo, IIf(rs!EsEdad = 1, True, False), IIf(rs!idTipoRelacion = 1, eRelacion.eAfectado, eRelacion.eSolicitante), "" & rs!DescripcionRelacion, rs!Telefono)
        
    Set LoadById = ae
    Exit Function
errman:
    ErrorLog "InvolucradoManager", "LoadById", err.Description
End Function

Friend Sub Save(trans As Transaccion, pSiniestro As AtencionB)
    Dim mInv As Involucrado
    For Each mInv In Me
        'internamente sabe si es un insert o un update
        mInv.Save trans, pSiniestro.Id
    Next
End Sub
