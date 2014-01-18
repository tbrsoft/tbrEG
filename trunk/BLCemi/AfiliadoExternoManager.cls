VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AfiliadoExternoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event ItemAdded(pAfiliadoExterno As AfiliadoExterno)
Public Event ItemRemoved(pAfiliadoExterno As AfiliadoExterno)
Public Event ItemModified(pAfiliadoExterno As AfiliadoExterno)

Private mParent As Object 'puede ser OS, SE o AreaProt.

Private mCol As Collection

'##ModelId=46A80F8302AF
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("AfiliadoExternoManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A80F8501B5
Public Sub AddItem(pAfiliadoExterno As AfiliadoExterno)
    On Error GoTo AddItemErr

    mCol.Add pAfiliadoExterno, "k" + Str(pAfiliadoExterno.Id)

    Exit Sub
AddItemErr:
    ErrorLog "AfiliadoExternoManager", "AddItem", Err.Description
End Sub

Public Function Nuevo(pApellido As String, pDireccion As Direccion, pFechaNacimiento As Date, pNroAfiliado As String, pNombre As String, pNrodoc As Long, pObraSocial As ObraSocial, pObservaciones As String, pSexo As Integer, pTipoDoc As TipoDoc, pTopeAtenciones As Integer, pHistoriaClinica As HistoriaClinica, pTelefonos As TelefonoManager, pEsEdad As Boolean)
Dim a As New AfiliadoExterno

Set a.Parent = mParent

a.Apellido = pApellido
'a.Atenciones un AfiliadoExterno nuevo todavia no tiene atenciones
Set a.Direccion = pDireccion
a.FechaNacimiento = pFechaNacimiento
Set a.HistoriaClinica = pHistoriaClinica
a.NroAfiliado = pNroAfiliado
a.Nombre = pNombre
a.NroDoc = pNrodoc
Set a.ObraSocial = pObraSocial
a.Observaciones = pObservaciones
a.Sexo = pSexo
Set a.TipoDoc = pTipoDoc
a.TopeAtenciones = pTopeAtenciones
Set a.Telefonos = pTelefonos
a.EsEdad = pEsEdad
a.Insert mParent

AddItem a

Set Nuevo = a
End Function

Friend Function Add(pId As Long, pApellido As String, pNombre As String, pTipoDoc As TipoDoc, pNrodoc As Long, pIdDireccion As Long, pFechaNacimiento As Date, pObraSocial As ObraSocial, pObservaciones As String, pSexo As Integer, pTopeAtenciones As Integer, pNroAfiliado As String, pEsEdad As Boolean)
Dim a As New AfiliadoExterno

Set a.Parent = mParent

a.Apellido = pApellido
'a.Atenciones ver
a.Direccion.Cargar pIdDireccion
a.FechaNacimiento = pFechaNacimiento
'Set a.HistoriaClinica = pHistoriaClinica  on demand
a.Id = pId
a.Nombre = pNombre
a.NroAfiliado = pNroAfiliado
a.NroDoc = pNrodoc
Set a.ObraSocial = pObraSocial
a.Observaciones = pObservaciones
a.Sexo = pSexo
Set a.TipoDoc = pTipoDoc
a.TopeAtenciones = pTopeAtenciones
a.EsEdad = pEsEdad

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

Public Property Get Parent() As Object
    Set Parent = mParent
End Property

Public Property Get Item(pId As Long) As AfiliadoExterno
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Friend Sub CargarAfiliadoExternos(pParent As Object)

Set mParent = pParent
Dim pTabla As String
Dim campoID As String

If TypeOf mParent Is AreaProtegida Then
    pTabla = "AfiliadoExternoXAreaProtegida"
    campoID = "idAreaProtegida"
ElseIf TypeOf mParent Is ServicioEmergencia Then
    pTabla = "AfiliadoExternoXServicioEmergencia"
    campoID = "idServicioEmergencia"
ElseIf TypeOf mParent Is ObraSocial Then
    pTabla = "AfiliadoExternoXObraSocial"
    campoID = "idObraSocial"
End If

Dim rs

Dim mObraSocialAux As ObraSocial

'levanto solo los afiliados q hacen falta
Set rs = BD.SelectSQL("SELECT * FROM AfiliadoExterno WHERE id IN (select idAfiliadoExterno from " + pTabla + " where " + campoID + " = " + Str(mParent.Id) + ")")

While Not rs.EOF

    If Not IsNull(rs!idObraSocial) Then
        Set mObraSocialAux = ObrasSocialesLocal.Item(rs!idObraSocial)
    Else
        Set mObraSocialAux = Nothing
    End If
    Add rs!Id, rs!Apellido, rs!Nombre, TiposDocumentoLocal.Item(rs!TipoDoc), rs!NroDoc, rs!idDireccion, rs!FechaNacimiento, mObraSocialAux, "" & rs!Observaciones, rs!Sexo, rs!TopeAtenciones, rs!NroAfiliado, IIf(rs!EsEdad = 1, True, False)
    rs.MoveNext
Wend

Exit Sub
errMaq:
ErrorLog "AfiliadoExternosManager", "cargarAfiliadoExternos ", Err.Description

End Sub

Public Function LoadById(pId As Long) As AfiliadoExterno
'en este punto no se si es de una obra social o un serv emergencia...
    On Error GoTo errman
    Dim ae As AfiliadoExterno
    Dim rs
    Set rs = BD.SelectSQL("SELECT * FROM AfiliadoExterno WHERE id=" + Str(pId))
    
    'esta es la obra social, no es el parent
    Dim mObraSocialAux As ObraSocial
    If Not IsNull(rs!idObraSocial) Then
        Set mObraSocialAux = ObrasSocialesLocal.Item(rs!idObraSocial)
    Else
        Set mObraSocialAux = Nothing
    End If
    
    Set ae = Add(rs!Id, rs!Apellido, rs!Nombre, TiposDocumentoLocal.Item(rs!TipoDoc), rs!NroDoc, rs!idDireccion, rs!FechaNacimiento, mObraSocialAux, "" & rs!Observaciones, rs!Sexo, rs!TopeAtenciones, rs!NroAfiliado, IIf(rs!EsEdad = 1, True, False))
    Set rs = BD.SelectSQL("select idObraSocial from AfiliadoExternoXObraSocial where idAfiliadoExterno = " + Str(pId))
    
    If Not rs.EOF Then 'si existe un registro...
        Set ae.Parent = ObrasSocialesLocal.Item(rs!idObraSocial)
    Else
        Set rs = BD.SelectSQL("select idServicioEmergencia from AfiliadoExternoXServicioEmergencia where idAfiliadoExterno = " + Str(pId))
        If Not rs.EOF Then
            Set ae.Parent = ServiciosEmergenciaLocal.Item(rs!idServicioEmergencia)
        End If
    End If
    
    Set LoadById = ae
    Exit Function
errman:
    ErrorLog "AfiliadoExternoManager", "LoadById", Err.Description
End Function

'Friend Sub Guardar(trans As Transaccion, pParent As AfiliadoExterno)
'    Dim AF As AfiliadoExterno
'    For Each AF In Me
'        AF.InsertACargo trans, pParent
'    Next
'End Sub
'
'Friend Sub GuardarModificaciones(trans As Transaccion)
'    Dim AF As AfiliadoExterno
'    For Each AF In Me
'        AF.GuardarModificacionesACargo trans
'    Next
'End Sub

'Public Function ExisteId(pId As Long) As Boolean
'Dim rs
'Set rs = BD.SelectSQL("select id from AfiliadoExterno where id=" + Str(pId))
'ExisteId = Not (rs.EOF)
'End Function

