VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AreaProtegidaManager"
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
    Call ErrorLog("AreaProtegidaManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pAreaProtegida As AreaProtegida)
    On Error GoTo AddItemErr

    mCol.Add pAreaProtegida, "k" + Str(pAreaProtegida.Id)

    Exit Sub
AddItemErr:
    ErrorLog "AreaProtegidaManager", "AddItem", err.Description
End Sub

Friend Function Add(pId As Long, pNombreArea As String, pNombreResp As String, pApellidoResp As String, pCobrador As Empleado, pIdDireccion As Long, pFechaInscripcion As Date, pInicioPrestacion As Date, pNroDocResp As Long, pObservaciones As String, pTipoDocResp As TipoDoc, pTopeAtenciones As Integer, pImporte As Currency) As AreaProtegida

    Dim ap As New AreaProtegida
    '.Afiliados on demand
    ap.ApellidoResp = pApellidoResp
    'Atenciones on dem
    Set ap.Cobrador = pCobrador
    ap.Direccion.Cargar pIdDireccion
    ap.FechaInscripcion = pFechaInscripcion
    ap.Id = pId
    ap.InicioPrestacion = pInicioPrestacion
    ap.NombreArea = pNombreArea
    ap.NombreResp = pNombreResp
    ap.NroDocResp = pNroDocResp
    ap.Observaciones = pObservaciones
    'pagos on demand
    '.Telefonos on demand
    Set ap.TipoDocResp = pTipoDocResp
    ap.TopeAtenciones = pTopeAtenciones
    ap.Importe = pImporte
    
    AddItem ap
    Set Add = ap
    Set ap = Nothing
    
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tAreaProtegida)
While Not rs.EOF
        
    Add rs!Id, rs!NombreArea, rs!NombreResp, rs!ApellidoResp, EmpleadosLocal.Item(rs!IdEmpleado), rs!idDireccion, rs!FechaInscripcion, rs!InicioPrestacion, rs!NroDocResp, rs!Observaciones, TiposDocumentoLocal.Item(rs!TipoDocResp), rs!TopeAtenciones, rs!Importe
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "AreaProtegidasManager", "cargarTodos ", err.Description
End Sub

Public Function Nuevo(pNombreArea As String, pNombreResp As String, pApellidoResp As String, pCobrador As Empleado, pDireccion As Direccion, pFechaInscripcion As Date, pInicioPrestacion As Date, pNroDocResp As Long, pObservaciones As String, pTipoDocResp As TipoDoc, pTopeAtenciones As Integer, pTelefonos As TelefonoManager, pImporte As Currency) As AreaProtegida
Dim ap As New AreaProtegida
With ap
    '.Afiliados se van agregando a medida q se van realizando las atenciones
    .ApellidoResp = pApellidoResp
    'todavia no hay Atenciones
    Set .Cobrador = pCobrador
    Set .Direccion = pDireccion
    .FechaInscripcion = pFechaInscripcion
    'id se lo asigno en guardar
    .InicioPrestacion = pInicioPrestacion
    .NombreArea = pNombreArea
    .NombreResp = pNombreResp
    .NroDocResp = pNroDocResp
    .Observaciones = pObservaciones
    'todavia no hay Pagos
    Set .Telefonos = pTelefonos
    Set .TipoDocResp = pTipoDocResp
    .TopeAtenciones = pTopeAtenciones
    .Importe = pImporte
    .Guardar
    
End With

AddItem ap
Set Nuevo = ap

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

Public Property Get Item(pId As Long) As AreaProtegida
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property
