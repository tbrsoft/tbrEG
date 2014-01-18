VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodigoEmergenciaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

'Private mCodigosOriginal As Collection

'##ModelId=46A8DFD60242
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("CodigoEmergenciaManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A8DFD8000F
Public Sub AddItem(pCodigoEmergencia As CodigoEmergencia)
    On Error GoTo AddItemErr

    mCol.Add pCodigoEmergencia, "k" + Str(pCodigoEmergencia.Id)

    Exit Sub
AddItemErr:
    ErrorLog "CodigoEmergenciaManager", "AddItem", err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String, pVencimiento As Integer, pBold As Boolean, pColorFuente As Long) As CodigoEmergencia
Dim c As New CodigoEmergencia
c.Id = pId
c.Nombre = pNombre
c.Vencimiento = pVencimiento
c.Bold = pBold
c.ColorFuente = pColorFuente

mCol.Add c, "k" + Str(c.Id)
Set Add = c
Set c = Nothing
End Function

Friend Sub CargarTodos()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tCodigoEmergencia)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre, IIf(IsNull(rs!Vencimiento), 0, rs!Vencimiento), IIf(rs!Bold = "True", True, False), rs!ColorFuente
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "CodigoEmergenciaManager", "cargarTodos ", err.Description
End Sub

Public Property Get Item(pId As Long) As CodigoEmergencia
Attribute Item.VB_UserMemId = 0
    On Error Resume Next
    Set Item = mCol("k" + Str(pId))
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

Public Function Exists(pId As Long) As Boolean
    Dim t As CodigoEmergencia
    Set t = Me.Item(pId)
    Exists = (Not t Is Nothing)
End Function

