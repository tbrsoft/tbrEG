VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SintomaManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A8167C00EA"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private mCol As Collection

'##ModelId=46A81690009C
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    '## Your code goes here ...

    Exit Sub
RemoveErr:
    Call ErrorLog("SintomaManager", "Remove Sub", err.Description)
End Sub

'##ModelId=46A81691038A
Public Sub AddItem(pSintoma As Sintoma)
    On Error GoTo AddItemErr

    mCol.Add pSintoma, "k" + Str(pSintoma.Id)

    Exit Sub
AddItemErr:
    ErrorLog "SintomaManager", "AddItem", err.Description
End Sub

Friend Function Add(pId As Long, pNombre As String, pParent As CodigoEmergencia)
Dim b As New Sintoma
b.Id = pId
b.Nombre = pNombre
Set b.Parent = pParent

mCol.Add b, "k" + Str(b.Id)
End Function

Public Function NuevoSintoma(pNombre As String, pParent As CodigoEmergencia) As Sintoma
'implementar, si hace falta!
End Function

Friend Sub CargarSintomas()
'muy probablemente esto deberia ir en la capa de datos...
On Error GoTo errMaq

Dim rs
Set rs = BD.SelectTabla(tSintoma)
While Not rs.EOF
        
    Add rs!Id, rs!Nombre, CodigoEmergenciaLocal.Item(rs!idCodigoEmergencia)
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
ErrorLog "sintomasManager", "cargarPorcodemergencia ", err.Description
End Sub

Public Function GetByCodEmergencia(pCodigo As CodigoEmergencia) As SintomaManager
On Error GoTo errMaq
    Dim sm As New SintomaManager
    Dim s As Sintoma
    For Each s In Me
        If s.Parent.Id = pCodigo.Id Then sm.AddItem s
    Next
    Set GetByCodEmergencia = sm
    Set sm = Nothing
Exit Function
errMaq:
ErrorLog "SintomasManager", "cargarPorcodemergencia ", err.Description
End Function

Public Property Get Item(pId As Long) As Sintoma
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
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

