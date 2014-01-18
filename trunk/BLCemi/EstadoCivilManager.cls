VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EstadoCivilManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mCol As Collection

'Public Function getEstadoCivil(pId As Integer) As EstadoCivil
'Set getEstadoCivil = mCol.Item("k" + Str(pId))
'End Function

Friend Sub cargarEstadoCivil()

    Add 1, "Soltero"
    Add 2, "Casado"
    Add 3, "Divorciado"

End Sub

Friend Function Add(pId As Integer, pNombre As String) As EstadoCivil
    'crear un nuevo objeto
    Dim objNewMember As New EstadoCivil
    
    'establecer las propiedades que se transfieren al método
    objNewMember.Nombre = pNombre
    objNewMember.id = pId
        
    mCol.Add objNewMember, "k" + Str(pId)
    
    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(pId As Long) As EstadoCivil
    Set Item = mCol("k" + Str(pId))
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub


