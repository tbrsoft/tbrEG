VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TipoDocManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TipoDoc"
Attribute VB_Ext_KEY = "Member0" ,"TipoDoc"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"4491C02200CB"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Private mCol As Collection
Attribute mCol.VB_VarDescription = "variable local para contener colección"

'Public Function getTipoDoc(pId As Integer) As TipoDoc
'Set getTipoDoc = mCol.Item("k" + Str(pId))
'End Function

Friend Sub cargarTipoDoc()

    Add 1, "DNI"
    Add 2, "LC"
    Add 3, "LE"

End Sub

Friend Function Add(pId As Integer, pNombre As String) As TipoDoc
    'crear un nuevo objeto
    Dim objNewMember As New TipoDoc
    
    'establecer las propiedades que se transfieren al método
    objNewMember.Nombre = pNombre
    objNewMember.id = pId
        
    mCol.Add objNewMember, "k" + Str(pId)
    
    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(pId As Long) As TipoDoc
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

