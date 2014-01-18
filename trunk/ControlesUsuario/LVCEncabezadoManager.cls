VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LVCEncabezadoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"LVCEncabezado"
Attribute VB_Ext_KEY = "Member0" ,"LVCEncabezado"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'variable local para contener colección
Public Event EncabezadoAgregado()

Private mCol As Collection

Public Function Add(nombre As String, miembro As String, ancho As Integer) As LVCEncabezado
    'crear un nuevo objeto
    On Error GoTo e
    Dim objNewMember As LVCEncabezado
    Set objNewMember = New LVCEncabezado


    'establecer las propiedades que se transfieren al método
    objNewMember.miembro = miembro
    objNewMember.ancho = ancho
    objNewMember.nombre = nombre
100 mCol.Add objNewMember, miembro  'estaba nombre, lo cambie por miembro para probar...
    
    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing
    RaiseEvent EncabezadoAgregado
    Exit Function
e:
    If Erl = 100 Then
        'cambiado 17/3/11
        tErr.AppendLog "daer922", "la clave " + miembro + " ya esta asociada a la coleccion..."
        'Err.Raise 457 'la clave ya esta asociada a la coleccion...
        Resume Next
    End If
End Function

Public Property Get Item(pKey As String) As LVCEncabezado
Attribute Item.VB_UserMemId = 0
    On Error GoTo errman
    Set Item = mCol(pKey)
    Exit Property
errman:
    If Err.Number = 5 Then Err.Raise 2011, , "No se encontro el encabezado del LVWConsulta."
    
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(pKey As String)
    mCol.Remove pKey
End Sub

Public Sub Clear()
    Set mCol = Nothing
    Set mCol = New Collection
End Sub

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

