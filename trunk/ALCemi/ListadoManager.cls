VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListadoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mCol As Collection

Public Sub Add(pPath As String)
    Dim l As New Listado
    l.Load (pPath)
    mCol.Add l, LCase(pPath)
End Sub

Public Sub AddItem(pListado As Listado)
    mCol.Add pListado, LCase(pListado.Path)
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

Public Property Get Item(pPath As String) As Listado
    Set Item = mCol.Item(LCase(pPath))
End Property

Public Sub Load(pFolder As String)
    Dim D As New FileSystemObject
    Dim f As File
    If D.FolderExists(pFolder) Then
        For Each f In D.GetFolder(pFolder).Files
             If Right(f.Name, 3) = "egr" Then
                Add (f.Path)
             End If
        Next
    End If
End Sub
