VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColeccionGenerica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim mCol As Collection

Public Sub Fill(rs As Object)
    Dim objG As ObjetoGenerico
    Dim f As Object
    While Not rs.EOF
        Set objG = New ObjetoGenerico
        For Each f In rs.Fields
            objG.AddProperty f.Name, f.Value
        Next
        mCol.Add objG
        rs.MoveNext
    Wend
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub
