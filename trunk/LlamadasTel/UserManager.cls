VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'variable local para contener colección
Private mCol As Collection

Public Function Add(pUsername As String, pPass As String) As User
    'crear un nuevo objeto
    Dim objNewMember As User
    Set objNewMember = New User
    
    objNewMember.UserName = pUsername
    objNewMember.Pass = pPass
    
    mCol.Add objNewMember, pUsername
    
    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Function NewUser(pUsername As String, pPass As String) As User
    
    If Me.Item(pUsername) Is Nothing And pPass <> "" Then
        'crear un nuevo objeto
        Dim objNewMember As User
        Set objNewMember = New User
        objNewMember.UserName = pUsername
        objNewMember.Pass = pPass
        objNewMember.Save
        mCol.Add objNewMember, pUsername
        
        'devolver el objeto creado
        Set NewUser = objNewMember
        Set objNewMember = Nothing
    End If
End Function

Public Function Login(pUsername As String, pPass As String) As User
Dim u As User
Set u = colUsers.Item(pUsername)
If u.Pass = pPass Then
    Set Login = u
Else
    Set Login = Nothing
End If
End Function

Friend Sub LoadUsers()

On Error GoTo errMaq

Dim rs As New ADODB.Recordset
Set rs = bd.select_("select * from User")
While Not rs.EOF
    
    Add rs!Login, "" & rs!Pass
    rs.MoveNext
    
Wend
Exit Sub
errMaq:
MsgBox "Error coleccion clientes, sub cargarclientes " + Err.Description
End Sub

Public Property Get Item(pUsername As String) As User
  Set Item = mCol(pUsername)
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


