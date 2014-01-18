VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TelefonoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A802F30196"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Private mCol As Collection
Private mIdOwner As Long
Private mOwnerType As eOwnerType

'##ModelId=46A80B0F01A5
Public Sub Remove(pId As Long)
    On Error GoTo RemoveErr

    mCol.Remove "k" + Str(pId)

    Exit Sub
RemoveErr:
    Call ErrorLog("TelefonoManager", "Remove Sub", Err.Description)
End Sub

'##ModelId=46A80B3C0261
Public Sub AddItem(pTelefono As Telefono)
    On Error GoTo AddItemErr

    mCol.Add pTelefono, "k" + Str(pTelefono.Id)

    Exit Sub
AddItemErr:
    ErrorLog "TelefonoManager", "AddItem", Err.Description
End Sub

Friend Function Add(pId As Long, pNumero As String, pTipo As TipoTelefono, pObservaciones As String, pOwnerId As Long, pOwnerType As eOwnerType) As Telefono
    Dim tel As Telefono
    Set tel = New Telefono
    tel.Id = pId
    tel.Numero = pNumero
    tel.Observaciones = pObservaciones
    Set tel.tipo = pTipo
    tel.OwnerId = pOwnerId
    tel.OwnerType = pOwnerType
    
    mCol.Add tel, "k" + Str(tel.Id)
    Set Add = tel
    Set tel = Nothing
End Function

Public Function Nuevo(pNumero As String, pTipo As TipoTelefono, pObservaciones As String) As Telefono
    'este nuevo no guarda el registro en la bd
    'habria q preguntar por ballowmodify
    Dim tel As Telefono
    Set tel = New Telefono
    'tel.id = IdProvisorio
    tel.Numero = pNumero
    tel.Observaciones = pObservaciones
    Set tel.tipo = pTipo
    
    'cuando estoy creando al dueño, no estan establecidos. cuando los agrego despues si
    tel.OwnerId = mIdOwner
    tel.OwnerType = mOwnerType
    
    mCol.Add tel ', "k" + Str(tel.id)
    Set Nuevo = tel
    Set tel = Nothing
End Function

'devuelve el dueño del telefono, puede ser un afiliado, una os,un empleado, etc...
Public Property Get OwnerId() As Long
    OwnerId = mIdOwner
End Property

'devuelve el tipo del dueño del telefono, puede ser un afiliado, una os,un empleado, etc...
Public Property Get OwnerType() As eOwnerType
    OwnerType = mOwnerType
End Property

Friend Property Let OwnerId(pValue As Long)
     mIdOwner = pValue
End Property

Friend Property Let OwnerType(pValue As eOwnerType)
    mOwnerType = pValue
End Property

Public Property Get Item(pId As Long) As Telefono
    'si no existe devuelve nothing
    On Error Resume Next
    Set Item = mCol.Item("k" + Str(pId))
End Property

Public Property Get ItemByTelNumber(pNumber As String) As Telefono
Attribute ItemByTelNumber.VB_Description = "Devuelve un telefono que se corresponde con el numero puro del telefono (sin guiones ni parentesis)"
    'si no existe devuelve nothing
    On Error Resume Next
    Dim t As Telefono
    Dim aux As String
    aux = Replace(pNumber, "(", "")
    aux = Replace(aux, ")", "")
    aux = Replace(aux, " ", "")
    aux = Replace(aux, "-", "")
    For Each t In Me
        If t.NumeroPuro = aux Then
            Set ItemByTelNumber = t
            Exit For
        End If
    Next
         
End Property


Public Function Exists(pId As Long) As Boolean
    Dim t As Telefono
    Set t = Me.Item(pId)
    Exists = (Not t Is Nothing)
End Function

'----funciones auxiliares privadas----------------------

Private Function IdProvisorio() As Long
Static idProv As Long
'aca tengo q obtener el maximo id de los q ya tengo, y de ahi ir incrementando
idProv = idProv + 1
IdProvisorio = idProv
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

'----------------------------BASE DE DATOS----------------------

Friend Function GetByOwner(pOwnerId As Long, pOwnerType As eOwnerType) As TelefonoManager

Dim tels As New TelefonoManager
Dim t As Telefono
'levanto solo los telefonos q hacen falta
For Each t In Me
    If t.OwnerId = pOwnerId And t.OwnerType = pOwnerType Then tels.AddItem t
Next
tels.OwnerId = pOwnerId
tels.OwnerType = pOwnerType

Set GetByOwner = tels
End Function

'los necesito cargados para cuando me llaman saber quien es
Friend Function CargarTodos()
On Error GoTo errman:
Dim rs
Set rs = BD.SelectSQL("SELECT * FROM TELEFONO")

While Not rs.EOF
    Add rs!Id, rs!Numero, TiposTelefonoLocal.Item(rs!idTipoTelefono), rs!Observaciones, rs!idOwner, rs!tipoOwner
    rs.MoveNext
Wend
Exit Function
errman:
ErrorLog "TelefonosManager", "CargarTodos ", Err.Description
End Function

Friend Function GuardarTelefonos(pOwnerId As Long, pOwnerType As eOwnerType, pTrans As Transaccion)

mOwnerType = pOwnerType
mIdOwner = pOwnerId

Dim t As Telefono

For Each t In Me
    
        t.OwnerId = pOwnerId
        t.OwnerType = pOwnerType
        t.Insert pTrans
    
Next

End Function

'muy probablemente se pueda hacer una sola funcion para guardar y savechanges
Friend Sub SaveChanges(pTrans As Transaccion)
            
Dim t As Telefono
Delete pTrans 'borro todos y los vuelvo a agregar
GuardarTelefonos mIdOwner, mOwnerType, pTrans
    
End Sub

'me borra todos los telefonos de este dueño, puede servir para la funcion eliminar
Private Sub Delete(pTrans As Transaccion)
    Dim params As ParametrosManager
    Set params = New ParametrosManager
    params.Add "idOwner", mIdOwner, True
    params.Add "tipoOwner", mOwnerType, True
    pTrans.AddDeleteCommand tTelefono, params
End Sub

