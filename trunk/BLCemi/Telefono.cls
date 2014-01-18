VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Telefono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A802E9005D"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Class Module"
Option Explicit

Public Enum eOwnerType
    eOTAfiliado = 1
    eOTAreaProtegida = 2
    eOTServicioEmergencia = 3
    eOTObraSocial = 4
    eOTEmpleado = 5
    eOTAfiliadoExterno = 6
End Enum

'##ModelId=46A804980290
Private mId As Long

'##ModelId=46A8095403D8
Private mTipo As TipoTelefono

'##ModelId=46A809550399
Private mNumero As String

'##ModelId=46A8189602BF
Private mObservaciones As String

Private mIdOwner As Long
Private mOwnerType As eOwnerType 'el tipo de dueño

'##ModelId=46A978940153
Public Property Get Id() As Long
    On Error GoTo idErr

    '## Generated default body ...
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Telefono", "Id Property Get", Err.Description)
End Property

'##ModelId=46A978940154
Public Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    '## Generated default body ...
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Telefono", "Id Property Let", Err.Description)
End Property

'##ModelId=46A978940158
Public Property Get tipo() As TipoTelefono
    On Error GoTo TipoErr

    '## Generated default body ...
    Set tipo = mTipo

    Exit Property
TipoErr:
    Call ErrorLog("Telefono", "Tipo Property Get", Err.Description)
End Property

'##ModelId=46A97894015B
Public Property Set tipo(ByVal vNewValue As TipoTelefono)
    On Error GoTo TipoErr

    '## Generated default body ...
    Set mTipo = vNewValue

    Exit Property
TipoErr:
    Call ErrorLog("Telefono", "Tipo Property Set", Err.Description)
End Property


Public Property Get NumeroPuro() As String
    On Error GoTo NumeroErr

        
    Dim aux As String
    aux = Replace(mNumero, "(", "")
    aux = Replace(aux, ")", "")
    aux = Replace(aux, " ", "")
    aux = Replace(aux, "-", "")
    
    NumeroPuro = aux

    Exit Property
NumeroErr:
    Call ErrorLog("Telefono", "Numero Property Get", Err.Description)
End Property


'##ModelId=46A97894015D
Public Property Get Numero() As String
    On Error GoTo NumeroErr

    '## Generated default body ...
    Numero = mNumero

    Exit Property
NumeroErr:
    Call ErrorLog("Telefono", "Numero Property Get", Err.Description)
End Property

'##ModelId=46A97894015E
Public Property Let Numero(ByVal vNewValue As String)
    On Error GoTo NumeroErr

    '## Generated default body ...
    mNumero = vNewValue

    Exit Property
NumeroErr:
    Call ErrorLog("Telefono", "Numero Property Let", Err.Description)
End Property

'##ModelId=46A978940162
Public Property Get Observaciones() As String
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    Observaciones = mObservaciones

    Exit Property
ObservacionesErr:
    Call ErrorLog("Telefono", "Observaciones Property Get", Err.Description)
End Property

'##ModelId=46A978940163
Public Property Let Observaciones(ByVal vNewValue As String)
    On Error GoTo ObservacionesErr

    '## Generated default body ...
    mObservaciones = vNewValue

    Exit Property
ObservacionesErr:
    Call ErrorLog("Telefono", "Observaciones Property Let", Err.Description)
End Property

Friend Function Equal(tDestino As Telefono) As Boolean
Equal = (tDestino.Numero = Me.Numero) And (tDestino.Observaciones = Me.Observaciones) And (tDestino.tipo.Id = Me.tipo.Id)
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

Public Function GetProperty(pPropertyName As String) As String
    pPropertyName = LCase(pPropertyName)
    Select Case pPropertyName
        Case "tipo"
            GetProperty = mTipo.Nombre
        Case "numero"
            GetProperty = Me.Numero
        Case "observaciones"
            GetProperty = Me.Observaciones
    End Select
End Function

Friend Function Clone() As Telefono
Dim t As New Telefono
t.Id = Me.Id
t.Numero = Me.Numero
t.Observaciones = Me.Observaciones
Set t.tipo = Me.tipo
Set Clone = t
End Function

'les tengo q mandar el id y tipo porq no los se hasta q los guardo
Friend Sub Insert(pTrans As Transaccion)
    Dim params As New ParametrosManager
    'el id q tienen los tel hasta aca es provisorio, aca le asigno uno valido
    If mId = 0 Then mId = BD.getIdDisponible(tTelefono)
    params.Add "id", mId
    params.Add "numero", mNumero
    params.Add "idTipoTelefono", mTipo.Id
    params.Add "observaciones", mObservaciones
    params.Add "idOwner", mIdOwner
    params.Add "tipoOwner", mOwnerType
    
    pTrans.AddInsertCommand tTelefono, params
End Sub

Friend Sub Delete(pTrans As Transaccion)
    Dim params As New ParametrosManager
    params.Add "id", mId, True
    pTrans.AddDeleteCommand tTelefono, params
End Sub

'Friend Sub Update(pTrans As Transaccion)
'    Dim params As New ParametrosManager
'    params.Add "id", mId, True
'    params.Add "numero", mNumero
'    params.Add "idTipoTelefono", mTipo.id
'    params.Add "observaciones", mObservaciones
'
'    pTrans.AddUpdateCommand tTelefono, params
'End Sub
