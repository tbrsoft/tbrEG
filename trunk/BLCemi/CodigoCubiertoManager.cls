VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodigoCubiertoManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mCol As Collection

Friend Function Add(pCodigo As CodigoEmergencia, pTipo As TipoCodigo, pCoseguro As Currency, pServicio As Currency) As CodigoCubierto
    Dim c As New CodigoCubierto
    Set c.Codigo = pCodigo
    c.Coseguro = pCoseguro
    c.Servicio = pServicio
    Set c.TipoCodigo = pTipo
    Dim clave As String
    If pTipo Is Nothing Then
        clave = Str(c.Codigo.Id * 100)
    Else
        clave = Str(c.Codigo.Id * 100 + c.TipoCodigo.Id)
    End If
    mCol.Add c, "k" + clave
    Set Add = c
    Set c = Nothing
End Function

Public Property Get Item(pIdCodigo As Long, pIdTipo As Long) As CodigoCubierto
    On Error Resume Next
    Dim cc As CodigoCubierto
    For Each cc In Me
        'si no usa tipos devuelvo el primero q aparezca con el codigo
        If pIdTipo = 0 Then
            If cc.Codigo.Id = pIdCodigo Then
                Set Item = cc
                Exit For
            End If
        Else
            If cc.Codigo.Id = pIdCodigo And cc.TipoCodigo.Id = pIdTipo Then
                Set Item = cc
                Exit For
            End If
        End If
    Next
'    Dim clave As String
'    clave = Str(IIf(pIdTipo = 0, pIdCodigo * 100, pIdCodigo * 100 + pIdTipo))
'    Set Item = mCol("k" + clave)
End Property

Public Sub Remove(pCodigoCubierto As CodigoCubierto)
    Dim clave As String
    If pCodigoCubierto.TipoCodigo Is Nothing Then
        clave = pCodigoCubierto.Codigo.Id * 100
    Else
        clave = pCodigoCubierto.Codigo.Id * 100 + pCodigoCubierto.TipoCodigo.Id
    End If
    mCol.Remove "k" + clave
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

Public Function Exists(pIdCodigo As Long, pIdTipo As Long) As Boolean
    Dim t As CodigoEmergencia
    Set t = Me.Item(pIdCodigo, pIdTipo)
    Exists = (Not t Is Nothing)
End Function

Public Sub AddItem(pCodigoCubierto As CodigoCubierto)
    On Error GoTo AddItemErr
    Dim clave As String
    If pCodigoCubierto.TipoCodigo Is Nothing Then
        clave = pCodigoCubierto.Codigo.Id * 100
    Else
        clave = pCodigoCubierto.Codigo.Id * 100 + pCodigoCubierto.TipoCodigo.Id
    End If
    mCol.Add pCodigoCubierto, "k" + clave

    Exit Sub
AddItemErr:
    ErrorLog "CodigoCubiertoManager", "AddItem", err.Description
End Sub

Friend Sub Cargar(pTipo As eDestinoCodigo, pId As Long)
On Error GoTo errMaq

Dim tipoAux As TipoCodigo
Dim rs
Dim tipo As String
Set rs = BD.SelectSQL("Select * from CodigoXEmpresa where idEmpresa=" + Str(pId) + " and tipoEmpresa=" + Str(pTipo))
While Not rs.EOF
    If IsNull(rs!idTipoCodigo) Then
        Set tipoAux = Nothing
    Else
        Set tipoAux = TiposCodigoLocal.Item(rs!idTipoCodigo)
    End If
    Add CodigoEmergenciaLocal.Item(rs!idCodigo), tipoAux, rs!Coseguro, rs!Servicio
    rs.MoveNext
Wend
Exit Sub
errMaq:
ErrorLog "CodigoCubiertoManager", "cargar", err.Description

End Sub

Friend Sub Guardar(pTipo As eDestinoCodigo, pTrans As Transaccion, idEmpresa As Long)

Dim cod As CodigoCubierto
For Each cod In Me
    cod.Save pTrans, pTipo, idEmpresa
Next

End Sub


Friend Sub SaveChanges(pTipo As eDestinoCodigo, pTrans As Transaccion, idEmpresa As Long)
    'borro todos los Codigos de esta OS
    Dim params As New ParametrosManager
    params.Add "idEmpresa", idEmpresa, True
    params.Add "tipoEmpresa", pTipo, True
    pTrans.AddDeleteCommand tCodigoXEmpresa, params
    'los vuelvo a guardar
    Me.Guardar pTipo, pTrans, idEmpresa
End Sub

Public Function GetCadenaCodigos() As String
On Error GoTo errman
Dim aux As String
Dim cod As CodigoCubierto
For Each cod In Me
    aux = aux + Trim(Str(cod.Codigo.Id)) + ", "
Next
GetCadenaCodigos = Left(aux, Len(aux) - 2)
Exit Function
errman:
GetCadenaCodigos = ""
End Function



'Public Sub BeginEdit()
'    Clone
'End Sub
'
'Public Sub CancelChanges()
'    If Not mCodigosOriginal Is Nothing Then
'        Set mCol = mCodigosOriginal
'        Set mCodigosOriginal = Nothing
'    End If
'End Sub
'Private Sub Clone()
'    Dim cm As Collection
'    Set cm = New Collection
'    Dim c As CodigoEmergencia
'    For Each c In Me
'        cm.Add c, "k" + Str(c.id) ' c.Clone, "k" + Str(c.id), no se si hace falta clonarlo
'    Next
'    Set mCodigosOriginal = cm
'End Sub
