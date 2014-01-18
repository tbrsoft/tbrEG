VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Campo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum eTipoDato
    etdString = 1
    etdInteger = 2
    etdDate = 3
End Enum

Private vTop As Single
Private vLeft As Single
Private vWidth As Single
Private vHeight As Single
Private vFontSize As Integer
Private vFont As String
Private vNombre As String
Private vId As Integer
Private vDescripcion As String
Private vOrden As Integer
Private vValue As String
Private vMaxLength As Integer
Private vErrorMessage As String
Private vDataType As eTipoDato
'ver despues si se utiliza para datos complejos
Private vFormatString As String

Private vTieneCambios As Boolean

Public Event ErrorDataType()

'------------------------- property get----------------------
Friend Property Get TieneCambios() As Boolean
    TieneCambios = vTieneCambios
End Property

Public Property Get Top() As Single
    Top = vTop
End Property

Public Property Get Left() As Single
    Left = vLeft
End Property

Public Property Get Width() As Single
    Width = vWidth
End Property

Public Property Get Height() As Single
    Height = vHeight
End Property

Public Property Get MaxLength() As Integer
    MaxLength = vMaxLength
End Property

Public Property Get FontSize() As Integer
    FontSize = vFontSize
End Property

Public Property Get ErrorMessage() As String
    ErrorMessage = vErrorMessage
End Property

Public Property Get FontName() As String
    FontName = vFont
End Property

Public Property Get Nombre() As String
    Nombre = vNombre
End Property

Public Property Get Descripcion() As String
    Descripcion = vDescripcion
End Property

Public Property Get Orden() As Integer
    Orden = vOrden
End Property

Public Property Get Id() As Integer
    Id = vId
End Property

Public Property Get DataType() As eTipoDato
    DataType = vDataType
End Property

Public Property Get Valor() As String
    Valor = vValue
End Property

'------------------------- property let----------------------

Public Property Let Top(pValue As Single)
    vTop = pValue
    vTieneCambios = True
End Property

Public Property Let Left(pValue As Single)
    vLeft = pValue
    vTieneCambios = True
End Property

Public Property Let Width(pValue As Single)
    vWidth = pValue
    vTieneCambios = True
End Property

Public Property Let Height(pValue As Single)
    vHeight = pValue
    vTieneCambios = True
End Property

Public Property Let MaxLength(pValue As Integer)
    vMaxLength = pValue
    vTieneCambios = True
End Property

Public Property Let FontSize(pValue As Integer)
    vFontSize = pValue
    vTieneCambios = True
End Property

Public Property Let ErrorMessage(pValue As String)
    vErrorMessage = pValue
    vTieneCambios = True
End Property

Public Property Let FontName(pValue As String)
    vFont = pValue
    vTieneCambios = True
End Property

Public Property Let Nombre(pValue As String)
    vNombre = pValue
    vTieneCambios = True
End Property

Public Property Let Descripcion(pValue As String)
    vDescripcion = pValue
    vTieneCambios = True
End Property

Public Property Let Orden(pValue As Integer)
    vOrden = pValue
    vTieneCambios = True
End Property

Public Property Let Id(pValue As Integer)
    vId = pValue
    vTieneCambios = True
End Property

Public Property Let DataType(pValue As eTipoDato)
    vDataType = pValue
    vTieneCambios = True
End Property

Public Property Let Valor(pValue As String)
    If DatosCorrectos(pValue) Then
        vValue = pValue
    Else
        'ver q numero de error mas coherente se puede usar
        'si le quiero asignar un valor no valido levanto un error
'        Err.Clear
'        Err.Raise 2007, "tbrCamposImpresion -> Campo", IIf(vErrorMessage <> "<Sin Mensaje de Error>", vErrorMessage, "Se intento asignar un valor no valido al campo.")
    RaiseEvent ErrorDataType
    End If
End Property

Friend Property Let TieneCambios(pValue As Boolean)
    vTieneCambios = pValue
End Property
'-----------------------funciones--------------------------------

Public Function DataTypeToString() As String
    DataTypeToString = Choose(vDataType, "String", "Integer", "Date")
End Function

Private Function DatosCorrectos(pNewValue As String) As Boolean
    On Error GoTo e
    Select Case vDataType
        Case etdString
            If vMaxLength <> 0 Then 'si maxlength es 0 se permite cualquier largo
                If vMaxLength >= Len(pNewValue) Then DatosCorrectos = True Else DatosCorrectos = False
            Else
                DatosCorrectos = True
            End If
        Case etdDate
            If IsDate(pNewValue) Then DatosCorrectos = True Else DatosCorrectos = False
        Case etdInteger
            If IsNumeric(pNewValue) Then DatosCorrectos = True Else DatosCorrectos = False
    End Select
    Exit Function
e:
    DatosCorrectos = False
End Function

Friend Function CampoToString() As String
'convierte el contenido del campo en string para guardarlo en un archivo
    Dim aux As String
    aux = aux + "BEGIN CAMPO" + vbCrLf
    aux = aux + vbTab + "datatype=" + Str(Me.DataType) + vbCrLf
    aux = aux + vbTab + "descripcion=" + Me.Descripcion + vbCrLf
    aux = aux + vbTab + "errormessage=" + Me.ErrorMessage + vbCrLf
    aux = aux + vbTab + "fontname=" + Me.FontName + vbCrLf
    aux = aux + vbTab + "fontsize=" + Str(Me.FontSize) + vbCrLf
    aux = aux + vbTab + "height=" + Str(Me.Height) + vbCrLf
    aux = aux + vbTab + "width=" + Str(Me.Width) + vbCrLf
    aux = aux + vbTab + "id=" + Str(Me.Id) + vbCrLf
    aux = aux + vbTab + "left=" + Str(Me.Left) + vbCrLf
    aux = aux + vbTab + "maxlength=" + Str(Me.MaxLength) + vbCrLf
    aux = aux + vbTab + "nombre=" + Me.Nombre + vbCrLf
    aux = aux + vbTab + "orden=" + Str(Me.Orden) + vbCrLf
    aux = aux + vbTab + "top=" + Str(Me.Top) + vbCrLf
    aux = aux + "END CAMPO" + vbCrLf
    CampoToString = aux
End Function
