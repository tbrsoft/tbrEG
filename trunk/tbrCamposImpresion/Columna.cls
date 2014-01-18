VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Columna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eAlineacion
    eIzquierda = 1
    eCentrado = 2
    eDerecha = 3
End Enum

Private vWidth As Single
Private vFontSize As Integer
Private vFont As String
Private vFontBold As Boolean
Private vFontItalic As Boolean
Private vForeColor As Long
Private vNombre As String
Private vId As Integer
Private vDescripcion As String
Private vOrden As Integer
Private vMaxLength As Integer
Private vErrorMessage As String
Private vDataType As eTipoDato
Private vAlineacion As eAlineacion
'ver despues si se utiliza para datos complejos
Private vFormatString As String
Private vTieneCambios As Boolean

Public Event ErrorDataType()

'------------------------- property get----------------------
Friend Property Get TieneCambios() As Boolean
    TieneCambios = vTieneCambios
End Property

Public Property Get Width() As Single
    Width = vWidth
End Property

Public Property Get MaxLength() As Integer
    MaxLength = vMaxLength
End Property

Public Property Get FontSize() As Integer
    FontSize = vFontSize
End Property

Public Property Get ForeColor() As Long
    ForeColor = vForeColor
End Property

Public Property Get FontBold() As Boolean
    FontBold = vFontBold
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = vFontItalic
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

Public Property Get Alineacion() As eAlineacion
    Alineacion = vAlineacion
End Property

'------------------------- property let----------------------

Public Property Let Width(pValue As Single)
    vWidth = pValue
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

Public Property Let FontBold(pValue As Boolean)
    vFontBold = pValue
    vTieneCambios = True
End Property

Public Property Let FontItalic(pValue As Boolean)
    vFontItalic = pValue
    vTieneCambios = True
End Property

Public Property Let ForeColor(pValue As Long)
    vForeColor = pValue
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

Public Property Let Alineacion(pValue As eAlineacion)
    vAlineacion = pValue
    vTieneCambios = True
End Property

Friend Property Let TieneCambios(pValue As Boolean)
    vTieneCambios = pValue
End Property
'-----------------------funciones--------------------------------

Public Function DataTypeToString() As String
    DataTypeToString = Choose(vDataType, "String", "Integer", "Date")
End Function

Public Function AlineacionToString() As String
    AlineacionToString = Choose(vAlineacion, "Izquierda", "Centrado", "Derecha")
End Function

Friend Function ColumnaToString() As String
    Dim aux As String
    aux = "BEGIN COLUMNA" + vbCrLf
    aux = aux + vbTab + "datatype=" + Trim(Str(Me.DataType)) + vbCrLf
    aux = aux + vbTab + "descripcion=" + Me.Descripcion + vbCrLf
    aux = aux + vbTab + "errormessage=" + Me.ErrorMessage + vbCrLf
    aux = aux + vbTab + "fontbold=" + Trim(Str(Me.FontBold)) + vbCrLf
    aux = aux + vbTab + "fontitalic=" + Trim(Str(Me.FontItalic)) + vbCrLf
    aux = aux + vbTab + "fontname=" + Me.FontName + vbCrLf
    aux = aux + vbTab + "fontsize=" + Trim(Str(Me.FontSize)) + vbCrLf
    aux = aux + vbTab + "forecolor=" + Trim(Str(Me.ForeColor)) + vbCrLf
    aux = aux + vbTab + "id=" + Trim(Str(Me.Id)) + vbCrLf
    aux = aux + vbTab + "maxlength=" + Trim(Str(Me.MaxLength)) + vbCrLf
    aux = aux + vbTab + "nombre=" + Me.Nombre + vbCrLf
    aux = aux + vbTab + "orden=" + Trim(Str(Me.Orden)) + vbCrLf
    aux = aux + vbTab + "width=" + Trim(Str(Me.Width)) + vbCrLf
    aux = aux + vbTab + "alineacion=" + Trim(Str(Me.Alineacion)) + vbCrLf
    aux = aux + "END COLUMNA" + vbCrLf
    ColumnaToString = aux
End Function

Friend Function DatosCorrectos(pNewValue As String) As Boolean
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
