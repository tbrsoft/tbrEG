VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tabla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Type EstadoImpresor
    DrawStyle As Integer
    FontName As String
    FontSize As Single
    Bold As Boolean
    Italic As Boolean
    ForeColor As Long
End Type


Private vTop As Single
Private vLeft As Single
Private vWidth As Single
Private vHeight As Single
Private vRows As Integer
Private vFontSize As Integer
Private vFont As String
Private vNombre As String
Private vId As Integer
Private vDescripcion As String
Private vOrden As Integer
Private vErrorMessage As String
Private vEstilo As Integer

Private vTieneCambios As Boolean
Private vValue() As String
Private WithEvents vColumnas As ColumnasManager
Attribute vColumnas.VB_VarHelpID = -1

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

Public Property Get Columnas() As ColumnasManager
    Set Columnas = vColumnas
End Property

Public Property Get Rows() As Integer
    Rows = vRows
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

Public Property Get Celda(pFila As Integer, pColumna As Integer) As String
    If pFila < 0 Or pFila > Rows Or pColumna < 0 Or pColumna > Columnas.Count Then
        Err.Raise 9 'el elemento esta fuera del intervalo
        Exit Property
    End If
    Celda = vValue(pFila, pColumna)
End Property

Public Property Get EstiloTabla() As Integer
    EstiloTabla = vEstilo
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

Public Property Let Columnas(pValue As ColumnasManager)
    Set vColumnas = pValue
    vTieneCambios = True
End Property

Public Property Let Rows(pValue As Integer)
    On Error Resume Next
    vRows = pValue
    ReDim vValue(1 To vRows, 1 To Columnas.Count)
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

Friend Property Let TieneCambios(pValue As Boolean)
    vTieneCambios = pValue
End Property

Public Property Let EstiloTabla(pValue As Integer)
    vEstilo = pValue
    vTieneCambios = True
End Property

Public Property Let Celda(pFila As Integer, pColumna As Integer, pValue As String)
    If pFila < 0 Or pFila > Rows Or pColumna < 0 Or pColumna > Columnas.Count Then
        Err.Raise 9 'el elemento esta fuera del intervalo
        Exit Property
    End If
    If Columnas.ItemByOrden(pColumna).DatosCorrectos(pValue) Then
        vValue(pFila, pColumna) = pValue
    Else
        RaiseEvent ErrorDataType
    End If
End Property

'Public Property Let Valor(pValue As String)
'    If DatosCorrectos(pValue) Then
'        vValue = pValue
'    Else
'        'ver q numero de error mas coherente se puede usar
'        'si le quiero asignar un valor no valido levanto un error
''        Err.Clear
''        Err.Raise 2007, "tbrCamposImpresion -> Campo", IIf(vErrorMessage <> "<Sin Mensaje de Error>", vErrorMessage, "Se intento asignar un valor no valido al campo.")
'    RaiseEvent ErrorDataType
'    End If
'End Property

'-----------------------funciones--------------------------------
Public Function EstiloTablaToString() As String
    EstiloTablaToString = Choose(vEstilo + 1, "Transparent", "Solid", "Dash", "Dot", "DashDot", "DashDotDot")
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

Public Sub Dibujar(pic As Object, pLeft As Single, pTop As Single)
'pic.Cls
Dim j As Integer
Dim col As Integer
Dim fila As Integer
Dim altoFila As Single
Dim c As Columna
Dim anchoAcum As Single
altoFila = Height / vRows
'anchocolumna = pic.Width / vColumns
'setear tipo de linea tabla

pic.DrawStyle = IIf(EstiloTabla = 0, 5, EstiloTabla - 1)
Debug.Print pic.DrawStyle
For i = 0 To vRows
    pic.Line (pLeft, (altoFila * i) + pTop)-(Width + pLeft, (altoFila * i) + pTop)
Next

For j = 1 To Columnas.Count
    Set c = Columnas.ItemByOrden(j)
    anchoAcum = anchoAcum + c.Width
    pic.Line (anchoAcum + pLeft, pTop)-(anchoAcum + pLeft, pTop + Height)
Next

'primera linea vertical
pic.Line (pLeft, pTop)-(pLeft, pTop + Height)

'imprimo el contenido de la tabla

For fila = 1 To vRows
    anchoAcum = 0
    For col = 1 To Columnas.Count
        Set c = Columnas.ItemByOrden(col)
            
        'ImprimirTexto pic, Str(fila) + "," + Str(col), c, anchoAcum + pLeft, altoFila * (fila - 1) + pTop, altoFila
        ImprimirTexto pic, Celda(fila, col), c, anchoAcum + pLeft, altoFila * (fila - 1) + pTop, altoFila
        
        anchoAcum = anchoAcum + c.Width
    Next
Next

End Sub

Private Sub ImprimirTexto(pic As Object, texto As String, col As Columna, x As Single, y As Single, cHeight As Single)
    
    'guardo el estado del picture o la impresora
    Dim estado As EstadoImpresor
    GuardarEstado pic, estado
        
    'seteo la impresora con los valores de la columna
    pic.FontName = col.FontName
    pic.FontSize = col.FontSize
    pic.FontBold = col.FontBold
    pic.FontItalic = col.FontItalic
    pic.ForeColor = col.ForeColor
    
    'calculo el tamaño del texto
    anchoTexto = pic.TextWidth(texto)
    altoTexto = pic.TextHeight(texto)
    
    Select Case col.Alineacion
        Case eAlineacion.eCentrado
            pic.CurrentX = x + (col.Width - anchoTexto) / 2
        Case eAlineacion.eDerecha
            pic.CurrentX = x + (col.Width - anchoTexto)
        Case eAlineacion.eIzquierda
            pic.CurrentX = x
    End Select
           
    pic.CurrentY = y + (cHeight - altoTexto) / 2
    pic.Print texto
    
    'restauro el estado
    RestaurarEstado pic, estado
End Sub

Friend Function TablaToString() As String
    Dim aux As String
    aux = "BEGIN TABLA" + vbCrLf
    aux = aux + vbTab + "descripcion=" + Me.Descripcion + vbCrLf
    aux = aux + vbTab + "errormessage=" + Me.ErrorMessage + vbCrLf
    aux = aux + vbTab + "fontname=" + Me.FontName + vbCrLf
    aux = aux + vbTab + "fontsize=" + Trim(Str(Me.FontSize)) + vbCrLf
    aux = aux + vbTab + "height=" + Trim(Str(Me.Height)) + vbCrLf
    aux = aux + vbTab + "id=" + Trim(Str(Me.Id)) + vbCrLf
    aux = aux + vbTab + "left=" + Trim(Str(Me.Left)) + vbCrLf
    aux = aux + vbTab + "nombre=" + Me.Nombre + vbCrLf
    aux = aux + vbTab + "orden=" + Trim(Str(Me.Orden)) + vbCrLf
    aux = aux + vbTab + "rows=" + Trim(Str(Me.Rows)) + vbCrLf
    aux = aux + vbTab + "top=" + Trim(Str(Me.Top)) + vbCrLf
    aux = aux + vbTab + "width=" + Trim(Str(Me.Width)) + vbCrLf
    aux = aux + vbTab + "tablestyle=" + Trim(Str(Me.EstiloTabla)) + vbCrLf
    aux = aux + vbTab + Me.Columnas.ColumnasToString + vbCrLf
    aux = aux + "END TABLA" + vbCrLf
    TablaToString = aux
End Function

Private Sub Class_Initialize()
    Set vColumnas = New ColumnasManager
End Sub

Private Sub vColumnas_ColumnCountChanged()
    On Error Resume Next
    
    ReDim vValue(1 To vRows, 1 To Columnas.Count)
End Sub

Private Sub GuardarEstado(obj As Object, estado As EstadoImpresor)
    estado.Bold = obj.FontBold
    estado.DrawStyle = obj.DrawStyle
    estado.FontName = obj.FontName
    estado.FontSize = obj.FontSize
    estado.ForeColor = obj.ForeColor
    estado.Italic = obj.FontItalic
End Sub

Private Sub RestaurarEstado(obj As Object, estado As EstadoImpresor)
    obj.FontBold = estado.Bold
    obj.DrawStyle = estado.DrawStyle
    obj.FontName = estado.FontName
    obj.FontSize = estado.FontSize
    obj.ForeColor = estado.ForeColor
    obj.FontItalic = estado.Italic
End Sub
