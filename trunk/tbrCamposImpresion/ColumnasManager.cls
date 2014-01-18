VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColumnasManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event ColumnCountChanged()
Private mcol As New Collection
Private vTieneCambios As Boolean

Friend Function AddColumnaFromString(cadena As String) As Columna
    Dim aux() As String
    Dim par() As String
    cadena = Replace(cadena, "END COLUMNA", "")
    'aca tengo todas las propiedades
    aux = Split(cadena, vbCrLf)
    'cargo todas las propiedades en un diccionario
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    For I = 1 To UBound(aux) - 2
        If aux(I) <> "" Then
            par = Split(aux(I), "=")
            d.Add par(0), par(1)
        End If
    Next
    
    Dim c As New Columna
     
    c.Id = d("id")
    c.Nombre = d("nombre")
    c.Width = CSng(d("width"))
    c.FontName = d("fontname")
    c.FontSize = CInt(d("fontsize"))
    c.FontBold = CBool(d("fontbold"))
    c.FontItalic = CBool(d("fontitalic"))
    c.Descripcion = d("descripcion")
    c.MaxLength = CInt(d("maxlength"))
    c.ErrorMessage = d("errormessage")
    c.DataType = CInt(d("datatype"))
    c.Alineacion = CInt(d("alineacion"))
    c.Orden = CInt(d("orden"))
    c.TieneCambios = False
    mcol.Add c, "k" + Str(c.Id)
    
    Set AddColumnaFromString = c
    RaiseEvent ColumnCountChanged
End Function

Public Function Nueva(pName As String, pDescripcion As String, pErrorMessage As String, pWidth As Single, pFontSize As Integer, pFontName As String, pFontBold As Boolean, pFontItalic As Boolean, pMaxLength As Integer, pDataType As eTipoDato, pAlineacion As eAlineacion) As Campo
    Dim c As New Columna
    c.Id = GetUltimoId
    c.Nombre = pName
    c.Width = pWidth
    c.FontName = pFontName
    c.FontSize = pFontSize
    c.FontBold = pFontBold
    c.FontItalic = pFontItalic
    c.Descripcion = pDescripcion
    c.MaxLength = pMaxLength
    c.ErrorMessage = pErrorMessage
    c.DataType = pDataType
    c.Alineacion = pAlineacion
    c.Orden = Me.Count + 1
    mcol.Add c, "k" + Str(c.Id)
    
    Set Add = c
    vTieneCambios = True
    RaiseEvent ColumnCountChanged
End Function

Public Property Get Item(Id As Integer) As Campo
    Set Item = mcol.Item("k" + Str(Id))
End Property

Public Property Get ItemByOrden(pOrden As Integer) As Columna
    On Error Resume Next
    Dim c As Columna
    For Each c In Me
        If c.Orden = pOrden Then
            Set ItemByOrden = c
            Exit For
        End If
    Next
End Property

Public Property Get ItemByName(pName As String) As Columna
    On Error Resume Next
    Dim c As Columna
    For Each c In Me
        If LCase(c.Nombre) = LCase(pName) Then
            Set ItemByName = c
            Exit For
        End If
    Next
End Property
Friend Property Get TieneCambios() As Boolean
    If vTieneCambios Then
        TieneCambios = True
    Else
        Dim c As Columna
        For Each c In Me
            If c.TieneCambios Then
                TieneCambios = True
                Exit For
            End If
        Next
    End If
End Property

Public Sub Clear()
    Set mcol = Nothing
    Set mcol = New Collection
End Sub

Public Sub Remove(Id As Integer)
    Dim ordenEliminado As Integer
    ordenEliminado = Item(Id).Orden
    mcol.Remove "k" + Str(Id)
    Reordenar ordenEliminado
    RaiseEvent ColumnCountChanged
End Sub

Private Sub Reordenar(ordenEliminado As Integer)
    Dim c As Columna
    For Each c In Me
        If c.Orden > ordenEliminado Then c.Orden = c.Orden - 1
    Next
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mcol.[_NewEnum]
End Property

Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la colección.
    Count = mcol.Count
End Property

Friend Sub ResetValues()
    Dim c As Columna
    For Each c In Me
        'c.Valor = ""
    Next
End Sub

Private Function GetUltimoId() As Integer
    Dim c As Columna
    Dim max As Integer
    For Each c In Me
        If c.Id > max Then
            max = c.Id
        End If
    Next
    GetUltimoId = max + 1
    
End Function

Friend Function ColumnasToString() As String
    Dim c As Columna
    Dim aux As String
    For Each c In Me
        aux = aux + c.ColumnaToString
    Next
    ColumnasToString = aux
End Function
