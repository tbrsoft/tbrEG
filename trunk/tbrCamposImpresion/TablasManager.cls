VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TablasManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mcol As New Collection
Private vTieneCambios As Boolean

Friend Function AddTablaFromString(cadena As String) As Tabla
    Dim aux() As String
    Dim par() As String
    cadena = Replace(cadena, "END TABLA", "")
    'aca tengo todas las propiedades
    aux = Split(cadena, vbCrLf)
    'cargo todas las propiedades en un diccionario
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    For I = 1 To UBound(aux) - 1
        If aux(I) <> "" Then
            par = Split(aux(I), "=")
            d.Add par(0), par(1)
        End If
    Next
    
    Dim c As New Tabla
    'revisar
    c.Id = d("id")
    c.Nombre = d("nombre")
    c.Height = CSng(d("height"))
    c.Width = CSng(d("width"))
    c.Top = CSng(d("top"))
    c.Left = CSng(d("left"))
    c.Rows = CInt(d("rows"))
    c.FontName = d("fontname")
    c.FontSize = CInt(d("fontsize"))
    c.Descripcion = d("descripcion")
    c.ErrorMessage = d("errormessage")
    c.Orden = CInt(d("orden"))
    c.EstiloTabla = CInt(d("tablestyle"))
    c.TieneCambios = False
    mcol.Add c, "k" + Str(c.Id)
    
    Set AddTablaFromString = c

End Function

Public Function Add(pId As Integer, pOrden As Integer, pName As String, pDescripcion As String, pErrorMessage As String, pWidth As Single, pHeight As Single, pTop As Single, pLeft As Single, pRows As Integer, pFontSize As Integer, pFontName As String, pTableStyle As Integer) As Tabla
    Dim c As New Tabla
    c.Id = pId
    c.Nombre = pName
    c.Height = pHeight
    c.Width = pWidth
    c.Top = pTop
    c.Left = pLeft
    'c.Columns = pColumns
    c.Rows = pRows
    c.FontName = pFontName
    c.FontSize = pFontSize
    c.Descripcion = pDescripcion
    c.ErrorMessage = pErrorMessage
    c.EstiloTabla = pTableStyle
    c.Orden = pOrden
    mcol.Add c, "k" + Str(pId)
    
    Set Add = c
    vTieneCambios = True
End Function

Public Property Get Item(Id As Integer) As Tabla
    Set Item = mcol.Item("k" + Str(Id))
End Property

Public Property Get ItemByOrden(pOrden As Integer) As Tabla
    On Error Resume Next
    Dim c As Tabla
    For Each c In Me
        If c.Orden = pOrden Then
            Set ItemByOrden = c
            Exit For
        End If
    Next
End Property

Public Property Get ItemByName(pName As String) As Tabla
    On Error Resume Next
    Dim c As Tabla
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
        Dim c As Tabla
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
End Sub

Private Sub Reordenar(ordenEliminado As Integer)
    Dim c As Tabla
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
    Dim c As Tabla
    For Each c In Me
        'resetear
        'c.Valor = ""
    Next
End Sub

Friend Function TablasToString() As String
    Dim t As Tabla
    Dim aux As String
    For Each t In Me
        aux = aux + t.TablaToString + vbCrLf
    Next
    TablasToString = aux
End Function
