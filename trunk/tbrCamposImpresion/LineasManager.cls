VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineasManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mcol As New Collection
Private vTieneCambios As Boolean

Friend Function AddLineaFromString(cadena As String) As Linea
    Dim aux() As String
    Dim par() As String
    cadena = Replace(cadena, "END LINEA", "")
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
    
    Dim c As New Linea
    
    c.Id = d("id")
    c.X1 = d("x1")
    c.X2 = d("x2")
    c.Y1 = d("y1")
    c.Y2 = d("y2")
    c.Color = CLng(d("color"))
    c.Estilo = CInt(d("estilo"))
    c.TieneCambios = False
    mcol.Add c, "k" + Str(c.Id)
    
    Set AddLineaFromString = c

End Function

Public Function Add(pId As Integer, pX1 As Single, pX2 As Single, pY1 As Single, pY2 As Single, pEstilo As Integer, pColor As Long) As Linea
    Dim c As New Linea
    c.Id = pId
    c.X1 = pX1
    c.X2 = pX2
    c.Y1 = pY1
    c.Y2 = pY2
    c.Color = pColor
    c.Estilo = pEstilo
    mcol.Add c, "k" + Str(pId)
    
    Set Add = c
    vTieneCambios = True
End Function

Public Property Get Item(Id As Integer) As Linea
    Set Item = mcol.Item("k" + Str(Id))
End Property

'Public Property Get ItemByOrden(pOrden As Integer) As Linea
'    On Error Resume Next
'    Dim c As Linea
'    For Each c In Me
'        If c.Orden = pOrden Then
'            Set ItemByOrden = c
'            Exit For
'        End If
'    Next
'End Property

Friend Property Get TieneCambios() As Boolean
    If vTieneCambios Then
        TieneCambios = True
    Else
        Dim c As Linea
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
    mcol.Remove "k" + Str(Id)
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

Friend Function LineasToString() As String
    Dim l As Linea
    Dim aux As String
    For Each l In Me
        aux = aux + l.LineaToString + vbCrLf
    Next
    LineasToString = aux
End Function

