VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CamposCheckManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mcol As New Collection
Private vTieneCambios As Boolean

Friend Function AddCampoCheckFromString(cadena As String) As CampoCheck
    Dim aux() As String
    Dim par() As String
    cadena = Replace(cadena, "END CAMPOCHECK", "")
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
    
    Dim c As New CampoCheck
    
    c.Id = d("id")
    c.Nombre = d("nombre")
    c.Height = CSng(d("height"))
    c.Width = CSng(d("width"))
    c.Top = CSng(d("top"))
    c.Left = CSng(d("left"))
    c.DefaultValue = CBool(d("defaultvalue"))
    c.Descripcion = d("descripcion")
    c.Simbolo = d("simbolo")
    c.Orden = CInt(d("orden"))
    c.TieneCambios = False
    mcol.Add c, "k" + Str(c.Id)
    
    Set AddCampoCheckFromString = c

End Function

Public Function Add(pId As Integer, pOrden As Integer, pName As String, pDescripcion As String, pWidth As Single, pHeight As Single, pTop As Single, pLeft As Single, pDefaultValue As Boolean, pSimbolo As String) As CampoCheck
    Dim c As New CampoCheck
    c.Id = pId
    c.Nombre = pName
    c.Height = pHeight
    c.Width = pWidth
    c.Top = pTop
    c.Left = pLeft
    c.Descripcion = pDescripcion
    c.DefaultValue = pDefaultValue
    c.Orden = pOrden
    c.Simbolo = pSimbolo
    mcol.Add c, "k" + Str(pId)
    
    Set Add = c
    vTieneCambios = True
End Function

Public Property Get Item(Id As Integer) As CampoCheck
Attribute Item.VB_UserMemId = 0
    Set Item = mcol.Item("k" + Str(Id))
End Property

Public Property Get ItemByOrden(pOrden As Integer) As CampoCheck
    On Error Resume Next
    Dim c As CampoCheck
    For Each c In Me
        If c.Orden = pOrden Then
            Set ItemByOrden = c
            Exit For
        End If
    Next
End Property

Public Property Get ItemByName(pName As String) As CampoCheck
    On Error Resume Next
    Dim c As CampoCheck
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
        Dim c As CampoCheck
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
    Dim c As CampoCheck
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
    Dim c As CampoCheck
    For Each c In Me
        c.Valor = c.DefaultValue
    Next
End Sub

Friend Function CamposCheckToString() As String
    Dim c As CampoCheck
    Dim aux As String
    For Each c In Me
        aux = aux + c.CampoCheckToString
    Next
    CamposCheckToString = aux
End Function

