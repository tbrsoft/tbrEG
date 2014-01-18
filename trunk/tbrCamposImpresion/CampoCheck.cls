VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CampoCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private vTop As Single
Private vLeft As Single
Private vWidth As Single
Private vHeight As Single
Private vNombre As String
Private vId As Integer
Private vOrden As Integer
Private vValue As Boolean
Private vSimbolo As String
Private vDescripcion As String
Private vDefaultValue As Boolean

Private vTieneCambios As Boolean

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

Public Property Get Valor() As Boolean
    Valor = vValue
End Property

Public Property Get Simbolo() As String
    Simbolo = vSimbolo
End Property

Public Property Get DefaultValue() As Boolean
    DefaultValue = vDefaultValue
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

Public Property Let DefaultValue(pValue As Boolean)
    vDefaultValue = pValue
    vTieneCambios = True
End Property

Public Property Let Simbolo(pValue As String)
    vSimbolo = pValue
    vTieneCambios = True
End Property

Public Property Let Valor(pValue As Boolean)
    vValue = pValue
End Property

Friend Property Let TieneCambios(pValue As Boolean)
    vTieneCambios = pValue
End Property
'-----------------------funciones--------------------------------

Friend Function CampoCheckToString() As String
'convierte el contenido del campo en string para guardarlo en un archivo
    Dim aux As String
    aux = aux + "BEGIN CAMPOCHECK" + vbCrLf
    aux = aux + vbTab + "descripcion=" + Me.Descripcion + vbCrLf
    aux = aux + vbTab + "height=" + Str(Me.Height) + vbCrLf
    aux = aux + vbTab + "width=" + Str(Me.Width) + vbCrLf
    aux = aux + vbTab + "id=" + Str(Me.Id) + vbCrLf
    aux = aux + vbTab + "left=" + Str(Me.Left) + vbCrLf
    aux = aux + vbTab + "nombre=" + Me.Nombre + vbCrLf
    aux = aux + vbTab + "orden=" + Str(Me.Orden) + vbCrLf
    aux = aux + vbTab + "top=" + Str(Me.Top) + vbCrLf
    aux = aux + vbTab + "defaultvalue=" + Str(Me.DefaultValue) + vbCrLf
    aux = aux + vbTab + "simbolo=" + Me.Simbolo + vbCrLf
    aux = aux + "END CAMPOCHECK" + vbCrLf
    CampoCheckToString = aux
End Function
