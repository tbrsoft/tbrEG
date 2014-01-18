VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CampoLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
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
Private vText As String

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

Public Property Get FontSize() As Integer
    FontSize = vFontSize
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

Public Property Get Id() As Integer
    Id = vId
End Property

Public Property Get Text() As String
   Text = vText
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

Public Property Let FontSize(pValue As Integer)
    vFontSize = pValue
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

Public Property Let Id(pValue As Integer)
    vId = pValue
    vTieneCambios = True
End Property

Public Property Let Text(pValue As String)
    vText = pValue
    vTieneCambios = True
End Property

Friend Property Let TieneCambios(pValue As Boolean)
    vTieneCambios = pValue
End Property
'-----------------------funciones--------------------------------

Friend Function CampoLabelToString() As String
'convierte el contenido del campo en string para guardarlo en un archivo
    Dim aux As String
    aux = aux + "BEGIN CAMPOLABEL" + vbCrLf
    aux = aux + vbTab + "descripcion=" + Me.Descripcion + vbCrLf
    aux = aux + vbTab + "fontname=" + Me.FontName + vbCrLf
    aux = aux + vbTab + "fontsize=" + Str(Me.FontSize) + vbCrLf
    aux = aux + vbTab + "height=" + Str(Me.Height) + vbCrLf
    aux = aux + vbTab + "width=" + Str(Me.Width) + vbCrLf
    aux = aux + vbTab + "id=" + Str(Me.Id) + vbCrLf
    aux = aux + vbTab + "left=" + Str(Me.Left) + vbCrLf
    aux = aux + vbTab + "nombre=" + Me.Nombre + vbCrLf
    aux = aux + vbTab + "top=" + Str(Me.Top) + vbCrLf
    aux = aux + vbTab + "text=" + Me.Text + vbCrLf
    aux = aux + "END CAMPOLABEL" + vbCrLf
    CampoLabelToString = aux
End Function

