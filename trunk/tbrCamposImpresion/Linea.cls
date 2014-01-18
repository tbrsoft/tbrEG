VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Linea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private vId As Integer
Private vX1 As Single
Private vX2 As Single
Private vY1 As Single
Private vY2 As Single
Private vColor As Long
Private vEstilo As Integer

Private vTieneCambios As Boolean

Public Property Get Id() As Integer
    Id = vId
End Property

Public Property Get X1() As Single
    X1 = vX1
End Property

Public Property Get X2() As Single
    X2 = vX2
End Property

Public Property Get Y1() As Single
    Y1 = vY1
End Property

Public Property Get Y2() As Single
    Y2 = vY2
End Property

Public Property Get Color() As Long
    Color = vColor
End Property

Public Property Get Estilo() As Integer
    Estilo = vEstilo
End Property

'--------------property let-----------------------
Friend Property Get TieneCambios() As Boolean
    TieneCambios = vTieneCambios
End Property

Public Property Let Id(pValue As Integer)
    vId = pValue
End Property

Public Property Let X1(pValue As Single)
    vX1 = pValue
    vTieneCambios = True
End Property

Public Property Let X2(pValue As Single)
    vX2 = pValue
    vTieneCambios = True
End Property

Public Property Let Y1(pValue As Single)
    vY1 = pValue
    vTieneCambios = True
End Property

Public Property Let Y2(pValue As Single)
    vY2 = pValue
    vTieneCambios = True
End Property

Public Property Let Color(pValue As Long)
    vColor = pValue
    vTieneCambios = True
End Property

Public Property Let Estilo(pValue As Integer)
    vEstilo = pValue
    vTieneCambios = True
End Property

Friend Property Let TieneCambios(pValue As Boolean)
    vTieneCambios = pValue
End Property

Public Function EstiloToString() As String
    EstiloToString = Choose(vEstilo + 1, "Transparent", "Solid", "Dash", "Dot", "DashDot", "DashDotDot")
End Function

Friend Function ImprimirLinea(pic As Object)
    pic.DrawStyle = IIf(Estilo = 0, 5, Estilo - 1)
    pic.Line (X1, Y1)-(X2, Y2), Color
End Function

Friend Function LineaToString() As String
    Dim aux As String
    aux = "BEGIN LINEA" + vbCrLf
    aux = aux + vbTab + "id=" + Trim(Str(Me.Id)) + vbCrLf
    aux = aux + vbTab + "x1=" + Trim(Str(Me.X1)) + vbCrLf
    aux = aux + vbTab + "x2=" + Trim(Str(Me.X2)) + vbCrLf
    aux = aux + vbTab + "y1=" + Trim(Str(Me.Y1)) + vbCrLf
    aux = aux + vbTab + "y2=" + Trim(Str(Me.Y2)) + vbCrLf
    aux = aux + vbTab + "color=" + Trim(Str(Me.Color)) + vbCrLf
    aux = aux + vbTab + "estilo=" + Trim(Str(Me.Estilo)) + vbCrLf
    aux = aux + "END LINEA"
    LineaToString = aux
End Function
