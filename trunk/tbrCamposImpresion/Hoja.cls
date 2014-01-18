VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Enum eOrientacion
    eHorizontal = 0
    eVertical = 1
End Enum

Private vOrientacion As eOrientacion
Private vCampos As New CamposManager
Private vTablas As New TablasManager
Private vLineas As New LineasManager
Private vCamposCheck As New CamposCheckManager
Private vCamposLabel As New CamposLabelManager

Private vNumero As Integer
Private vPathImagen As String
Private vWidth As Single
Private vHeight As Single

Public Property Get TieneCambios() As Boolean
    TieneCambios = vCampos.TieneCambios Or vCamposCheck.TieneCambios Or vLineas.TieneCambios Or vTablas.TieneCambios Or vCamposLabel.TieneCambios
End Property

Public Property Get Width() As Single
    Width = vWidth
End Property

Public Property Get Height() As Single
    Height = vHeight
End Property

Public Property Let Width(pValue As Single)
    vWidth = pValue
End Property

Public Property Let Height(pValue As Single)
    vHeight = pValue
End Property

Public Property Get PathImagen() As String
    PathImagen = vPathImagen
End Property

Public Property Let PathImagen(pPathImagen As String)
    vPathImagen = pPathImagen
End Property

Public Property Get Orientacion() As eOrientacion
    Orientacion = vOrientacion
End Property

Public Property Let Orientacion(pOrientacion As eOrientacion)
    vOrientacion = pOrientacion
End Property

Public Property Get Campos() As CamposManager
    Set Campos = vCampos
End Property

Public Property Get Tablas() As TablasManager
    Set Tablas = vTablas
End Property

Public Property Get Lineas() As LineasManager
    Set Lineas = vLineas
End Property

Public Property Get CamposCheck() As CamposCheckManager
    Set CamposCheck = vCamposCheck
End Property

Public Property Get CamposLabel() As CamposLabelManager
    Set CamposLabel = vCamposLabel
End Property

Public Property Get Numero() As Integer
    Numero = vNumero
End Property

Public Property Let Numero(pNumero As Integer)
    vNumero = pNumero
End Property

Public Sub Imprimir()
Dim impr As New Impresor
impr.ImprimirHoja Me
End Sub

Friend Function HojaToString() As String
    Dim aux As String
    
    aux = aux + vbTab + "BEGIN HOJA" + vbCrLf + vbTab + vbTab
    aux = aux + "height=" + Str(Me.Height) + vbCrLf + vbTab + vbTab
    aux = aux + "width=" + Str(Me.Width) + vbCrLf + vbTab + vbTab
    aux = aux + "numero=" + Str(Me.Numero) + vbCrLf + vbTab + vbTab
    aux = aux + "orientacion=" + Str(Me.Orientacion) + vbCrLf + vbTab + vbTab
    aux = aux + "pathimagen=" + Me.PathImagen + vbCrLf
    aux = aux + Campos.CamposToString
    aux = aux + Tablas.TablasToString
    aux = aux + Lineas.LineasToString
    aux = aux + CamposCheck.CamposCheckToString
    aux = aux + CamposLabel.CamposToString
    aux = aux + "END HOJA"
    
    HojaToString = aux
End Function
