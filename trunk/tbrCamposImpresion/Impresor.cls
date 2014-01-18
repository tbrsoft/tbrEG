VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Impresor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type EstadoImpresor
    DrawStyle As Integer
    FontName As String
    FontSize As Single
    Bold As Boolean
    Italic As Boolean
    ForeColor As Long
End Type

Public Event Error(pDescripcion As String)
Dim mDisp As Object

'Public Property Set Impresora(pmDisp As mDisp)
'    Set vmDisp = pmDisp
'End Property
'
'Public Property Get Impresora() As mDisp
'    Set Impresora = vmDisp
'End Property

Public Sub ImprimirFormulario(pFormulario As Formulario, Dispositivo As Object, Optional pEndDoc As Boolean = True)
On Error GoTo e
    If Not Dispositivo Is Nothing Then
        Set mDisp = Dispositivo
        Dim h As Hoja
        For Each h In pFormulario.Hojas
            ImprimirHoja h, , pEndDoc
        Next
    Else
        RaiseEvent Error("Debe setear una impresora.")
    End If
    Exit Sub
e:
    RaiseEvent Error(Err.Description)
End Sub

Public Sub ImprimirHoja(pHoja As Hoja, Optional Dispositivo As Object, Optional pEndDoc As Boolean = True)
'ver como solucionar el tipo de letra q no esta instalado
'On Error GoTo e
    If Not Dispositivo Is Nothing Then Set mDisp = Dispositivo
    
    Dim estado As EstadoImpresor
'    mDisp.Width = pHoja.Width
On Error Resume Next
    
    GuardarEstado Printer, estado

    'mDisp.TrackDefault = False
    Load Form1
    SelectForm pHoja, Form1.hwnd

    Dim c As Campo
    For Each c In pHoja.Campos
        Printer.CurrentX = c.Left
        Printer.CurrentY = c.Top
        Printer.Font.Name = c.FontName
        Printer.Font.Size = c.FontSize
        Printer.Print c.Valor
    Next
    RestaurarEstado Printer, estado

    Dim k As CampoCheck
    For Each k In pHoja.CamposCheck
        If k.Valor = True Then
            Printer.CurrentX = k.Left
            Printer.CurrentY = k.Top
            Printer.Print k.Simbolo
        End If
    Next

    GuardarEstado Printer, estado

    Dim l As Linea
    For Each l In pHoja.Lineas
        l.ImprimirLinea Printer
    Next

    RestaurarEstado Printer, estado

    Dim t As Tabla
    For Each t In pHoja.Tablas
        t.Dibujar Printer, t.Left, t.Top
    Next

    GuardarEstado Printer, estado

    Dim cl As CampoLabel
    For Each cl In pHoja.CamposLabel
        Printer.CurrentX = cl.Left
        Printer.CurrentY = cl.Top
        Printer.Font.Name = cl.FontName
        Printer.Font.Size = cl.FontSize
        Printer.Print cl.Text
    Next

    RestaurarEstado Printer, estado

    
    
    If pEndDoc Then Printer.EndDoc Else Printer.NewPage
    
    Unload Form1
    Exit Sub
e:
    RaiseEvent Error(Err.Description)
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
