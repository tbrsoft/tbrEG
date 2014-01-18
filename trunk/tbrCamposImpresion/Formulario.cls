VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Formulario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private WithEvents vImpresor As Impresor
Attribute vImpresor.VB_VarHelpID = -1
Private vPath As String
Public Nombre As String
Private vHojas As New HojasManager

Public Property Get TieneCambios() As Boolean
    Dim aux As Boolean
    aux = False
    Dim h As Hoja
    For Each h In vHojas
        If h.TieneCambios Then
            aux = True
            Exit For
        End If
    Next
    TieneCambios = aux
End Property

Public Property Get path() As String
    path = vPath
End Property

Public Property Let path(pPath As String)
    vPath = pPath
End Property

Public Property Get Hojas() As HojasManager
    Set Hojas = vHojas
End Property

Public Sub Guardar(Optional pPath As String = "")
'si path = "" => es guardar.
'si path <>"" => es guardar como...
Dim file As New Archivo
If pPath = "" Then
    file.EscribirArchivo vPath, FormularioToString
Else
    vPath = pPath
    file.EscribirArchivo pPath, FormularioToString
End If
'restaurar tiene cambios a false
Dim h As Hoja
Dim c As Campo
Dim k As CampoCheck
Dim l As Linea
Dim t As Tabla
Dim cl As CampoLabel

For Each h In Me.Hojas
    For Each c In h.Campos
        c.TieneCambios = False
    Next
    For Each k In h.CamposCheck
        k.TieneCambios = False
    Next
    For Each l In h.Lineas
        l.TieneCambios = False
    Next
    For Each t In h.Tablas
        t.TieneCambios = False
    Next
    For Each cl In h.CamposLabel
        cl.TieneCambios = False
    Next
    
Next

End Sub

Public Sub Abrir(path As String)
'On Error GoTo e
    Dim file As New Archivo

    If path <> "" Then
        vPath = path
        CargarFormulario file.LeerArchivo(path)
    Else
    'error
    End If
    'cuando abrimos un archivo setea q todos los objetos tienen cambios, por eso
    'hay q ponerle q no hubo cambios
   

Exit Sub
e:

End Sub

Public Sub Imprimir(Dispositivo As Object, Optional pEndDoc As Boolean = True)

Set vImpresor = New Impresor
vImpresor.ImprimirFormulario Me, Dispositivo, pEndDoc

End Sub

Private Sub vImpresor_Error(pDescripcion As String)
    'hacer algo
End Sub

Private Function FormularioToString() As String
    Dim aux As String
    aux = "BEGIN FORMULARIO" + vbCrLf + vbTab
    aux = aux + "nombre=" + Me.Nombre + vbCrLf
    aux = aux + Hojas.HojasTostring
    aux = aux + "END FORMULARIO"
    FormularioToString = aux
   
End Function

Private Sub CargarFormulario(cadena As String)

Dim aux() As String
cadena = Replace(cadena, vbTab, "")
cadena = Replace(cadena, "END FORMULARIO", "")
aux = Split(cadena, "BEGIN HOJA")
'en aux(0) quedan las propiedades del formulario
Dim par() As String
aux(0) = Replace(aux(0), vbCrLf, "")
par = Split(aux(0), "=")
Me.Nombre = par(1)
For I = 1 To UBound(aux)
   Me.Hojas.AddHojaFromString aux(I)
Next
aux = Split(cadena, vbCrLf)

End Sub

Public Sub ResetearValoresDeCampos()
    Me.Hojas.EliminarValoresDeCampos
End Sub
