VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HojasManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mcol As New Collection

Friend Sub AddHoja(pHoja As Hoja)
    mcol.Add pHoja, "k" + Str(pHoja.Numero)
End Sub

Friend Function AddHojaFromString(cadena As String) As Hoja
    Dim aux() As String
    Dim par() As String
    Dim misProp() As String
    cadena = Replace(cadena, "END HOJA", "")
    cadena = Replace(cadena, "END TABLA", "")
    
    aux = Split(cadena, "BEGIN ")
    
    'en aux(0) quedan las propiedades de la hoja
    misProp = Split(aux(0), vbCrLf)
    
    'cargo todas las propiedades en un diccionario
    Dim d
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(misProp) - 1
        If misProp(i) <> "" Then
            par = Split(misProp(i), "=")
            d.Add par(0), par(1)
        End If
    Next
    
    Dim h As New Hoja
    
    h.Numero = CInt(d("numero"))
    h.Orientacion = CInt(d("orientacion"))
    h.Height = CSng(d("height"))
    h.Width = CSng(d("width"))
    h.PathImagen = d("pathimagen")
    
    Dim tAux As Tabla
    'aca cargo los campos, tablas, etc...
    For i = 1 To UBound(aux)
        tipo = Left(aux(i), InStr(aux(i), vbCrLf) - 1)
        Select Case tipo
            Case "CAMPO"
                h.Campos.AddCampoFromString aux(i)
            Case "CAMPOCHECK"
                h.CamposCheck.AddCampoCheckFromString aux(i)
            Case "LINEA"
                h.Lineas.AddLineaFromString aux(i)
            Case "TABLA"
                'guardo la tabla q acabo de crear, lo q sigue son las columnas
                Set tAux = h.Tablas.AddTablaFromString(aux(i))
            Case "COLUMNA"
                'agrego las columnas a la tabla q cargue en el paso anterior
                tAux.Columnas.AddColumnaFromString aux(i)
            Case "CAMPOLABEL"
                h.CamposLabel.AddCampoFromString aux(i)
        End Select
    Next
    
    mcol.Add h, "k" + Str(h.Numero)
    
    Set Add = h
    
End Function

Public Function NuevaHoja(Optional pOrientacion As eOrientacion = eVertical) As Hoja
    Dim h As New Hoja
    h.Numero = mcol.Count + 1
    h.Orientacion = pOrientacion
    mcol.Add h, "k" + Str(h.Numero)
    Set NuevaHoja = h
End Function

Public Sub Eliminar(pNumero As Integer)
    mcol.Remove "k" + Str(pNumero)
End Sub

Public Property Get Item(pNumero As Integer) As Hoja
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = mcol.Item("k" + Str(pNumero))
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mcol.[_NewEnum]
End Property

Friend Sub EliminarValoresDeCampos()
    Dim h As Hoja
    For Each h In Me
        h.Campos.ResetValues
    Next
End Sub

Friend Function HojasTostring() As String
    Dim h As Hoja
    Dim aux As String
    For Each h In Me
        aux = aux + h.HojaToString + vbCrLf
    Next
    HojasTostring = aux
End Function
