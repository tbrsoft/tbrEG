VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mEncabezados As ControlesPOO.LVCEncabezadoManager
Private mTitulo As String
Private mDescripcion As String
Private mSQL As String
Private mMensajeError As String
Private mParametros As LParameterManager
Private mPath As String

Public Function GetColeccion() As ColeccionGenerica
    Dim colG As ColeccionGenerica
    If mMensajeError = "" Then 'si se pudo cargar...
        Dim rs As Object
        If mParametros Is Nothing Then
            Set rs = GBL.ExecuteSQL(mSQL)
        Else
            'esta funcion incluye los parametros
            Set rs = GBL.ExecuteSQL(ConvertSQL)
        End If
        Set colG = New ColeccionGenerica
        colG.Fill rs
    End If
    Set GetColeccion = colG
End Function


Public Sub Load(pPath As String)
    On Error GoTo errman
    mPath = mPath
    mTitulo = Leer_Ini(pPath, "Listado", "Titulo", "")
    mDescripcion = Leer_Ini(pPath, "Listado", "Descripcion", "")
     Dim aux As String
    aux = Leer_Ini(pPath, "Listado", "SQL", "")
    mSQL = unSwap(Mid(aux, 2, Len(aux) - 2))
    Set mEncabezados = GetEncabezados(pPath)
    Set mParametros = GetParametros(pPath)
    Exit Sub
errman:
    mMensajeError = Err.Description
End Sub

'son todas property get
Public Property Get MensajeError() As String
    MensajeError = mMensajeError
End Property

Public Property Get Encabezados() As ControlesPOO.LVCEncabezadoManager
    Set Encabezados = mEncabezados
End Property

Public Property Get Titulo() As String
    Titulo = mTitulo
End Property

Public Property Get SQL() As String
    SQL = mSQL
End Property

Public Property Get Path() As String
    Path = mPath
End Property

Public Property Get Parametros() As LParameterManager
    Set Parametros = mParametros
End Property

Private Function ConvertSQL() As String
Dim lp As LParameter
Dim aux As String
aux = mSQL
For Each lp In mParametros
    aux = Replace(aux, "@" + lp.Nombre, lp.Valor)
Next
ConvertSQL = aux
End Function

'encriptacion
Function swap(cadena As String, Optional seed As Long = 127986) As String

Dim longArray() As Long
Dim largo As Long
Dim temp As String
Dim aux As String
Dim cadenaArray() As String

largo = Len(cadena)

ReDim longArray(largo * 3)
ReDim cadenaArray(largo)

For I = 1 To largo

cadenaArray(I) = Mid(cadena, I, 1)


Next
Rnd -1
Randomize seed
For I = 1 To largo * 3

    longArray(I) = Int(largo * Rnd) + 1

Next

For I = 1 To largo

    temp = cadenaArray(I)
    cadenaArray(I) = cadenaArray(longArray(I))
    cadenaArray(longArray(I)) = temp
Next



For I = 1 To largo
aux = aux + cadenaArray(I)
'aux = Trim(aux + cadenaArray(i))

Next
swap = aux


End Function


Function unSwap(cadena As String, Optional seed As Long = 127986) As String
Dim longArray() As Long
Dim largo As Long
Dim temp As String
Dim aux As String
Dim cadenaArray() As String

largo = Len(cadena)

ReDim longArray(largo * 3)
ReDim cadenaArray(largo)

For I = 1 To largo

cadenaArray(I) = Mid(cadena, I, 1)


Next
Rnd -1
Randomize seed

For I = 1 To largo * 3

    longArray(I) = Int(largo * Rnd) + 1

Next

For I = largo To 1 Step -1

    temp = cadenaArray(I)
    cadenaArray(I) = cadenaArray(longArray(I))
    cadenaArray(longArray(I)) = temp
Next



For I = 1 To largo
aux = aux + cadenaArray(I)
'aux = Trim(aux + cadenaArray(i))

Next
unSwap = aux


End Function
