VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuerpoBomberos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mId As Long

Private mNombre As String

Private mResponsables As ResponsableCuerpoManager

Private mUnidades As UnidadCuerpoManager

Public Property Get Responsables() As ResponsableCuerpoManager
    If mResponsables Is Nothing Then
        Set mResponsables = New ResponsableCuerpoManager
        mResponsables.CargarPorCuerpo Me.Id
    End If
    Set Responsables = mResponsables
End Property

Public Property Set Responsables(pResponsables As ResponsableCuerpoManager)
    Set mResponsables = pResponsables
End Property

Public Property Get Unidades() As UnidadCuerpoManager
    If mUnidades Is Nothing Then
        Set mUnidades = New UnidadCuerpoManager
        mUnidades.CargarPorCuerpo Me.Id
    End If
    Set Unidades = mUnidades
End Property

Public Property Set Unidades(pUnidades As UnidadCuerpoManager)
    Set mUnidades = pUnidades
End Property

Public Property Get Id() As Long
    On Error GoTo idErr

    Id = mId

    Exit Property
idErr:
    Call ErrorLog("CuerpoBomberos", "Id Property Get", err.Description)
End Property

Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr

    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("CuerpoBomberos", "Id Property Let", err.Description)
End Property

Public Property Get Nombre() As String
    On Error GoTo nombreErr

    Nombre = mNombre

    Exit Property
nombreErr:
    Call ErrorLog("CuerpoBomberos", "Nombre Property Get", err.Description)
End Property

Public Property Let Nombre(ByVal vNewValue As String)
    On Error GoTo nombreErr

    '## Generated default body ...
    mNombre = vNewValue

    Exit Property
nombreErr:
    Call ErrorLog("CuerpoBomberos", "Nombre Property Let", err.Description)
End Property

Public Property Get GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "pgresponsables"
            GetProperty = Responsables.ResponsablesToString
        Case "pgunidades"
            GetProperty = Unidades.UnidadesToString
    End Select
End Property

Friend Sub Guardar()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId
    parametros.Add "Nombre", mNombre
    
    Responsables.Guardar Id
    Unidades.Guardar Id
    
    BD.Insertar tCuerpoBomberos, parametros
End Sub

Public Sub GuardarModificaciones()
    Dim parametros As New ParametrosManager
    parametros.Add "Id", mId, True
    parametros.Add "Nombre", mNombre
    'warning: implementar eliminar
    'adentro se da cuenta si tiene q hacer insert o update
    Responsables.Guardar Id
    Unidades.Guardar Id
    
    BD.Update tCuerpoBomberos, parametros
End Sub


