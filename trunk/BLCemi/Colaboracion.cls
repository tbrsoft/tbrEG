VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Colaboracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim mCuerpo As CuerpoBomberos
Dim mResponsable As ResponsableCuerpo
Dim mUnidad As UnidadCuerpo
Dim mCantidadEfectivos As Integer
Dim mObservaciones As String
Dim mId As Long

Public Property Get Unidad() As UnidadCuerpo
    Set Unidad = mUnidad
End Property

Public Property Set Unidad(vNewValue As UnidadCuerpo)
    Set mUnidad = vNewValue
End Property

Public Property Get Responsable() As ResponsableCuerpo
    Set Responsable = mResponsable
End Property

Public Property Set Responsable(vNewValue As ResponsableCuerpo)
    Set mResponsable = vNewValue
End Property

Public Property Get Cuerpo() As CuerpoBomberos
    Set Cuerpo = mCuerpo
End Property

Public Property Set Cuerpo(vNewValue As CuerpoBomberos)
    Set mCuerpo = vNewValue
End Property

Public Property Get CantidadEfectivos() As Integer
    CantidadEfectivos = mCantidadEfectivos
End Property

Public Property Let CantidadEfectivos(ByVal vNewValue As Integer)
    mCantidadEfectivos = vNewValue
End Property

Public Property Get Id() As Long
    Id = mId
End Property

Public Property Let Id(ByVal vNewValue As Long)
    mId = vNewValue
End Property

Public Property Get Observaciones() As String
    Observaciones = mObservaciones
End Property

Public Property Let Observaciones(ByVal vNewValue As String)
    mObservaciones = vNewValue
End Property

Public Property Get GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "pgcuerpo"
            If Cuerpo Is Nothing Then
                GetProperty = "-Sin asignar-"
            Else
                GetProperty = Cuerpo.Nombre
            End If
        Case "pgresponsable"
            If Responsable Is Nothing Then
                GetProperty = "-Sin asignar-"
            Else
                GetProperty = Responsable.Nombre
            End If
        Case "pgunidad"
            If Unidad Is Nothing Then
                GetProperty = "-Sin asignar-"
            Else
                GetProperty = Unidad.Nombre
            End If
    End Select
End Property

Public Sub Guardar(pIdSiniestro As Long)
On Error GoTo errman
    Dim parametros As New ParametrosManager
    
    If Not mUnidad Is Nothing Then parametros.Add "idUnidad", mUnidad.Id
    If Not mResponsable Is Nothing Then parametros.Add "idResponsable", mResponsable.Id
    If Not mCuerpo Is Nothing Then parametros.Add "idCuerpo", mCuerpo.Id
    parametros.Add "Observaciones", mObservaciones
    parametros.Add "CantidadEfectivos", mCantidadEfectivos
        
    If mId = 0 Then
        mId = BD.getIdDisponible(tColaboracionCuerpo)
        parametros.Add "IdSiniestro", pIdSiniestro
        parametros.Add "Id", mId
        BD.Insertar tColaboracionCuerpo, parametros
    Else
        parametros.Add "Id", mId, True
        BD.Update tColaboracionCuerpo, parametros
    End If
    
    Exit Sub
errman:
    ErrorLog "Colaboracion", "Guardar", Err.Description
End Sub
