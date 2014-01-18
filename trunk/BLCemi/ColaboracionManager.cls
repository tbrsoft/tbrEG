VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColaboracionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mCol As Collection

'Public Sub Remove(pId As Long)
'    On Error GoTo RemoveErr
'
'    '## Your code goes here ...
'
'    Exit Sub
'RemoveErr:
'    Call ErrorLog(MyUnhandledError, "Remove Sub")
'End Sub

Public Sub AddItem(pColaboracion As Colaboracion)
    On Error GoTo AddItemErr

    mCol.Add pColaboracion ', "k" + Str(pColaboracion.Id)

    Exit Sub
AddItemErr:
    ErrorLog "ColaboracionManager", "AddItem", err.Description
End Sub

Friend Function Add(pId As Long, pCuerpos As CuerpoBomberos, pResponsable As ResponsableCuerpo, pUnidad As UnidadCuerpo, pCantidadEfectivos As Integer, pObservaciones As String) As Colaboracion
    Dim c As New Colaboracion
    c.Id = pId
    c.CantidadEfectivos = pCantidadEfectivos
    Set c.Cuerpo = pCuerpos
    c.Observaciones = pObservaciones
    Set c.Responsable = pResponsable
    Set c.Unidad = pUnidad
    
    AddItem c
    Set Add = c
    Set c = Nothing
End Function

Public Sub CargarPorSiniestro(pIdSiniestro As Long)
    On Error GoTo errMaq
    
    Dim mCuerpo As CuerpoBomberos
    Dim mUnidad As UnidadCuerpo
    Dim mResponsable As ResponsableCuerpo
        
    Dim rs
    Set rs = BD.SelectSQL("Select * from ColaboracionCuerpo where idSiniestro=" + Str(pIdSiniestro))
    While Not rs.EOF
            
        Set mCuerpo = CuerposDeBomberosLocal.Item(rs!idCuerpo)
            
        If Not IsNull(rs!idResponsable) Then
            Set mResponsable = mCuerpo.Responsables.Item(rs!idResponsable)
        Else
            Set mResponsable = Nothing
        End If
        If Not IsNull(rs!idUnidad) Then
            Set mUnidad = mCuerpo.Unidades.ItemById(rs!idUnidad)
         Else
            Set mUnidad = Nothing
        End If
        
        Add rs!Id, mCuerpo, mResponsable, mUnidad, rs!CantidadEfectivos, rs!Observaciones
        rs.MoveNext
        
    Wend
    Exit Sub
errMaq:
    ErrorLog "ColaboracionManager", "CargarPorCuerpo ", err.Description

End Sub

'falta cargar uno solo

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

'lo desabilito porq tengo q agregar items con id=0...
'Public Property Get Item(pId As Long) As Colaboracion
'    On Error Resume Next
'    Set Item = mCol.Item("k" + Str(pId))
'End Property

'Public Function Nuevo(pNombre As String) As Colaboracion
'    Dim c As New Colaboracion
'    c.Id = BD.getIdDisponible(tColaboracion)
'    c.Nombre = pNombre
'    c.Guardar
'    Me.AddItem c
'    Set Nuevo = c
'End Function


Public Sub Guardar(pIdSiniestro As Long)
    Dim c As Colaboracion
    For Each c In Me
        c.Guardar pIdSiniestro
    Next
End Sub
