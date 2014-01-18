VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Equipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mId As Long
Private mMovil As Movil
Private mDotacion As EmpleadoManager

Public Property Get Id() As Long
    On Error GoTo idErr
    
    Id = mId

    Exit Property
idErr:
    Call ErrorLog("Equipo", "Id Property Get", err.Description)
End Property

Friend Property Let Id(ByVal vNewValue As Long)
    On Error GoTo idErr
   
    mId = vNewValue

    Exit Property
idErr:
    Call ErrorLog("Equipo", "Id Property Let", err.Description)
End Property

Public Property Get Movil() As Movil
    Set Movil = mMovil
End Property

Public Property Set Movil(pMovil As Movil)
    Set mMovil = pMovil
End Property

Public Property Get Dotacion() As EmpleadoManager
    If mDotacion Is Nothing Then
        Set mDotacion = New EmpleadoManager
        mDotacion.CargarEmpleadosXEquipo Me.Id
    End If
    Set Dotacion = mDotacion
End Property

Public Property Set Dotacion(pDotacion As EmpleadoManager)
    Set mDotacion = pDotacion
End Property

Public Function GetProperty(pNombre As String) As String
    Dim aux As String
    aux = LCase(pNombre)
    Select Case aux
        Case "nombremovil"
            GetProperty = Movil.Nombre
        Case "pgdotacion"
            Dim e As Empleado
            Dim auxD As String
            For Each e In Dotacion
                auxD = auxD + e.NombreCompleto + "; "
            Next
            GetProperty = auxD
        End Select
End Function

Friend Sub Insert()
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
    
    Id = BD.getIdDisponible(tEquipo)
    
    params.Add "id", Id
    params.Add "idMovil", Movil.Id
    params.Add "eliminado", False
    
    trans.AddInsertCommand tEquipo, params
    
    Me.Dotacion.GuardarXEquipo trans, Me.Id
    trans.Execute
End Sub

Public Sub GuardarCambios()
    Dim trans As Transaccion
    Set trans = BD.Transaccion
    
    Dim params As New ParametrosManager
          
    params.Add "id", Id, True
    params.Add "idMovil", Movil.Id
        
    trans.AddUpdateCommand tEquipo, params
    
    Me.Dotacion.GuardarCambios trans, Me.Id
    trans.Execute
End Sub

'Indica si el equipo esta referenciado en alguna otra tabla
Public Function HasReferences() As Boolean
    Dim rs
    Set rs = BD.SelectSQL("select idEquipo from equipoxatencion where idEquipo=" + Str(mId))
    HasReferences = Not rs.EOF
    Set rs = Nothing
End Function

Friend Sub DarDeBaja()
    Dim c As New Parametro
    c.Nombre = "id"
    c.Valor = mId
    BD.BorradoLogico tEquipo, c
End Sub

Friend Sub Restaurar()
    Dim c As New Parametro
    c.Nombre = "id"
    c.Valor = mId
    BD.RestaurarElemento tEquipo, c
End Sub
