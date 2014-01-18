VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPermisos 
   Caption         =   "Propiedades para el uso del Sistema"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   6795
   ScaleWidth      =   4185
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResetPass 
      Caption         =   "Nuevo Password (En caso de olvido)"
      Height          =   300
      Left            =   1080
      TabIndex        =   4
      Top             =   480
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   1935
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4048
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox txtLogin 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblPermisos 
      AutoSize        =   -1  'True
      Caption         =   "Permisos"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   630
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Login:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPermisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event PropiedadesModificadas(pLogin As String, pPass As String)

Dim mLogin As String
Dim mPass As String
Dim mPermisos As blcemi.PermisoManager

Dim mPuedeModificarPermisos As Boolean
Dim bAceptarCambios As Boolean

Private Sub cmdAceptar_Click()
bAceptarCambios = True
Me.Hide
End Sub

Private Sub cmdCancelar_Click()
bAceptarCambios = False
Me.Hide
End Sub

Private Sub cmdResetPass_Click()
cmdResetPass.Visible = False
txtPass = ""
txtPass.SetFocus
End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon

Select Case CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento

Case eModoFuncionamiento.eMFEmergencia:

tvw.Nodes.Add(, , "sistema", "Sistema").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "recepcion", "Recepcion").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "administracion", "Administracion").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "mantenimiento", "Mantenimiento").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "configuracion", "Configuracion").Expanded = True

tvw.Nodes.Add "recepcion", tvwChild, "k 25", "Consultar Atencion"
tvw.Nodes.Add "k 25", tvwChild, "k 26", "Registrar Atencion"
tvw.Nodes.Add "k 25", tvwChild, "k 28", "Modificar Atencion"
tvw.Nodes.Add "k 25", tvwChild, "k 27", "Anular Atencion"

tvw.Nodes.Add "recepcion", tvwChild, "k 29", "Consultar Dotacion"
tvw.Nodes.Add "k 29", tvwChild, "k 30", "Registrar Dotacion"
'agregado 09/01/2009
tvw.Nodes.Add "k 29", tvwChild, "k 80", "Eliminar Dotacion"
tvw.Nodes.Add "k 29", tvwChild, "k 31", "Modificar Dotacion"

tvw.Nodes.Add "recepcion", tvwChild, "k 32", "Consultar Movil"
tvw.Nodes.Add "k 32", tvwChild, "k 33", "Alta Movil"
tvw.Nodes.Add "k 32", tvwChild, "k 34", "Eliminar Movil"
tvw.Nodes.Add "k 32", tvwChild, "k 35", "Modificar Movil"

tvw.Nodes.Add "administracion", tvwChild, "k 79", "Setear Numero Recibo"
tvw.Nodes.Add "administracion", tvwChild, "k 84", "Registrar Guardia Empleado"
tvw.Nodes.Add "administracion", tvwChild, "k 85", "Liquidar Sueldo a Empleados"
tvw.Nodes.Add "administracion", tvwChild, "k 86", "Liquidar Servicios a Empresas"
tvw.Nodes.Add "administracion", tvwChild, "k 89", "Consultar Liquidaciones a Empleados"
tvw.Nodes.Add "administracion", tvwChild, "k 90", "Consultar Liquidaciones a Empresas"
tvw.Nodes.Add "administracion", tvwChild, "k 91", "Ver/Modificar Información Contable de Atención"


tvw.Nodes.Add "administracion", tvwChild, "k 1", "Consultar Afiliado"
tvw.Nodes.Add "k 1", tvwChild, "k 2", "Alta Afiliado"
tvw.Nodes.Add "k 1", tvwChild, "k 3", "Modificar Afiliado"
tvw.Nodes.Add "k 1", tvwChild, "k 4", "Eliminar Afiliado"

tvw.Nodes.Add "administracion", tvwChild, "k 5", "Consultar Area Protegida"
tvw.Nodes.Add "k 5", tvwChild, "k 6", "Alta Area Protegida"
tvw.Nodes.Add "k 5", tvwChild, "k 8", "Modificar Area Protegida"
tvw.Nodes.Add "k 5", tvwChild, "k 7", "Eliminar Area Protegida"

tvw.Nodes.Add "administracion", tvwChild, "k 14", "Consultar Obra Social"
tvw.Nodes.Add "k 14", tvwChild, "k 15", "Alta Obra Social"
tvw.Nodes.Add "k 14", tvwChild, "k 17", "Modificar Obra Social"
tvw.Nodes.Add "k 14", tvwChild, "k 16", "Eliminar Obra Social"

tvw.Nodes.Add "administracion", tvwChild, "k 21", "Consultar Servicio de Emergencia"
tvw.Nodes.Add "k 21", tvwChild, "k 22", "Alta Servicio de Emergencia"
tvw.Nodes.Add "k 21", tvwChild, "k 24", "Modificar Servicio de Emergencia"
tvw.Nodes.Add "k 21", tvwChild, "k 23", "Eliminar Servicio de Emergencia"

tvw.Nodes.Add "administracion", tvwChild, "k 9", "Consultar Empleado"
tvw.Nodes.Add "k 9", tvwChild, "k 10", "Alta Empleado"
tvw.Nodes.Add "k 9", tvwChild, "k 11", "Eliminar Empleado"
tvw.Nodes.Add "k 9", tvwChild, "k 12", "Modificar Empleado"
'tvw.Nodes.Add "k 9", tvwChild, "k 13", "Asignar Permisos Empleado"

tvw.Nodes.Add "administracion", tvwChild, "k 18", "Consultar Pagos"
tvw.Nodes.Add "k 18", tvwChild, "k 19", "Registrar Pago"
tvw.Nodes.Add "k 18", tvwChild, "k 20", "Registrar Devolucion Recibos Anulados"
'tvw.Nodes.Add "k 18", tvwChild, "k 20", "Anular Pago" ver anular recibo, q no es lo mismo


'a 1 porq no tengo consulta para estos
tvw.Nodes.Add "mantenimiento", tvwChild, "a 1", "Medicamentos"
tvw.Nodes.Add "a 1", tvwChild, "k 39", "Alta Medicamento"
tvw.Nodes.Add "a 1", tvwChild, "k 40", "Eliminar Medicamento"
tvw.Nodes.Add "a 1", tvwChild, "k 41", "Modificar Medicamento"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 2", "Enfermedades"
tvw.Nodes.Add "a 2", tvwChild, "k 42", "Alta Enfermedad"
tvw.Nodes.Add "a 2", tvwChild, "k 43", "Eliminar Enfermedad"
tvw.Nodes.Add "a 2", tvwChild, "k 44", "Modificar Enfermedad"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 3", "Alergias"
tvw.Nodes.Add "a 3", tvwChild, "k 45", "Alta Alergia"
tvw.Nodes.Add "a 3", tvwChild, "k 46", "Eliminar Alergia"
tvw.Nodes.Add "a 3", tvwChild, "k 47", "Modificar Alergia"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 9", "Tipos de Codigo de Emergencia"
tvw.Nodes.Add "a 9", tvwChild, "k 88", "Alta Codigo de Emergencia"
tvw.Nodes.Add "a 9", tvwChild, "k 81", "Eliminar Codigo de Emergencia"
tvw.Nodes.Add "a 9", tvwChild, "k 82", "Modificar Codigo de Emergencia"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 4", "Tipos de Telefono"
tvw.Nodes.Add "a 4", tvwChild, "k 48", "Alta Tipo de Telefono"
tvw.Nodes.Add "a 4", tvwChild, "k 49", "Eliminar Tipo de Telefono"
tvw.Nodes.Add "a 4", tvwChild, "k 50", "Modificar Tipo de Telefono"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 5", "Ocupaciones"
tvw.Nodes.Add "a 5", tvwChild, "k 51", "Alta Ocupacion"
tvw.Nodes.Add "a 5", tvwChild, "k 52", "Eliminar Ocupacion"
tvw.Nodes.Add "a 5", tvwChild, "k 53", "Modificar Ocupacion"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 6", "Parentezcos"
tvw.Nodes.Add "a 6", tvwChild, "k 54", "Alta Parentezco"
tvw.Nodes.Add "a 6", tvwChild, "k 55", "Eliminar Parentezco"
tvw.Nodes.Add "a 6", tvwChild, "k 56", "Modificar Parentezco"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 7", "Cargos"
tvw.Nodes.Add "a 7", tvwChild, "k 36", "Alta Cargo"
tvw.Nodes.Add "a 7", tvwChild, "k 37", "Eliminar Cargo"
tvw.Nodes.Add "a 7", tvwChild, "k 38", "Modificar Cargo"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 8", "Ciudades y Barrios"
tvw.Nodes.Add "a 8", tvwChild, "k 60", "Alta Ciudad y Barrios"
tvw.Nodes.Add "a 8", tvwChild, "k 61", "Eliminar Ciudad y Barrios"
tvw.Nodes.Add "a 8", tvwChild, "k 62", "Modificar Ciudad y Barrios"

tvw.Nodes.Add "configuracion", tvwChild, "k 75", "Configurar Apariencia"
tvw.Nodes.Add "configuracion", tvwChild, "k 76", "Configurar Red"
tvw.Nodes.Add "configuracion", tvwChild, "k 77", "Configurar Base de Datos"
tvw.Nodes.Add "configuracion", tvwChild, "k 78", "Configurar Comportamiento"
tvw.Nodes.Add "configuracion", tvwChild, "k 83", "Configurar Codigos"
tvw.Nodes.Add "configuracion", tvwChild, "k 87", "Configurar Predeterminados"

Case eModoFuncionamiento.eMFBomberos

'--------------------------------------------------------------------------------------------
'                                       BOMBEROS
'--------------------------------------------------------------------------------------------

tvw.Nodes.Add(, , "sistema", "Sistema").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "recepcion", "Recepcion").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "administracion", "Administracion").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "mantenimiento", "Mantenimiento").Expanded = True
tvw.Nodes.Add("sistema", tvwChild, "configuracion", "Configuracion").Expanded = True

tvw.Nodes.Add "recepcion", tvwChild, "k 25", "Consultar Siniestro"
tvw.Nodes.Add "k 25", tvwChild, "k 26", "Registrar Siniestro"
tvw.Nodes.Add "k 25", tvwChild, "k 28", "Modificar Siniestro"
'ver, creo q no se puede tvw.Nodes.Add "k 25", tvwChild, "k 27", "Anular Atencion"

tvw.Nodes.Add "recepcion", tvwChild, "k 29", "Consultar Dotacion"
tvw.Nodes.Add "k 29", tvwChild, "k 30", "Registrar Dotacion"
'agregado 09/01/2009
tvw.Nodes.Add "k 29", tvwChild, "k 80", "Eliminar Dotacion"
tvw.Nodes.Add "k 29", tvwChild, "k 31", "Modificar Dotacion"

tvw.Nodes.Add "recepcion", tvwChild, "k 32", "Consultar Movil"
tvw.Nodes.Add "k 32", tvwChild, "k 33", "Alta Movil"
tvw.Nodes.Add "k 32", tvwChild, "k 34", "Eliminar Movil"
tvw.Nodes.Add "k 32", tvwChild, "k 35", "Modificar Movil"

tvw.Nodes.Add "administracion", tvwChild, "k 84", "Registrar Guardia Empleado"
tvw.Nodes.Add "administracion", tvwChild, "k 85", "Liquidar Sueldo a Empleados"
tvw.Nodes.Add "administracion", tvwChild, "k 89", "Consultar Liquidaciones a Empleados"

tvw.Nodes.Add "administracion", tvwChild, "k 9", "Consultar Empleado"
tvw.Nodes.Add "k 9", tvwChild, "k 10", "Alta Empleado"
tvw.Nodes.Add "k 9", tvwChild, "k 11", "Eliminar Empleado"
tvw.Nodes.Add "k 9", tvwChild, "k 12", "Modificar Empleado"
'tvw.Nodes.Add "k 9", tvwChild, "k 13", "Asignar Permisos Empleado"

'ver si hacen falta
tvw.Nodes.Add "administracion", tvwChild, "k 18", "Consultar Pagos"
tvw.Nodes.Add "k 18", tvwChild, "k 19", "Registrar Pago"
tvw.Nodes.Add "k 18", tvwChild, "k 20", "Registrar Devolucion Recibos Anulados"
'tvw.Nodes.Add "k 18", tvwChild, "k 20", "Anular Pago" ver anular recibo, q no es lo mismo


'a 1 porq no tengo consulta para estos

tvw.Nodes.Add "mantenimiento", tvwChild, "a 4", "Tipos de Telefono"
tvw.Nodes.Add "a 4", tvwChild, "k 48", "Alta Tipo de Telefono"
tvw.Nodes.Add "a 4", tvwChild, "k 49", "Eliminar Tipo de Telefono"
tvw.Nodes.Add "a 4", tvwChild, "k 50", "Modificar Tipo de Telefono"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 7", "Cargos"
tvw.Nodes.Add "a 7", tvwChild, "k 36", "Alta Cargo"
tvw.Nodes.Add "a 7", tvwChild, "k 37", "Eliminar Cargo"
tvw.Nodes.Add "a 7", tvwChild, "k 38", "Modificar Cargo"

tvw.Nodes.Add "mantenimiento", tvwChild, "a 8", "Ciudades y Barrios"
tvw.Nodes.Add "a 8", tvwChild, "k 60", "Alta Ciudad y Barrios"
tvw.Nodes.Add "a 8", tvwChild, "k 61", "Eliminar Ciudad y Barrios"
tvw.Nodes.Add "a 8", tvwChild, "k 62", "Modificar Ciudad y Barrios"

tvw.Nodes.Add "configuracion", tvwChild, "k 75", "Configurar Apariencia"
tvw.Nodes.Add "configuracion", tvwChild, "k 76", "Configurar Red"
tvw.Nodes.Add "configuracion", tvwChild, "k 77", "Configurar Base de Datos"
tvw.Nodes.Add "configuracion", tvwChild, "k 78", "Configurar Comportamiento"
tvw.Nodes.Add "configuracion", tvwChild, "k 83", "Configurar Codigos"
tvw.Nodes.Add "configuracion", tvwChild, "k 87", "Configurar Predeterminados"

End Select


'    AltaLugar = 57
'    BajaLugar = 58
'    ModificacionLugar = 59

End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.WindowState = vbMaximized Or Me.WindowState = vbNormal Then
    If tvw.Visible Then
        tvw.Width = Me.Width - 300
        cmdAceptar.Top = Me.ScaleHeight - cmdAceptar.Height - 50
        cmdCancelar.Top = cmdAceptar.Top
        cmdCancelar.Left = Me.Width - cmdCancelar.Width - 150
        cmdAceptar.Left = cmdCancelar.Left - cmdAceptar.Width - 120
        tvw.Height = cmdAceptar.Top - 150 - tvw.Top
    Else
        cmdCancelar.Left = Me.Width - cmdCancelar.Width - 150
        cmdAceptar.Left = cmdCancelar.Left - cmdAceptar.Width - 120
        cmdAceptar.Top = txtPass.Height + txtPass.Top + 100
        cmdCancelar.Top = cmdAceptar.Top
        Me.Height = 1700
    End If
End If
End Sub

Public Sub Cargar(pLogin As String, pPass As String, pPermisos As blcemi.PermisoManager, BotonResetearPass As Boolean)
    txtLogin = pLogin
    txtPass = pPass
    cmdResetPass.Visible = BotonResetearPass
    Set mPermisos = pPermisos
    CargarPermisos
    
    Me.Show vbModal
    
    If bAceptarCambios Then
        If mPuedeModificarPermisos Then LlenarPermisos
        RaiseEvent PropiedadesModificadas(txtLogin, txtPass)
    End If
    Unload Me
    'los permisos ya estan referenciados
    
End Sub

Private Sub tvw_NodeCheck(ByVal Node As MSComctlLib.Node)
On Error Resume Next
Dim N As Node
'For Each N In Node.Children
'    'If N.Parent.Key = Node.Key Then
'    N.Checked = Node.Checked
'Next
Set N = Node.Child
While Not N Is Nothing
    N.Checked = Node.Checked
    tvw_NodeCheck N
    Set N = N.Next
Wend

End Sub

Private Sub LlenarPermisos()
Dim N As Node
Dim idPermiso As Long

For Each N In tvw.Nodes
    'si hay una k es un permiso, sino es otra hoja del arbol
    If InStr(1, N.Key, "k") Then
        idPermiso = CLng(Right(N.Key, Len(N.Key) - 2))
        If N.Checked Then
            mPermisos.Grant idPermiso
        Else
            mPermisos.Revoke idPermiso
        End If
    End If
Next

End Sub

Private Sub CargarPermisos()
If UsuarioActual Is Nothing Then 'es el primer uso del sistema
    mPuedeModificarPermisos = False
    tvw.Visible = False
    lblPermisos.Visible = False
Else
    If UsuarioActual.Permisos.EsSuperUsuario And Not mPermisos.EsSuperUsuario Then
        mPuedeModificarPermisos = True
        Dim N As Node
        Dim idPermiso As Long
        
        For Each N In tvw.Nodes
            'si hay una k es un permiso, sino es otra hoja del arbol
            If InStr(1, N.Key, "k") Then
                idPermiso = CLng(Right(N.Key, Len(N.Key) - 2))
                N.Checked = mPermisos.Can(idPermiso)
            End If
        Next
    Else
        mPuedeModificarPermisos = False
        tvw.Visible = False
        lblPermisos.Visible = False
    End If
End If
End Sub
