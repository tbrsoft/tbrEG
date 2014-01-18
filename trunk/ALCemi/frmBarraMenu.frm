VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBarraMenu 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.TreeView tvw 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   9763
      _Version        =   393217
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmBarraMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Actualizar()
Form_Load
End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon
Me.Left = 0
Me.Top = 0
Me.Height = MDI.ScaleHeight
Set tvw.ImageList = MDI.il16
tvw.Nodes.Clear
If Not UsuarioActual Is Nothing Then
    With UsuarioActual.Permisos
        tvw.Nodes.Add(, , "Archivo", "Archivo").Expanded = True
        tvw.Nodes("Archivo").Bold = True
        tvw.Nodes.Add "Archivo", tvwChild, "sesion", "Cerrar Sesion"
        If .Can(blcemi.ConsultarAfiliados) Then tvw.Nodes.Add "Archivo", tvwChild, , "Afiliados", "afiliados"
        If .Can(blcemi.ConsultarAreaProtegida) Then tvw.Nodes.Add "Archivo", tvwChild, , "Areas Protegidas"
        If .Can(blcemi.ConsultarEmpleado) Then tvw.Nodes.Add "Archivo", tvwChild, , "Empleados", "empleados"
        If .Can(blcemi.ConsultarObraSocial) Then tvw.Nodes.Add "Archivo", tvwChild, , "Obras Sociales"
        If .Can(blcemi.ConsultarServicioEmergencia) Then tvw.Nodes.Add "Archivo", tvwChild, , "Servicios de Emergencia"
        If .Can(blcemi.ConsultarMovil) Then tvw.Nodes.Add "Archivo", tvwChild, , "Moviles"
        If .Can(blcemi.ConsultarEquipo) Then tvw.Nodes.Add "Archivo", tvwChild, , "Dotaciones"

        tvw.Nodes.Add "Archivo", tvwChild, , "Salir"
        
        tvw.Nodes.Add(, , "Atencion", "Atencion").Expanded = True
        tvw.Nodes("Atencion").Bold = True
        
        If .Can(blcemi.AltaAtencion) Then tvw.Nodes.Add "Atencion", tvwChild, , "Registrar Atencion", "atencion"
        If .Can(blcemi.ConsultarAtencion) Then tvw.Nodes.Add "Atencion", tvwChild, , "Listado de Atenciones Pendientes"
        If .Can(blcemi.ConsultarAtencion) Then tvw.Nodes.Add "Atencion", tvwChild, , "Listado de Atenciones"
        
        'tvw.Nodes.Add "Administracion"
        'tvw.Nodes.Add "Registrar Pago"
        
        tvw.Nodes.Add(, , "Mantenimiento", "Mantenimiento").Expanded = True
        tvw.Nodes("Mantenimiento").Bold = True
        
        tvw.Nodes.Add "Mantenimiento", tvwChild, , "Cargos"
        tvw.Nodes.Add "Mantenimiento", tvwChild, , "Ocupaciones"
        tvw.Nodes.Add "Mantenimiento", tvwChild, , "Parentezcos"
        tvw.Nodes.Add "Mantenimiento", tvwChild, , "Alergias"
        tvw.Nodes.Add "Mantenimiento", tvwChild, , "Enfermedades"
        tvw.Nodes.Add "Mantenimiento", tvwChild, , "Medicamentos"
        tvw.Nodes.Add "Mantenimiento", tvwChild, , "Tipos de Telefono"
        tvw.Nodes.Add(, , , "Actualizar").Bold = True
    
    End With
Else
    tvw.Nodes.Add(, , "Archivo", "Archivo").Expanded = True
    tvw.Nodes.Add "Archivo", tvwChild, "sesion", "Iniciar Sesion..."
    tvw.Nodes.Add("Archivo", tvwChild, , "Salir").Bold = True
End If
End Sub

Private Sub Form_Resize()
If MDI.WindowState <> vbMinimized Then
    tvw.Height = Me.ScaleHeight
    tvw.Width = Me.Width - 100
End If
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
Select Case Node.Text
    Case "Iniciar Sesion...", "Cerrar Sesion"
        MDI.mnuInicioSesion_Click
    Case "Afiliados"
        MDI.mnuAfiliados_Click
    Case "Areas Protegidas"
        MDI.mnuAreasProtegidas_Click
    Case "Empleados"
        MDI.mnuEmpleados_Click
    Case "Obras Sociales"
        MDI.mnuObrasSociales_Click
    Case "Servicios de Emergencia"
        MDI.mnuServiciosEmergencia_Click
    Case "Moviles"
        MDI.mnuMovil_Click
    Case "Dotaciones"
        MDI.mnuDotaciones_Click
    Case "Salir"
        MDI.mnuSalir_Click
    Case "Registrar Atencion"
        MDI.mnuRegistrarAtencion_Click
    Case "Listado de Atenciones Pendientes"
        MDI.mnuConsultarAtencionesPendientes_Click
    Case "Listado de Atenciones"
        MDI.mnuListadoAtenciones_Click
    'Case "Administracion"
    'Case "Registrar Pago"
    
    Case "Cargos"
        MDI.mnuCargos_Click
    Case "Ocupaciones"
        MDI.mnuOcupaciones_Click
    Case "Parentezcos"
        MDI.mnuParentezcos_Click
    Case "Alergias"
        MDI.mnuAlergias_Click
    Case "Enfermedades"
        MDI.mnuEnfermedades_Click
    Case "Medicamentos"
        MDI.mnuMedicamentos_Click
    Case "Tipos de Telefono"
        MDI.mnuTipoTelefono_Click
    Case "Actualizar"
        MDI.mnuActualizar_Click
End Select
End Sub
