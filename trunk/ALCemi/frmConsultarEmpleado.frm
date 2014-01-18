VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarEmpleado 
   Caption         =   "Consultar Empleados"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   9180
   Begin VB.TextBox txtFiltro 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   615
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Nombre"
      MEncabezado0    =   "nombre"
      AEncabezado0    =   30
      NEncabezado1    =   "Apellido"
      MEncabezado1    =   "apellido"
      AEncabezado1    =   30
      NEncabezado2    =   "Cargos"
      MEncabezado2    =   "cargostostring"
      AEncabezado2    =   40
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo Empleado..."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar datos del empleado seleccionado"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Muestra los detalles del empleado"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eliminar"
            Object.ToolTipText     =   "Envia el empleado seleccionado a la papelera"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "papelera"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime una lista de los empleados"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Exporta el listado a MS Word"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Exporta el listado a MS Excel"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "writer"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Writer"
            Object.Width           =   2000
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "calc"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Calc"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "html"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "registrarcobro"
            Object.ToolTipText     =   "Registrar el cobro de cuotas"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "recibosanulados"
            Object.ToolTipText     =   "Registrar Devolucion de Recibos Anulados"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardia"
            Object.ToolTipText     =   "Registrar Guardia"
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "liquidacion"
            Object.ToolTipText     =   "Registrar liquidacion de sueldo"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
            Object.ToolTipText     =   "Envia el empleado marcado al formulario anterior"
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultarEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Private Tipo As eTipoFormulario
Public Event EmpleadoSeleccionado(pEmpleado As blcemi.Empleado)
Private WithEvents frmABM As frmABMEmpleado
Attribute frmABM.VB_VarHelpID = -1

Private WithEvents mEmpleados As blcemi.EmpleadoManager
Attribute mEmpleados.VB_VarHelpID = -1

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2010, , "No se puede mostrar el formulario con el metodo Show, utilice la funcion Consultar."
On Error Resume Next
Set tBar.ImageList = MDI.il32 'ver si esta o otra il
Dim b As Button
For Each b In tBar.Buttons
    If b.Style = tbrDefault Then b.Image = b.Key
Next

If Tipo = etConRetorno Then tBar.Buttons("aceptar").Visible = True
AplicarConfiguracion
AplicarPermisos
Set lvw.Coleccion = mEmpleados
tBar.Buttons("papelera").Image = CStr(IIf(GBL.EmpleadosGBL.GetEliminados.Count = 0, "papeleravacia", "papelerallena"))
Set Me.Icon = MDI.Icon

MDI.SetStatusBarText Trim(Str(mEmpleados.Count)) + " Empleados registrados."

lvw.Encabezados.Item("nombre").filtrar = True
lvw.Encabezados.Item("apellido").filtrar = True

End Sub

Private Sub AplicarConfiguracion()
   lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
   tBar.Buttons("word").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWord
   tBar.Buttons("excel").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToExcel
   tBar.Buttons("calc").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToCalc
   tBar.Buttons("writer").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWrite
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "empleados"
End Function

Public Sub Refrescar()
    lvw.filtrar txtFiltro
    tBar.Buttons("papelera").Image = CStr(IIf(GBL.EmpleadosGBL.GetEliminados.Count = 0, "papeleravacia", "papelerallena"))
    AplicarConfiguracion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        'RaiseEvent SeleccionCancelada
        Unload Me
    Case vbKeyA To vbKeyZ
        If ActiveControl.Name <> "txtFiltro" Then
            txtFiltro = txtFiltro + Chr$(KeyCode)
            txtFiltro.SelStart = Len(txtFiltro)
            txtFiltro.SetFocus
        End If
    Case vbKeyF10
        'registrar cobro
        If tBar.Buttons("registrarcobro").Enabled Then tBar_ButtonClick tBar.Buttons("registrarcobro")
     Case vbKeyF11
        'registrar devolucion recibos anulados
        If tBar.Buttons("recibosanulados").Enabled Then tBar_ButtonClick tBar.Buttons("recibosanulados")
     End Select
End Sub

Private Sub mEmpleados_ItemAdded(pEmpleado As blcemi.Empleado)
    lvw.Refresh
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    lvw.SetFocus
End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN
        
        txtFiltro.Top = tBar.Height
        lvw.Top = txtFiltro.Top + txtFiltro.Height
        lvw.Height = Me.ScaleHeight - lvw.Top
        txtFiltro.Width = Me.Width - 100
        lvw.Width = Me.Width - 100
        
        DistribuirBotones tBar
    End If
End Sub

Public Sub Consultar(pEmpleados As blcemi.EmpleadoManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
    Set mEmpleados = pEmpleados
    Me.Show
    'AplicarPermisos lo puse en el load, ver como funciona
End Sub

Private Sub AplicarPermisos()
    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(blcemi.AltaEmpleado)
    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(blcemi.ModificacionEmpleado)
    'tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(BajaEmpleado)
    'me fijo en itemgotfocus para desabilitarlo en caso de administrador
    tBar.Buttons.Item("guardia").Enabled = UsuarioActual.Permisos.Can(blcemi.RegistrarGuardia)
    tBar.Buttons.Item("liquidacion").Enabled = UsuarioActual.Permisos.Can(blcemi.LiquidacionEmpleado)

    'no seria necesario porq pregunto cada vez q se posiciona en un item
    'tBar.Buttons.Item("registrarcobro").Enabled = UsuarioActual.Permisos.Can(AltaPago)

End Sub

Private Sub frmABM_EmpleadoEliminado(pEmpleado As blcemi.Empleado)
lvw.Refresh
End Sub

Private Sub frmABM_EmpleadoModificado(pEmpleado As blcemi.Empleado)
    lvw.Refresh
    Set lvw.SelectedItem = pEmpleado
End Sub

Private Sub frmABM_NuevoEmpleado(pEmpleado As blcemi.Empleado)
    lvw.Refresh
    Set lvw.SelectedItem = pEmpleado
End Sub

Private Sub lvw_ItemGotFocus(Item As Object)
    Dim e As blcemi.Empleado
    Set e = Item
    tBar.Buttons("registrarcobro").Enabled = e.Cargos.Exists(4) And UsuarioActual.Permisos.Can(blcemi.AltaPago)
    tBar.Buttons("recibosanulados").Enabled = e.Cargos.Exists(4) And UsuarioActual.Permisos.Can(blcemi.RegistrarDevolucionRecibosAnulados)
    tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(blcemi.BajaEmpleado) And Not e.Permisos.EsSuperUsuario

End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    Aceptar Item
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    Aceptar Item
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "nuevo"
        Set frmABM = New frmABMEmpleado
        frmABM.Nuevo mEmpleados
        
    Case "modificar"
        'ver si hay q preguntar por canmodify de empleado
        If Not lvw.SelectedItem Is Nothing Then
            Set frmABM = New frmABMEmpleado
            frmABM.Modificar lvw.SelectedItem
        End If
    Case "eliminar"
        If MsgBox("Esta seguro que desea dar de baja al empleado?", vbQuestion + vbYesNo) = vbYes Then
            GBL.EmpleadosGBL.DarItemDeBaja lvw.SelectedItem.id
            Me.Refrescar
        End If
    Case "detalles"
        VerDetalles lvw.SelectedItem
    Case "imprimir"
    Case "papelera"
        Dim frmP As New frmPapelera
        frmP.Mostrar GBL.EmpleadosGBL.GetEliminados, lvw.Encabezados
    Case "word"
       lvw.ExportToWord "Empleados", , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
       lvw.ExportToExcel "Empleados"
    Case "writer"
       lvw.ExportToOOWriter "Empleados", CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
       lvw.ExportToOOCalc "Empleados"
    Case "html"
       'GuardarReporte lvw.ExportToHtml("Empleados"), "Empleados.html"
    Case "registrarcobro"
        frmRegistrarCobroXCobrador.MostrarListadoCuotasImpagas lvw.SelectedItem
    Case "recibosanulados"
        frmRegistrarRecibosAnulados.MostrarListadoRecibosAnulados lvw.SelectedItem
    Case "guardia"
        If Not lvw.SelectedItem Is Nothing Then
            Dim frmRG As New frmRegistrarGuardia
            Load frmRG
            frmRG.NuevaGuardia lvw.SelectedItem
        End If
    Case "liquidacion"
        If Not lvw.SelectedItem Is Nothing Then
            Dim frmLE As New frmLiquidacionEmpleado
            Load frmLE
            frmLE.NuevaLiquidacion lvw.SelectedItem
        End If
    Case "aceptar"
        Aceptar lvw.SelectedItem
    Case "cancelar"
        Unload Me
End Select

End Sub

Private Sub VerDetalles(pEmpleado As blcemi.Empleado)
    If Not pEmpleado Is Nothing Then
        Set frmABM = New frmABMEmpleado
        frmABM.VerDatos pEmpleado
    End If
End Sub

Private Sub Aceptar(pEmpleado As blcemi.Empleado)
    If Tipo = etConRetorno Then
        RaiseEvent EmpleadoSeleccionado(pEmpleado)
        Unload Me
    Else
        VerDetalles pEmpleado
    End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

