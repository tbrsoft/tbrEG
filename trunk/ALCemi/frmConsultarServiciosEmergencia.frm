VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarServiciosEmergencia 
   Caption         =   "Consulta de Servicios de Emergencia"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5310
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   5310
   Begin VB.TextBox txtFiltro 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   1695
      _ExtentX        =   2990
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
      AEncabezado0    =   100
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
      TabIndex        =   1
      Top             =   0
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo Servicio de Emergencia..."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar Servicio de Emergencia seleccionado"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Muestra los detalles del Servicio de Emergencia"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Envia el elemento seleccionado a la papelera"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime una lista de los Servicios de Emergencia"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Exporta el listado a MS Word"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Exporta el listado a MS Excel"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "writer"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Writer"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "calc"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Calc"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   2000
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "liquidacion"
            Object.ToolTipText     =   "Registrar liquidacion de servicios"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "afiliados"
            Object.ToolTipText     =   "Listado de afiliados"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
            Object.ToolTipText     =   "Envia el elemento marcado al formulario anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultarServiciosEmergencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 5000
Private Const ALTOMIN = 5000

Public Event ServicioEmergenciaSeleccionado(pServicioEmergencia As blcemi.ServicioEmergencia)
Public Event SeleccionCancelada()

Private WithEvents frmABMSE As frmABMServicioEmergencia
Attribute frmABMSE.VB_VarHelpID = -1

Private mServiciosEmergencia As blcemi.ServicioEmergenciaManager 'withevents

Private Tipo As eTipoFormulario

Private Sub Form_Load()

'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2010, , "No se puede mostrar el formulario con el metodo Show, utilice la funcion Consultar."

On Error Resume Next
Set tBar.ImageList = MDI.il32 'ver si esta o otra il
Dim b As Button
For Each b In tBar.Buttons
    If b.Style = tbrDefault Then b.Image = b.Key
Next
Set Me.Icon = MDI.Icon

If Tipo = etConRetorno Then tBar.Buttons("aceptar").Visible = True

Set lvw.Coleccion = mServiciosEmergencia
MDI.SetStatusBarText Trim(Str(mServiciosEmergencia.Count)) + " Servicios de Emergencia registrados."

lvw.Encabezados.Item("nombre").filtrar = True
AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()
    lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
    tBar.Buttons("word").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWord
    tBar.Buttons("excel").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToExcel
    tBar.Buttons("calc").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToCalc
    tBar.Buttons("writer").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWrite
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "consultas"
End Function

Public Sub Refrescar()
    lvw.filtrar txtFiltro
    AplicarConfiguracion
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

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    RaiseEvent SeleccionCancelada
    Unload Me
End If
End Sub

Public Sub Consultar(pServiciosEmergencia As blcemi.ServicioEmergenciaManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
    Set mServiciosEmergencia = pServiciosEmergencia
    Me.Show
    AplicarPermisos
End Sub

Private Sub AplicarPermisos()
    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(blcemi.AltaServicioEmergencia)
    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(blcemi.ModificacionServicioEmergencia)
    tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(blcemi.BajaServicioEmergencia)
    tBar.Buttons.Item("liquidacion").Enabled = UsuarioActual.Permisos.Can(blcemi.LiquidacionEmpresa)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

Private Sub frmABMSE_NuevoServicioEmergencia(pServicioEmergencia As blcemi.ServicioEmergencia)
    lvw.Refresh
    Set lvw.SelectedItem = pServicioEmergencia
End Sub

Private Sub frmABMSE_ServicioEmergenciaModificado(pServicioEmergencia As blcemi.ServicioEmergencia)
    lvw.Refresh
    Set lvw.SelectedItem = pServicioEmergencia
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
   RetornarElemento Item
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    RetornarElemento Item
End Sub

Private Sub RetornarElemento(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent ServicioEmergenciaSeleccionado(Item)
        Unload Me
    Else
        VerDetalles Item
    End If
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "nuevo"
        Set frmABMSE = New frmABMServicioEmergencia
        frmABMSE.Nuevo mServiciosEmergencia
        
    Case "modificar"
        'ver si hay q preguntar por canmodify
        If Not lvw.SelectedItem Is Nothing Then
            Set frmABMSE = New frmABMServicioEmergencia
            frmABMSE.Modificar lvw.SelectedItem
        End If
    Case "eliminar"
    Case "detalles"
        VerDetalles lvw.SelectedItem
    Case "imprimir"
    Case "word"
       lvw.ExportToWord "Servicios de Emergencia", , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
       lvw.ExportToExcel "Servicios de Emergencia"
    Case "writer"
       lvw.ExportToOOWriter "Servicios de Emergencia", CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
       lvw.ExportToOOCalc "Servicios de Emergencia"
    Case "liquidacion"
       If Not lvw.SelectedItem Is Nothing Then
            Dim frmL As New frmLiquidacionEmpresa
            frmL.LiquidarServiciosSE lvw.SelectedItem
        End If
    Case "afiliados"
        If Not lvw.SelectedItem Is Nothing Then
            frmConsultarAfiliadoExterno.Consultar lvw.SelectedItem.Afiliados
        End If
    Case "aceptar"
        If Not lvw.SelectedItem Is Nothing Then
            RetornarElemento lvw.SelectedItem
        End If
    Case "cancelar"
        Unload Me
End Select

End Sub

Private Sub VerDetalles(pServicioE As blcemi.ServicioEmergencia)
    If Not pServicioE Is Nothing Then
        Set frmABMSE = New frmABMServicioEmergencia
        frmABMSE.VerDatos pServicioE
    End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub
