VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarObraSocial 
   Caption         =   "Form2"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   9180
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
      Width           =   7815
      _ExtentX        =   13785
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
      NEncabezado1    =   "Servicio Emergencia"
      MEncabezado1    =   "servicio"
      AEncabezado1    =   30
      NEncabezado2    =   "Coseguro"
      MEncabezado2    =   "coseguro"
      AEncabezado2    =   20
      NEncabezado3    =   "Codigos"
      MEncabezado3    =   "codigos"
      AEncabezado3    =   20
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
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nueva Obra Social..."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar datos de la Obra Social seleccionado"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Muestra los detalles de la Obra Social"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Envia la Obra Social seleccionado a la papelera"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime una lista de las Obras Sociales"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Exporta el contenido del listado a Word"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Exporta el contenido del listado a Excel"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "writer"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Writer"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "calc"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Calc"
            Object.Width           =   2000
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "liquidacion"
            Object.ToolTipText     =   "Registrar liquidacion de servicios."
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "afiliados"
            Object.ToolTipText     =   "Ver listado de afiliados"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
            Object.ToolTipText     =   "Envia la Obra Social marcada al formulario anterior"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il 
      Left            =   8040
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaObraSocial.frx":0000
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaObraSocial.frx":6862
            Key             =   "aceptar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaObraSocial.frx":15404
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaObraSocial.frx":187F6
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaObraSocial.frx":2F890
            Key             =   "modificar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultarObraSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Public Event ObraSocialSeleccionada(pObraSocial As blcemi.ObraSocial)
Public Event SeleccionCancelada()

Private WithEvents frmABMOS As frmABMObraSocial
Attribute frmABMOS.VB_VarHelpID = -1

Private mObrasSociales As blcemi.ObraSocialManager

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

If Tipo = etConRetorno Then tBar.Buttons("aceptar").Visible = True
AplicarConfiguracion
Set lvw.Coleccion = mObrasSociales
MDI.SetStatusBarText Trim(Str(mObrasSociales.Count)) + " Obras Sociales registradas."
Set Me.Icon = MDI.Icon

lvw.Encabezados.Item("nombre").filtrar = True
lvw.Encabezados.Item("servicio").filtrar = True
lvw.Encabezados.Item("codigos").filtrar = True

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
    End Select
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

Public Sub Consultar(pObrasSociales As blcemi.ObraSocialManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
    Set mObrasSociales = pObrasSociales
    Me.Show
    Me.Caption = "Consulta de Obras Sociales"
    AplicarPermisos
End Sub

Private Sub AplicarPermisos()
    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(blcemi.AltaObraSocial)
    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(blcemi.ModificacionObraSocial)
    tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(blcemi.BajaObraSocial)
    tBar.Buttons.Item("liquidacion").Enabled = UsuarioActual.Permisos.Can(blcemi.LiquidacionEmpresa)
End Sub

Private Sub frmABMOS_NuevaObraSocial(pObraSocial As blcemi.ObraSocial)
    lvw.Refresh
    Set lvw.SelectedItem = pObraSocial
End Sub

Private Sub frmABMOS_ObraSocialModificada(pObraSocial As blcemi.ObraSocial)
    lvw.Refresh
    Set lvw.SelectedItem = pObraSocial
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    RetornarObraSocialSeleccionada Item
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    RetornarObraSocialSeleccionada Item
End Sub

Private Sub RetornarObraSocialSeleccionada(pObraSocial As blcemi.ObraSocial)
    If Tipo = etConRetorno Then
        RaiseEvent ObraSocialSeleccionada(pObraSocial)
        Unload Me
    Else
        VerDetalles pObraSocial
    End If
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "nuevo"
        Set frmABMOS = New frmABMObraSocial
        frmABMOS.Nuevo mObrasSociales
        
    Case "modificar"
        'ver si hay q preguntar por canmodify
        If Not lvw.SelectedItem Is Nothing Then
            Set frmABMOS = New frmABMObraSocial
            frmABMOS.Modificar lvw.SelectedItem
        End If
    Case "eliminar"
    Case "detalles"
        VerDetalles lvw.SelectedItem
    Case "word"
        lvw.ExportToWord "Obras Sociales", , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
        lvw.ExportToExcel "Obras Sociales"
    Case "writer"
        lvw.ExportToOOWriter "Obras Sociales", CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
        lvw.ExportToOOCalc "Obras Sociales"
    Case "liquidacion"
        If Not lvw.SelectedItem Is Nothing Then
            Dim frmL As New frmLiquidacionEmpresa
            frmL.LiquidarServiciosOS lvw.SelectedItem
        End If
    Case "afiliados"
        If Not lvw.SelectedItem Is Nothing Then
            frmConsultarAfiliadoExterno.Consultar lvw.SelectedItem.Afiliados
        End If
    Case "aceptar"
        RetornarObraSocialSeleccionada lvw.SelectedItem
    Case "cancelar"
        Unload Me
End Select


End Sub

Private Sub VerDetalles(pObraSocial As blcemi.ObraSocial)
    If Not pObraSocial Is Nothing Then
        Set frmABMOS = New frmABMObraSocial
        frmABMOS.VerDatos pObraSocial
    End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

