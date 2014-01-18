VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarAreaProtegida 
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
      Top             =   600
      Width           =   615
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2990
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Nombre del Area"
      MEncabezado0    =   "nombrearea"
      AEncabezado0    =   35
      NEncabezado1    =   "Apellido y Nombre del Responsable"
      MEncabezado1    =   "NombreCompleto"
      AEncabezado1    =   35
      NEncabezado2    =   "Cant. Atenc."
      MEncabezado2    =   "cantatenciones"
      AEncabezado2    =   15
      NEncabezado3    =   "Estado"
      MEncabezado3    =   "estadopagos"
      AEncabezado3    =   15
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
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nueva Area Protegida..."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar datos del Area Protegida seleccionada"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Muestra los detalles del Area Protegida"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Envia el Area Protegida seleccionada a la papelera"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime una lista de las Areas Protegidas"
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
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "registrarcobro"
            Object.ToolTipText     =   "Registrar el cobro de cuotas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2000
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
            Object.ToolTipText     =   "Envia el area seleccionada al formulario anterior"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultarAreaProtegida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Public Event AreaProtegidaSeleccionada(pAreaProtegida As blcemi.AreaProtegida)
Public Event SeleccionCancelada()

Private Tipo As eTipoFormulario
Private WithEvents frmABM As frmABMAreaProtegida
Attribute frmABM.VB_VarHelpID = -1
Private WithEvents frmRC As frmRegistrarCobro
Attribute frmRC.VB_VarHelpID = -1

Private mAreasProtegidas As blcemi.AreaProtegidaManager

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
Set lvw.Coleccion = mAreasProtegidas
MDI.SetStatusBarText Trim(Str(mAreasProtegidas.Count)) + " Areas Protegidas registradas."
Set Me.Icon = MDI.Icon

lvw.Encabezados.Item("nombrearea").filtrar = True
lvw.Encabezados.Item("nombrecompleto").filtrar = True

End Sub

Private Sub AplicarConfiguracion()
    lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
    tBar.Buttons("word").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWord
    tBar.Buttons("excel").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToExcel
    tBar.Buttons("calc").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToCalc
    tBar.Buttons("writer").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWrite
End Sub

Public Sub Refrescar()
    lvw.filtrar txtFiltro
    AplicarConfiguracion
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        RaiseEvent SeleccionCancelada
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

Public Sub Consultar(pAreasProtegidas As blcemi.AreaProtegidaManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
    Set mAreasProtegidas = pAreasProtegidas
    Me.Show
    Me.Caption = "Consulta de Areas Protegidas"
    AplicarPermisos
End Sub

Private Sub AplicarPermisos()
    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(blcemi.AltaAreaProtegida)
    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(blcemi.ModificacionAreaProtegida)
    tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(blcemi.BajaAreaProtegida)
End Sub

Private Sub frmABM_AreaModificada(pArea As blcemi.AreaProtegida)
    lvw.Refresh
    Set lvw.SelectedItem = pArea
End Sub

Private Sub frmABM_NuevaAreaProtegida(pArea As blcemi.AreaProtegida)
    lvw.Refresh
    Set lvw.SelectedItem = pArea
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent AreaProtegidaSeleccionada(Item)
        Unload Me
    Else
        VerDetalles Item
    End If
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent AreaProtegidaSeleccionada(Item)
        Unload Me
    Else
        VerDetalles Item
    End If
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "nuevo"
        Set frmABM = New frmABMAreaProtegida
        frmABM.Nuevo mAreasProtegidas
        
    Case "modificar"
'        'ver si hay q preguntar por canmodify de empleado
        If Not lvw.SelectedItem Is Nothing Then
            Set frmABM = New frmABMAreaProtegida
            frmABM.Modificar lvw.SelectedItem
        End If
    Case "eliminar"
    Case "detalles"
        VerDetalles lvw.SelectedItem
    Case "word"
        lvw.ExportToWord "Areas Protegidas", , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
        lvw.ExportToExcel "Areas Protegidas"
    Case "writer"
        lvw.ExportToOOWriter "Areas Protegidas", CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
        lvw.ExportToOOCalc "Areas Protegidas"
    Case "registrarcobro"
        If Not lvw.SelectedItem Is Nothing Then
            Set frmRC = New frmRegistrarCobro
            frmRC.RegistrarPagoAreaProtegida lvw.SelectedItem
        End If
    Case "aceptar"
        RaiseEvent AreaProtegidaSeleccionada(lvw.SelectedItem)
        Unload Me
    Case "cancelar"
        Unload Me
End Select

End Sub

Private Sub VerDetalles(pAreaProtegida As blcemi.AreaProtegida)
    If Not lvw.SelectedItem Is Nothing Then
        Set frmABM = New frmABMAreaProtegida
        frmABM.VerDatos pAreaProtegida
    End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

