VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarAfiliado 
   Caption         =   "Form2"
   ClientHeight    =   5475
   ClientLeft      =   2715
   ClientTop       =   2970
   ClientWidth     =   9210
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   9210
   Begin VB.TextBox txtFiltro 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   2566
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Nro Afiliado"
      MEncabezado0    =   "idCompleto"
      AEncabezado0    =   10
      NEncabezado1    =   "Apellido"
      MEncabezado1    =   "apellido"
      AEncabezado1    =   20
      NEncabezado2    =   "Nombre"
      MEncabezado2    =   "Nombre"
      AEncabezado2    =   20
      NEncabezado3    =   "Edad"
      MEncabezado3    =   "Edad"
      AEncabezado3    =   10
      NEncabezado4    =   "Af. a Cargo"
      MEncabezado4    =   "aac"
      AEncabezado4    =   15
      NEncabezado5    =   "Cant. Atenc."
      MEncabezado5    =   "cantatenciones"
      AEncabezado5    =   12
      NEncabezado6    =   "Estado"
      MEncabezado6    =   "estadopagos"
      AEncabezado6    =   13
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
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9210
      _ExtentX        =   16245
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo Afiliado"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar los datos del Afiliado seleccionado"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Ver detalles del afiliado"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Elimina los datos del afiliado seleccionado"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "papelera"
            Object.ToolTipText     =   "Muestra elementos eliminados"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Exporta el listado a MS Word"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Exporta el listado a MS Excel "
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
            Object.Visible         =   0   'False
            Key             =   "imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "registrarcobro"
            Object.ToolTipText     =   "Registrar el cobro de cuotas"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra este formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultarAfiliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event AfiliadoSeleccionado(pAfiliado As blcemi.Afiliado)
Public Event SeleccionCancelada()


'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Private Tipo As eTipoFormulario

Private mAfiliados As blcemi.AfiliadoManager
Private WithEvents frmABM As frmABMAfiliado
Attribute frmABM.VB_VarHelpID = -1
Private WithEvents frmRC As frmRegistrarCobro
Attribute frmRC.VB_VarHelpID = -1

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
    Case vbKeyF9
        RegistrarCobro
    Case vbKeyF2
        Modificar
    Case vbKeyF3
        Nuevo
    End Select
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    lvw.SetFocus
End If
End Sub

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
Set lvw.Coleccion = mAfiliados
tBar.Buttons("papelera").Image = CStr(IIf(GBL.AfiliadosGBL.GetEliminados.Count = 0, "papeleravacia", "papelerallena"))
MDI.SetStatusBarText Trim(Str(mAfiliados.Count)) + " afiliados registrados."

'Set Me.Icon = MDI.il16.ListImages("afiliados").Picture
Set Me.Icon = MDI.Icon

lvw.Encabezados.Item("nombre").filtrar = True
lvw.Encabezados.Item("apellido").filtrar = True
lvw.Encabezados.Item("idCompleto").filtrar = True

End Sub

Private Sub AplicarConfiguracion()
   lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
   tBar.Buttons("word").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWord
   tBar.Buttons("excel").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToExcel
   tBar.Buttons("calc").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToCalc
   tBar.Buttons("writer").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWrite
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "afiliados"
End Function

Public Sub Refrescar()
    Set lvw.Coleccion = mAfiliados
    lvw.filtrar txtFiltro
    tBar.Buttons("papelera").Image = CStr(IIf(GBL.AfiliadosGBL.GetEliminados.Count = 0, "papeleravacia", "papelerallena"))
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

Public Sub Consultar(pAfiliados As blcemi.AfiliadoManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
   ' Set cmdAceptar.Picture = MDI.il16.ListImages("aceptar").Picture
    Set mAfiliados = pAfiliados
    Me.Show
    Me.Caption = "Consulta de afiliados"
    AplicarPermisos
End Sub

Private Sub AplicarPermisos()
    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(blcemi.AltaAfiliado)
    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(blcemi.ModificacionAfiliado)
    tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(blcemi.BajaAfiliado)
    'tBar.Buttons.Item("registrarcobro").Enabled = UsuarioActual.Permisos.Can(AltaPago)
End Sub

Private Sub lvw_ItemGotFocus(Item As Object)
    tBar.Buttons.Item("registrarcobro").Enabled = UsuarioActual.Permisos.Can(blcemi.AltaPago) And Item.TipoAfiliado = blcemi.eTipoAfiliado.eTitular
End Sub

'Private Sub DistribuirBotones()
'fijarse si hace falta
'End Sub

Private Sub frmABM_AfiliadoModificado(pAfiliado As blcemi.Afiliado)
    lvw.Refresh
    Set lvw.SelectedItem = pAfiliado
End Sub

Private Sub frmABM_NuevoAfiliado(pAfiliado As blcemi.Afiliado)
    lvw.Refresh
    Set lvw.SelectedItem = pAfiliado
End Sub

Private Sub frmRC_CobroRegistrado()
    Dim pAfiliado As blcemi.Afiliado
    Set pAfiliado = lvw.SelectedItem
    lvw.Refresh
    Set lvw.SelectedItem = pAfiliado
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent AfiliadoSeleccionado(Item)
        Unload Me
    Else
        VerDetalles Item
    End If
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent AfiliadoSeleccionado(Item)
        Unload Me
    Else
        VerDetalles Item
    End If
End Sub

Private Sub Modificar()
    If UsuarioActual.Permisos.Can(blcemi.ModificacionAfiliado) And Not lvw.SelectedItem Is Nothing Then
        Set frmABM = New frmABMAfiliado
        frmABM.Modificar lvw.SelectedItem
    End If
End Sub

Private Sub Nuevo()
    If UsuarioActual.Permisos.Can(blcemi.AltaAfiliado) Then
        Set frmABM = New frmABMAfiliado
        frmABM.Nuevo mAfiliados
    End If
End Sub

Private Sub RegistrarCobro()
    
    If UsuarioActual.Permisos.Can(blcemi.AltaPago) And Not lvw.SelectedItem Is Nothing Then
        Set frmRC = New frmRegistrarCobro
        frmRC.RegistrarPagoAfiliado lvw.SelectedItem
    End If
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim fuenteTitulo As StdFont
    Dim fuenteContenidos As StdFont
    Set fuenteTitulo = GetFont("Times New Roman", 14, True)
    Set fuenteContenidos = GetFont("Times New Roman", 10)
Select Case Button.Key
    
    Case "nuevo"
        Nuevo
    Case "modificar"
        Modificar
    Case "eliminar"
        If Not lvw.SelectedItem Is Nothing Then
            If MsgBox("Esta seguro que desea dar de baja al afiliado?", vbQuestion + vbYesNo) = vbYes Then
    '            gbl.AfiliadosGBL.Item(lvw.SelectedItem.id).DarDeBaja
                GBL.AfiliadosGBL.DarItemDeBaja lvw.SelectedItem.id
                mAfiliados.Remove (lvw.SelectedItem.id) 'para q se mantenga actualizado el listado
                Me.Refrescar
            End If
        End If
    Case "detalles"
        VerDetalles lvw.SelectedItem
    Case "imprimir"
    Case "papelera"
        Dim frmP As New frmPapelera
        frmP.Mostrar GBL.AfiliadosGBL.GetEliminados, lvw.Encabezados
    Case "word"
       lvw.ExportToWord "Listado de Afiliados", , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
       lvw.ExportToExcel "Listado de Afiliados"
    Case "writer"
        lvw.ExportToOOWriter "Listado de Afiliados", CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
       lvw.ExportToOOCalc "Listado de Afiliados"
    Case "registrarcobro"
        RegistrarCobro
    Case "aceptar"
        RaiseEvent AfiliadoSeleccionado(lvw.SelectedItem)
        Unload Me
    Case "cancelar"
        RaiseEvent SeleccionCancelada
        Unload Me
End Select

End Sub

Private Sub VerDetalles(pAfiliado As blcemi.Afiliado)
    If Not pAfiliado Is Nothing Then
        Set frmABM = New frmABMAfiliado
        frmABM.VerDatos pAfiliado
    End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub


Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

