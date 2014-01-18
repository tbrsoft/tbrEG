VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarAfiliadoExterno 
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4950
   ScaleWidth      =   8745
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
      TabIndex        =   0
      Top             =   960
      Width           =   5175
      _ExtentX        =   9128
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
      NEncabezado0    =   "NroAfiliado"
      MEncabezado0    =   "nroafiliado"
      AEncabezado0    =   15
      NEncabezado1    =   "Apellido"
      MEncabezado1    =   "Apellido"
      AEncabezado1    =   20
      NEncabezado2    =   "Nombre"
      MEncabezado2    =   "nombre"
      AEncabezado2    =   20
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
      TabIndex        =   1
      Top             =   0
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo Afiliado"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar datos del Afiliado seleccionado"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "writer"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Writer"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "calc"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Calc"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultarAfiliadoExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event AfiliadoExternoSeleccionado(pAfiliadoExterno As blcemi.AfiliadoExterno)
Public Event SeleccionCancelada()


'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Private Tipo As eTipoFormulario

Private mAfiliados As blcemi.AfiliadoExternoManager

Private WithEvents frmABM As frmABMAfiliadoExterno
Attribute frmABM.VB_VarHelpID = -1

Private Sub cmdCancelar_Click()
    RaiseEvent SeleccionCancelada
    Unload Me
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
Set Me.Icon = MDI.Icon

If Tipo = etConRetorno Then tBar.Buttons("aceptar").Visible = True
Set lvw.Coleccion = mAfiliados
AplicarConfiguracion
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
                                                                
Public Sub Consultar(pAfiliadoExternos As blcemi.AfiliadoExternoManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
   ' Set cmdAceptar.Picture = MDI.il16.ListImages("aceptar").Picture
    Set mAfiliados = pAfiliadoExternos
    Me.Show
    If TypeOf mAfiliados.Parent Is blcemi.AreaProtegida Then
        Me.Caption = "Consulta de Afiliados de " + mAfiliados.Parent.NombreArea
    Else
        Me.Caption = "Consulta de Afiliados de " + mAfiliados.Parent.Nombre
    End If
End Sub

Private Sub frmABM_AfiliadoModificado(pAfiliado As blcemi.AfiliadoExterno)
    lvw.Refresh
    Set lvw.SelectedItem = pAfiliado
End Sub

Private Sub frmABM_NuevoAfiliado(pAfiliado As blcemi.AfiliadoExterno)
    lvw.Refresh
    Set lvw.SelectedItem = pAfiliado
End Sub

Private Sub lvw_ItemClick(Item As Object)
'
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    Aceptar
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    Aceptar
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "nuevo"
        Set frmABM = New frmABMAfiliadoExterno
        frmABM.Nuevo mAfiliados
        
    Case "modificar"
        If Not lvw.SelectedItem Is Nothing Then
            Set frmABM = New frmABMAfiliadoExterno
            frmABM.Modificar lvw.SelectedItem
        End If
    Case "eliminar"
    Case "detalles"
        VerDatos
    Case "word"
        lvw.ExportToWord Me.Caption, , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
        lvw.ExportToExcel Me.Caption
    Case "writer"
        lvw.ExportToOOWriter Me.Caption, CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
        lvw.ExportToOOCalc Me.Caption
    Case "aceptar"
        Aceptar
    Case "cancelar"
        RaiseEvent SeleccionCancelada
        Unload Me
End Select

End Sub

Private Sub VerDatos()
    If Not lvw.SelectedItem Is Nothing Then
        Set frmABM = New frmABMAfiliadoExterno
        frmABM.VerDatos lvw.SelectedItem
    End If
End Sub

Private Sub Aceptar()
    If Tipo = etConRetorno Then
        If Not lvw.SelectedItem Is Nothing Then
            RaiseEvent AfiliadoExternoSeleccionado(lvw.SelectedItem)
            Unload Me
        Else
            MsgBox "No se selecciono ningun Afiliado.", vbExclamation
        End If
    Else
        VerDatos
    End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

