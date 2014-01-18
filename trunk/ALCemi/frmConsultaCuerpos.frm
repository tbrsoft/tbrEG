VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultaCuerpos 
   Caption         =   "Consultar Cuerpos de Bomberos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7590
   Begin VB.TextBox txtFiltro 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   615
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   1695
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   6615
      _ExtentX        =   11668
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
      AEncabezado0    =   20
      NEncabezado1    =   "Responsables"
      MEncabezado1    =   "pgresponsables"
      AEncabezado1    =   40
      NEncabezado2    =   "Unidades"
      MEncabezado2    =   "pgunidades"
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
      TabIndex        =   2
      Top             =   0
      Width           =   7590
      _ExtentX        =   13388
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo movil..."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar datos del movil seleccionado"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Envia el movil seleccionado a la papelera"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime una lista de los moviles"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   2000
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
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
            Object.ToolTipText     =   "Envia el movil marcado al formulario anterior"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaCuerpos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event CuerpoSeleccionado(pCuerpo As blcemi.CuerpoBomberos)

Private Const ANCHOMIN = 6700
Private Const ALTOMIN = 5000

Private Tipo As eTipoFormulario

Private mCuerpos As blcemi.CuerpoBomberosManager

Private WithEvents frmABM As frmABMCuerpoBomberos
Attribute frmABM.VB_VarHelpID = -1

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
    Set lvw.Coleccion = mCuerpos
    MDI.SetStatusBarText Trim(Str(mCuerpos.Count)) + " Cuerpos de Bomberos registrados."
Set Me.Icon = MDI.Icon

    lvw.Encabezados.Item("nombre").filtrar = True
    lvw.Encabezados.Item("pgresponsables").filtrar = True
    lvw.Encabezados.Item("pgunidades").filtrar = True

End Sub

Private Sub AplicarConfiguracion()
   lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
   tBar.Buttons("calc").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToCalc
   tBar.Buttons("writer").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWrite
   tBar.Buttons("word").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWord
   tBar.Buttons("excel").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToExcel
End Sub

Public Sub Refrescar()
    lvw.filtrar txtFiltro
    AplicarConfiguracion
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    'RaiseEvent SeleccionCancelada
    Unload Me
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

Public Sub Consultar(pCuerpoBomberos As blcemi.CuerpoBomberosManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
Set mCuerpos = pCuerpoBomberos
Tipo = pTipo
Me.Show
AplicarPermisos
End Sub

Private Sub AplicarPermisos()
    'warning: ver permisos
'    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(AltaCuerpoBomberos)
'    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(ModificacionCuerpoBomberos)
'    tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(BajaCuerpoBomberos)
    tBar.Buttons.Item("nuevo").Enabled = True
    tBar.Buttons.Item("modificar").Enabled = True
    tBar.Buttons.Item("eliminar").Enabled = True
End Sub

Private Sub frmABM_CuerpoModificado(pCuerpo As blcemi.CuerpoBomberos)
    lvw.Refresh
    Set lvw.SelectedItem = pCuerpo
End Sub

Private Sub frmABM_NuevoCuerpo(pCuerpo As blcemi.CuerpoBomberos)
    lvw.Refresh
    Set lvw.SelectedItem = pCuerpo
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent CuerpoSeleccionado(Item)
        Unload Me
    Else
       ' VerDetalles Item
    End If
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent CuerpoSeleccionado(Item)
        Unload Me
    Else
        'VerDetalles Item
    End If
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "nuevo"
        Set frmABM = New frmABMCuerpoBomberos
        frmABM.Nuevo mCuerpos
        
    Case "modificar"
        'ver si hay q preguntar por canmodify
        If Not lvw.SelectedItem Is Nothing Then
            Set frmABM = New frmABMCuerpoBomberos
            frmABM.Modificar lvw.SelectedItem
        End If
    Case "eliminar"
    Case "word"
        lvw.ExportToWord "CuerpoBomberos", , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
        lvw.ExportToExcel "CuerpoBomberos"
    Case "writer"
        lvw.ExportToOOWriter "CuerpoBomberos", CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
        lvw.ExportToOOCalc "CuerpoBomberos"
    Case "aceptar"
        If Not lvw.SelectedItem Is Nothing Then
            If Tipo = etConRetorno Then
                RaiseEvent CuerpoSeleccionado(lvw.SelectedItem)
                Unload Me
            Else
                'VerDetalles Item
            End If
        End If
    Case "cancelar"
        Unload Me
End Select

End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub


