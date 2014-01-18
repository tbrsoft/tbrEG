VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmPapelera 
   Caption         =   "Papelera"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   7500
   Begin VB.TextBox txtFiltro 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6800
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
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
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "restaurar"
            Object.ToolTipText     =   "Restaura el elemento"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Elimina el elemento, no se pueden recuperar ls datos "
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmPapelera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 6700
Private Const ALTOMIN = 5000

Private Tipo As Integer
Private mEncabezados As Object

Private mCol As Object

Private Sub Form_Load()
    'levanta un error si quiere usar el metodo show
    If Tipo = 0 Then Err.Raise 2010, , "No se puede mostrar el formulario con el metodo Show, utilice la funcion Mostrar."
    On Error Resume Next
    Set tBar.ImageList = MDI.il32 'ver si esta o otra il
    Dim b As Button
    For Each b In tBar.Buttons
        If b.Style = tbrDefault Then b.Image = b.Key
    Next
      
    AplicarConfiguracion
    Set lvw.Encabezados = mEncabezados
    Set lvw.Coleccion = mCol
    MDI.SetStatusBarText Trim(Str(mCol.Count)) + " Elementos."
    Set Me.Icon = MDI.Icon

End Sub

Private Sub AplicarConfiguracion()
   lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
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

Public Sub Mostrar(pCol As Object, pEncabezados As Object)
Set mCol = pCol
Tipo = 1 'para q no salte error en el load
Set mEncabezados = pEncabezados
Me.Show
AplicarPermisos
End Sub

Private Sub AplicarPermisos()
'ver lo de los permisos
'    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(AltaMovil)
'    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(ModificacionMovil)
'    tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(BajaMovil)
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "restaurar"
        If Not lvw.SelectedItem Is Nothing Then
            lvw.Coleccion.RestaurarItem lvw.SelectedItem.id
            Unload Me
        End If
    Case "eliminar"
    Case "imprimir"
        lvw.ExportToWord "Elementos eliminados", , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    
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


