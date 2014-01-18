VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultaGenerico 
   Caption         =   "Form2"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4635
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   4635
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
      Top             =   840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2990
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   -1  'True
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   -1  'True
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
      TabIndex        =   1
      Top             =   0
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nuevo..."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprime una lista."
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Exporta el listado a MS Word"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Exporta el listado a MS Excel"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "writer"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Writer"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "calc"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Calc"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ph"
            Style           =   4
            Object.Width           =   2000
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
            Object.ToolTipText     =   "Retorna el elemento seleccionado"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il 
      Left            =   3240
      Top             =   480
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
            Picture         =   "frmConsultaGenerico.frx":0000
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaGenerico.frx":6862
            Key             =   "aceptar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaGenerico.frx":15404
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaGenerico.frx":187F6
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaGenerico.frx":2F890
            Key             =   "modificar"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ANCHOMIN = 5000
Private Const ALTOMIN = 5000

Public Event ItemSeleccionado(pItem As Object)
Public Event ItemsSeleccionados(pColItems As Collection)
Public Event SeleccionCancelada()

Private Tipo As eTipoFormulario

Private mColeccion As Object

Private mTituloFormABM As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF2
        lvw.Editar
        Me.KeyPreview = False
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
Me.Width = 5100
Me.Height = 5100
Set Me.Icon = MDI.Icon

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

Public Sub Consultar(pColeccion As Object, TituloForm As String, TituloFormABM As String, pPuedeAgregar As Boolean, pPuedeModificar As Boolean, pPuedeEliminar As Boolean, Optional NombreCampo As String = "Nombre", Optional pTipo As eTipoFormulario = etSinRetorno)
Set mColeccion = pColeccion
Tipo = pTipo
mTituloFormABM = TituloFormABM
Me.Show
Me.Caption = TituloForm
lvw.Encabezados.Add(NombreCampo, "Nombre", 100).filtrar = True

Set lvw.Coleccion = mColeccion
MDI.SetStatusBarText Trim(Str(mColeccion.Count)) + " elementos en total."

tBar.Buttons.Item("nuevo").Enabled = pPuedeAgregar
tBar.Buttons.Item("modificar").Enabled = pPuedeModificar
tBar.Buttons.Item("eliminar").Enabled = pPuedeEliminar
lvw.AllowModify = pPuedeModificar
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent ItemSeleccionado(Item)
        Unload Me
    End If
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    If Tipo = etConRetorno Then
        RaiseEvent ItemSeleccionado(Item)
        Unload Me
    End If
End Sub

Private Sub lvw_ItemEdited(Item As Object, pNewValue As String, pCancel As Boolean)
Me.KeyPreview = True
If pNewValue <> "" Then
    Item.Nombre = pNewValue
    Item.GuardarModificaciones
Else
    pCancel = True
End If

End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
   
    Case "nuevo"
        Dim frmGen As frmABMGenerico
        Set frmGen = New frmABMGenerico
        Dim Nombre As String
        Nombre = frmGen.Nuevo(mTituloFormABM)
        'validar que no exista un item con el mismo nombre
        If Nombre <> "" Then
            mColeccion.Nuevo Nombre
            lvw.Refresh
        End If
    Case "modificar"
        Me.KeyPreview = False
        lvw.Editar
    Case "eliminar"
    Dim res As VbMsgBoxResult
        res = MsgBox("Esta seguro que desea eliminar el elemento: " + lvw.SelectedItem.Nombre, vbQuestion + vbYesNo)
        If res = vbYes Then
            MsgBox "Implementar!"
        End If
    Case "imprimir"
    Case "word"
        lvw.ExportToWord Me.Caption, , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "excel"
        lvw.ExportToExcel Me.Caption
    Case "write"
        lvw.ExportToOOWriter Me.Caption, CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    Case "calc"
        lvw.ExportToOOCalc Me.Caption
    Case "aceptar"
        If lvw.SelectedItems.Count <> 0 Then
            RaiseEvent ItemsSeleccionados(lvw.SelectedItems)
        End If
        Unload Me
    Case "cancelar"
        RaiseEvent SeleccionCancelada
        Unload Me
End Select
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

