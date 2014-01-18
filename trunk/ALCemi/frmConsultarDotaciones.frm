VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarDotaciones 
   Caption         =   "Form2"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8475
   KeyPreview      =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   8475
   Begin ControlesPOO.ListViewConsulta lvwAux 
      Height          =   975
      Left            =   4200
      TabIndex        =   2
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      CampoKey        =   ""
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Movil"
      MEncabezado0    =   "nombremovil"
      AEncabezado0    =   30
      NEncabezado1    =   "Dotacion"
      MEncabezado1    =   "pgdotacion"
      AEncabezado1    =   70
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
   Begin ControlesPOO.TreeViewConsulta tvw 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   3735
      _ExtentX        =   6376
      _ExtentY        =   6800
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCheckBoxes  =   -1  'True
      Indentation     =   299,906
      LineStyle       =   1
      Nodo.BackColor0 =   " 16777215"
      Nodo.Bold0      =   "False"
      Nodo.ChildCollectionField0=   "Dotacion"
      Nodo.Expanded0  =   "False"
      Nodo.ForeColor0 =   " 0"
      Nodo.IdField0   =   "id"
      Nodo.TextField0 =   "NombreMovil"
      Nodo.BackColor1 =   " 16777215"
      Nodo.Bold1      =   "False"
      Nodo.ChildCollectionField1=   ""
      Nodo.Expanded1  =   "False"
      Nodo.ForeColor1 =   " 0"
      Nodo.IdField1   =   "id"
      Nodo.TextField1 =   "NombreCompleto"
      Nodo.BackColor0 =   "0"
      Nodo.Bold0      =   "False"
      Nodo.ChildCollectionField0=   ""
      Nodo.Expanded0  =   "False"
      Nodo.ForeColor0 =   "0"
      Nodo.IdField0   =   ""
      Nodo.TextField0 =   ""
      Nodo.BackColor1 =   "0"
      Nodo.Bold1      =   "False"
      Nodo.ChildCollectionField1=   ""
      Nodo.Expanded1  =   "False"
      Nodo.ForeColor1 =   "0"
      Nodo.IdField1   =   ""
      Nodo.TextField1 =   ""
      Nodo.BackColor2 =   "0"
      Nodo.Bold2      =   "False"
      Nodo.ChildCollectionField2=   ""
      Nodo.Expanded2  =   "False"
      Nodo.ForeColor2 =   "0"
      Nodo.IdField2   =   ""
      Nodo.TextField2 =   ""
      Nodo.BackColor3 =   "0"
      Nodo.Bold3      =   "False"
      Nodo.ChildCollectionField3=   ""
      Nodo.Expanded3  =   "False"
      Nodo.ForeColor3 =   "0"
      Nodo.IdField3   =   ""
      Nodo.TextField3 =   ""
      Nodo.BackColor4 =   "0"
      Nodo.Bold4      =   "False"
      Nodo.ChildCollectionField4=   ""
      Nodo.Expanded4  =   "False"
      Nodo.ForeColor4 =   "0"
      Nodo.IdField4   =   ""
      Nodo.TextField4 =   ""
      Nodo.BackColor5 =   "0"
      Nodo.Bold5      =   "False"
      Nodo.ChildCollectionField5=   ""
      Nodo.Expanded5  =   "False"
      Nodo.ForeColor5 =   "0"
      Nodo.IdField5   =   ""
      Nodo.TextField5 =   ""
      Nodo.BackColor6 =   "0"
      Nodo.Bold6      =   "False"
      Nodo.ChildCollectionField6=   ""
      Nodo.Expanded6  =   "False"
      Nodo.ForeColor6 =   "0"
      Nodo.IdField6   =   ""
      Nodo.TextField6 =   ""
      Nodo.BackColor7 =   "0"
      Nodo.Bold7      =   "False"
      Nodo.ChildCollectionField7=   ""
      Nodo.Expanded7  =   "False"
      Nodo.ForeColor7 =   "0"
      Nodo.IdField7   =   ""
      Nodo.TextField7 =   ""
      Nodo.BackColor8 =   "0"
      Nodo.Bold8      =   "False"
      Nodo.ChildCollectionField8=   ""
      Nodo.Expanded8  =   "False"
      Nodo.ForeColor8 =   "0"
      Nodo.IdField8   =   ""
      Nodo.TextField8 =   ""
      Nodo.BackColor9 =   "0"
      Nodo.Bold9      =   "False"
      Nodo.ChildCollectionField9=   ""
      Nodo.Expanded9  =   "False"
      Nodo.ForeColor9 =   "0"
      Nodo.IdField9   =   ""
      Nodo.TextField9 =   ""
      Nodo.BackColor10=   "0"
      Nodo.Bold10     =   "False"
      Nodo.ChildCollectionField10=   ""
      Nodo.Expanded10 =   "False"
      Nodo.ForeColor10=   "0"
      Nodo.IdField10  =   ""
      Nodo.TextField10=   ""
      Nodo.BackColor11=   "0"
      Nodo.Bold11     =   "False"
      Nodo.ChildCollectionField11=   ""
      Nodo.Expanded11 =   "False"
      Nodo.ForeColor11=   "0"
      Nodo.IdField11  =   ""
      Nodo.TextField11=   ""
      Nodo.BackColor12=   "0"
      Nodo.Bold12     =   "False"
      Nodo.ChildCollectionField12=   ""
      Nodo.Expanded12 =   "False"
      Nodo.ForeColor12=   "0"
      Nodo.IdField12  =   ""
      Nodo.TextField12=   ""
      Nodo.BackColor13=   "0"
      Nodo.Bold13     =   "False"
      Nodo.ChildCollectionField13=   ""
      Nodo.Expanded13 =   "False"
      Nodo.ForeColor13=   "0"
      Nodo.IdField13  =   ""
      Nodo.TextField13=   ""
      Nodo.BackColor14=   "0"
      Nodo.Bold14     =   "False"
      Nodo.ChildCollectionField14=   ""
      Nodo.Expanded14 =   "False"
      Nodo.ForeColor14=   "0"
      Nodo.IdField14  =   ""
      Nodo.TextField14=   ""
      Nodo.BackColor15=   "0"
      Nodo.Bold15     =   "False"
      Nodo.ChildCollectionField15=   ""
      Nodo.Expanded15 =   "False"
      Nodo.ForeColor15=   "0"
      Nodo.IdField15  =   ""
      Nodo.TextField15=   ""
      Nodo.BackColor16=   "0"
      Nodo.Bold16     =   "False"
      Nodo.ChildCollectionField16=   ""
      Nodo.Expanded16 =   "False"
      Nodo.ForeColor16=   "0"
      Nodo.IdField16  =   ""
      Nodo.TextField16=   ""
      Nodo.BackColor17=   "0"
      Nodo.Bold17     =   "False"
      Nodo.ChildCollectionField17=   ""
      Nodo.Expanded17 =   "False"
      Nodo.ForeColor17=   "0"
      Nodo.IdField17  =   ""
      Nodo.TextField17=   ""
      Nodo.BackColor18=   "0"
      Nodo.Bold18     =   "False"
      Nodo.ChildCollectionField18=   ""
      Nodo.Expanded18 =   "False"
      Nodo.ForeColor18=   "0"
      Nodo.IdField18  =   ""
      Nodo.TextField18=   ""
      Nodo.BackColor19=   "0"
      Nodo.Bold19     =   "False"
      Nodo.ChildCollectionField19=   ""
      Nodo.Expanded19 =   "False"
      Nodo.ForeColor19=   "0"
      Nodo.IdField19  =   ""
      Nodo.TextField19=   ""
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Nueva Dotacion"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Modificar los datos de la Dotaicion seleccionada"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "detalles"
            Object.ToolTipText     =   "Ver detalles de la dotacion"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eliminar"
            Object.ToolTipText     =   "Elimina los datos de la dotacion seleccionada"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "papelera"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "imprimir"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "word"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "excel"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "aceptar"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra este formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultarDotaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Private Tipo As eTipoFormulario
Public Event EquipoSeleccionado(pEquipo As blcemi.Equipo)
Public Event EquiposSeleccionados(pEquipos As blcemi.EquipoManager)

Private WithEvents frmABM As frmABMEquipo
Attribute frmABM.VB_VarHelpID = -1

Private WithEvents mEquipos As blcemi.EquipoManager
Attribute mEquipos.VB_VarHelpID = -1

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2010, , "No se puede mostrar el formulario con el metodo Show, utilice la funcion Consultar."
On Error Resume Next
Set tBar.ImageList = MDI.il32 'ver si esta o otra il
Dim b As Button
For Each b In tBar.Buttons
    If b.Style = tbrDefault Then b.Image = b.Key
Next
tBar.Buttons("papelera").Image = CStr(IIf(GBL.EquiposGBL.GetEliminados.Count = 0, "papeleravacia", "papelerallena"))

If Tipo = etConRetorno Then
    tBar.Buttons("aceptar").Visible = True
Else
    tvw.ShowCheckBoxes = False
End If
Set Me.Icon = MDI.Icon

AplicarConfiguracion
AplicarPermisos
Set tvw.Coleccion = mEquipos
MDI.SetStatusBarText Trim(Str(mEquipos.Count)) + " Dotaciones Registradas."

'lvw.Encabezados.Item("nombre").filtrar = True
'lvw.Encabezados.Item("apellido").filtrar = True

End Sub

Private Sub AplicarConfiguracion()
  ' lvw.GridLines = ccffgg.configuracion.Apariencia.GridLinesConsultas
   'tBar.Buttons("word").Visible = ccffgg.configuracion.Comportamiento.AllowExportToWord
   'tBar.Buttons("excel").Visible = ccffgg.configuracion.Comportamiento.AllowExportToExcel
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "dotaciones"
End Function

Public Sub Refrescar()
'    lvw.filtrar txtFiltro
    tvw.Refresh
    AplicarConfiguracion
    tBar.Buttons("papelera").Image = CStr(IIf(GBL.EquiposGBL.GetEliminados.Count = 0, "papeleravacia", "papelerallena"))

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        'RaiseEvent SeleccionCancelada
        Unload Me
    Case vbKeyF2
        Modificar
    Case vbKeyF3
        Nuevo
    End Select
End Sub

'Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
'    lvw.SetFocus
'End If
'End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN

        'txtFiltro.Top = tBar.Height
        tvw.Top = tBar.Height
        tvw.Height = Me.ScaleHeight - tvw.Top
        'txtFiltro.Width = Me.Width - 100
        tvw.Width = Me.Width - 100
        
        DistribuirBotones tBar
    End If
End Sub

Public Sub Consultar(pEquipos As blcemi.EquipoManager, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
    Set mEquipos = pEquipos
    Me.Show
    Me.Caption = "Consulta de Dotaciones"
End Sub

Private Sub AplicarPermisos()
    tBar.Buttons.Item("nuevo").Enabled = UsuarioActual.Permisos.Can(blcemi.AltaEquipo)
    tBar.Buttons.Item("modificar").Enabled = UsuarioActual.Permisos.Can(blcemi.ModificacionEquipo)
    'tBar.Buttons.Item("eliminar").Enabled = UsuarioActual.Permisos.Can(blcemi.BajaEquipo)
End Sub

Private Sub frmABM_EquipoEliminado(pEquipo As blcemi.Equipo)
lvw.Refresh
End Sub

Private Sub frmABM_EquipoModificado(pEquipo As blcemi.Equipo)
    tvw.Refresh
    'Set tvw.SelectedItem = pEquipo IMPLEMENTAR!!
End Sub

Private Sub frmABM_NuevoEquipo(pEquipo As blcemi.Equipo)
    tvw.Refresh
    Set tvw.SelectedItem = pEquipo
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    Aceptar
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    Aceptar
End Sub

Private Sub mEquipos_ItemAdded(pEquipo As blcemi.Equipo)
    tvw.Refresh
End Sub

Private Sub mEquipos_ItemRemoved(pEquipo As blcemi.Equipo)
    tvw.Refresh
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    
    Case "nuevo"
        Nuevo
    Case "modificar"
        'ver si hay q preguntar por canmodify de Equipo
        Modificar
    Case "eliminar"
        Eliminar
    Case "imprimir"
    Case "papelera"
        Dim frmP As New frmPapelera
        frmP.Mostrar GBL.EquiposGBL.GetEliminados, lvwAux.Encabezados

    Case "word"
      ' lvw.ExportToWord "Equipos", , ccffgg.configuracion.Apariencia.ContentsFont, ccffgg.configuracion.Apariencia.TitleFont
    Case "excel"
      ' lvw.ExportToExcel "Equipos"
    Case "aceptar"
        Aceptar
    Case "cancelar"
        Unload Me
End Select

End Sub
Private Sub Nuevo()
    If UsuarioActual.Permisos.Can(blcemi.AltaEquipo) Then
        Set frmABM = New frmABMEquipo
        frmABM.Nuevo mEquipos
    End If
End Sub
Private Sub Modificar()

    If UsuarioActual.Permisos.Can(blcemi.ModificacionEquipo) Then
        
        If Not tvw.SelectedItem Is Nothing Then
            If TypeOf tvw.SelectedItem Is blcemi.Equipo Then
                If Not tvw.SelectedItem.HasReferences Then
                    Set frmABM = New frmABMEquipo
                    frmABM.Modificar tvw.SelectedItem
                Else
                    MsgBox "No se permite modificar la dotacion porque esta relacionada a una Atencion.", vbInformation + vbOKOnly
                End If
            Else
                    MsgBox "Seleccione un movil para poder realizar modificaciones.", vbInformation + vbOKOnly
            End If
        End If
    End If
End Sub

Private Sub Eliminar()
'no hace falta preguntar si tiene referencias porq de todos modos queda guardada la dotacion
    If UsuarioActual.Permisos.Can(blcemi.BajaEquipo) Then
        If Not tvw.SelectedItem Is Nothing Then
            If TypeOf tvw.SelectedItem Is blcemi.Equipo Then
                If MsgBox("Esta seguro que desea dar de baja a la Dotacion?", vbQuestion + vbYesNo) = vbYes Then
                    GBL.EquiposGBL.DarItemDeBaja tvw.SelectedItem.id
                    Me.Refrescar
                End If
            Else
                MsgBox "Seleccione un movil para poder realizar modificaciones.", vbInformation + vbOKOnly
            End If
        End If
    End If
End Sub

Private Sub VerDetalles(pEquipo As blcemi.Equipo)
'    If Not pEquipo Is Nothing Then
'        Set frmABM = New frmABMEquipo
'        frmABM.VerDatos pEquipo
'    End If
End Sub

Private Sub Aceptar()
    If Tipo = etConRetorno Then
        'no lo deberia hacer, pero no puedo descheckear un nodo
        Dim e As blcemi.Equipo
        Dim em As New blcemi.EquipoManager
        Dim obj As Object
        
        For Each obj In tvw.CheckedItems
            If TypeOf obj Is blcemi.Equipo Then
                'este lo uso si es un solo equipo seleccionado
                Set e = obj
                em.AddItem obj
            End If
        Next
        
        'aca tengo los equipos seleccionados
        If em.Count = 1 Then
            RaiseEvent EquipoSeleccionado(e)
            Unload Me
        ElseIf em.Count > 1 Then
            RaiseEvent EquiposSeleccionados(em)
            Unload Me
        Else
            MsgBox "No se ha seleccionado ninguna dotacion.", vbExclamation
        End If
                
    Else
        VerDetalles e
    End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MDI.SetStatusBarText ""
End Sub

Private Sub tvw_ItemCheck(Item As Object, pCancel As Boolean)
    If TypeOf Item Is blcemi.Empleado Then pCancel = True
End Sub

Private Sub tvw_ItemKeyDeletePressed(Item As Object)
    Eliminar
End Sub

Private Sub tvw_ItemKeyEnterPressed(Item As Object)
    Aceptar
End Sub
