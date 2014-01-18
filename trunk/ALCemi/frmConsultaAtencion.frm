VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultaAtencion 
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
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   7200
      Top             =   600
   End
   Begin MSComctlLib.ImageList iListCodigos 
      Left            =   7800
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   13027014
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAtencion.frx":0000
            Key             =   "Amarillo"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAtencion.frx":0C54
            Key             =   "Rojo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAtencion.frx":18A8
            Key             =   "Verde"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmConsultaAtencion.frx":24FC
            Key             =   "Azul"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFiltro 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   615
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   2566
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   "codigo"
      NEncabezado0    =   "Codigo"
      MEncabezado0    =   "codigo"
      AEncabezado0    =   8
      NEncabezado1    =   "Sintoma"
      MEncabezado1    =   "sintoma"
      AEncabezado1    =   19
      NEncabezado2    =   "Afiliado"
      MEncabezado2    =   "afiliado"
      AEncabezado2    =   18
      NEncabezado3    =   "Hora Llamado"
      MEncabezado3    =   "horallamada"
      AEncabezado3    =   11
      NEncabezado4    =   "Vence"
      MEncabezado4    =   "GetVencimiento"
      AEncabezado4    =   8
      NEncabezado5    =   "Despachador"
      MEncabezado5    =   "despachador"
      AEncabezado5    =   18
      NEncabezado6    =   "QTH"
      MEncabezado6    =   "qth"
      AEncabezado6    =   7
      NEncabezado7    =   "VL"
      MEncabezado7    =   "vl"
      AEncabezado7    =   7
      NEncabezado8    =   "Movil"
      MEncabezado8    =   "movil"
      AEncabezado8    =   10
      NEncabezado9    =   "Direccion"
      MEncabezado9    =   "pgdireccion"
      AEncabezado9    =   40
      NEncabezado10   =   "Transcurrido"
      MEncabezado10   =   "transcurrido"
      AEncabezado10   =   20
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
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "registraratencion"
            Object.ToolTipText     =   "Registrar Nueva Atencion"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "modificar"
            Object.ToolTipText     =   "Completar los datos de la atencion selecionada"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "salidap"
            Object.ToolTipText     =   "Registrar Salida Preinspeccion"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "llegadap"
            Object.ToolTipText     =   "Registrar Llegada Preinspeccion"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "salidad"
            Object.ToolTipText     =   "Registrar Salida Dotacion"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "qth"
            Object.ToolTipText     =   "Registrar QTH (F3)"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "vl"
            Object.ToolTipText     =   "Registrar VL (F4)"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "configurar"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra este formulario"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabS 
      Height          =   2775
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   4895
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "En curso"
            Key             =   "asignadas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sin movil asignado"
            Key             =   "sinasignar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Finalizadas"
            Key             =   "listasparacerrar"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Event AfiliadoSeleccionado(pAfiliado as blcemi.Afiliado)
Public Event AtencionesModificadas(pCantidadAtencionesPendientes As Integer)

'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Private Tipo As eTipoFormulario

Private WithEvents mAtenciones As blcemi.AtencionManager
Attribute mAtenciones.VB_VarHelpID = -1
Private WithEvents mAtencionesB As blcemi.AtencionBManager
Attribute mAtencionesB.VB_VarHelpID = -1

Private frm As frmAtencion
Attribute frm.VB_VarHelpID = -1
Private frmB As frmAtencionBomberos

Private WithEvents frmConfigEnc As frmConfigEncabezados
Attribute frmConfigEnc.VB_VarHelpID = -1

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        'RaiseEvent SeleccionCancelada
        Unload Me
    Case vbKeyF3
        If tBar.Buttons("qth").Enabled Then tBar_ButtonClick tBar.Buttons("qth")
    Case vbKeyF4
        If tBar.Buttons("vl").Enabled Then tBar_ButtonClick tBar.Buttons("vl")
    Case vbKeyA To vbKeyZ
        If ActiveControl.Name <> "txtFiltro" Then
            txtFiltro = txtFiltro + Chr$(KeyCode)
            txtFiltro.SelStart = Len(txtFiltro)
            txtFiltro.SetFocus
        End If
End Select
End Sub


Private Sub lvw_ItemDataBound(Item As Object, listItem As MSComctlLib.listItem)
'warning: revisar esto
Dim a As Object 'Atencion
Set a = Item
Dim aux As String
On Error GoTo errman:
Dim lsi As ListSubItem
If a.Sintoma.Parent.ColorFuente <> 0 Then
    listItem.ForeColor = a.Sintoma.Parent.ColorFuente
    For Each lsi In listItem.ListSubItems
        lsi.ForeColor = a.Sintoma.Parent.ColorFuente
    Next
End If
If a.Sintoma.Parent.Bold <> 0 Then
    listItem.Bold = a.Sintoma.Parent.Bold
    For Each lsi In listItem.ListSubItems
        lsi.Bold = a.Sintoma.Parent.Bold
    Next
End If
Exit Sub
errman:
GBL.PrintToErrorLog "frmConsultaAtencion", "DataBound: ", Err.Description
End Sub

Private Sub mAtenciones_HasChanged()
    Dim at As blcemi.Atencion
    Set at = lvw.SelectedItem
    'MostrarAtenciones
    'Set lvw.Coleccion = mAtenciones
    'lvw.Refresh
    RaiseEvent AtencionesModificadas(mAtenciones.Count)
    Set lvw.SelectedItem = at
    lvw_ItemClick at
End Sub

Private Sub mAtencionesB_HasChanged()
    Dim at As blcemi.AtencionB
    Set at = lvw.SelectedItem
    'MostrarAtenciones
    'Set lvw.Coleccion = mAtenciones
    'lvw.Refresh
    RaiseEvent AtencionesModificadas(mAtencionesB.Count)
    Set lvw.SelectedItem = at
    lvw_ItemClick at
End Sub

Private Sub tabS_Click()
    MostrarAtenciones
End Sub

Private Sub MostrarAtenciones()
    Dim a As Object
    Set a = lvw.SelectedItem
    
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            If CCFFGG.Configuracion.Comportamiento.SepararAtenciones Then
                mAtencionesB.Reload
                Select Case tabS.SelectedItem.Key
                    Case "asignadas"
                        Set lvw.Coleccion = mAtencionesB.GetAsignadas
                    Case "sinasignar"
                        Set lvw.Coleccion = mAtencionesB.GetSinAsignar
                    Case "listasparacerrar"
                        Set lvw.Coleccion = mAtencionesB.GetByEstado(blcemi.eListaParaCerrar)
                End Select
                tabS.tabS(1).Caption = "En curso (" + Trim(Str(mAtencionesB.GetAsignadas.Count)) + ")"
                tabS.tabS(2).Caption = "Sin movil asignado (" + Trim(Str(mAtencionesB.GetSinAsignar.Count)) + ")"
                tabS.tabS(3).Caption = "Finalizadas (" + Trim(Str(mAtencionesB.GetByEstado(blcemi.eListaParaCerrar).Count)) + ")"
                
                'lvw_ItemClick lvw.SelectedItem
            Else
                Set lvw.Coleccion = mAtencionesB
            End If
            
        Case eModoFuncionamiento.eMFEmergencia:
            If CCFFGG.Configuracion.Comportamiento.SepararAtenciones Then
            mAtenciones.Reload
            Select Case tabS.SelectedItem.Key
                Case "asignadas"
                    Set lvw.Coleccion = mAtenciones.GetAsignadas
                Case "sinasignar"
                    Set lvw.Coleccion = mAtenciones.GetSinAsignar
                Case "listasparacerrar"
                    Set lvw.Coleccion = mAtenciones.GetByEstado(blcemi.eListaParaCerrar)
            End Select
            tabS.tabS(1).Caption = "En curso (" + Trim(Str(mAtenciones.GetAsignadas.Count)) + ")"
            tabS.tabS(2).Caption = "Sin movil asignado (" + Trim(Str(mAtenciones.GetSinAsignar.Count)) + ")"
            tabS.tabS(3).Caption = "Finalizadas (" + Trim(Str(mAtenciones.GetByEstado(blcemi.eListaParaCerrar).Count)) + ")"
            
            'lvw_ItemClick lvw.SelectedItem
        Else
            Set lvw.Coleccion = mAtenciones
        End If
    End Select
    
    On Error Resume Next
    Set lvw.SelectedItem = a
    lvw_ItemClick lvw.SelectedItem
End Sub

Private Sub Timer1_Timer()
'    Dim seleccionado As Object 'Atencion
'    Set seleccionado = lvw.SelectedItem
'    lvw.Refresh
'    Set lvw.SelectedItem = seleccionado
MostrarAtenciones
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
Set lvw.ListImage = iListCodigos
'Set lvw.Coleccion = mAtenciones
AplicarConfiguracion
lvw_ItemClick lvw.SelectedItem
Me.Caption = "Atenciones Pendientes"
Set Me.Icon = MDI.Icon

End Sub

Private Sub AplicarConfiguracion()
    On Error GoTo errman
    Dim encs As New ControlesPOO.LVCEncabezadoManager
    Set encs = GetEncabezados(APh + "frmConsultaAtencion.lvw.encs")
    If Not encs Is Nothing Then
        If Not encs.Count = 0 Then
            Set lvw.Encabezados = encs
        Else
        Select Case modoSoftware
            Case eModoFuncionamiento.eMFBomberos:
                Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesBPendientes)
            Case eModoFuncionamiento.eMFEmergencia:
                Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesPendientes)
                tBar.Buttons("salidap").Visible = False
                tBar.Buttons("llegadap").Visible = False
                tBar.Buttons("salidad").Visible = False
        End Select
        End If
    End If
    lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
    tabS.Visible = CCFFGG.Configuracion.Comportamiento.SepararAtenciones
    Form_Resize
    MostrarAtenciones
    Exit Sub
errman:
'si hubo problemas
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesBPendientes)
        Case eModoFuncionamiento.eMFEmergencia:
            Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesPendientes)
    End Select
    'Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesPendientes)
    GBL.PrintToErrorLog "frmConsultaAtencion", "AplicarConfiguracion - Error cargando los encabezados", Err.Description
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN
                
        If tabS.Visible Then
            
            tabS.Top = tBar.Height
            txtFiltro.Top = tBar.Height + txtFiltro.Height
            lvw.Top = txtFiltro.Top + txtFiltro.Height
            lvw.Height = Me.ScaleHeight - lvw.Top - 100
            txtFiltro.Width = Me.Width - 300
            lvw.Width = Me.Width - 300
            tabS.Width = Me.Width - 100
            tabS.Height = Me.ScaleHeight - tabS.Top
            DistribuirBotones tBar
        Else
            txtFiltro.Left = 50
            lvw.Left = 50
            txtFiltro.Top = tBar.Height
            lvw.Top = txtFiltro.Top + txtFiltro.Height
            lvw.Height = Me.ScaleHeight - lvw.Top
            txtFiltro.Width = Me.Width - 100
            lvw.Width = Me.Width - 100
            DistribuirBotones tBar
        End If
    End If
End Sub

Public Sub Inicializar()
    Tipo = etSinRetorno
   ' Set cmdAceptar.Picture = MDI.il16.ListImages("aceptar").Picture
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            Set mAtencionesB = GBL.AtencionesBGBL.GetByEstado(blcemi.ePendiente)
        Case eModoFuncionamiento.eMFEmergencia:
            Set mAtenciones = GBL.AtencionesGBL.GetByEstado(blcemi.ePendiente)
    End Select
End Sub

Public Sub Mostrar()
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            Me.Caption = "Emergencias Pendientes"
        Case eModoFuncionamiento.eMFEmergencia:
            Me.Caption = "Atenciones Pendientes"
    End Select
    Me.Show
End Sub

'esto me pinta q no se ejecuta nunca
Private Sub frmAtencion_NuevaAtencion(pAtencion As blcemi.Atencion)
    lvw.Refresh
    Set lvw.SelectedItem = pAtencion
End Sub

Private Sub lvw_ItemGotFocus(Item As Object)
    lvw_ItemClick Item
End Sub

Private Sub lvw_ItemClick(Item As Object)
    On Error GoTo errman
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            Dim ItemS As blcemi.AtencionB
            Set ItemS = Item 'super casting...
            'habilitado si no salio la preinsp y si no salio ya la dotacion
            tBar.Buttons("salidap").Enabled = (ItemS.SalidaPreInspeccion = "") And (ItemS.SalidaDotacion = "")
            'si ya salio, si no llego y si no salio la dotacion...
            tBar.Buttons("llegadap").Enabled = (ItemS.LlegadaPreInspeccion = "") And (ItemS.SalidaPreInspeccion <> "") And (ItemS.SalidaDotacion = "")
            'habilitado si no salio ya la dotacion
            tBar.Buttons("salidad").Enabled = (ItemS.SalidaDotacion = "")
            'habilito solo si salio la dotacion, ya no controlo la preinspeccion
            tBar.Buttons("qth").Enabled = (ItemS.QTH = "") And (ItemS.SalidaDotacion <> "")
            'vl tiene q ser despues de qth
            tBar.Buttons("vl").Enabled = (ItemS.VL = "") And (ItemS.QTH <> "")
        Case eModoFuncionamiento.eMFEmergencia:
            tBar.Buttons("qth").Enabled = (Item.QTH = "")
            'vl tiene q ser despues de qth
            tBar.Buttons("vl").Enabled = (Item.VL = "") And (Item.QTH <> "")
    End Select
    
    Exit Sub
errman:
    'es porq no hay ningun item en el lvw
    tBar.Buttons("salidap").Enabled = False
    tBar.Buttons("llegadap").Enabled = False
    tBar.Buttons("salidad").Enabled = False
    tBar.Buttons("qth").Enabled = False
    tBar.Buttons("vl").Enabled = False
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
   ModificarAtencion Item
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
   ModificarAtencion Item
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
'implementar
Select Case Button.Key
    Dim at As Object ' Atencion, es choto, pero es lo que hay...
    Case "registraratencion"
        Select Case modoSoftware
            Case eModoFuncionamiento.eMFBomberos:
                Set frmB = New frmAtencionBomberos
                frmB.NuevaAtencion Me
            Case eModoFuncionamiento.eMFEmergencia:
                Set frm = New frmAtencion
                frm.NuevaAtencion Me
        End Select
    Case "modificar"
        If Not lvw.SelectedItem Is Nothing Then
            ModificarAtencion lvw.SelectedItem
        End If
    Case "eliminar"
    Case "detalles"
        
    Case "imprimir"
    Case "aceptar"
       ' RaiseEvent AfiliadoSeleccionado(lvw.SelectedItem)
        Unload Me
    Case "cancelar"
      '  RaiseEvent SeleccionCancelada
        Unload Me
    Case "refresh"
        Refrescar
    Case "salidap"
        Set at = lvw.SelectedItem
        at.SalidaPreInspeccion = Time
        at.GuardarModificaciones UsuarioActual
        Refrescar
    Case "llegadap"
        Set at = lvw.SelectedItem
        at.LlegadaPreInspeccion = Time
        at.GuardarModificaciones UsuarioActual
        Refrescar
    Case "salidad"
        Set at = lvw.SelectedItem
        at.SalidaDotacion = Time
        at.GuardarModificaciones UsuarioActual
        Refrescar
    Case "qth"
        Set at = lvw.SelectedItem
        at.QTH = Time
        at.GuardarModificaciones UsuarioActual
        Refrescar
    Case "vl"
        Set at = lvw.SelectedItem
        at.VL = Time
        at.GuardarModificaciones UsuarioActual
        Refrescar
    Case "configurar"
        ConfigurarListado
End Select

End Sub

Private Sub ModificarAtencion(Item As Object)
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            Set frmB = New frmAtencionBomberos
            frmB.ModificarAtencion Item, Me
        Case eModoFuncionamiento.eMFEmergencia:
            Set frm = New frmAtencion
            frm.ModificarAtencion Item, Me
    End Select
End Sub
Public Function GetHelpContext() As String
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            GetHelpContext = "consultasiniestros"
        Case eModoFuncionamiento.eMFEmergencia:
            GetHelpContext = "consultaatenciones"
    End Select
End Function

Public Sub Refrescar()
'    If Not mAtenciones Is Nothing Then mAtenciones.Reload
'    If Not mAtencionesB Is Nothing Then mAtencionesB.Reload
    
    'AplicarConfiguracion
    MostrarAtenciones
End Sub

Private Sub ConfigurarListado()
    Set frmConfigEnc = New frmConfigEncabezados
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            frmConfigEnc.ConfigurarColumnas GetEncabezadosDisponibles(eListadoAtencionesBPendientes), lvw.Encabezados
        Case eModoFuncionamiento.eMFEmergencia:
            frmConfigEnc.ConfigurarColumnas GetEncabezadosDisponibles(eListadoAtencionesPendientes), lvw.Encabezados
    End Select
    End Sub

Private Sub frmConfigEnc_ColumnasSeleccionadas(pSeleccionadas As ControlesPOO.LVCEncabezadoManager)
    SaveEncabezados APh + "frmConsultaAtencion.lvw.encs", pSeleccionadas
    AplicarConfiguracion
End Sub

Private Sub Form_Unload(Cancel As Integer)
lvw.ActualizarAnchos
    SaveEncabezados APh + "frmConsultaAtencion.lvw.encs", lvw.Encabezados
End Sub

