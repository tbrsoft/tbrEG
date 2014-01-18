VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmListadoAtenciones 
   Caption         =   "Listado de atenciones"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9825
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   9825
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   1455
      Left            =   30
      TabIndex        =   0
      Top             =   870
      Width           =   9135
      _ExtentX        =   16113
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
      NEncabezado0    =   "Fecha"
      MEncabezado0    =   "fecha"
      AEncabezado0    =   10
      NEncabezado1    =   "Codigo"
      MEncabezado1    =   "codigo"
      AEncabezado1    =   10
      NEncabezado2    =   "Sintoma"
      MEncabezado2    =   "sintoma"
      AEncabezado2    =   20
      NEncabezado3    =   "Afiliado"
      MEncabezado3    =   "afiliado"
      AEncabezado3    =   20
      NEncabezado4    =   "Hora Llamado"
      MEncabezado4    =   "horallamada"
      AEncabezado4    =   10
      NEncabezado5    =   "Despachador"
      MEncabezado5    =   "despachador"
      AEncabezado5    =   20
      NEncabezado6    =   "Nro Inc."
      MEncabezado6    =   "nroincidente"
      AEncabezado6    =   10
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
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Object.ToolTipText     =   "Actualiza el contenido del listado."
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Ver detalles de la atencion"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Exporta el listado de atenciones a MS Word"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Exporta el listado de atenciones a MS Excel"
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
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "configurar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra este formulario"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblDetalles 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "frmListadoAtenciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mAtenciones As blcemi.AtencionManager
Dim mAtencionesB As blcemi.AtencionBManager

Dim Tipo As eTipoFormulario
Private WithEvents frmConfigEnc As frmConfigEncabezados
Attribute frmConfigEnc.VB_VarHelpID = -1

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN
                 
        lblDetalles.Top = tBar.Height
        lvw.Top = lblDetalles.Top + lblDetalles.Height
        lvw.Height = Me.ScaleHeight - lvw.Top
        lblDetalles.Width = Me.Width - 100
        lvw.Width = Me.Width - 100
        
        DistribuirBotones tBar

    End If
End Sub

Public Sub Consultar(pAtenciones As blcemi.AtencionManager, pDescripcionListado As String, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
   ' Set cmdAceptar.Picture = MDI.il16.ListImages("aceptar").Picture
    Set mAtenciones = pAtenciones
    Me.Caption = "Listado de Atenciones"
    If Tipo = etConRetorno Then
        Me.Show
'        cmdCancelar.Caption = "Cerrar"
'        cmdAceptar.Enabled = False
'
    ElseIf Tipo = etSinRetorno Then
        Me.Show
    End If
    lblDetalles = pDescripcionListado
End Sub

Public Sub ConsultarB(pAtenciones As blcemi.AtencionBManager, pDescripcionListado As String, Optional pTipo As eTipoFormulario = eTipoFormulario.etSinRetorno)
    Tipo = pTipo
   ' Set cmdAceptar.Picture = MDI.il16.ListImages("aceptar").Picture
    Set mAtencionesB = pAtenciones
    Me.Caption = "Listado de Emergencias"
    If Tipo = etConRetorno Then
        Me.Show
'        cmdCancelar.Caption = "Cerrar"
'        cmdAceptar.Enabled = False
'
    ElseIf Tipo = etSinRetorno Then
        Me.Show
    End If
    lblDetalles = pDescripcionListado
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
    Set lvw.Coleccion = IIf(mAtenciones Is Nothing, mAtencionesB, mAtenciones)
    AplicarConfiguracion
    Set Me.Icon = MDI.Icon

End Sub

Public Function GetHelpContext() As String
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            GetHelpContext = "filtrosiniestros"
        Case eModoFuncionamiento.eMFEmergencia:
            GetHelpContext = "filtroatenciones"
    End Select
End Function

Private Sub AplicarConfiguracion()
    lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    tBar.Buttons("word").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWord
    tBar.Buttons("excel").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToExcel
    tBar.Buttons("calc").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToCalc
    tBar.Buttons("writer").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWrite
    
    On Error GoTo errman
    Dim encs As New ControlesPOO.LVCEncabezadoManager
    Set encs = GetEncabezados(APh + "frmListadoAtenciones.lvw.encs")
    If Not encs Is Nothing Then
        If Not encs.Count = 0 Then
            Set lvw.Encabezados = encs
        Else
             Select Case modoSoftware
                Case eModoFuncionamiento.eMFBomberos:
                    Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesBGeneral)
                Case eModoFuncionamiento.eMFEmergencia:
                    Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesGeneral)
             End Select
        End If
    End If
    lvw.Refresh
    Exit Sub
errman:
    Select Case modoSoftware
       Case eModoFuncionamiento.eMFBomberos:
           Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesBGeneral)
       Case eModoFuncionamiento.eMFEmergencia:
           Set lvw.Encabezados = GetEncabezadosDefault(eListadoAtencionesGeneral)
    End Select
End Sub

Public Sub Refrescar()
    AplicarConfiguracion
    lvw.Refresh
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    MostrarDetalles Item
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    MostrarDetalles Item
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "word"
            lvw.ExportToWord lblDetalles, eHorizontal, CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
        Case "excel"
            lvw.ExportToExcel lblDetalles
        Case "writer"
           lvw.ExportToOOWriter lblDetalles, CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
        Case "calc"
            lvw.ExportToOOCalc lblDetalles
        Case "refresh"
            Refrescar
        Case "detalles"
            If Not lvw.SelectedItem Is Nothing Then
                MostrarDetalles lvw.SelectedItem
            End If
        Case "cancelar"
            Unload Me
        Case "configurar"
            ConfigurarListado
    End Select
End Sub

Private Sub MostrarDetalles(Item As Object)
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
           Dim frmDAB As New frmDetalleAtencionB
           frmDAB.VerDetalleAtencionB Item
        Case eModoFuncionamiento.eMFEmergencia:
           Dim frmDA As New frmDetalleAtencion
           frmDA.VerDetalleAtencion Item
    End Select
End Sub

Private Sub ConfigurarListado()
    Set frmConfigEnc = New frmConfigEncabezados
    
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            frmConfigEnc.ConfigurarColumnas GetEncabezadosDisponibles(eListadoAtencionesBGeneral), lvw.Encabezados
        Case eModoFuncionamiento.eMFEmergencia:
            frmConfigEnc.ConfigurarColumnas GetEncabezadosDisponibles(eListadoAtencionesGeneral), lvw.Encabezados
    End Select
End Sub

Private Sub frmConfigEnc_ColumnasSeleccionadas(pSeleccionadas As ControlesPOO.LVCEncabezadoManager)
    SaveEncabezados APh + "frmListadoAtenciones.lvw.encs", pSeleccionadas
    Refrescar 'incluye aplicar configuracion
End Sub
