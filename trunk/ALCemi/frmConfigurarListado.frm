VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultaListado 
   Caption         =   "Listados disponibles"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3540
   ScaleWidth      =   4095
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   2775
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4895
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      CampoKey        =   ""
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Listado"
      MEncabezado0    =   "titulo"
      AEncabezado0    =   100
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
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "listado"
            Object.ToolTipText     =   "Muestra el listado"
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra este formulario"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmConsultaListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ANCHOMIN = 5000
Private Const ALTOMIN = 5000

Dim WithEvents frm As frmConfigEncabezados
Attribute frm.VB_VarHelpID = -1
Dim listados As ListadoManager

Private Sub Form_Load()
    On Error Resume Next
    Set tBar.ImageList = MDI.il32
    Dim b As Button
    For Each b In tBar.Buttons
        If b.Style = tbrDefault Then b.Image = b.Key
    Next
    
    Set listados = New ListadoManager
    listados.Load ("d:\informes\")
    Set lvw.Coleccion = listados
    Set Me.Icon = MDI.Icon
    Me.Width = 7000
    Me.Height = 7000
    AplicarConfiguracion
End Sub

Public Sub Refrescar()
    Set listados = New ListadoManager
    listados.Load ("d:\informes\")
    Set lvw.Coleccion = listados
End Sub

Public Sub AplicarConfiguracion()
    lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
End Sub
'Private Sub Command1_Click()
'    Dim encs As New ControlesPOO.LVCEncabezadoManager
'    encs.Add "primer", "primer", 10
'    encs.Add "segundo", "segundo", 10
'    encs.Add "tercer", "tercer", 10
'    encs.Add "cuarto", "cuarto", 10
'    encs.Add "quinto", "quinto", 10
'    encs.Add "sexto", "sexto", 10
'    encs.Add "septimo", "septimo", 10
'    encs.Add "octavo", "octavo", 10
'
'    Dim encs2 As New ControlesPOO.LVCEncabezadoManager
'    encs2.Add "quinto", "quinto", 10
'    encs2.Add "sexto", "sexto", 10
'    encs2.Add "septimo", "septimo", 10
'    encs2.Add "octavo", "octavo", 10
'    Set frm = New frmConfigEncabezados
'    frm.ConfigurarColumnas encs, encs2
'End Sub
'Private Sub frm_ColumnasSeleccionadas(pSeleccionadas As ControlesPOO.LVCEncabezadoManager)
'Set lvw.Encabezados = pSeleccionadas
'End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN
                
        lvw.Top = tBar.Height
        lvw.Height = Me.ScaleHeight - lvw.Top
        lvw.Width = Me.Width - 100
        
        DistribuirBotones tBar

    End If
End Sub


Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "listado"
            If Not lvw.SelectedItem Is Nothing Then
                frmListadoGenerico.MostrarListado lvw.SelectedItem
            End If
        Case "cancelar"
            Unload Me
    End Select
End Sub
