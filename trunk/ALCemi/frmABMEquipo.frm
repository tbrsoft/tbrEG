VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMEquipo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6600
   Begin ALCemi.GraphicButton cmdQuitar 
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin ALCemi.GraphicButton cmdAgregar 
      Height          =   375
      Left            =   5280
      TabIndex        =   9
      Top             =   2520
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   5040
      Width           =   1695
   End
   Begin ControlesPOO.Combo cmbMovil 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   556
      Enabled         =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABMEquipo.frx":0000
            Key             =   "medico"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABMEquipo.frx":005E
            Key             =   "agregar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABMEquipo.frx":01B8
            Key             =   "quitar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmABMEquipo.frx":0312
            Key             =   "auto"
         EndProperty
      EndProperty
   End
   Begin ControlesPOO.ListViewConsulta lvwMedicos 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   3413
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   -1  'True
      CampoImage      =   ""
      NEncabezado0    =   "Apellido"
      MEncabezado0    =   "apellido"
      AEncabezado0    =   30
      NEncabezado1    =   "Nombre"
      MEncabezado1    =   "nombre"
      AEncabezado1    =   30
      NEncabezado2    =   "Cargos"
      MEncabezado2    =   "CargosToString"
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
   Begin MSComctlLib.TreeView tvw 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin ControlesPOO.ListViewConsulta lvwAsignados 
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   2778
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   -1  'True
      CampoImage      =   ""
      NEncabezado0    =   "Apellido"
      MEncabezado0    =   "apellido"
      AEncabezado0    =   30
      NEncabezado1    =   "Nombre"
      MEncabezado1    =   "nombre"
      AEncabezado1    =   30
      NEncabezado2    =   "Cargos"
      MEncabezado2    =   "CargosToString"
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
   Begin VB.Label Label2 
      Caption         =   "Seleccione el movil de la dotacion:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label D 
      Caption         =   "Dotaciones Similares:"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Personal:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   855
   End
End
Attribute VB_Name = "frmABMEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar

Public Event NuevoEquipo(pEquipo As blcemi.Equipo)
Public Event EquipoModificado(pEquipo As blcemi.Equipo)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mEquipo As blcemi.Equipo
Private mEquipos As blcemi.EquipoManager
'Private mMedicosAsignados As New EmpleadoManager

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Set mEquipo = mEquipos.Nuevo(cmbMovil.SelectedItem, lvwAsignados.Coleccion)
                RaiseEvent NuevoEquipo(mEquipo)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            If DatosCorrectos Then
                mEquipo.GuardarCambios
                RaiseEvent EquipoModificado(mEquipo)
                Unload Me
            End If
    End Select
    
End Sub

Private Sub cmdAgregar_Click()
    Dim e As blcemi.Empleado
    For Each e In lvwMedicos.SelectedItems
        lvwAsignados.Coleccion.AddItem e
    Next
    lvwAsignados.Refresh
End Sub

Private Sub lvwMedicos_ItemDblClick(Item As Object)
    lvwAsignados.Coleccion.AddItem Item
    lvwAsignados.Refresh
End Sub

Private Sub cmdCancelar_Click()
    If Tipo = etMODIFICACION Then
        mEquipo.Dotacion.CancelChanges
    End If
    Unload Me
End Sub

Public Sub Nuevo(pEquipos As blcemi.EquipoManager)
    'implementar
    Tipo = etALTA
    Me.Show
    Me.Caption = "Nueva Dotacion"
    Set mEquipos = pEquipos
    Set lvwAsignados.Coleccion = New blcemi.EmpleadoManager
End Sub

Public Sub Modificar(pEquipo As blcemi.Equipo)
'implementar
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar Dotacion"

Set mEquipo = pEquipo
Set cmbMovil.SelectedItem = mEquipo.Movil

mEquipo.Dotacion.BeginChanges
Set lvwAsignados.Coleccion = mEquipo.Dotacion

End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
End Sub

'Public Sub VerDatos(pEquipo As blcemi.Equipo)
''implementar
'tipo = etCONSULTA
'Me.Show
'Me.Caption = "Ver Datos de la Obra Social"
'Set mEquipo = pEquipo
'Set ctlTel.Telefonos = mEquipo.Telefonos
'Set lvwCodigos.Coleccion = mEquipo.CodigosCubiertos
'ctlTel.SoloConsulta = True
'ctlDir.SoloConsulta = True
'cmdAgregarCodigo.Enabled = False
'cmdQuitarCodigo.Enabled = False
'cmdAceptar.Enabled = False
'cmdCancelar.Caption = "Cerrar"
'cmdCancelar.Cancel = True
'LlenarCampos
'End Sub

Private Function DatosCorrectos() As Boolean
Dim msj As String

If cmbMovil.SelectedItem Is Nothing Then msj = msj + "Seleccione un Movil" + vbCrLf
    
If lvwAsignados.Coleccion.Count = 0 Then msj = msj + "La dotacion debe incluir al menos un integrante." + vbCrLf

If msj = "" Then
    DatosCorrectos = True
Else
    MsgBox "Faltan los siguientes datos:" + vbCrLf + msj, vbExclamation
    DatosCorrectos = False
End If

End Function

Private Sub cmdQuitar_Click()
Dim e As blcemi.Empleado
For Each e In lvwAsignados.SelectedItems
    lvwAsignados.Coleccion.Remove e.id
Next
lvwAsignados.Refresh
End Sub

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
Set Me.Icon = MDI.Icon

Set cmbMovil.Coleccion = GBL.MovilesGBL

Set cmdAgregar.Picture = ImageList1.ListImages("agregar").Picture
Set cmdQuitar.Picture = ImageList1.ListImages("quitar").Picture

Dim medYParamed As blcemi.EmpleadoManager
'cargo primero los paramedicos
Set medYParamed = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eParamedico)
'ahora agrego los medicos
Dim e As blcemi.Empleado
For Each e In GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eMedico)
    medYParamed.AddItem e
Next
'por ultimo los choferes, habria q ver de mandarle una coleccion de cargos y q devuelva todo junto
For Each e In GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eChofer)
    medYParamed.AddItem e
Next

Set lvwMedicos.Coleccion = medYParamed
'setear icono form

AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()
    lvwMedicos.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    lvwAsignados.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
End Sub

Public Sub Refrescar()
    cmbMovil.Refresh
    lvwMedicos.Refresh
    lvwAsignados.Refresh
    AplicarConfiguracion
End Sub

