VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMEmpleado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6405
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6405
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4560
      TabIndex        =   9
      Top             =   5640
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Frame fraDatos 
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtMP 
         Height          =   285
         Left            =   4200
         MaxLength       =   12
         TabIndex        =   20
         Top             =   1440
         Width           =   1935
      End
      Begin ControlesPOO.Combo cmbTipoDoc 
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdSistema 
         Caption         =   "Propiedades para el uso del Sistema..."
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   4920
         Width           =   6015
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   4815
      End
      Begin VB.TextBox txtApellido 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   315
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1080
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpFechaNac 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45809665
         CurrentDate     =   39292
      End
      Begin TabDlg.SSTab sTabDatos 
         Height          =   2775
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2040
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Cargos"
         TabPicture(0)   =   "frmABMEmpleado.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "fra"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Direccion"
         TabPicture(1)   =   "frmABMEmpleado.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ctlDir"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Telefonos"
         TabPicture(2)   =   "frmABMEmpleado.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ctlTel"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin VB.Frame fra 
            Height          =   2295
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   5805
            Begin ALCemi.GraphicButton cmdEliminarCargo 
               Height          =   495
               Left            =   5160
               TabIndex        =   22
               ToolTipText     =   "Quitar el cargo seleccionado"
               Top             =   840
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ALCemi.GraphicButton cmdAgregarCargo 
               Height          =   495
               Left            =   5160
               TabIndex        =   21
               ToolTipText     =   "Agregar uno o varios cargos"
               Top             =   240
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ControlesPOO.ListViewConsulta lvwCargos 
               Height          =   1935
               Left            =   120
               TabIndex        =   4
               Top             =   240
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   3413
               HideSelection   =   0   'False
               HideEncabezados =   0   'False
               GridLines       =   -1  'True
               FullRowSelection=   -1  'True
               AutoDistribuirColumnas=   -1  'True
               AllowModify     =   0   'False
               ShowCheckBoxes  =   0   'False
               MultiSelect     =   0   'False
               CampoImage      =   ""
               NEncabezado0    =   "Cargo"
               MEncabezado0    =   "Nombre"
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
         End
         Begin ALCemi.ctlTelefonos ctlTel 
            Height          =   2235
            Left            =   -74880
            TabIndex        =   6
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3942
            Caption         =   ""
            SoloConsulta    =   0   'False
         End
         Begin ALCemi.ctlDireccion ctlDir 
            Height          =   2265
            Left            =   -74880
            TabIndex        =   5
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3995
            ProvinciaVisible=   -1  'True
            Caption         =   ""
            CanDragDrop     =   0   'False
            SoloConsulta    =   0   'False
            EntrecallesVisible=   0   'False
         End
      End
      Begin VB.Label Label2 
         Caption         =   "MP:"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Apellidos:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Doc:"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   15
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Doc:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Nac:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         ToolTipText     =   "Fecha Nacimiento:"
         Top             =   1440
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmABMEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Implementar las funciones Modificar, Eliminar

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Public Event NuevoEmpleado(pEmpleado As blcemi.Empleado)
Public Event EmpleadoModificado(pEmpleado As blcemi.Empleado)
Public Event EmpleadoEliminado(pEmpleado As blcemi.Empleado)

Dim WithEvents mFrmPermisos As frmPermisos
Attribute mFrmPermisos.VB_VarHelpID = -1
Dim WithEvents frmCons As frmConsultaGenerico
Attribute frmCons.VB_VarHelpID = -1

Private mEmpleadoActual As blcemi.Empleado
Private mEmpleados As blcemi.EmpleadoManager

'aca guardo lo q me devuelve el form permisos, solo cuando es un nuevo empleado, para modificar los paso byref
Private mPass As String
Private mLogin As String
Private mPermisos As New blcemi.PermisoManager

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
 'faltan permisos, login y pass
                Set mEmpleadoActual = mEmpleados.Nuevo(txtApellido, txtNombre, lvwCargos.Coleccion, ctlDir.MiDireccion, dtpFechaNac.Value, mLogin, mPass, cmbTipoDoc.SelectedItem, CLng(txtNroDoc), mPermisos, ctlTel.Telefonos, txtMP)
                RaiseEvent NuevoEmpleado(mEmpleadoActual)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            If DatosCorrectos Then
                mEmpleadoActual.Apellido = txtApellido
                mEmpleadoActual.FechaNacimiento = dtpFechaNac.Value
  'faltan permisos, login y pass
                'mEmpleadoActual.Login = ""
                mEmpleadoActual.Nombre = txtNombre
                mEmpleadoActual.NroDoc = CLng(txtNroDoc)
'                mEmpleadoActual.Pass = mPass            los lleno por referencia
'                mEmpleadoActual.Permisos = mLogin
                Set mEmpleadoActual.TipoDoc = cmbTipoDoc.SelectedItem
                Set mEmpleadoActual.Direccion = ctlDir.MiDireccion
                mEmpleadoActual.MP = txtMP
                mEmpleadoActual.GuardarModificaciones
                
                RaiseEvent EmpleadoModificado(mEmpleadoActual)
                Unload Me
            End If
       
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean
    'implementar, preguntar por pass y login tmb
Dim msj As String
Dim msj2 As String 'para los datos no obligatorios
Dim msjDir As String

If Not TextBoxValidado(txtNombre, eString) Then msj = msj + "Ingrese el nombre del empleado." + vbCrLf
If Not TextBoxValidado(txtApellido, eString) Then msj = msj + "Ingrese el apellido del empleado." + vbCrLf
If Not TextBoxValidado(txtNroDoc, eLong) Then msj = msj + "Ingrese el numero de documento." + vbCrLf
If Not ctlDir.DireccionCompleta(msjDir) Then msj = msj + msjDir

If cmbTipoDoc.SelectedItem Is Nothing Then msj = msj + "Seleccione un Tipo de Documento" + vbCrLf
If lvwCargos.Coleccion.Count = 0 Then msj = msj + "El Empleado debe tener algun cargo." + vbCrLf
  
If ctlTel.Telefonos.Count = 0 Then msj2 = "Esta seguro que el empleado no tiene telefonos?" + vbCrLf
If Not TextBoxValidado(txtMP, eString) Then msj2 = "Esta seguro que el empleado no tiene Matricula Profesional (MP)?" + vbCrLf
'veo si es un empleado nuevo me fijo q le hayan asignado login y pass, VER ESTO!!! no me acuerdo...
'If mEmpleadoActual Is Nothing Then
'        mLogin
'        mPass
'End If

If UsuarioActual Is Nothing Then 'en el primer uso si o si tiene q poner login y pass
    If mLogin = "" Then msj = msj + "Ingrese su Login (Boton ""Propiedades para el uso del Sistema"")" + vbCrLf
    If mPass = "" Then msj = msj + "Ingrese su Password (Boton ""Propiedades para el uso del Sistema"")" + vbCrLf
End If

If msj2 <> "" And CCFFGG.Configuracion.Comportamiento.MostrarSugerenciasDatosFaltantes Then
    res = MsgBox(msj2, vbOKCancel + vbQuestion)
    If res = vbCancel Then
        DatosCorrectos = False
        Exit Function
    End If
End If

If msj = "" Then
    DatosCorrectos = True
Else
    MsgBox "Faltan los siguientes datos:" + vbCrLf + msj, vbExclamation
    DatosCorrectos = False
End If
End Function

Private Sub cmdAgregarCargo_Click()
    Set frmCons = New frmConsultaGenerico
    If UsuarioActual Is Nothing Then 'es la primera vez q se usa el sistema
        frmCons.Consultar GBL.CargosGBL, "Consulta de Cargos", "Nuevo Cargo", True, True, True, , etConRetorno
    Else
        frmCons.Consultar GBL.CargosGBL, "Consulta de Cargos", "Nuevo Cargo", UsuarioActual.Permisos.Can(blcemi.AltaCargo), UsuarioActual.Permisos.Can(blcemi.ModificacionCargo), UsuarioActual.Permisos.Can(blcemi.BajaCargo), , etConRetorno
    End If
End Sub

Private Sub cmdCancelar_Click()
    If Tipo = etMODIFICACION Then
        Set mEmpleadoActual.Telefonos = Nothing
        mEmpleadoActual.Cargos.CancelChanges
    End If
    Unload Me
End Sub

Public Sub Nuevo(pEmpleados As blcemi.EmpleadoManager)

Tipo = etALTA
Me.Show
Me.Caption = "Nuevo Empleado"

Set mEmpleados = pEmpleados

If UsuarioActual Is Nothing Then 'primer uso del sistema
    mPermisos.EsSuperUsuario = True
End If

Set ctlTel.Telefonos = New blcemi.TelefonoManager
Set lvwCargos.Coleccion = New blcemi.CargoManager
Set ctlDir.MiDireccion = New blcemi.Direccion

End Sub

Public Sub Modificar(pEmpleado As blcemi.Empleado)  'mandar como parametro el elemento a modificar
'implementar
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar Empleado"

Set mEmpleadoActual = pEmpleado
Set ctlTel.Telefonos = mEmpleadoActual.Telefonos
Set lvwCargos.Coleccion = mEmpleadoActual.Cargos
mEmpleadoActual.Cargos.BeginEdit

LlenarCampos
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
Me.Caption = "Eliminar Empleado"
End Sub

Public Sub VerDatos(pEmpleado As blcemi.Empleado)
Tipo = etCONSULTA
Me.Show
Me.Caption = "Detalles de Empleado"
Set mEmpleadoActual = pEmpleado
LlenarCampos
Set ctlTel.Telefonos = mEmpleadoActual.Telefonos
ctlTel.SoloConsulta = True
ctlDir.SoloConsulta = True
Set lvwCargos.Coleccion = mEmpleadoActual.Cargos
lvwCargos.Width = cmdAgregarCargo.Left + cmdAgregarCargo.Width - lvwCargos.Left
cmdAgregarCargo.Visible = False
cmdEliminarCargo.Visible = False
cmdAceptar.Visible = False
cmdCancelar.Caption = "Cerrar"
cmdCancelar.Cancel = True
cmdSistema.Enabled = False
BloquearTextBoxes True, Me.Controls

End Sub

Private Sub LlenarCampos()
txtApellido = mEmpleadoActual.Apellido
txtNombre = mEmpleadoActual.Nombre
txtNroDoc = Trim(Str(mEmpleadoActual.NroDoc))
dtpFechaNac.Value = mEmpleadoActual.FechaNacimiento
Set cmbTipoDoc.SelectedItem = mEmpleadoActual.TipoDoc
Set ctlDir.MiDireccion = mEmpleadoActual.Direccion 'creo q no hay problema de mandar la dir directamente

txtMP = mEmpleadoActual.MP
'mEmpleadoActual.Id

'mEmpleadoActual.Login
'mEmpleadoActual.Pass
'mEmpleadoActual.Permisos

End Sub

Private Sub cmdEliminarCargo_Click()

If Not lvwCargos.SelectedItem Is Nothing Then
    lvwCargos.Coleccion.Remove lvwCargos.SelectedItem.id
    lvwCargos.Refresh
End If

End Sub

Private Sub cmdSistema_Click()
    Set mFrmPermisos = New frmPermisos
    If mEmpleadoActual Is Nothing Then
        mFrmPermisos.Cargar mLogin, mPass, mPermisos, False
    Else
        mFrmPermisos.Cargar mEmpleadoActual.Login, mEmpleadoActual.Pass, mEmpleadoActual.Permisos, True
    End If

End Sub

Private Sub Form_Load()
    'levanta un error si quiere usar el metodo show
    If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
    
    Set cmbTipoDoc.Coleccion = GBL.TiposDocumentoGBL
    Set cmbTipoDoc.SelectedItem = GBL.TiposDocumentoGBL.Item(1)
    
    'setear icono form
    Set Me.Icon = MDI.Icon

    Set cmdAgregarCargo.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdEliminarCargo.Picture = MDI.il32.ListImages("eliminar").Picture
    
    Set ctlTel.BotonAgregar.Picture = MDI.il32.ListImages("agregar").Picture
    Set ctlTel.BotonEliminar.Picture = MDI.il32.ListImages("eliminar").Picture
    Set ctlTel.BotonModificar.Picture = MDI.il32.ListImages("modificar").Picture

    InicializarDireccion ctlDir
    AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()
    lvwCargos.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
End Sub

Public Sub Refrescar()
lvwCargos.Refresh
AplicarConfiguracion
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mFrmPermisos = Nothing
    Set frmCons = Nothing
End Sub

Private Sub frmCons_ItemSeleccionado(pItem As Object)
lvwCargos.Coleccion.AddItem pItem
lvwCargos.Refresh
End Sub

Private Sub frmCons_ItemsSeleccionados(pColItems As Collection)
    Dim c As blcemi.Cargo
    For Each c In pColItems
        lvwCargos.Coleccion.AddItem c
    Next
    lvwCargos.Refresh
End Sub

Private Sub mFrmPermisos_PropiedadesModificadas(pLogin As String, pPass As String)
    If mEmpleadoActual Is Nothing Then
        mLogin = pLogin
        mPass = pPass
    Else
        mEmpleadoActual.Login = pLogin
        mEmpleadoActual.Pass = pPass
    End If
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub

'-----------------taborder------------------

Private Sub lvwCargos_GotFocus()
sTabDatos.Tab = 0
End Sub

Private Sub ctlDir_GotFocus()
sTabDatos.Tab = 1
End Sub

Private Sub ctlTel_GotFocus()
sTabDatos.Tab = 2
End Sub

