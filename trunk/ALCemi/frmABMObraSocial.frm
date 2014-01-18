VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMObraSocial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5265
   ClientLeft      =   1935
   ClientTop       =   525
   ClientWidth     =   6450
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6450
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4680
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2760
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame fraDatos 
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdContable 
         Caption         =   "Datos Contables"
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdListadoAfiliados 
         Caption         =   "Listado de Afiliados"
         Height          =   315
         Left            =   4080
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin ControlesPOO.Combo cmbServicioEmergencia 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         NuevoEnabled    =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   4215
      End
      Begin TabDlg.SSTab sTabDatos 
         Height          =   2775
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1680
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "Direccion"
         TabPicture(0)   =   "frmABMObraSocial.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ctlDir"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Telefonos"
         TabPicture(1)   =   "frmABMObraSocial.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ctlTel"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Observaciones"
         TabPicture(2)   =   "frmABMObraSocial.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtObservaciones"
         Tab(2).ControlCount=   1
         Begin VB.TextBox txtObservaciones 
            Height          =   2175
            Left            =   -74880
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   480
            Width           =   5775
         End
         Begin ALCemi.ctlDireccion ctlDir 
            Height          =   2265
            Left            =   120
            TabIndex        =   3
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
         Begin ALCemi.ctlTelefonos ctlTel 
            Height          =   2235
            Left            =   -74880
            TabIndex        =   4
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3942
            Caption         =   ""
            SoloConsulta    =   0   'False
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Servicio de Emergencia:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1725
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmABMObraSocial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar

Public Event NuevaObraSocial(pObraSocial As blcemi.ObraSocial)
Public Event ObraSocialModificada(pObraSocial As blcemi.ObraSocial)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mObraSocial As blcemi.ObraSocial
Private mObrasSociales As blcemi.ObraSocialManager

Private mInfoContable As blcemi.InfoContableEmp
Private mCodigos As blcemi.CodigoCubiertoManager

'Private WithEvents frmConsultaC As frmConsultaGenerico
Private WithEvents frmConsultaSE As frmConsultarServiciosEmergencia
Attribute frmConsultaSE.VB_VarHelpID = -1
Private WithEvents frmDetallesC As frmDetallesContables
Attribute frmDetallesC.VB_VarHelpID = -1

Private Sub cmbServicioEmergencia_NuevoSeleccionado()
    Set frmConsultaSE = New frmConsultarServiciosEmergencia
    frmConsultaSE.Consultar GBL.ServiciosEmergenciaGBL, etConRetorno
End Sub

Private Sub cmdContable_Click()
    Set frmDetallesC = New frmDetallesContables
    frmDetallesC.MostrarDetalles mInfoContable, mCodigos, Tipo
End Sub

Private Sub frmConsultaSE_ServicioEmergenciaSeleccionado(pServicioEmergencia As blcemi.ServicioEmergencia)
    cmbServicioEmergencia.Refresh
    Set cmbServicioEmergencia.SelectedItem = pServicioEmergencia
End Sub

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then                              'CCur(txtCoseguro)
                Set mObraSocial = mObrasSociales.Nuevo(mCodigos, 0, ctlDir.MiDireccion, txtNombre, txtObservaciones, cmbServicioEmergencia.SelectedItem, ctlTel.Telefonos, mInfoContable)
                RaiseEvent NuevaObraSocial(mObraSocial)
                Unload Me
            End If
            'implementar
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            If DatosCorrectos Then
                Set mObraSocial.CodigosCubiertos = mCodigos
                mObraSocial.Coseguro = 0 'CCur(txtCoseguro.Text) debe haber quedado de antes
                Set mObraSocial.Direccion = ctlDir.MiDireccion
                mObraSocial.Nombre = txtNombre
                mObraSocial.Observaciones = txtObservaciones
                Set mObraSocial.ServicioEmergencia = cmbServicioEmergencia.SelectedItem
                mObraSocial.GuardarModificaciones
                RaiseEvent ObraSocialModificada(mObraSocial)
                Unload Me
            End If
        Case etCONSULTA
            'implementar
    End Select
    
End Sub

Private Sub cmdCancelar_Click()
    If Tipo = etMODIFICACION Then
        Set mObraSocial.Telefonos = Nothing
        Set mObraSocial.CodigosCubiertos = Nothing
    End If
    Unload Me
End Sub

Public Sub Nuevo(pObrasSociales As blcemi.ObraSocialManager)
    'implementar
    Tipo = etALTA
    Me.Show
    Me.Caption = "Nueva Obra Social"
    Set mObrasSociales = pObrasSociales
    
    Set ctlDir.MiDireccion = New blcemi.Direccion
    Set ctlTel.Telefonos = New blcemi.TelefonoManager
    Set mCodigos = New blcemi.CodigoCubiertoManager
    Set mInfoContable = New blcemi.InfoContableEmp
    cmdListadoAfiliados.Enabled = False

End Sub

Public Sub Modificar(pObraSocial As blcemi.ObraSocial)
'implementar
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar Obra Social"

Set mObraSocial = pObraSocial
Set ctlTel.Telefonos = mObraSocial.Telefonos
Set mCodigos = mObraSocial.CodigosCubiertos
Set mInfoContable = mObraSocial.InfoContable
'mObraSocial.CodigosCubiertos.BeginEdit
cmdListadoAfiliados.Enabled = True
LlenarCampos
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
End Sub

Public Sub VerDatos(pObraSocial As blcemi.ObraSocial)
'implementar
Tipo = etCONSULTA
Me.Show
Me.Caption = "Ver Datos de la Obra Social"
Set mObraSocial = pObraSocial
Set ctlTel.Telefonos = mObraSocial.Telefonos
Set mCodigos = mObraSocial.CodigosCubiertos
Set mInfoContable = mObraSocial.InfoContable
ctlTel.SoloConsulta = True
ctlDir.SoloConsulta = True
cmdAceptar.Enabled = False
cmdCancelar.Caption = "Cerrar"
cmdCancelar.Cancel = True
LlenarCampos
BloquearTextBoxes True, Me.Controls

End Sub

Private Sub LlenarCampos()
txtNombre = mObraSocial.Nombre
Set ctlDir.MiDireccion = mObraSocial.Direccion
'mObraSocial.Afiliados
'mObraSocial.Telefonos
Set cmbServicioEmergencia.SelectedItem = mObraSocial.ServicioEmergencia
'txtCoseguro = mObraSocial.Coseguro
txtObservaciones = mObraSocial.Observaciones

'mObraSocial.id

End Sub

Private Function DatosCorrectos() As Boolean
Dim msj As String
Dim msj2 As String 'para los datos no obligatorios
Dim msjDir As String

If Not TextBoxValidado(txtNombre, eString) Then msj = msj + "Ingrese el nombre de la Obra Social." + vbCrLf
'If Not TextBoxValidado(txtCoseguro, eLong) Then msj = msj + "Ingrese el coseguro de la Obra Social." + vbCrLf
If Not ctlDir.DireccionCompleta(msjDir) Then msj = msj + msjDir

If cmbServicioEmergencia.SelectedItem Is Nothing Then msj = msj + "Seleccione un Servicio de Emergencias" + vbCrLf
    
If ctlTel.Telefonos.Count = 0 Then msj2 = "Esta seguro que la Obra Social no tiene telefonos?" + vbCrLf

If CCFFGG.Configuracion.Codigo.ExigirCodigos Then
    If mCodigos.Count = 0 Then msj = msj + "La Obra Social debe cubrir al menos un Codigo de Emergencia." + vbCrLf
End If
'la info contable no se como validarla...

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

Private Sub cmdListadoAfiliados_Click()
    frmConsultarAfiliadoExterno.Consultar mObraSocial.Afiliados
End Sub

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
Set cmbServicioEmergencia.Coleccion = GBL.ServiciosEmergenciaGBL

'setear icono form
Set Me.Icon = MDI.Icon

Set ctlTel.BotonAgregar.Picture = MDI.il32.ListImages("agregar").Picture
Set ctlTel.BotonEliminar.Picture = MDI.il32.ListImages("eliminar").Picture
Set ctlTel.BotonModificar.Picture = MDI.il32.ListImages("modificar").Picture

InicializarDireccion ctlDir
AplicarConfiguracion
End Sub


Private Sub AplicarConfiguracion()
    
End Sub

Public Sub Refrescar()
    cmbServicioEmergencia.Refresh
    AplicarConfiguracion
End Sub

'-------taborder------------------------

Private Sub ctlDir_GotFocus()
sTabDatos.Tab = 0
End Sub

Private Sub ctlTel_GotFocus()
sTabDatos.Tab = 1
End Sub

'Private Sub lvwCodigos_GotFocus()
'sTabDatos.Tab = 2
'End Sub

Private Sub frmDetallesC_DatosModificados(pCodigos As blcemi.CodigoCubiertoManager)
Set mCodigos = pCodigos
End Sub
