VERSION 5.00
Begin VB.Form frmABMServicioEmergencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5205
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   5205
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3480
      TabIndex        =   6
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Frame fraDatos 
      Height          =   6135
      Left            =   50
      TabIndex        =   4
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdDatosContables 
         Caption         =   "Datos Contables"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   4695
      End
      Begin VB.CommandButton cmdListadoAfiliados 
         Caption         =   "Ver listado de Afiliados"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   4695
      End
      Begin ALCemi.ctlDireccion ctlDir 
         Height          =   2265
         Left            =   120
         TabIndex        =   2
         Top             =   1680
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   3995
         ProvinciaVisible=   0   'False
         Caption         =   "Direccion"
         CanDragDrop     =   0   'False
         SoloConsulta    =   0   'False
         EntrecallesVisible=   0   'False
      End
      Begin ALCemi.ctlTelefonos ctlTel 
         DragIcon        =   "frmABMServicioEmergencia.frx":0000
         Height          =   1995
         Left            =   120
         TabIndex        =   3
         Top             =   3960
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   3519
         Caption         =   "Telefonos"
         SoloConsulta    =   0   'False
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   3855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmABMServicioEmergencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar
Public Event NuevoServicioEmergencia(pServicioEmergencia As blcemi.ServicioEmergencia)
Public Event ServicioEmergenciaModificado(pServicioEmergencia As blcemi.ServicioEmergencia)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mServicioE As blcemi.ServicioEmergencia
Private mServiciosEmergencia As blcemi.ServicioEmergenciaManager

Private mInfoContable As blcemi.InfoContableEmp
Private mCodigos As blcemi.CodigoCubiertoManager
Private WithEvents frmDetallesC As frmDetallesContables
Attribute frmDetallesC.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Set mServicioE = mServiciosEmergencia.Nuevo(txtNombre, ctlDir.MiDireccion, ctlTel.Telefonos, mCodigos, mInfoContable)
                RaiseEvent NuevoServicioEmergencia(mServicioE)
                Unload Me
            End If
            
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            If DatosCorrectos Then
                Set mServicioE.CodigosCubiertos = mCodigos
                Set mServicioE.InfoContable = mInfoContable
                mServicioE.Nombre = txtNombre
                Set mServicioE.Direccion = ctlDir.MiDireccion
                mServicioE.GuardarModificaciones
                RaiseEvent ServicioEmergenciaModificado(mServicioE)
                Unload Me
            End If
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean

Dim msj As String
Dim msjDir As String

If Not TextBoxValidado(txtNombre, eString) Then msj = msj + "Ingrese el nombre del Servicio de Emergencia." + vbCrLf
If Not ctlDir.DireccionCompleta(msjDir) Then msj = msj + msjDir

If CCFFGG.Configuracion.Codigo.ExigirCodigos Then
    If mCodigos.Count = 0 Then msj = msj + "El Servicio de Emergencia debe cubrir al menos un Codigo de Emergencia." + vbCrLf
End If

If ctlTel.Telefonos.Count = 0 And CCFFGG.Configuracion.Comportamiento.MostrarSugerenciasDatosFaltantes Then
    res = MsgBox("Esta seguro que el Servicio de Emergencia no tiene telefonos?", vbOKCancel + vbQuestion)
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

Private Sub cmdCancelar_Click()
'implementar
If Tipo = etMODIFICACION Then
    Set mServicioE.Telefonos = Nothing
    Set mServicioE.CodigosCubiertos = Nothing
End If
Unload Me
End Sub

Public Sub Nuevo(pServiciosEmergencia As blcemi.ServicioEmergenciaManager)

Tipo = etALTA
Me.Show
Me.Caption = "Nuevo Servicio de Emergencia"
Set mServiciosEmergencia = pServiciosEmergencia

Set ctlDir.MiDireccion = New blcemi.Direccion
Set ctlTel.Telefonos = New blcemi.TelefonoManager
Set mCodigos = New blcemi.CodigoCubiertoManager
Set mInfoContable = New blcemi.InfoContableEmp
cmdListadoAfiliados.Enabled = False

End Sub

Public Sub Modificar(pServicioE As blcemi.ServicioEmergencia)
'implementar
Tipo = etMODIFICACION
Me.Show
Set mServicioE = pServicioE
Me.Caption = "Modificar Servicio de Emergencia"

Set ctlDir.MiDireccion = mServicioE.Direccion
txtNombre = mServicioE.Nombre
Set ctlTel.Telefonos = mServicioE.Telefonos
Set mCodigos = mServicioE.CodigosCubiertos
Set mInfoContable = mServicioE.InfoContable
cmdListadoAfiliados.Enabled = True
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
End Sub

Public Sub VerDatos(pServicioE As blcemi.ServicioEmergencia)
Tipo = etCONSULTA
Me.Show
Me.Caption = "Ver detalles del Servicio de Emergencia"

cmdListadoAfiliados.Enabled = False
Set mServicioE = pServicioE
Set ctlDir.MiDireccion = mServicioE.Direccion
txtNombre = mServicioE.Nombre
Set ctlTel.Telefonos = mServicioE.Telefonos
Set mCodigos = mServicioE.CodigosCubiertos
Set mInfoContable = mServicioE.InfoContable
'lvwCodigos.Width = 4575
ctlTel.SoloConsulta = True
ctlDir.SoloConsulta = True
'cmdAgregarCodigo.Visible = False
'cmdQuitarCodigo.Visible = False

cmdAceptar.Enabled = False
cmdCancelar.Caption = "Cerrar"
cmdCancelar.Cancel = True
cmdListadoAfiliados.Enabled = True

End Sub

Private Sub cmdDatosContables_Click()
    Set frmDetallesC = New frmDetallesContables
    frmDetallesC.MostrarDetalles mInfoContable, mCodigos, Tipo
End Sub

Private Sub frmDetallesC_DatosModificados(pCodigos As blcemi.CodigoCubiertoManager)
    Set mCodigos = pCodigos
End Sub

Private Sub cmdListadoAfiliados_Click()
    frmConsultarAfiliadoExterno.Consultar mServicioE.Afiliados
End Sub

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
Set Me.Icon = MDI.Icon

InicializarDireccion ctlDir
Set ctlTel.BotonAgregar.Picture = MDI.il32.ListImages("agregar").Picture
Set ctlTel.BotonEliminar.Picture = MDI.il32.ListImages("eliminar").Picture
Set ctlTel.BotonModificar.Picture = MDI.il32.ListImages("modificar").Picture
AplicarConfiguracion

End Sub

Public Sub Refrescar()
AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()
   
End Sub

