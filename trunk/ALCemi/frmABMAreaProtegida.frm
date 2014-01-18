VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMAreaProtegida 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10500
   Begin VB.Frame Frame1 
      Caption         =   "Datos de Afiliacion"
      Height          =   4335
      Left            =   5400
      TabIndex        =   20
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton cmdListadoAfiliados 
         Caption         =   "Ver Listado de Afiliados"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   4815
      End
      Begin VB.TextBox txtTope 
         Height          =   285
         Left            =   4320
         TabIndex        =   9
         Text            =   "10"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   1455
         Left            =   120
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2760
         Width           =   4815
      End
      Begin VB.TextBox txtImporte 
         Height          =   285
         Left            =   4320
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45613057
         CurrentDate     =   39293
      End
      Begin MSComCtl2.DTPicker dtpInscripcion 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45613057
         CurrentDate     =   39293
      End
      Begin ControlesPOO.Combo cmbCobrador 
         Height          =   315
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         AtributoAMostrar=   "nombreCompleto"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Cobrador:"
         Height          =   195
         Left            =   720
         TabIndex        =   29
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Inscripcion:"
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   1305
      End
      Begin VB.Label Label7 
         Caption         =   "Inicio Prestacion:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblTope 
         Caption         =   "Tope Atenciones:"
         Height          =   195
         Left            =   3000
         TabIndex        =   26
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label9 
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   1110
      End
      Begin VB.Label Label10 
         Caption         =   "Importe:"
         Height          =   195
         Left            =   3690
         TabIndex        =   24
         Top             =   1320
         Width           =   570
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8760
      TabIndex        =   14
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   6960
      TabIndex        =   13
      Top             =   4560
      Width           =   1695
   End
   Begin VB.Frame fraDatos 
      Height          =   5055
      Left            =   50
      TabIndex        =   15
      Top             =   0
      Width           =   5295
      Begin TabDlg.SSTab sTabDatos 
         Height          =   3015
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1920
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Direccion"
         TabPicture(0)   =   "frmABMAreaProtegida.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ctlDir"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Telefonos"
         TabPicture(1)   =   "frmABMAreaProtegida.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ctlTel"
         Tab(1).ControlCount=   1
         Begin ALCemi.ctlTelefonos ctlTel 
            Height          =   2595
            Left            =   -74880
            TabIndex        =   5
            Top             =   360
            Width           =   4815
            _ExtentX        =   8493
            _ExtentY        =   4577
            Caption         =   ""
            SoloConsulta    =   0   'False
         End
         Begin ALCemi.ctlDireccion ctlDir 
            Height          =   2565
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   3995
            ProvinciaVisible=   0   'False
            Caption         =   ""
            CanDragDrop     =   0   'False
            SoloConsulta    =   0   'False
            EntrecallesVisible=   -1  'True
         End
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   315
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtApellido 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   1080
         Width           =   3255
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtNombreArea 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   3255
      End
      Begin ControlesPOO.Combo cmbTipoDoc 
         Height          =   315
         Left            =   1920
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Doc:"
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   22
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Doc:"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   17
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Apellido Responsable:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Responsable:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre Area:"
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmABMAreaProtegida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar
Public Event NuevaAreaProtegida(pArea As blcemi.AreaProtegida)
Public Event AreaModificada(pArea As blcemi.AreaProtegida)
Public Event AreaEliminada(pArea As blcemi.AreaProtegida)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mAreaProtegida As blcemi.AreaProtegida
Private mAreasProtegidas As blcemi.AreaProtegidaManager

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Set mAreaProtegida = mAreasProtegidas.Nuevo(txtNombreArea, txtNombre, txtApellido, cmbCobrador.SelectedItem, ctlDir.MiDireccion, dtpInscripcion.Value, dtpInicio.Value, txtNroDoc, txtObservaciones, cmbTipoDoc.SelectedItem, CInt(txtTope), ctlTel.Telefonos, CCur(txtImporte))
                RaiseEvent NuevaAreaProtegida(mAreaProtegida)
                Unload Me
            End If
            'implementar
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            If DatosCorrectos Then
                mAreaProtegida.ApellidoResp = txtApellido
                Set mAreaProtegida.Cobrador = cmbCobrador.SelectedItem
                Set mAreaProtegida.Direccion = ctlDir.MiDireccion
                mAreaProtegida.FechaInscripcion = dtpInscripcion.Value
                mAreaProtegida.Importe = txtImporte
                mAreaProtegida.InicioPrestacion = dtpInicio.Value
                mAreaProtegida.NombreArea = txtNombreArea
                mAreaProtegida.NombreResp = txtNombre
                mAreaProtegida.NroDocResp = txtNroDoc
                mAreaProtegida.Observaciones = txtObservaciones
                Set mAreaProtegida.TipoDocResp = cmbTipoDoc.SelectedItem
                mAreaProtegida.TopeAtenciones = txtTope
                
                mAreaProtegida.GuardarModificaciones
                RaiseEvent AreaModificada(mAreaProtegida)
                Unload Me
            End If
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean

Dim msj As String
Dim msj2 As String 'para los datos no obligatorios
Dim msjDir As String 'por las dudas tenga incompleta la direccion

If Not TextBoxValidado(txtNombre, eString) Then msj = msj + "Ingrese el nombre del Responsable." + vbCrLf
If Not TextBoxValidado(txtApellido, eString) Then msj = msj + "Ingrese el apellido del Responsable." + vbCrLf
If Not TextBoxValidado(txtNombreArea, eString) Then msj = msj + "Ingrese el nombre del Area Protegida." + vbCrLf

If CCFFGG.Configuracion.Requeridos.ExigirDNIRespArea Then
    If Not TextBoxValidado(txtNroDoc, eLong) Then msj = msj + "Ingrese el numero de documento del Responsable." + vbCrLf
End If

If Not ctlDir.DireccionCompleta(msjDir) Then msj = msj + msjDir

If cmbTipoDoc.SelectedItem Is Nothing Then msj = msj + "Seleccione un Tipo de Documento" + vbCrLf

If CCFFGG.Configuracion.Requeridos.UsarTopeAtencArea Then
    If Not TextBoxValidado(txtTope, eString) Then msj = msj + "Ingrese un Tope de Atenciones" + vbCrLf
    If Not TextBoxValidado(txtTope, eInteger) Then msj = msj + "Ingrese un Tope de Atenciones entero y sin letras" + vbCrLf
End If

If Not TextBoxValidado(txtImporte, eMoneda) Then msj = msj + "Ingrese un importe a cobrar." + vbCrLf
If cmbCobrador.SelectedItem Is Nothing Then msj = msj + "Seleccione un cobrador." + vbCrLf

If ctlTel.Telefonos.Count = 0 Then msj2 = "Esta seguro que el Area Protegida no tiene telefonos?" + vbCrLf

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

Private Sub cmdCancelar_Click()
'implementar
If Tipo = etMODIFICACION Then
    Set mAreaProtegida.Telefonos = Nothing
End If
Unload Me
End Sub

Public Sub Nuevo(pAreasProtegidas As blcemi.AreaProtegidaManager)

Tipo = etALTA
Me.Show
Me.Caption = "Nueva Area Protegida"

Set mAreasProtegidas = pAreasProtegidas

Set ctlDir.MiDireccion = New blcemi.Direccion
Set ctlTel.Telefonos = New blcemi.TelefonoManager
cmdListadoAfiliados.Enabled = False
dtpInicio.Value = Date
dtpInscripcion.Value = Date
End Sub

Public Sub Modificar(pAreaProtegida As blcemi.AreaProtegida)
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar Area Protegida"
Set mAreaProtegida = pAreaProtegida
cmdListadoAfiliados.Enabled = True
LlenarCampos
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
Me.Caption = "Eliminar Area Protegida"
End Sub

Public Sub VerDatos(pAreaProtegida As blcemi.AreaProtegida)
Tipo = etCONSULTA
Me.Show
Me.Caption = "Ver detalles del Area Protegida"
Set mAreaProtegida = pAreaProtegida
cmdListadoAfiliados.Enabled = True
ctlTel.SoloConsulta = True
ctlDir.SoloConsulta = True
cmdAceptar.Enabled = False
cmdCancelar.Caption = "Cerrar"
cmdCancelar.Cancel = True
LlenarCampos
BloquearTextBoxes True, Me.Controls

End Sub

Private Sub LlenarCampos()
    txtApellido = mAreaProtegida.ApellidoResp
    txtNombreArea = mAreaProtegida.NombreArea
    txtNombre = mAreaProtegida.NombreResp
    Set cmbTipoDoc.SelectedItem = mAreaProtegida.TipoDocResp
    txtNroDoc = mAreaProtegida.NroDocResp
    Set ctlDir.MiDireccion = mAreaProtegida.Direccion
    Set cmbCobrador.SelectedItem = mAreaProtegida.Cobrador
    dtpInscripcion.Value = mAreaProtegida.FechaInscripcion
    dtpInicio.Value = mAreaProtegida.InicioPrestacion
    txtObservaciones = mAreaProtegida.Observaciones
    txtTope = mAreaProtegida.TopeAtenciones
    txtImporte = mAreaProtegida.Importe
    Set ctlTel.Telefonos = mAreaProtegida.Telefonos
    'mAreaProtegida.Pagos
    'mAreaProtegida.Atenciones
End Sub

Private Sub cmdListadoAfiliados_Click()
    frmConsultarAfiliadoExterno.Consultar mAreaProtegida.Afiliados
End Sub

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."

Set cmbCobrador.Coleccion = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eCobrador)
Set cmbTipoDoc.Coleccion = GBL.TiposDocumentoGBL
Set cmbTipoDoc.SelectedItem = GBL.TiposDocumentoGBL.Item(1)
InicializarDireccion ctlDir
Set ctlTel.BotonAgregar.Picture = MDI.il32.ListImages("agregar").Picture
Set ctlTel.BotonEliminar.Picture = MDI.il32.ListImages("eliminar").Picture
Set ctlTel.BotonModificar.Picture = MDI.il32.ListImages("modificar").Picture
Set Me.Icon = MDI.Icon

If Not CCFFGG.Configuracion.Requeridos.UsarTopeAtencAP Then
    txtTope.Visible = False
    lblTope.Visible = False
End If
End Sub

Public Function GetHelpContext() As String
    'habilitar cuando agregue estas paginas a la ayuda
    GetHelpContext = "" '"abmareaProtegida"
End Function

Public Sub Refrescar()
    
End Sub

Private Sub ctlDir_GotFocus()
sTabDatos.Tab = 0
End Sub

Private Sub ctlTel_GotFocus()
sTabDatos.Tab = 1
End Sub

