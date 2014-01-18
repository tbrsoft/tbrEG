VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMInvolucrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   10305
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8520
      TabIndex        =   9
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   6600
      TabIndex        =   8
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame fraDatos 
      Height          =   3855
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   10215
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1200
         MaxLength       =   25
         TabIndex        =   4
         Top             =   2400
         Width           =   3855
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sexo"
         Height          =   615
         Left            =   5160
         TabIndex        =   22
         Top             =   2400
         Width           =   4935
         Begin VB.OptionButton optMasculino 
            Caption         =   "Masculino"
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optFemenino 
            Caption         =   "Femenino"
            Height          =   255
            Left            =   1920
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Edad"
         Height          =   855
         Left            =   120
         TabIndex        =   18
         Top             =   2880
         Width           =   4935
         Begin VB.TextBox txtEdad 
            Height          =   285
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   5
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton optEdad 
            Caption         =   "Edad"
            Height          =   255
            Left            =   2880
            TabIndex        =   20
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optFechaNac 
            Caption         =   "Fecha Nac:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpFechaNac 
            Height          =   315
            Left            =   1320
            TabIndex        =   21
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   45481985
            CurrentDate     =   39292
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Relacion con el Siniestro"
         Height          =   975
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   4935
         Begin VB.TextBox txtRelacion 
            Height          =   645
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   0
            Top             =   240
            Width           =   4695
         End
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   615
         Left            =   6360
         MaxLength       =   254
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   3120
         Width           =   3735
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   315
         Left            =   3600
         MaxLength       =   9
         TabIndex        =   3
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtApellido 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1320
         Width           =   3855
      End
      Begin ControlesPOO.Combo cmbTipoDoc 
         Height          =   315
         Left            =   1200
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin ALCemi.ctlDireccion ctlDir 
         Height          =   2265
         Left            =   5160
         TabIndex        =   6
         Top             =   120
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   3995
         ProvinciaVisible=   0   'False
         Caption         =   "Direccion"
         CanDragDrop     =   0   'False
         SoloConsulta    =   0   'False
         EntrecallesVisible=   0   'False
      End
      Begin VB.Label Label2 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   5160
         TabIndex        =   16
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Doc:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Doc:"
         Height          =   255
         Index           =   2
         Left            =   2640
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Apellidos:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   1320
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmABMInvolucrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar

Public Event NuevoInvolucrado(pInvolucrado As blcemi.Involucrado)
Public Event InvolucradoModificado(pInvolucrado As blcemi.Involucrado)
Public Event InvolucradoEliminado(pInvolucrado As blcemi.Involucrado)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mInvolucrado As blcemi.Involucrado

Private mInvolucrados As blcemi.InvolucradoManager

Private mTipoRelacion As blcemi.eRelacion

Private WithEvents frmConsulta As frmConsultaGenerico
Attribute frmConsulta.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    On Error GoTo errman:
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Dim mFechaNac As Date
                If optEdad.Value Then
                    mFechaNac = CDate("01/01/" + Trim(Str(Year(Date) - CInt(txtEdad))))
                Else
                    mFechaNac = dtpFechaNac.Value
                End If
                Set mInvolucrado = mInvolucrados.Nuevo(txtApellido, ctlDir.MiDireccion, mFechaNac, txtNombre, CLng(txtNroDoc), txtObservaciones, CInt(IIf(optMasculino.Value = True, "1", "0")), cmbTipoDoc.SelectedItem, optEdad.Value, mTipoRelacion, txtRelacion, txtTelefono)
                RaiseEvent NuevoInvolucrado(mInvolucrado)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            If DatosCorrectos Then
                LlenarObjeto
                'mInvolucrado.GuardarModificaciones
                RaiseEvent InvolucradoModificado(mInvolucrado)
                Unload Me
            End If
    End Select
    
    Exit Sub
errman:
    GBL.PrintToErrorLog "frmABMInvolucrado", "cmdAceptar", Err.Description
End Sub

Private Function DatosCorrectos() As Boolean
'COMPLETAR!!!
Dim msj As String
Dim msj2 As String 'para los datos no obligatorios
Dim msjDir As String

If Not TextBoxValidado(txtRelacion, eString) Then msj = msj + "Ingrese la relación del Involucrado con el Siniestro." + vbCrLf
If Not TextBoxValidado(txtNombre, eString) Then msj = msj + "Ingrese el nombre del Involucrado." + vbCrLf
If Not TextBoxValidado(txtApellido, eString) Then msj = msj + "Ingrese el apellido del Involucrado." + vbCrLf

If mTipoRelacion = eSolicitante Then
    If Not TextBoxValidado(txtTelefono, eString) Then msj = msj + "Ingrese el telefono del Involucrado." + vbCrLf
End If
'warning: ver de q sea configurable exigir doc
'If ccffgg.configuracion.Requeridos.ExigirDNIAE Then
'    If Not TextBoxValidado(txtNroDoc, eLong) Then msj = msj + "Ingrese el numero de documento." + vbCrLf
'End If

'por el momento le meto un 0 si no cargo el doc
If Not TextBoxValidado(txtNroDoc, eLong) Then txtNroDoc = "0"

If Not ctlDir.DireccionCompleta(msjDir) Then msj = msj + msjDir

If cmbTipoDoc.SelectedItem Is Nothing Then msj = msj + "Seleccione un Tipo de Documento." + vbCrLf
'solo si es tipo edad controlo
If optEdad.Value Then If Not TextBoxValidado(txtEdad, eInteger) Then msj = msj + "Ingrese la edad." + vbCrLf

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
    Unload Me
End Sub

Public Sub Nuevo(pInvolucrados As blcemi.InvolucradoManager, pTipoRelacion As blcemi.eRelacion)
    mTipoRelacion = pTipoRelacion
    Tipo = etALTA
    Set mInvolucrados = pInvolucrados
    Me.Show
    Me.Caption = "Registrar Nuevo " + IIf(pTipoRelacion = eAfectado, "Afectado", "Solicitante")
        
    Set ctlDir.MiDireccion = New blcemi.Direccion
End Sub

Public Sub Modificar(pInvolucrado As blcemi.Involucrado, pTipoRelacion As blcemi.eRelacion)
'implementar
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar datos de " + IIf(pTipoRelacion = eAfectado, "Afectado", "Solicitante")
Set mInvolucrado = pInvolucrado
LlenarCampos
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
Me.Caption = "Eliminar Involucrado"

End Sub

Public Sub VerDatos(pInvolucrado As blcemi.Involucrado)
'implementar
Tipo = etCONSULTA
Me.Show
Set mInvolucrado = pInvolucrado
LlenarCampos
Me.Caption = "Consultar Detalles de " + IIf(pTipoRelacion = eAfectado, "Afectado", "Solicitante")
ctlDir.SoloConsulta = True
cmbTipoDoc.Enabled = False
BloquearTextBoxes True, Me.Controls

cmdAceptar.Visible = False
cmdCancelar.Caption = "Cerrar"

End Sub

'se la utiliza para modificar
Private Sub LlenarObjeto()
    mInvolucrado.Apellido = txtApellido
    mInvolucrado.Nombre = txtNombre
    mInvolucrado.NroDoc = txtNroDoc
    Set mInvolucrado.TipoDoc = cmbTipoDoc.SelectedItem
    Set mInvolucrado.Direccion = ctlDir.MiDireccion
    mInvolucrado.FechaNacimiento = dtpFechaNac.Value
    mInvolucrado.Sexo = CInt(IIf(optMasculino.Value = True, "1", "0"))
    mInvolucrado.DescripcionRelacion = txtRelacion
    mInvolucrado.Telefono = txtTelefono
    mInvolucrado.Observaciones = txtObservaciones
End Sub

Private Sub LlenarCampos()
    txtApellido = mInvolucrado.Apellido
    txtNombre = mInvolucrado.Nombre
    txtNroDoc = mInvolucrado.NroDoc
    txtRelacion = mInvolucrado.DescripcionRelacion
    txtTelefono = mInvolucrado.Telefono
    Set cmbTipoDoc.SelectedItem = mInvolucrado.TipoDoc
    Set ctlDir.MiDireccion = mInvolucrado.Direccion
   
    optMasculino.Value = IIf(mInvolucrado.Sexo = 1, True, False)
    
    txtObservaciones = mInvolucrado.Observaciones
    
    optEdad.Value = mInvolucrado.EsEdad
    If mInvolucrado.EsEdad Then
        txtEdad.Text = mInvolucrado.Edad
    Else
        dtpFechaNac.Value = mInvolucrado.FechaNacimiento
    End If
End Sub


Private Sub Form_Load()

    'levanta un error si quiere usar el metodo show
    If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
    
    Set cmbTipoDoc.Coleccion = GBL.TiposDocumentoGBL
    Set cmbTipoDoc.SelectedItem = GBL.TiposDocumentoGBL.Item(1) 'para q seleccione dni predeterminado
    'setear icono form
    Set Me.Icon = MDI.Icon

    InicializarDireccion ctlDir
    AplicarConfiguracion
End Sub

Public Sub Refrescar()
    AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()

End Sub

Private Sub txtEdad_Change()
    optEdad.Value = True
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub

