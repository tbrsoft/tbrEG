VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMAfiliadoExterno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   10545
   Begin VB.Frame fraDatos 
      Height          =   5775
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   10335
      Begin VB.Frame Frame3 
         Caption         =   "Historia Clinica"
         Height          =   3495
         Left            =   5400
         TabIndex        =   26
         Top             =   2160
         Width           =   4815
         Begin TabDlg.SSTab sTabHC 
            Height          =   3015
            Left            =   120
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   360
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   5318
            _Version        =   393216
            Style           =   1
            Tabs            =   4
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Enfermedades"
            TabPicture(0)   =   "frmABMAfiliadoExterno.frx":0000
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lvwEnfermedad"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdAgregarEnfermedad"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "cmdQuitarEnfermedad"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Alergias"
            TabPicture(1)   =   "frmABMAfiliadoExterno.frx":001C
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "cmdQuitarAlergia"
            Tab(1).Control(1)=   "cmdAgregarAlergia"
            Tab(1).Control(2)=   "lvwAlergia"
            Tab(1).ControlCount=   3
            TabCaption(2)   =   "Medicamentos"
            TabPicture(2)   =   "frmABMAfiliadoExterno.frx":0038
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "cmdQuitarMedicamento"
            Tab(2).Control(1)=   "cmdAgregarMedicamento"
            Tab(2).Control(2)=   "lvwMedicamento"
            Tab(2).ControlCount=   3
            TabCaption(3)   =   "Observaciones"
            TabPicture(3)   =   "frmABMAfiliadoExterno.frx":0054
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "txtObservaciones"
            Tab(3).ControlCount=   1
            Begin VB.TextBox txtObservaciones 
               Height          =   2415
               Left            =   -74880
               MaxLength       =   254
               MultiLine       =   -1  'True
               TabIndex        =   33
               Top             =   480
               Width           =   4335
            End
            Begin ALCemi.GraphicButton cmdQuitarMedicamento 
               Height          =   495
               Left            =   -71040
               TabIndex        =   27
               ToolTipText     =   "Quita el medicamento seleccionado"
               Top             =   1080
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ALCemi.GraphicButton cmdAgregarMedicamento 
               Height          =   495
               Left            =   -71040
               TabIndex        =   28
               ToolTipText     =   "Agregar uno o varios medicamentos"
               Top             =   480
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ALCemi.GraphicButton cmdQuitarAlergia 
               Height          =   495
               Left            =   -71040
               TabIndex        =   29
               ToolTipText     =   "Quita la alergia seleccionada"
               Top             =   1080
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ALCemi.GraphicButton cmdAgregarAlergia 
               Height          =   495
               Left            =   -71040
               TabIndex        =   30
               ToolTipText     =   "Agregar una o varias alergias"
               Top             =   480
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ALCemi.GraphicButton cmdQuitarEnfermedad 
               Height          =   495
               Left            =   3960
               TabIndex        =   31
               ToolTipText     =   "Quita la enfermedad seleccionada"
               Top             =   1080
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ALCemi.GraphicButton cmdAgregarEnfermedad 
               Height          =   495
               Left            =   3960
               TabIndex        =   32
               ToolTipText     =   "Agregar una o varias enfermedades"
               Top             =   480
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   873
            End
            Begin ControlesPOO.ListViewConsulta lvwEnfermedad 
               Height          =   2415
               Left            =   120
               TabIndex        =   34
               Top             =   480
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   4260
               HideSelection   =   0   'False
               HideEncabezados =   0   'False
               GridLines       =   -1  'True
               FullRowSelection=   0   'False
               AutoDistribuirColumnas=   -1  'True
               AllowModify     =   0   'False
               ShowCheckBoxes  =   0   'False
               MultiSelect     =   0   'False
               CampoImage      =   ""
               NEncabezado0    =   "Nombre"
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
            Begin ControlesPOO.ListViewConsulta lvwAlergia 
               Height          =   2415
               Left            =   -74880
               TabIndex        =   35
               Top             =   480
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   4260
               HideSelection   =   0   'False
               HideEncabezados =   0   'False
               GridLines       =   0   'False
               FullRowSelection=   0   'False
               AutoDistribuirColumnas=   -1  'True
               AllowModify     =   0   'False
               ShowCheckBoxes  =   0   'False
               MultiSelect     =   0   'False
               CampoImage      =   ""
               NEncabezado0    =   "Nombre"
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
            Begin ControlesPOO.ListViewConsulta lvwMedicamento 
               Height          =   2415
               Left            =   -74880
               TabIndex        =   36
               Top             =   480
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   4260
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
               NEncabezado0    =   "Nombre"
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
      End
      Begin VB.Frame Frame2 
         Caption         =   "Edad"
         Height          =   855
         Left            =   5400
         TabIndex        =   25
         Top             =   360
         Width           =   4815
         Begin MSComCtl2.DTPicker dtpFechaNac 
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   330
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   45547521
            CurrentDate     =   39292
         End
         Begin VB.TextBox txtEdad 
            Height          =   285
            Left            =   3720
            MaxLength       =   2
            TabIndex        =   39
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton optEdad 
            Caption         =   "Edad"
            Height          =   255
            Left            =   2880
            TabIndex        =   38
            Top             =   360
            Width           =   735
         End
         Begin VB.OptionButton optFechaNac 
            Caption         =   "Fecha Nac:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   3015
         Left            =   120
         TabIndex        =   8
         Top             =   2640
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Direccion"
         TabPicture(0)   =   "frmABMAfiliadoExterno.frx":0070
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ctlDir"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Telefonos"
         TabPicture(1)   =   "frmABMAfiliadoExterno.frx":008C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ctlTel"
         Tab(1).ControlCount=   1
         Begin ALCemi.ctlDireccion ctlDir 
            Height          =   2565
            Left            =   120
            TabIndex        =   0
            Top             =   360
            Width           =   4815
            _ExtentX        =   8916
            _ExtentY        =   3995
            ProvinciaVisible=   0   'False
            Caption         =   "Direccion"
            CanDragDrop     =   0   'False
            SoloConsulta    =   0   'False
            EntrecallesVisible=   -1  'True
         End
         Begin ALCemi.ctlTelefonos ctlTel 
            Height          =   2595
            Left            =   -74880
            TabIndex        =   7
            Top             =   360
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   3942
            Caption         =   ""
            SoloConsulta    =   0   'False
         End
      End
      Begin VB.TextBox txtTope 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "5"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.TextBox txtApellido 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   315
         Left            =   3720
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox txtNroAfiliado 
         Height          =   315
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sexo"
         Height          =   615
         Left            =   5400
         TabIndex        =   17
         Top             =   1320
         Width           =   4815
         Begin VB.OptionButton optMasculino 
            Caption         =   "Masculino"
            Height          =   255
            Left            =   480
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optFemenino 
            Caption         =   "Femenino"
            Height          =   255
            Left            =   1920
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
      End
      Begin ControlesPOO.Combo cmbObraSocial 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin ControlesPOO.Combo cmbTipoDoc 
         Height          =   315
         Left            =   1320
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTope 
         Caption         =   "Tope Atenciones:"
         Height          =   195
         Left            =   45
         TabIndex        =   21
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Apellidos:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   24
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Doc:"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   23
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Doc:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   19
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Nº de afiliado:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Obra Social:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   6720
      TabIndex        =   13
      Top             =   5880
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8640
      TabIndex        =   14
      Top             =   5880
      Width           =   1695
   End
End
Attribute VB_Name = "frmABMAfiliadoExterno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar

Public Event NuevoAfiliado(pAfiliado As blcemi.AfiliadoExterno)
Public Event AfiliadoModificado(pAfiliado As blcemi.AfiliadoExterno)
Public Event AfiliadoEliminado(pAfiliado As blcemi.AfiliadoExterno)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mAfiliado As blcemi.AfiliadoExterno

Private mAfiliados As blcemi.AfiliadoExternoManager

Private WithEvents frmConsulta As frmConsultaGenerico
Attribute frmConsulta.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Dim mHistoriaClinica As New blcemi.HistoriaClinica
                mHistoriaClinica.Inicializar lvwAlergia.Coleccion, lvwEnfermedad.Coleccion, lvwMedicamento.Coleccion
                Dim mFechaNac As Date
                If optEdad.Value Then
                    mFechaNac = CDate("01/01/" + Trim(Str(Year(Date) - CInt(txtEdad))))
                Else
                    mFechaNac = dtpFechaNac.Value
                End If
                Set mAfiliado = mAfiliados.Nuevo(txtApellido, ctlDir.MiDireccion, mFechaNac, txtNroAfiliado, txtNombre, CLng(txtNroDoc), cmbObraSocial.SelectedItem, txtObservaciones, CInt(IIf(optMasculino.Value = True, "1", "0")), cmbTipoDoc.SelectedItem, CInt(txtTope), mHistoriaClinica, ctlTel.Telefonos, optEdad.Value)
                RaiseEvent NuevoAfiliado(mAfiliado)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            If DatosCorrectos Then
                LlenarObjeto
                mAfiliado.GuardarModificaciones
                RaiseEvent AfiliadoModificado(mAfiliado)
                Unload Me
            End If
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean
'COMPLETAR!!!
Dim msj As String
Dim msj2 As String 'para los datos no obligatorios
Dim msjDir As String

If Not TextBoxValidado(txtNroAfiliado, eString) Then msj = msj + "Ingrese el Numero de Afiliado." + vbCrLf
If Not TextBoxValidado(txtNombre, eString) Then msj = msj + "Ingrese el nombre del afiliado." + vbCrLf
If Not TextBoxValidado(txtApellido, eString) Then msj = msj + "Ingrese el apellido del afiliado." + vbCrLf

If CCFFGG.Configuracion.Requeridos.ExigirDNIAE Then
    If Not TextBoxValidado(txtNroDoc, eLong) Then msj = msj + "Ingrese el numero de documento." + vbCrLf
End If

If Not ctlDir.DireccionCompleta(msjDir) Then msj = msj + msjDir

If cmbObraSocial.SelectedItem Is Nothing Then msj = msj + "Seleccione una Obra Social." + vbCrLf
If cmbTipoDoc.SelectedItem Is Nothing Then msj = msj + "Seleccione un Tipo de Documento." + vbCrLf
'solo si es tipo edad controlo
If optEdad.Value Then If Not TextBoxValidado(txtEdad, eInteger) Then msj = msj + "Ingrese la edad." + vbCrLf

If CCFFGG.Configuracion.Requeridos.UsarTopeAtencAE Then
    If Not TextBoxValidado(txtTope, eString) Then msj = msj + "Ingrese un Tope de Atenciones." + vbCrLf
End If
    
If lvwAlergia.Coleccion.Count = 0 Then msj2 = msj2 + "Esta seguro que el afiliado no tiene alergias?" + vbCrLf
If lvwMedicamento.Coleccion.Count = 0 Then msj2 = msj2 + "Esta seguro que el afiliado no tiene recetados medicamentos?" + vbCrLf
If lvwEnfermedad.Coleccion.Count = 0 Then msj2 = msj2 + "Esta seguro que el afiliado no tiene enfermedades?" + vbCrLf

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
    If Tipo = etMODIFICACION Then
        mAfiliado.HistoriaClinica.CancelChanges
        Set mAfiliado.Telefonos = Nothing 'asi lo recarga  y me evito el beginedit
    End If
    Unload Me
End Sub

Public Sub Nuevo(pAfiliados As blcemi.AfiliadoExternoManager)
    
    Tipo = etALTA
    Set mAfiliados = pAfiliados
    Me.Show
    Me.Caption = "Nuevo Afiliado Externo"
    
    Set lvwEnfermedad.Coleccion = New blcemi.EnfermedadManager
    Set lvwAlergia.Coleccion = New blcemi.AlergiaManager
    Set lvwMedicamento.Coleccion = New blcemi.MedicamentoManager
   
    Set ctlDir.MiDireccion = New blcemi.Direccion
    Set ctlTel.Telefonos = New blcemi.TelefonoManager
End Sub

Public Sub Modificar(pAfiliado As blcemi.AfiliadoExterno)
'implementar
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar Afiliado Externo"
Set mAfiliado = pAfiliado
mAfiliado.HistoriaClinica.BeginEdit
LlenarCampos
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
Me.Caption = "Eliminar Afiliado"

End Sub

Public Sub VerDatos(pAfiliado As blcemi.AfiliadoExterno)
'implementar
Tipo = etCONSULTA
Me.Show
Set mAfiliado = pAfiliado
LlenarCampos
Me.Caption = "Consultar Detalles Afiliado Externo"
cmdQuitarAlergia.Enabled = False
cmdAgregarAlergia.Enabled = False
cmdAgregarEnfermedad.Enabled = False
cmdAgregarMedicamento.Enabled = False
cmdQuitarEnfermedad.Enabled = False
cmdQuitarMedicamento.Enabled = False
ctlDir.SoloConsulta = True
cmbObraSocial.Enabled = False
cmbTipoDoc.Enabled = False
BloquearTextBoxes True, Me.Controls

cmdAceptar.Visible = False
cmdCancelar.Caption = "Cerrar"

End Sub

'se la utiliza para modificar
Private Sub LlenarObjeto()
    mAfiliado.Apellido = txtApellido
    mAfiliado.Nombre = txtNombre
    mAfiliado.NroDoc = txtNroDoc
    Set mAfiliado.TipoDoc = cmbTipoDoc.SelectedItem
    Set mAfiliado.Direccion = ctlDir.MiDireccion
    mAfiliado.NroAfiliado = txtNroAfiliado
    mAfiliado.FechaNacimiento = dtpFechaNac.Value
    mAfiliado.Sexo = CInt(IIf(optMasculino.Value = True, "1", "0"))
    
    'HistoriaClinica
    'Set lvwAlergia.Coleccion = mAfiliado.HistoriaClinica.Alergias
    'Set lvwEnfermedad.Coleccion = mAfiliado.HistoriaClinica.Enfermedades
    'Set lvwMedicamento.Coleccion = mAfiliado.HistoriaClinica.Medicamentos
    
    mAfiliado.Observaciones = txtObservaciones
    Set mAfiliado.ObraSocial = cmbObraSocial.SelectedItem
         
    mAfiliado.TopeAtenciones = txtTope
End Sub

Private Sub LlenarCampos()
    txtApellido = mAfiliado.Apellido
    txtNombre = mAfiliado.Nombre
    txtNroDoc = mAfiliado.NroDoc
    Set cmbTipoDoc.SelectedItem = mAfiliado.TipoDoc
    Set ctlDir.MiDireccion = mAfiliado.Direccion
    txtNroAfiliado = mAfiliado.NroAfiliado
   
    optMasculino.Value = IIf(mAfiliado.Sexo = 1, True, False)
    
    'HistoriaClinica
    Set lvwAlergia.Coleccion = mAfiliado.HistoriaClinica.Alergias
    Set lvwEnfermedad.Coleccion = mAfiliado.HistoriaClinica.Enfermedades
    Set lvwMedicamento.Coleccion = mAfiliado.HistoriaClinica.Medicamentos
    
    txtObservaciones = mAfiliado.Observaciones
    Set cmbObraSocial.SelectedItem = mAfiliado.ObraSocial
    
    Set ctlTel.Telefonos = mAfiliado.Telefonos
    'mAfiliado.Atenciones
    
    txtTope = mAfiliado.TopeAtenciones
    optEdad.Value = mAfiliado.EsEdad
    If mAfiliado.EsEdad Then
        txtEdad.Text = mAfiliado.Edad
    Else
        dtpFechaNac.Value = mAfiliado.FechaNacimiento
    End If
End Sub



Private Sub Form_Load()

    'levanta un error si quiere usar el metodo show
    If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
    
    Set cmbTipoDoc.Coleccion = GBL.TiposDocumentoGBL
    Set cmbTipoDoc.SelectedItem = GBL.TiposDocumentoGBL.Item(1) 'para q seleccione dni predeterminado
    Set cmbObraSocial.Coleccion = GBL.ObrasSocialesGBL
    'setear icono form
    Set Me.Icon = MDI.Icon

    Set cmdAgregarAlergia.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdAgregarEnfermedad.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdAgregarMedicamento.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdQuitarAlergia.Picture = MDI.il32.ListImages("eliminar").Picture
    Set cmdQuitarEnfermedad.Picture = MDI.il32.ListImages("eliminar").Picture
    Set cmdQuitarMedicamento.Picture = MDI.il32.ListImages("eliminar").Picture
    
    Set ctlTel.BotonAgregar.Picture = MDI.il32.ListImages("agregar").Picture
    Set ctlTel.BotonEliminar.Picture = MDI.il32.ListImages("eliminar").Picture
    Set ctlTel.BotonModificar.Picture = MDI.il32.ListImages("modificar").Picture
    
    InicializarDireccion ctlDir
    AplicarConfiguracion
End Sub

Public Function GetHelpContext() As String
    'habilitar cuando agregue estas paginas a la ayuda
    GetHelpContext = "" '"abmafiliadoexterno"
End Function

Public Sub Refrescar()
    cmbObraSocial.Refresh
    lvwAlergia.Refresh
    lvwMedicamento.Refresh
    lvwEnfermedad.Refresh
    AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()
    lvwAlergia.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    lvwEnfermedad.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    lvwMedicamento.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    If Not CCFFGG.Configuracion.Requeridos.UsarTopeAtencAP Then
        txtTope.Visible = False
        lblTope.Visible = False
    End If
End Sub


'------------manejo de alergias, enfermedades y medicamentos--------------------

Private Sub cmdAgregarEnfermedad_Click()
    Set frmConsulta = New frmConsultaGenerico
    frmConsulta.Consultar GBL.EnfermedadesGBL, "Seleccione una a varias enfermedades", "Nueva Enfermedad", UsuarioActual.Permisos.Can(blcemi.AltaEnfermedad), UsuarioActual.Permisos.Can(blcemi.ModificacionEnfermedad), UsuarioActual.Permisos.Can(blcemi.BajaEnfermedad), , etConRetorno
End Sub

Private Sub cmdAgregarAlergia_Click()
    Set frmConsulta = New frmConsultaGenerico
    frmConsulta.Consultar GBL.AlergiasGBL, "Seleccione una a varias alergia", "Nueva Alergia", UsuarioActual.Permisos.Can(blcemi.AltaAlergia), UsuarioActual.Permisos.Can(blcemi.ModificacionAlergia), UsuarioActual.Permisos.Can(blcemi.BajaAlergia), , etConRetorno
End Sub

Private Sub cmdAgregarMedicamento_Click()
    Set frmConsulta = New frmConsultaGenerico
    frmConsulta.Consultar GBL.MedicamentosGBL, "Seleccione uno o varios medicamentos", "Nuevo Medicamento", UsuarioActual.Permisos.Can(blcemi.AltaMedicamento), UsuarioActual.Permisos.Can(blcemi.ModificacionMedicamento), UsuarioActual.Permisos.Can(blcemi.BajaMedicamento), , etConRetorno
End Sub

Private Sub cmdQuitarAlergia_Click()
    If Not lvwAlergia.SelectedItem Is Nothing Then
        lvwAlergia.Coleccion.Remove lvwAlergia.SelectedItem.id
        lvwAlergia.Refresh
    End If
End Sub

Private Sub cmdQuitarEnfermedad_Click()
    If Not lvwEnfermedad.SelectedItem Is Nothing Then
        lvwEnfermedad.Coleccion.Remove lvwEnfermedad.SelectedItem.id
        lvwEnfermedad.Refresh
    End If
End Sub

Private Sub cmdQuitarMedicamento_Click()
    If Not lvwMedicamento.SelectedItem Is Nothing Then
        lvwMedicamento.Coleccion.Remove lvwMedicamento.SelectedItem.id
        lvwMedicamento.Refresh
    End If
End Sub

Private Sub frmConsulta_ItemSeleccionado(pItem As Object)
    'aca pueden venir objetos de tipo alergia, enfermedad o medicamento
    AgregarElemento(pItem).Refresh
    
End Sub

Private Sub frmConsulta_ItemsSeleccionados(pColItems As Collection)
    Dim lvw
    Dim v As Object
    For Each v In pColItems
         Set lvw = AgregarElemento(v)
    Next
    lvw.Refresh
End Sub

Private Function AgregarElemento(pItem As Object) As Object

    If TypeOf pItem Is blcemi.Alergia Then
        lvwAlergia.Coleccion.AddItem pItem
        Set AgregarElemento = lvwAlergia
    ElseIf TypeOf pItem Is blcemi.Enfermedad Then
        lvwEnfermedad.Coleccion.AddItem pItem
        Set AgregarElemento = lvwEnfermedad
    ElseIf TypeOf pItem Is blcemi.Medicamento Then
        lvwMedicamento.Coleccion.AddItem pItem
        Set AgregarElemento = lvwMedicamento
    End If

End Function

'--------tabOrder---------------------------

Private Sub lvwEnfermedad_GotFocus()
sTabHC.Tab = 0
End Sub

Private Sub lvwAlergia_GotFocus()
sTabHC.Tab = 1
End Sub

Private Sub lvwMedicamento_GotFocus()
sTabHC.Tab = 2
End Sub

Private Sub ctlTel_GotFocus()
SSTab1.Tab = 1
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

Private Sub txtTope_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub
