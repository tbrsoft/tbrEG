VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMAfiliado 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12735
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   12735
   Begin VB.Frame Frame3 
      Caption         =   "Datos de Afiliacion"
      Height          =   3135
      Left            =   6360
      TabIndex        =   46
      Top             =   0
      Width           =   6255
      Begin TabDlg.SSTab sTabAfiliacion 
         Height          =   2655
         Left            =   120
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4683
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Principal"
         TabPicture(0)   =   "frmABMAfiliado.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "txtImporte"
         Tab(0).Control(1)=   "txtTope"
         Tab(0).Control(2)=   "cmbCobrador"
         Tab(0).Control(3)=   "dtpInicio"
         Tab(0).Control(4)=   "dtpInscripcion"
         Tab(0).Control(5)=   "lblImporte"
         Tab(0).Control(6)=   "lblInscripcion"
         Tab(0).Control(7)=   "lblInicio"
         Tab(0).Control(8)=   "lblTope"
         Tab(0).Control(9)=   "lblCobrador"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Personas a Cargo"
         TabPicture(1)   =   "frmABMAfiliado.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "lvwACargo"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdAgregarACargo"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdEditarACargo"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "cmdVerDetallesACargo"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "cmdEliminarACargo"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).ControlCount=   5
         TabCaption(2)   =   "Relacion con el Titular"
         TabPicture(2)   =   "frmABMAfiliado.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label13"
         Tab(2).Control(1)=   "cmbParentezco"
         Tab(2).ControlCount=   2
         Begin ALCemi.GraphicButton cmdEliminarACargo 
            Height          =   495
            Left            =   5400
            TabIndex        =   66
            ToolTipText     =   "Dar de baja una persona a cargo"
            Top             =   2040
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin ALCemi.GraphicButton cmdVerDetallesACargo 
            Height          =   495
            Left            =   5400
            TabIndex        =   65
            ToolTipText     =   "Ver detalles de la presona a cargo."
            Top             =   1480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin ALCemi.GraphicButton cmdEditarACargo 
            Height          =   495
            Left            =   5400
            TabIndex        =   64
            ToolTipText     =   "Modificar los datos de la persona a cargo"
            Top             =   920
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin ALCemi.GraphicButton cmdAgregarACargo 
            Height          =   495
            Left            =   5400
            TabIndex        =   63
            ToolTipText     =   "Agregar una persona a cargo"
            Top             =   360
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin VB.TextBox txtImporte 
            Height          =   285
            Left            =   -69960
            TabIndex        =   19
            Top             =   1680
            Width           =   735
         End
         Begin ControlesPOO.Combo cmbParentezco 
            Height          =   315
            Left            =   -73680
            TabIndex        =   21
            Top             =   600
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            NuevoEnabled    =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.ListViewConsulta lvwACargo 
            Height          =   2175
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   3836
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
            NEncabezado0    =   "Nombre"
            MEncabezado0    =   "nombre"
            AEncabezado0    =   33
            NEncabezado1    =   "Apellido"
            MEncabezado1    =   "apellido"
            AEncabezado1    =   33
            NEncabezado2    =   "Parentezco"
            MEncabezado2    =   "parentezco"
            AEncabezado2    =   34
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
         Begin VB.TextBox txtTope 
            Height          =   285
            Left            =   -69960
            TabIndex        =   18
            Text            =   "5"
            Top             =   1140
            Width           =   735
         End
         Begin ControlesPOO.Combo cmbCobrador 
            Height          =   315
            Left            =   -73440
            TabIndex        =   15
            Top             =   660
            Width           =   4215
            _ExtentX        =   7435
            _ExtentY        =   556
            AtributoAMostrar=   "nombreCompleto"
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpInicio 
            Height          =   315
            Left            =   -73440
            TabIndex        =   17
            Top             =   1620
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   45678593
            CurrentDate     =   39293
         End
         Begin MSComCtl2.DTPicker dtpInscripcion 
            Height          =   315
            Left            =   -73440
            TabIndex        =   16
            Top             =   1140
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   45678593
            CurrentDate     =   39293
         End
         Begin VB.Label lblImporte 
            Caption         =   "Importe:"
            Height          =   195
            Left            =   -70680
            TabIndex        =   54
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label13 
            Caption         =   "Parentezco:"
            Height          =   195
            Left            =   -74760
            TabIndex        =   52
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblInscripcion 
            Caption         =   "Fecha Inscripcion:"
            Height          =   195
            Left            =   -74880
            TabIndex        =   51
            Top             =   1140
            Width           =   1305
         End
         Begin VB.Label lblInicio 
            Caption         =   "Inicio Prestacion:"
            Height          =   195
            Left            =   -74760
            TabIndex        =   50
            Top             =   1620
            Width           =   1215
         End
         Begin VB.Label lblTope 
            Caption         =   "Tope Atenciones:"
            Height          =   195
            Left            =   -71280
            TabIndex        =   49
            Top             =   1140
            Width           =   1260
         End
         Begin VB.Label lblCobrador 
            Caption         =   "Cobrador:"
            Height          =   195
            Left            =   -74280
            TabIndex        =   48
            Top             =   660
            Width           =   690
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Historia Clinica"
      Height          =   2535
      Left            =   6360
      TabIndex        =   44
      Top             =   3360
      Width           =   6255
      Begin TabDlg.SSTab sTabHistoriaClinica 
         Height          =   2175
         Left            =   120
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   3836
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Enfermedades"
         TabPicture(0)   =   "frmABMAfiliado.frx":0054
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdAgregarEnfermedad"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lvwEnfermedad"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdQuitarEnfermedad"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Alergias"
         TabPicture(1)   =   "frmABMAfiliado.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lvwAlergia"
         Tab(1).Control(1)=   "cmdAgregarAlergia"
         Tab(1).Control(2)=   "cmdQuitarAlergia"
         Tab(1).ControlCount=   3
         TabCaption(2)   =   "Medicamentos"
         TabPicture(2)   =   "frmABMAfiliado.frx":008C
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "lvwMedicamento"
         Tab(2).Control(1)=   "cmdAgregarMedicamento"
         Tab(2).Control(2)=   "cmdQuitarMedicamento"
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Observaciones"
         TabPicture(3)   =   "frmABMAfiliado.frx":00A8
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "txtObservaciones"
         Tab(3).ControlCount=   1
         Begin ALCemi.GraphicButton cmdQuitarMedicamento 
            Height          =   495
            Left            =   -69600
            TabIndex        =   62
            ToolTipText     =   "Quita el medicamento seleccionado"
            Top             =   1080
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin ALCemi.GraphicButton cmdAgregarMedicamento 
            Height          =   495
            Left            =   -69600
            TabIndex        =   61
            ToolTipText     =   "Agregar uno o varios medicamentos"
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin ALCemi.GraphicButton cmdQuitarAlergia 
            Height          =   495
            Left            =   -69600
            TabIndex        =   60
            ToolTipText     =   "Quita la alergia seleccionada"
            Top             =   1080
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin ALCemi.GraphicButton cmdAgregarAlergia 
            Height          =   495
            Left            =   -69600
            TabIndex        =   59
            ToolTipText     =   "Agregar una o varias alergias"
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin ALCemi.GraphicButton cmdQuitarEnfermedad 
            Height          =   495
            Left            =   5400
            TabIndex        =   58
            ToolTipText     =   "asdasd"
            Top             =   1080
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   1575
            Left            =   -74880
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   480
            Width           =   5775
         End
         Begin ControlesPOO.ListViewConsulta lvwEnfermedad 
            Height          =   1575
            Left            =   120
            TabIndex        =   22
            Top             =   480
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   2778
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
            Height          =   1575
            Left            =   -74880
            TabIndex        =   23
            Top             =   480
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   2778
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
            Height          =   1575
            Left            =   -74880
            TabIndex        =   24
            Top             =   480
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   2778
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
         Begin ALCemi.GraphicButton cmdAgregarEnfermedad 
            Height          =   495
            Left            =   5400
            TabIndex        =   57
            ToolTipText     =   "Agregar enfermedades"
            Top             =   480
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   873
         End
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   10920
      TabIndex        =   27
      Top             =   6000
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   9000
      TabIndex        =   26
      Top             =   6000
      Width           =   1695
   End
   Begin VB.Frame fraDatos 
      Height          =   6495
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtNroAfiliadoACargo 
         Height          =   315
         Left            =   4200
         MaxLength       =   5
         TabIndex        =   56
         Text            =   "0"
         Top             =   360
         Width           =   855
      End
      Begin ControlesPOO.Combo cmbEstadoCivil 
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   2160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin ControlesPOO.Combo cmbObraSocial 
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Top             =   2880
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin ControlesPOO.Combo cmbTipoDoc 
         Height          =   315
         Left            =   1320
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin TabDlg.SSTab sTabDatos 
         Height          =   3015
         Left            =   120
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3360
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Direccion"
         TabPicture(0)   =   "frmABMAfiliado.frx":00C4
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "ctlDir"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Telefonos"
         TabPicture(1)   =   "frmABMAfiliado.frx":00E0
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "ctlTel"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Vehiculo"
         TabPicture(2)   =   "frmABMAfiliado.frx":00FC
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraVehiculo"
         Tab(2).ControlCount=   1
         Begin ALCemi.ctlTelefonos ctlTel 
            Height          =   2595
            Left            =   -74880
            TabIndex        =   11
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   4577
            Caption         =   ""
            SoloConsulta    =   0   'False
         End
         Begin ALCemi.ctlDireccion ctlDir 
            Height          =   2565
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   3995
            ProvinciaVisible=   0   'False
            Caption         =   ""
            CanDragDrop     =   0   'False
            SoloConsulta    =   0   'False
            EntrecallesVisible=   -1  'True
         End
         Begin VB.Frame fraVehiculo 
            Height          =   2595
            Left            =   -74880
            TabIndex        =   39
            Top             =   360
            Width           =   5775
            Begin VB.TextBox txtPatente 
               Height          =   315
               Left            =   960
               TabIndex        =   14
               Top             =   1080
               Width           =   4695
            End
            Begin VB.TextBox txtModelo 
               Height          =   315
               Left            =   960
               TabIndex        =   13
               Top             =   720
               Width           =   4695
            End
            Begin VB.TextBox txtMarca 
               Height          =   315
               Left            =   960
               TabIndex        =   12
               Top             =   360
               Width           =   4695
            End
            Begin VB.Label Label9 
               Caption         =   "Patente:"
               Height          =   255
               Left            =   240
               TabIndex        =   42
               Top             =   1080
               Width           =   615
            End
            Begin VB.Label Label8 
               Caption         =   "Modelo:"
               Height          =   255
               Left            =   240
               TabIndex        =   41
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label7 
               Caption         =   "Marca:"
               Height          =   255
               Left            =   240
               TabIndex        =   40
               Top             =   360
               Width           =   615
            End
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Sexo"
         Height          =   615
         Left            =   3240
         TabIndex        =   37
         Top             =   1800
         Width           =   2895
         Begin VB.OptionButton optFemenino 
            Caption         =   "Femenino"
            Height          =   255
            Left            =   1560
            TabIndex        =   6
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optMasculino 
            Caption         =   "Masculino"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin ControlesPOO.Combo cmbOcupacion 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   2520
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   556
         NuevoEnabled    =   -1  'True
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker dtpFechaNac 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45678593
         CurrentDate     =   39292
      End
      Begin VB.TextBox txtNroAfiliado 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Text            =   "1"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtNroDoc 
         Height          =   315
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtApellido 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label lblNroAfCargo 
         Caption         =   "Nº A Cargo: "
         Height          =   255
         Left            =   3240
         TabIndex        =   55
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Obra Social:"
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   2880
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "Estado Civil:"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ocupacion:"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   2520
         Width           =   825
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Nac:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         ToolTipText     =   "Fecha Nacimiento:"
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Nº de afiliado:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Doc:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   32
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nro Doc:"
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   31
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Apellidos:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   30
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmABMAfiliado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar
Option Explicit

Public Event NuevoAfiliado(pAfiliado As blcemi.Afiliado)
Public Event AfiliadoModificado(pAfiliado As blcemi.Afiliado)
Public Event AfiliadoEliminado(pAfiliado As blcemi.Afiliado)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mTipoAfiliado As blcemi.eTipoAfiliado
Private mAfiliado As blcemi.Afiliado
Private mAfiliados As blcemi.AfiliadoManager
Private mVehiculo As blcemi.Vehiculo
Private frmAbmParent As frmABMAfiliado 'esto es porq no puedo agarrar directamente los eventos de este msmo form

Private WithEvents frmConsOS As frmConsultarObraSocial
Attribute frmConsOS.VB_VarHelpID = -1
Private WithEvents frmConsulta As frmConsultaGenerico
Attribute frmConsulta.VB_VarHelpID = -1

Public Property Get TipoABM() As eTipoAMB
    TipoABM = Tipo
End Property

Public Property Get AfiliadoTitular() As blcemi.Afiliado
    Set AfiliadoTitular = mAfiliado
End Property

Private Sub cmbObraSocial_NuevoSeleccionado()
    Set frmConsOS = New frmConsultarObraSocial
    frmConsOS.Consultar GBL.ObrasSocialesGBL, etConRetorno
End Sub

Private Sub frmConsOS_ObraSocialSeleccionada(pObraSocial As blcemi.ObraSocial)
    cmbObraSocial.Refresh
    Set cmbObraSocial.SelectedItem = pObraSocial
End Sub

Private Sub cmbOcupacion_NuevoSeleccionado()
    Dim aux As String
    
    aux = frmABMGenerico.Nuevo("Ingrese la ocupacion:")
    
    If aux <> "" Then
        Dim ocAux As blcemi.Ocupacion
        Dim newOc As blcemi.Ocupacion
        Set ocAux = GBL.OcupacionesGBL.ItemByName(aux)
        If ocAux Is Nothing Then
            Set newOc = GBL.OcupacionesGBL.Nuevo(aux)
            Set cmbOcupacion.Coleccion = GBL.OcupacionesGBL
            Set cmbOcupacion.SelectedItem = newOc
        End If
    End If

End Sub

Private Sub cmbParentezco_NuevoSeleccionado()
Dim aux As String
    
    aux = frmABMGenerico.Nuevo("Ingrese el parentezco:")
    
    If aux <> "" Then
        Dim ocAux As blcemi.Parentezco
        Dim newOc As blcemi.Parentezco
        Set ocAux = GBL.ParentezcosGBL.ItemByName(aux)
        If ocAux Is Nothing Then
            Set newOc = GBL.ParentezcosGBL.Nuevo(aux)
            cmbParentezco.Refresh
            Set cmbParentezco.SelectedItem = newOc
        End If
    End If
End Sub

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Dim mHistoriaClinica As New blcemi.HistoriaClinica
                mHistoriaClinica.Inicializar lvwAlergia.Coleccion, lvwEnfermedad.Coleccion, lvwMedicamento.Coleccion
                    
                If mTipoAfiliado = blcemi.eTitular Then
                    If txtPatente <> "" And txtModelo <> "" And txtMarca <> "" Then
                        Set mVehiculo = New blcemi.Vehiculo
                        mVehiculo.Marca = txtMarca
                        mVehiculo.Modelo = txtModelo
                        mVehiculo.Patente = txtPatente
                    End If
                                                                                                                                                                                         
                    Dim afAux As blcemi.Afiliado
                    For Each afAux In lvwACargo.Coleccion
                        afAux.id = CLng(txtNroAfiliado) * 1000 + afAux.id
                    Next                                                                                                                                                               'ver!!!
                    Set mAfiliado = mAfiliados.Nuevo(txtApellido, cmbCobrador.SelectedItem, ctlDir.MiDireccion, cmbEstadoCivil.SelectedItem, dtpInscripcion.Value, dtpFechaNac.Value, CLng(txtNroAfiliado) * 1000, dtpInicio.Value, txtNombre, CLng(txtNroDoc), cmbObraSocial.SelectedItem, txtObservaciones, cmbOcupacion.SelectedItem, lvwACargo.Coleccion, CInt(IIf(optMasculino.Value = True, "1", "0")), ctlTel.Telefonos, cmbTipoDoc.SelectedItem, CInt(txtTope), mVehiculo, mHistoriaClinica, CCur(txtImporte))
                    RaiseEvent NuevoAfiliado(mAfiliado)
                Else
                    'si estoy modificando el afiliado titular, hago el insert del child en el acto
                    
                    If frmAbmParent.TipoABM = etMODIFICACION Then
                        Set mAfiliado = mAfiliados.NuevoACargo(txtApellido, ctlDir.MiDireccion, cmbEstadoCivil.SelectedItem, dtpInscripcion.Value, dtpFechaNac.Value, frmAbmParent.AfiliadoTitular.id + CLng(txtNroAfiliadoACargo), dtpInicio.Value, txtNombre, CLng(txtNroDoc), cmbObraSocial.SelectedItem, txtObservaciones, cmbOcupacion.SelectedItem, IIf(optMasculino.Value = True, 1, 0), ctlTel.Telefonos, cmbTipoDoc.SelectedItem, CInt(txtTope), cmbParentezco.SelectedItem, mHistoriaClinica, frmAbmParent.AfiliadoTitular)
                    Else 'sino ,es uno nuevo...
                        Set mAfiliado = mAfiliados.NuevoACargo(txtApellido, ctlDir.MiDireccion, cmbEstadoCivil.SelectedItem, dtpInscripcion.Value, dtpFechaNac.Value, CLng(txtNroAfiliadoACargo), dtpInicio.Value, txtNombre, CLng(txtNroDoc), cmbObraSocial.SelectedItem, txtObservaciones, cmbOcupacion.SelectedItem, IIf(optMasculino.Value = True, 1, 0), ctlTel.Telefonos, cmbTipoDoc.SelectedItem, CInt(txtTope), cmbParentezco.SelectedItem, mHistoriaClinica, Nothing)
                    End If
                    frmAbmParent.RefrescarLvwAfiliados
                End If
                                
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            If DatosCorrectos Then
                LlenarObjeto
                If mAfiliado.TipoAfiliado = blcemi.eTitular Then
                    mAfiliado.GuardarModificaciones
                    RaiseEvent AfiliadoModificado(mAfiliado)
                Else
                    mAfiliado.GuardarModificacionesACargo
                    'aca me fijo si lo modifican desde frmabm o dsde frmconsulta
                    If Not frmAbmParent Is Nothing Then
                        'frmabm
                        frmAbmParent.RefrescarLvwAfiliados
                    Else
                        'frmconsulta
                        RaiseEvent AfiliadoModificado(mAfiliado)
                    End If
                End If
                            
                Unload Me
            End If
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean
'aca verificar de alguna forma q el id sea valido
If Tipo = etALTA Then 'si es un alta controlo el id, no se puede modificar una vez asignado
    If TextBoxValidado(txtNroAfiliado, eLong) Then 'si es correcto el formato
        ' ya esta modifcado p el nuevo tipo de id
        If mAfiliados.ExisteId(CLng(txtNroAfiliado) * 1000 + CLng(txtNroAfiliadoACargo.Text)) Then
            MsgBox "El Numero de afiliado ingresado ya existe en la Base de Datos, por favor ingrese otro.", vbInformation + vbOKOnly
            txtNroAfiliado.SetFocus
        Else 'si el nro de afiliado no existe controlo los demas campos
            DatosCorrectos = datosCorrectosAuxiliar
        End If
    Else
        MsgBox "El Numero de afiliado ingresado no cumple con el formato establecido, por favor ingrese un numero valido.", vbInformation + vbOKOnly
        DatosCorrectos = False
    End If
Else 'si no es un alta, controlo todos los campos menos el od q no se puede modifcar
    DatosCorrectos = datosCorrectosAuxiliar
End If
End Function

Private Function datosCorrectosAuxiliar() As Boolean
Dim msj As String
Dim msj2 As String 'para los datos no obligatorios
Dim msjDir As String 'por las dudas tenga incompleta la direccion

If Not TextBoxValidado(txtNombre, eString) Then msj = msj + "Ingrese el nombre del afiliado." + vbCrLf
If Not TextBoxValidado(txtApellido, eString) Then msj = msj + "Ingrese el apellido del afiliado." + vbCrLf

If CCFFGG.Configuracion.Requeridos.ExigirDNIAP Then
    If Not TextBoxValidado(txtNroDoc, eLong) Then msj = msj + "Ingrese el numero de documento." + vbCrLf
End If

If Not ctlDir.DireccionCompleta(msjDir) Then msj = msj + msjDir

If cmbEstadoCivil.SelectedItem Is Nothing Then msj = msj + "Seleccione un Estado Civil" + vbCrLf
If cmbObraSocial.SelectedItem Is Nothing Then msj = msj + "Seleccione una Obra Social" + vbCrLf
If cmbOcupacion.SelectedItem Is Nothing Then msj = msj + "Seleccione una Ocupacion" + vbCrLf
If cmbTipoDoc.SelectedItem Is Nothing Then msj = msj + "Seleccione un Tipo de Documento" + vbCrLf

'solo afiliados a cargo
If mTipoAfiliado = blcemi.eACargo Then If cmbParentezco.SelectedItem Is Nothing Then msj = msj + "Seleccione un Parentezco" + vbCrLf

'solo para afiliado titular
If mTipoAfiliado = blcemi.eTitular Then
    If CCFFGG.Configuracion.Requeridos.UsarTopeAtencAP Then
        If Not TextBoxValidado(txtTope, eString) Then msj = msj + "Ingrese un Tope de Atenciones" + vbCrLf
        If Not TextBoxValidado(txtTope, eInteger) Then msj = msj + "Ingrese un Tope de Atenciones numerico." + vbCrLf
    End If
    If Not TextBoxValidado(txtImporte, eMoneda) Then msj = msj + "Ingrese un importe a cobrar." + vbCrLf
    If cmbCobrador.SelectedItem Is Nothing Then msj = msj + "Seleccione un cobrador." + vbCrLf
    
    If txtMarca <> "" Or txtModelo <> "" Or txtPatente <> "" Then
        If Not TextBoxValidado(txtPatente, ePatenteAutomovil) Then msj = msj + "Ingrese la patente del vehiculo." + vbCrLf
        If Not TextBoxValidado(txtMarca, eString) Then msj = msj + "Ingrese la marca del vehiculo." + vbCrLf
        If Not TextBoxValidado(txtModelo, eString) Then msj = msj + "Ingrese el modelo del vehiculo." + vbCrLf
    Else
        msj2 = msj2 + "Esta seguro que el afiliado no tiene vehiculo?" + vbCrLf
    End If
    'recordatorio de datos no cargados pero no obligatorios
    If lvwACargo.Coleccion.Count = 0 Then msj2 = msj2 + "Esta seguro que el afiliado no tiene personas a cargo?" + vbCrLf
    
End If
    
    If ctlTel.Telefonos.Count = 0 Then msj2 = "Esta seguro que el afiliado no tiene telefonos?" + vbCrLf
    If lvwAlergia.Coleccion.Count = 0 Then msj2 = msj2 + "Esta seguro que el afiliado no tiene alergias?" + vbCrLf
    If lvwMedicamento.Coleccion.Count = 0 Then msj2 = msj2 + "Esta seguro que el afiliado no tiene recetados medicamentos?" + vbCrLf
    If lvwEnfermedad.Coleccion.Count = 0 Then msj2 = msj2 + "Esta seguro que el afiliado no tiene enfermedades?" + vbCrLf
    'falta vehiculo

If msj2 <> "" And CCFFGG.Configuracion.Comportamiento.MostrarSugerenciasDatosFaltantes Then
    Dim res As VbMsgBoxResult
    res = MsgBox(msj2, vbOKCancel + vbQuestion)
    If res = vbCancel Then
        datosCorrectosAuxiliar = False
        Exit Function
    End If
End If

If msj = "" Then
    datosCorrectosAuxiliar = True
Else
    MsgBox "Faltan los siguientes datos:" + vbCrLf + msj, vbExclamation
    datosCorrectosAuxiliar = False
End If
End Function

Private Sub cmdCancelar_Click()
    If Tipo = etMODIFICACION Then
        Set mAfiliado.Telefonos = Nothing
        mAfiliado.HistoriaClinica.CancelChanges
        'mAfiliado.PersonasACargo.cancelchanges hacer!!!
    End If
    Unload Me
End Sub

Public Sub Nuevo(pAfiliados As blcemi.AfiliadoManager)
    
    Tipo = etALTA
    Set mAfiliados = pAfiliados
    Me.Show
    Me.Caption = "Nuevo Afiliado"
    Set ctlTel.Telefonos = New blcemi.TelefonoManager
    txtNroAfiliado = mAfiliados.GetUltimoIdTitular + 1
    dtpInicio.Value = Date
    dtpInscripcion.Value = Date
    
    Set lvwEnfermedad.Coleccion = New blcemi.EnfermedadManager
    Set lvwAlergia.Coleccion = New blcemi.AlergiaManager
    Set lvwMedicamento.Coleccion = New blcemi.MedicamentoManager

    Set lvwACargo.Coleccion = New blcemi.AfiliadoManager

    Set ctlDir.MiDireccion = New blcemi.Direccion
    
    mTipoAfiliado = blcemi.eTitular
    mostrarTabs
End Sub
                                                             'esto es porq no se pueden interceptar los evento dentro del mismo form q los produce
Public Sub NuevoAfiliadoACargo(pAfiliados As blcemi.AfiliadoManager, frmParent As frmABMAfiliado, pDireccionParent As blcemi.Direccion)
    Tipo = etALTA
    Me.Show
    Set frmAbmParent = frmParent
    Me.Caption = "Nuevo Afiliado"
    
    Set mAfiliados = pAfiliados
    
    dtpInicio.Value = Date
    dtpInscripcion.Value = Date
    
    txtNroAfiliado = frmAbmParent.txtNroAfiliado.Text
   ' txtNroAfiliado = mAfiliados.GetUltimoId + mAfiliados.Count + 2 'mas o menos estimado
    If frmParent.AfiliadoTitular Is Nothing Then
        txtNroAfiliadoACargo = 1
    Else
        Dim idAux As Long
        idAux = mAfiliados.GetUltimoIdACargo(frmParent.AfiliadoTitular.id)
        txtNroAfiliadoACargo = IIf(mAfiliados.Count < idAux, idAux + 1, mAfiliados.Count + 1)
    End If
    Set ctlTel.Telefonos = New blcemi.TelefonoManager
    
    Set lvwEnfermedad.Coleccion = New blcemi.EnfermedadManager
    Set lvwAlergia.Coleccion = New blcemi.AlergiaManager
    Set lvwMedicamento.Coleccion = New blcemi.MedicamentoManager
    
    mTipoAfiliado = blcemi.eACargo
    
    mostrarTabs
   
   'clono la direccion de parent, despues me fijo si es la misma
    Set ctlDir.MiDireccion = pDireccionParent.Clone

End Sub

Private Sub mostrarTabs()
    If mTipoAfiliado = blcemi.eTitular Then
        sTabAfiliacion.TabVisible(2) = False
        sTabAfiliacion.TabVisible(1) = True
        lblNroAfCargo.Visible = False
        txtNroAfiliadoACargo.Visible = False
        'si no usa tope aca lo oculto
        If Not CCFFGG.Configuracion.Requeridos.UsarTopeAtencAP Then
            txtTope.Visible = False
            lblTope.Visible = False
        End If
    Else
        txtNroAfiliado.Locked = True
        sTabAfiliacion.TabVisible(0) = True
        sTabAfiliacion.TabVisible(1) = False
        sTabAfiliacion.TabVisible(2) = True
        
        sTabDatos.TabVisible(2) = False 'q no se vea vehiculo
        
        'todo lo q sigue es para que se vean solo la fecha de inicio de prestacion
        'y la fecha de inscripcion
        
        lblCobrador.Visible = False
        cmbCobrador.Visible = False
        lblTope.Visible = False
        txtTope.Visible = False
'        udTope.Visible = False
        txtImporte.Visible = False
        lblImporte.Visible = False
                
        lblInicio.Top = lblInscripcion.Top
        lblInscripcion.Top = cmbCobrador.Top
        dtpInscripcion.Top = lblInscripcion.Top
        dtpInicio.Top = lblInicio.Top
        
    End If
End Sub

Public Sub Modificar(pAfiliado As blcemi.Afiliado, Optional frmParent As frmABMAfiliado)
'implementar
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar Afiliado"
Set frmAbmParent = frmParent
Set mAfiliado = pAfiliado
mTipoAfiliado = mAfiliado.TipoAfiliado
mAfiliado.HistoriaClinica.BeginEdit
'mAfiliado.PersonasACargo.beginedit HACER!!!!
txtNroAfiliado.Locked = True 'no se puede modificar el nroAfiliado
LlenarCampos
mostrarTabs
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
Me.Caption = "Eliminar Afiliado"

End Sub

Public Sub VerDatos(pAfiliado As blcemi.Afiliado)
Tipo = etCONSULTA
Me.Show
Set mAfiliado = pAfiliado
mTipoAfiliado = mAfiliado.TipoAfiliado
ctlTel.SoloConsulta = True
ctlDir.SoloConsulta = True
'desabilito todos los botones agregar, etc
cmdAgregarACargo.Enabled = False
cmdQuitarAlergia.Enabled = False
cmdAgregarAlergia.Enabled = False
cmdAgregarEnfermedad.Enabled = False
cmdAgregarMedicamento.Enabled = False
cmdQuitarEnfermedad.Enabled = False
cmdQuitarMedicamento.Enabled = False
cmdEliminarACargo.Enabled = False
cmdEditarACargo.Enabled = False

LlenarCampos
BloquearTextBoxes True, Me.Controls
mostrarTabs
cmdAceptar.Visible = False
cmdCancelar.Caption = "Cerrar"
Me.Caption = "Ver detalles del Afiliado"
End Sub

'se la utiliza para modificar
Private Sub LlenarObjeto()
    mAfiliado.Apellido = txtApellido
    mAfiliado.Nombre = txtNombre
    mAfiliado.NroDoc = txtNroDoc
    Set mAfiliado.TipoDoc = cmbTipoDoc.SelectedItem
    Set mAfiliado.Direccion = ctlDir.MiDireccion
    'mAfiliado.id = txtNroAfiliado
    mAfiliado.FechaNacimiento = dtpFechaNac.Value
    Set mAfiliado.Ocupacion = cmbOcupacion.SelectedItem
    mAfiliado.Sexo = CInt(IIf(optMasculino.Value = True, "1", "0"))
    Set mAfiliado.EstadoCivil = cmbEstadoCivil.SelectedItem
   ' Set ctlTel.Telefonos = mAfiliado.Telefonos
    
    'HistoriaClinica
    'Set lvwAlergia.Coleccion = mAfiliado.HistoriaClinica.Alergias
    'Set lvwEnfermedad.Coleccion = mAfiliado.HistoriaClinica.Enfermedades
    'Set lvwMedicamento.Coleccion = mAfiliado.HistoriaClinica.Medicamentos
    
    mAfiliado.Observaciones = txtObservaciones
    mAfiliado.FechaInscripcion = dtpInscripcion.Value
    mAfiliado.InicioPrestacion = dtpInicio
    Set mAfiliado.ObraSocial = cmbObraSocial.SelectedItem
         
    If mTipoAfiliado = blcemi.eTitular Then
        If txtPatente <> "" And txtModelo <> "" And txtMarca <> "" Then
            If mAfiliado.Vehiculo Is Nothing Then Set mAfiliado.Vehiculo = New blcemi.Vehiculo
            mAfiliado.Vehiculo.Marca = txtMarca
            mAfiliado.Vehiculo.Modelo = txtModelo
            mAfiliado.Vehiculo.Patente = txtPatente
        End If
        'mAfiliado.PersonasACargo
        Set mAfiliado.Cobrador = cmbCobrador.SelectedItem
        mAfiliado.TopeAtenciones = txtTope
        
    Else
        'no deberia estar, lo dejo para acordarme...
        'mAfiliado.Parent
        Set mAfiliado.Parentezco = cmbParentezco.SelectedItem
    End If
End Sub

Private Sub LlenarCampos()
    txtApellido = mAfiliado.Apellido
    txtNombre = mAfiliado.Nombre
    txtNroDoc = mAfiliado.NroDoc
    Set cmbTipoDoc.SelectedItem = mAfiliado.TipoDoc
    Set ctlDir.MiDireccion = mAfiliado.Direccion
    
    dtpFechaNac.Value = mAfiliado.FechaNacimiento
    Set cmbOcupacion.SelectedItem = mAfiliado.Ocupacion
    optMasculino.Value = IIf(mAfiliado.Sexo = 1, True, False)
    Set cmbEstadoCivil.SelectedItem = mAfiliado.EstadoCivil
    Set ctlTel.Telefonos = mAfiliado.Telefonos
    
    'HistoriaClinica
    Set lvwAlergia.Coleccion = mAfiliado.HistoriaClinica.Alergias
    Set lvwEnfermedad.Coleccion = mAfiliado.HistoriaClinica.Enfermedades
    Set lvwMedicamento.Coleccion = mAfiliado.HistoriaClinica.Medicamentos
    
    txtObservaciones = mAfiliado.Observaciones
    dtpInscripcion.Value = mAfiliado.FechaInscripcion
    dtpInicio = mAfiliado.InicioPrestacion
    Set cmbObraSocial.SelectedItem = mAfiliado.ObraSocial
    
    'mAfiliado.Atenciones
    
    If mTipoAfiliado = blcemi.eTitular Then
        If Not mAfiliado.Vehiculo Is Nothing Then
            txtMarca = mAfiliado.Vehiculo.Marca
            txtModelo = mAfiliado.Vehiculo.Modelo
            txtPatente = mAfiliado.Vehiculo.Patente
        End If
        Set lvwACargo.Coleccion = mAfiliado.PersonasACargo
        Set cmbCobrador.SelectedItem = mAfiliado.Cobrador
        txtTope = mAfiliado.TopeAtenciones
        txtImporte = mAfiliado.Importe
        'mAfiliado.Pagos
        txtNroAfiliado = mAfiliado.IdF
    Else
        txtNroAfiliado = mAfiliado.Parent.IdF
        txtNroAfiliadoACargo = mAfiliado.IdF
        Set cmbParentezco.SelectedItem = mAfiliado.Parentezco
    End If
    
End Sub

Private Sub cmdEditarACargo_Click()
If Not lvwACargo.SelectedItem Is Nothing Then
    frmABMAfiliado.Modificar lvwACargo.SelectedItem, Me
End If
End Sub

Private Sub cmdAgregarACargo_Click()
    frmABMAfiliado.NuevoAfiliadoACargo lvwACargo.Coleccion, Me, ctlDir.MiDireccion
End Sub

Private Sub cmdVerDetallesACargo_Click()
If Not lvwACargo.SelectedItem Is Nothing Then
    frmABMAfiliado.VerDatos lvwACargo.SelectedItem
End If
End Sub

Private Sub cmdEliminarACargo_Click()
If Not lvwACargo.SelectedItem Is Nothing Then
    If MsgBox("Esta seguro que desea dar de baja al afiliado?", vbQuestion + vbYesNo) = vbYes Then
        GBL.AfiliadosGBL.DarItemDeBaja lvwACargo.SelectedItem.id
        lvwACargo.Coleccion.Remove lvwACargo.SelectedItem.id
        Me.Refrescar
    End If
End If
End Sub

Private Sub Form_Load()

'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."

Set cmbTipoDoc.Coleccion = GBL.TiposDocumentoGBL
Set cmbTipoDoc.SelectedItem = GBL.TiposDocumentoGBL.Item(1) 'para q seleccione dni predeterminado
Set cmbEstadoCivil.Coleccion = GBL.EstadosCivilesGBL
Set cmbOcupacion.Coleccion = GBL.OcupacionesGBL
Set cmbParentezco.Coleccion = GBL.ParentezcosGBL
Set cmbCobrador.Coleccion = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eCobrador)
Set cmbObraSocial.Coleccion = GBL.ObrasSocialesGBL

'setear icono form
Set Me.Icon = MDI.Icon

Set cmdAgregarAlergia.Picture = MDI.il32.ListImages("agregar").Picture
Set cmdAgregarEnfermedad.Picture = MDI.il32.ListImages("agregar").Picture
Set cmdAgregarMedicamento.Picture = MDI.il32.ListImages("agregar").Picture
Set cmdAgregarACargo.Picture = MDI.il32.ListImages("agregar").Picture
Set cmdQuitarAlergia.Picture = MDI.il32.ListImages("eliminar").Picture
Set cmdQuitarEnfermedad.Picture = MDI.il32.ListImages("eliminar").Picture
Set cmdQuitarMedicamento.Picture = MDI.il32.ListImages("eliminar").Picture
Set cmdEliminarACargo.Picture = MDI.il32.ListImages("eliminar").Picture
Set cmdEditarACargo.Picture = MDI.il32.ListImages("modificar").Picture
Set cmdVerDetallesACargo.Picture = MDI.il32.ListImages("detalles").Picture

Set ctlTel.BotonAgregar.Picture = MDI.il32.ListImages("agregar").Picture
Set ctlTel.BotonEliminar.Picture = MDI.il32.ListImages("eliminar").Picture
Set ctlTel.BotonModificar.Picture = MDI.il32.ListImages("modificar").Picture

InicializarDireccion ctlDir
AplicarPermisos
AplicarConfiguracion

End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "abmafiliado"
End Function

Public Sub Refrescar()
'On Error Resume Next
    cmbTipoDoc.Refresh
    cmbEstadoCivil.Refresh
    cmbOcupacion.Refresh
    cmbParentezco.Refresh
    
    Dim e As blcemi.Empleado
    Set e = cmbCobrador.SelectedItem
    Set cmbCobrador.Coleccion = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eCobrador)
    Set cmbCobrador.SelectedItem = e
    
    cmbObraSocial.Refresh
    lvwAlergia.Refresh
    lvwMedicamento.Refresh
    lvwEnfermedad.Refresh
    If Not mAfiliado Is Nothing Then
        Set mAfiliado.PersonasACargo = Nothing 'obligo a recargar
        Set lvwACargo.Coleccion = mAfiliado.PersonasACargo
    Else
        lvwACargo.Refresh 'si no es a cargo refresco, por el otro camino se refresca sola la lista
    End If
    
   ' ctlDir.Refresh
    AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()
    lvwACargo.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    lvwAlergia.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    lvwEnfermedad.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    lvwMedicamento.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
End Sub

Private Sub AplicarPermisos()
    cmbObraSocial.NuevoEnabled = UsuarioActual.Permisos.Can(blcemi.AltaObraSocial)
    cmbOcupacion.NuevoEnabled = UsuarioActual.Permisos.Can(blcemi.AltaOcupacion)
    cmbParentezco.NuevoEnabled = UsuarioActual.Permisos.Can(blcemi.AltaParentezco)
End Sub

'esto esta porq no puedo agarrar los eventos de este mismo formulario
Friend Sub RefrescarLvwAfiliados()
    lvwACargo.Refresh
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

'-------tabOrder-------------------

Private Sub ctlDir_GotFocus()
sTabDatos.Tab = 0
End Sub

Private Sub ctlTel_GotFocus()
sTabDatos.Tab = 1
End Sub

Private Sub lvwAlergia_GotFocus()
sTabHistoriaClinica.Tab = 1
End Sub

Private Sub lvwEnfermedad_GotFocus()
sTabHistoriaClinica.Tab = 0
End Sub

Private Sub lvwMedicamento_GotFocus()
sTabHistoriaClinica.Tab = 2
End Sub

Private Sub txtMarca_GotFocus()
sTabDatos.Tab = 2
End Sub

Private Sub cmbCobrador_GotFocus()
sTabAfiliacion.Tab = 0
End Sub

Private Sub lvwACargo_GotFocus()
sTabAfiliacion.Tab = 1
End Sub

Private Sub cmbParentezco_GotFocus()
On Error Resume Next
sTabAfiliacion.Tab = 2
End Sub

Private Sub txtNroAfiliadoACargo_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub

Private Sub txtNroDoc_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub

Private Sub txtObservaciones_Change()
sTabHistoriaClinica.Tab = 3
End Sub
