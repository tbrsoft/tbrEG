VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmAtencionBomberos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parte de Salida"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9585
   Begin TabDlg.SSTab sTab 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Principal"
      TabPicture(0)   =   "frmAtencionBomberos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tvw"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraCodigoEmerg"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdDotacion"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraComun"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraAfiliadoExterno"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtObservaciones"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtReseña"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Datos del Siniestro"
      TabPicture(1)   =   "frmAtencionBomberos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraSolicitantes"
      Tab(1).Control(1)=   "fraAtencion"
      Tab(1).Control(2)=   "ctlDirEmergencia"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Personas Afectadas"
      TabPicture(2)   =   "frmAtencionBomberos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Bienes Afectados"
      TabPicture(3)   =   "frmAtencionBomberos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraVehiculos"
      Tab(3).Control(1)=   "fraVivienda"
      Tab(3).Control(2)=   "fraCampo"
      Tab(3).Control(3)=   "fraComplementarios"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Colaboración"
      TabPicture(4)   =   "frmAtencionBomberos.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraBomberos"
      Tab(4).Control(1)=   "fraPolicia"
      Tab(4).Control(2)=   "Frame1"
      Tab(4).ControlCount=   3
      Begin VB.Frame fraSolicitantes 
         Caption         =   "Datos de los Solicitantes"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   105
         Top             =   480
         Width           =   9135
         Begin ControlesPOO.ListViewConsulta lvwSolicitantes 
            Height          =   2175
            Left            =   120
            TabIndex        =   106
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   3836
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   -1  'True
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            CampoKey        =   ""
            AllowModify     =   0   'False
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Apellido"
            MEncabezado0    =   "apellido"
            AEncabezado0    =   15
            NEncabezado1    =   "Nombre"
            MEncabezado1    =   "nombre"
            AEncabezado1    =   15
            NEncabezado2    =   "DNI"
            MEncabezado2    =   "nrodoc"
            AEncabezado2    =   15
            NEncabezado3    =   "Telefono"
            MEncabezado3    =   "telefono"
            AEncabezado3    =   15
            NEncabezado4    =   "Direccion"
            MEncabezado4    =   "pgdireccion"
            AEncabezado4    =   25
            NEncabezado5    =   "Relacion"
            MEncabezado5    =   "descripcionrelacion"
            AEncabezado5    =   15
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
         Begin ALCemi.GraphicButton cmdEliminarSolicitante 
            Height          =   375
            Left            =   8640
            TabIndex        =   107
            ToolTipText     =   "Quitar los datos del solicitante seleccionado."
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdEditarSolicitante 
            Height          =   375
            Left            =   8640
            TabIndex        =   108
            ToolTipText     =   "Modificar los datos del solicitante."
            Top             =   660
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdAgregarSolicitante 
            Height          =   375
            Left            =   8640
            TabIndex        =   109
            ToolTipText     =   "Agregar un solicitante al siniestro."
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   100
         Top             =   480
         Width           =   9135
         Begin ControlesPOO.ListViewConsulta lvwPersonas 
            Height          =   3255
            Left            =   120
            TabIndex        =   101
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   5741
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   -1  'True
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            CampoKey        =   ""
            AllowModify     =   0   'False
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Apellido"
            MEncabezado0    =   "apellido"
            AEncabezado0    =   15
            NEncabezado1    =   "Nombre"
            MEncabezado1    =   "nombre"
            AEncabezado1    =   15
            NEncabezado2    =   "DNI"
            MEncabezado2    =   "NroDoc"
            AEncabezado2    =   15
            NEncabezado3    =   "Direccion"
            MEncabezado3    =   "pgdireccion"
            AEncabezado3    =   20
            NEncabezado4    =   "Relacion c/Siniestro"
            MEncabezado4    =   "descripcionrelacion"
            AEncabezado4    =   35
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
         Begin ALCemi.GraphicButton cmdEliminarAfectado 
            Height          =   375
            Left            =   8640
            TabIndex        =   102
            ToolTipText     =   "Eliminar los datos de la persona seleccionada."
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdEditarAfectado 
            Height          =   375
            Left            =   8640
            TabIndex        =   103
            ToolTipText     =   "Modificar los datos de la persona seleccionada."
            Top             =   660
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdAgregarAfectado 
            Height          =   375
            Left            =   8640
            TabIndex        =   104
            ToolTipText     =   "Agregar una persona afectada por el siniestro."
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
      End
      Begin VB.Frame fraVehiculos 
         Caption         =   "Vehiculos"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   95
         Top             =   480
         Width           =   9135
         Begin ControlesPOO.ListViewConsulta lvwVehiculos 
            Height          =   1215
            Left            =   120
            TabIndex        =   96
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2143
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   -1  'True
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            AllowModify     =   0   'False
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Tipo"
            MEncabezado0    =   "tipo"
            AEncabezado0    =   10
            NEncabezado1    =   "Marca"
            MEncabezado1    =   "marca"
            AEncabezado1    =   10
            NEncabezado2    =   "Modelo"
            MEncabezado2    =   "modelo"
            AEncabezado2    =   10
            NEncabezado3    =   "Patente"
            MEncabezado3    =   "patente"
            AEncabezado3    =   10
            NEncabezado4    =   "Color"
            MEncabezado4    =   "color"
            AEncabezado4    =   10
            NEncabezado5    =   "Daños"
            MEncabezado5    =   "perjuicios"
            AEncabezado5    =   50
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
         Begin ALCemi.GraphicButton cmdEliminarVehiculo 
            Height          =   375
            Left            =   8640
            TabIndex        =   97
            ToolTipText     =   "Eliminar los datos del vehiculo seleccionado."
            Top             =   1080
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdEditarVehiculo 
            Height          =   375
            Left            =   8640
            TabIndex        =   98
            ToolTipText     =   "Modificar los datos del vehiculo afectado."
            Top             =   660
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdAgregarVehiculo 
            Height          =   375
            Left            =   8640
            TabIndex        =   99
            ToolTipText     =   "Agregar un vehiculo afectado por el siniestro."
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
      End
      Begin VB.Frame fraVivienda 
         Caption         =   "Vivienda"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   90
         Top             =   2160
         Width           =   4215
         Begin VB.TextBox txtAmbientes 
            Height          =   285
            Left            =   1920
            TabIndex        =   92
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtPerjuiciosVivienda 
            Height          =   1095
            Left            =   120
            TabIndex        =   91
            Top             =   840
            Width           =   3975
         End
         Begin VB.Label Label7 
            Caption         =   "Ambientes Afectados:"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label lblPerjuicios 
            Caption         =   "Descripción daños:"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   600
            Width           =   1575
         End
      End
      Begin VB.Frame fraCampo 
         Caption         =   "Campo - Baldío"
         Height          =   2055
         Left            =   -70560
         TabIndex        =   83
         Top             =   2160
         Width           =   4815
         Begin VB.TextBox txtHectareas 
            Height          =   285
            Left            =   1920
            TabIndex        =   86
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtPerjuiciosCampo 
            Height          =   735
            Left            =   120
            TabIndex        =   85
            Top             =   1200
            Width           =   4575
         End
         Begin VB.TextBox txtMaterialesCombustibles 
            Height          =   285
            Left            =   1920
            TabIndex        =   84
            Top             =   600
            Width           =   2775
         End
         Begin VB.Label Label8 
            Caption         =   "Hectareas Afectadas:"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Descripción daños:"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Materiales Combustibles:"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   600
            Width           =   1815
         End
      End
      Begin VB.Frame fraComplementarios 
         Caption         =   "Datos Complementarios"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   72
         Top             =   4320
         Width           =   9135
         Begin VB.CheckBox chkSeguro 
            Caption         =   "Seguro.       Aseguradora:"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   720
            Width           =   2115
         End
         Begin VB.TextBox txtAseguradora 
            Height          =   285
            Left            =   2280
            TabIndex        =   76
            Top             =   720
            Width           =   2895
         End
         Begin VB.TextBox txtPoliza 
            Height          =   285
            Left            =   6840
            TabIndex        =   75
            Top             =   720
            Width           =   2175
         End
         Begin VB.CheckBox chkPrevencionFuego 
            Caption         =   "Material de Prevención y Lucha contra el Fuego.         Descripcion:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1080
            Width           =   5055
         End
         Begin VB.TextBox txtMaterialPrevencionFuego 
            Height          =   285
            Left            =   5400
            TabIndex        =   73
            Top             =   1080
            Width           =   3615
         End
         Begin ControlesPOO.Combo cmbGas 
            Height          =   315
            Left            =   6840
            TabIndex        =   78
            Top             =   360
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbElectricidad 
            Height          =   315
            Left            =   2280
            TabIndex        =   79
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   556
            Enabled         =   -1  'True
         End
         Begin VB.Label Label16 
            Caption         =   "Instalación Eléctrica:"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label18 
            Caption         =   "Instalación de Gas:"
            Height          =   255
            Left            =   5400
            TabIndex        =   81
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "Poliza Nº:"
            Height          =   255
            Left            =   5400
            TabIndex        =   80
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame fraBomberos 
         Caption         =   "Colaboración de Otros Cuerpos"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   65
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtEquiposEspeciales 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   66
            Top             =   1800
            Width           =   8895
         End
         Begin ControlesPOO.ListViewConsulta lvwOtrosCuerpos 
            Height          =   1215
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   2143
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   -1  'True
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            AllowModify     =   0   'False
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Cuerpo"
            MEncabezado0    =   "pgcuerpo"
            AEncabezado0    =   25
            NEncabezado1    =   "A Cargo De"
            MEncabezado1    =   "pgresponsable"
            AEncabezado1    =   25
            NEncabezado2    =   "Efectivos"
            MEncabezado2    =   "cantidadefectivos"
            AEncabezado2    =   25
            NEncabezado3    =   "Unidad"
            MEncabezado3    =   "pgunidad"
            AEncabezado3    =   25
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
         Begin ALCemi.GraphicButton cmdEliminarCuerpo 
            Height          =   375
            Left            =   8640
            TabIndex        =   68
            ToolTipText     =   "Eliminar los datos del cuerpo seleccionado."
            Top             =   1080
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdEditarCuerpo 
            Height          =   375
            Left            =   8640
            TabIndex        =   69
            ToolTipText     =   "Modificar los datos del cuerpo seleccionado."
            Top             =   660
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdAgregarCuerpo 
            Height          =   375
            Left            =   8640
            TabIndex        =   70
            ToolTipText     =   "Agregar un cuerpo de bomberos."
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin VB.Label Label20 
            Caption         =   "Equipos especiales:"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   1560
            Width           =   2895
         End
      End
      Begin VB.Frame fraPolicia 
         Caption         =   "Policía"
         Height          =   855
         Left            =   -74880
         TabIndex        =   58
         Top             =   3240
         Width           =   9135
         Begin VB.TextBox txtPoliciaACargo 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   61
            Top             =   360
            Width           =   3495
         End
         Begin VB.TextBox txtPoliciaCantEfectivos 
            Height          =   285
            Left            =   6600
            TabIndex        =   60
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtPoliciaNroMovil 
            Height          =   285
            Left            =   8280
            MaxLength       =   50
            TabIndex        =   59
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label21 
            Caption         =   "A Cargo De:"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label22 
            Caption         =   "Cantidad de Efectivos:"
            Height          =   255
            Left            =   4800
            TabIndex        =   63
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label23 
            Caption         =   "Móvil Nº:"
            Height          =   255
            Left            =   7440
            TabIndex        =   62
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Servicios Emergencias"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   50
         Top             =   4320
         Width           =   9135
         Begin VB.TextBox txtMedicoNombre 
            Height          =   285
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   54
            Top             =   360
            Width           =   5535
         End
         Begin VB.TextBox txtMedicoMP 
            Height          =   285
            Left            =   8280
            MaxLength       =   25
            TabIndex        =   53
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtCentroAsistenciales 
            Height          =   285
            Left            =   1800
            MaxLength       =   255
            TabIndex        =   52
            Top             =   720
            Width           =   5535
         End
         Begin VB.CheckBox chkAmbulancias 
            Caption         =   "Ambulancias."
            Height          =   255
            Left            =   7440
            TabIndex        =   51
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label24 
            Caption         =   "Médico:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   735
         End
         Begin VB.Label lblMP 
            Caption         =   "MP:"
            Height          =   255
            Left            =   7440
            TabIndex        =   56
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label25 
            Caption         =   "Centros Asistenciales:"
            Height          =   255
            Left            =   120
            TabIndex        =   55
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.TextBox txtReseña 
         Height          =   855
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   44
         Top             =   4560
         Width           =   4935
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   4560
         Width           =   3975
      End
      Begin VB.Frame fraAtencion 
         Caption         =   "Referencias"
         Height          =   2415
         Left            =   -70200
         TabIndex        =   40
         Top             =   3120
         Width           =   4455
         Begin VB.TextBox txtAccesoPor 
            Height          =   645
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   49
            Top             =   1680
            Width           =   4215
         End
         Begin VB.TextBox txtReferencias 
            Height          =   765
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   47
            Top             =   600
            Width           =   4215
         End
         Begin VB.Label Label27 
            Caption         =   "Acceso por:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label Label26 
            Caption         =   "Referencia más próxima:"
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame fraAfiliadoExterno 
         Height          =   735
         Left            =   4320
         TabIndex        =   35
         Top             =   1320
         Width           =   4935
         Begin VB.TextBox txtNroIncidente 
            Height          =   285
            Left            =   2040
            TabIndex        =   37
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtNroInterno 
            Height          =   285
            Left            =   3840
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label32 
            Caption         =   "Nro. de Incidente Externo:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label17 
            Caption         =   "Interno:"
            Height          =   255
            Left            =   3120
            TabIndex        =   38
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame fraComun 
         Caption         =   "Tiempos"
         Height          =   2895
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   3975
         Begin VB.CommandButton cmdRegistrarArribo 
            Caption         =   "Arribo al Lugar (QTH)"
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   23
            Top             =   2160
            Width           =   3375
         End
         Begin VB.CommandButton cmdRegistrarLiberacion 
            Caption         =   "Registrar Liberacion (VL)"
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   22
            Top             =   2520
            Width           =   3375
         End
         Begin VB.CommandButton cmdArriboPreinspeccion 
            Caption         =   "Arribo Pre-Inspección"
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   1440
            Width           =   3375
         End
         Begin VB.CommandButton cmdSalidaDotacion 
            Caption         =   "Salida Dotacion"
            Height          =   285
            Left            =   240
            TabIndex        =   18
            Top             =   1800
            Width           =   3375
         End
         Begin VB.CommandButton cmdSalidaPreInspeccion 
            Caption         =   "Salida Pre-Inspección"
            Height          =   285
            Left            =   240
            TabIndex        =   16
            Top             =   1080
            Width           =   3375
         End
         Begin VB.TextBox txtFecha 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtHora 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtVL 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   25
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtQTH 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtPreInspeccion 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtSalidaDotacion 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   19
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtSalidaPreInspeccion 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label12 
            Caption         =   "Hora Llamado:"
            Height          =   195
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label14 
            Caption         =   "VL:"
            Height          =   195
            Left            =   240
            TabIndex        =   32
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "QTH:"
            Height          =   195
            Left            =   240
            TabIndex        =   31
            Top             =   2160
            Width           =   390
         End
         Begin VB.Label Label3 
            Caption         =   "Arribo Pre-Inspección:"
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Salida Dotación:"
            Height          =   255
            Left            =   240
            TabIndex        =   29
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label5 
            Caption         =   "Salida Pre-Inspección:"
            Height          =   255
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdDotacion 
         Caption         =   "Dotacion"
         Height          =   375
         Left            =   4320
         TabIndex        =   13
         Top             =   2160
         Width           =   4935
      End
      Begin VB.Frame fraCodigoEmerg 
         Caption         =   "Seleccione el Grado y Tipo del Siniestro"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtSintoma 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4680
            TabIndex        =   8
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdVerFichaPreArribo 
            Caption         =   "Ver ficha Pre-Arribo..."
            Height          =   375
            Left            =   6720
            TabIndex        =   6
            Top             =   480
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   720
            TabIndex        =   5
            Top             =   240
            Width           =   375
         End
         Begin ControlesPOO.Combo cmbSintoma 
            Height          =   315
            Left            =   5160
            TabIndex        =   7
            Top             =   240
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   556
            AtributoAMostrar=   "nombrecompuesto"
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbCodigo 
            Height          =   315
            Left            =   1200
            TabIndex        =   9
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            AtributoAMostrar=   "nombrecompuesto"
            Enabled         =   -1  'True
         End
         Begin ALCemi.GraphicButton cmdBuscarSintoma 
            Height          =   315
            Left            =   8730
            TabIndex        =   10
            Top             =   240
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   556
         End
         Begin VB.Label lblSintoma 
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   4080
            TabIndex        =   12
            Top             =   240
            Width           =   495
         End
         Begin VB.Label lblTipo 
            Caption         =   "Grado:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   615
         End
      End
      Begin ControlesPOO.TreeViewConsulta tvw 
         Height          =   1695
         Left            =   4320
         TabIndex        =   14
         Top             =   2520
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   2990
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Indentation     =   299,906
         LineStyle       =   1
         Nodo.BackColor0 =   " 16777215"
         Nodo.Bold0      =   "False"
         Nodo.ChildCollectionField0=   "Dotacion"
         Nodo.Expanded0  =   "False"
         Nodo.ForeColor0 =   " 0"
         Nodo.IdField0   =   "id"
         Nodo.TextField0 =   "NombreMovil"
         Nodo.BackColor1 =   " 16777215"
         Nodo.Bold1      =   "False"
         Nodo.ChildCollectionField1=   ""
         Nodo.Expanded1  =   "False"
         Nodo.ForeColor1 =   " 0"
         Nodo.IdField1   =   "id"
         Nodo.TextField1 =   "NombreCompleto"
         Nodo.BackColor0 =   "0"
         Nodo.Bold0      =   "False"
         Nodo.ChildCollectionField0=   ""
         Nodo.Expanded0  =   "False"
         Nodo.ForeColor0 =   "0"
         Nodo.IdField0   =   ""
         Nodo.TextField0 =   ""
         Nodo.BackColor1 =   "0"
         Nodo.Bold1      =   "False"
         Nodo.ChildCollectionField1=   ""
         Nodo.Expanded1  =   "False"
         Nodo.ForeColor1 =   "0"
         Nodo.IdField1   =   ""
         Nodo.TextField1 =   ""
         Nodo.BackColor2 =   "0"
         Nodo.Bold2      =   "False"
         Nodo.ChildCollectionField2=   ""
         Nodo.Expanded2  =   "False"
         Nodo.ForeColor2 =   "0"
         Nodo.IdField2   =   ""
         Nodo.TextField2 =   ""
         Nodo.BackColor3 =   "0"
         Nodo.Bold3      =   "False"
         Nodo.ChildCollectionField3=   ""
         Nodo.Expanded3  =   "False"
         Nodo.ForeColor3 =   "0"
         Nodo.IdField3   =   ""
         Nodo.TextField3 =   ""
         Nodo.BackColor4 =   "0"
         Nodo.Bold4      =   "False"
         Nodo.ChildCollectionField4=   ""
         Nodo.Expanded4  =   "False"
         Nodo.ForeColor4 =   "0"
         Nodo.IdField4   =   ""
         Nodo.TextField4 =   ""
         Nodo.BackColor5 =   "0"
         Nodo.Bold5      =   "False"
         Nodo.ChildCollectionField5=   ""
         Nodo.Expanded5  =   "False"
         Nodo.ForeColor5 =   "0"
         Nodo.IdField5   =   ""
         Nodo.TextField5 =   ""
         Nodo.BackColor6 =   "0"
         Nodo.Bold6      =   "False"
         Nodo.ChildCollectionField6=   ""
         Nodo.Expanded6  =   "False"
         Nodo.ForeColor6 =   "0"
         Nodo.IdField6   =   ""
         Nodo.TextField6 =   ""
         Nodo.BackColor7 =   "0"
         Nodo.Bold7      =   "False"
         Nodo.ChildCollectionField7=   ""
         Nodo.Expanded7  =   "False"
         Nodo.ForeColor7 =   "0"
         Nodo.IdField7   =   ""
         Nodo.TextField7 =   ""
         Nodo.BackColor8 =   "0"
         Nodo.Bold8      =   "False"
         Nodo.ChildCollectionField8=   ""
         Nodo.Expanded8  =   "False"
         Nodo.ForeColor8 =   "0"
         Nodo.IdField8   =   ""
         Nodo.TextField8 =   ""
         Nodo.BackColor9 =   "0"
         Nodo.Bold9      =   "False"
         Nodo.ChildCollectionField9=   ""
         Nodo.Expanded9  =   "False"
         Nodo.ForeColor9 =   "0"
         Nodo.IdField9   =   ""
         Nodo.TextField9 =   ""
         Nodo.BackColor10=   "0"
         Nodo.Bold10     =   "False"
         Nodo.ChildCollectionField10=   ""
         Nodo.Expanded10 =   "False"
         Nodo.ForeColor10=   "0"
         Nodo.IdField10  =   ""
         Nodo.TextField10=   ""
         Nodo.BackColor11=   "0"
         Nodo.Bold11     =   "False"
         Nodo.ChildCollectionField11=   ""
         Nodo.Expanded11 =   "False"
         Nodo.ForeColor11=   "0"
         Nodo.IdField11  =   ""
         Nodo.TextField11=   ""
         Nodo.BackColor12=   "0"
         Nodo.Bold12     =   "False"
         Nodo.ChildCollectionField12=   ""
         Nodo.Expanded12 =   "False"
         Nodo.ForeColor12=   "0"
         Nodo.IdField12  =   ""
         Nodo.TextField12=   ""
         Nodo.BackColor13=   "0"
         Nodo.Bold13     =   "False"
         Nodo.ChildCollectionField13=   ""
         Nodo.Expanded13 =   "False"
         Nodo.ForeColor13=   "0"
         Nodo.IdField13  =   ""
         Nodo.TextField13=   ""
         Nodo.BackColor14=   "0"
         Nodo.Bold14     =   "False"
         Nodo.ChildCollectionField14=   ""
         Nodo.Expanded14 =   "False"
         Nodo.ForeColor14=   "0"
         Nodo.IdField14  =   ""
         Nodo.TextField14=   ""
         Nodo.BackColor15=   "0"
         Nodo.Bold15     =   "False"
         Nodo.ChildCollectionField15=   ""
         Nodo.Expanded15 =   "False"
         Nodo.ForeColor15=   "0"
         Nodo.IdField15  =   ""
         Nodo.TextField15=   ""
         Nodo.BackColor16=   "0"
         Nodo.Bold16     =   "False"
         Nodo.ChildCollectionField16=   ""
         Nodo.Expanded16 =   "False"
         Nodo.ForeColor16=   "0"
         Nodo.IdField16  =   ""
         Nodo.TextField16=   ""
         Nodo.BackColor17=   "0"
         Nodo.Bold17     =   "False"
         Nodo.ChildCollectionField17=   ""
         Nodo.Expanded17 =   "False"
         Nodo.ForeColor17=   "0"
         Nodo.IdField17  =   ""
         Nodo.TextField17=   ""
         Nodo.BackColor18=   "0"
         Nodo.Bold18     =   "False"
         Nodo.ChildCollectionField18=   ""
         Nodo.Expanded18 =   "False"
         Nodo.ForeColor18=   "0"
         Nodo.IdField18  =   ""
         Nodo.TextField18=   ""
         Nodo.BackColor19=   "0"
         Nodo.Bold19     =   "False"
         Nodo.ChildCollectionField19=   ""
         Nodo.Expanded19 =   "False"
         Nodo.ForeColor19=   "0"
         Nodo.IdField19  =   ""
         Nodo.TextField19=   ""
      End
      Begin ALCemi.ctlDireccion ctlDirEmergencia 
         Height          =   2565
         Left            =   -74880
         TabIndex        =   45
         Top             =   3120
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   3995
         ProvinciaVisible=   0   'False
         Caption         =   "Ubicación del Siniestro"
         CanDragDrop     =   -1  'True
         SoloConsulta    =   0   'False
         EntrecallesVisible=   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Reseña de la Actuación:"
         Height          =   255
         Left            =   4320
         TabIndex        =   43
         Top             =   4320
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   4320
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar"
      Height          =   495
      Left            =   5520
      TabIndex        =   2
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   6840
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   6120
      Width           =   1215
   End
End
Attribute VB_Name = "frmAtencionBomberos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAtencion As BLCemi.AtencionB

Private mAfiliadoPropio As BLCemi.Afiliado

Private mInvolucrados As BLCemi.InvolucradoManager

Private WithEvents mFrmABMInvolucrado As frmABMInvolucrado
Attribute mFrmABMInvolucrado.VB_VarHelpID = -1
Private WithEvents mFrmABMVehiculo As frmABMVehiculo
Attribute mFrmABMVehiculo.VB_VarHelpID = -1

Private WithEvents mFrmConsultarDotacion As frmConsultarDotaciones
Attribute mFrmConsultarDotacion.VB_VarHelpID = -1

Private WithEvents mFrmColaboracion As frmColaboracion
Attribute mFrmColaboracion.VB_VarHelpID = -1

Private WithEvents mFrmBuscarSintoma As frmSeleccionarSintoma
Attribute mFrmBuscarSintoma.VB_VarHelpID = -1

Private mFrmParent As frmConsultaAtencion 'para enviarle eventos porq en gral se van a abrir varias frmAtencionBomberos

Private Sub cmdFinalizar_Click()
    If MsgBox("Esta seguro que desea finalizar la Atención? " + vbCrLf + "(Una vez cerrada la misma no se podrán realizar modificaciones).", vbQuestion + vbYesNo, "tbrEmergencyGroup") = vbYes Then
        Guardar True
    End If
End Sub

Private Sub cmdGuardar_Click()
    Guardar False
End Sub

Private Sub Guardar(cerrar As Boolean)
    'seteo todos, alguno debe ser, si estan todos vacios despues hay q preguntar
    'lo unico q necesito si o si es el sintoma
    If DatosBasicosCorrectos Then
        
        'los vehiculos, colaboracion de bomberos y equipos no los tengo q asignar porq los tengo referenciados.
        Set mAtencion.Direccion = ctlDirEmergencia.MiDireccion
        
        Set mAtencion.Despachador = UsuarioActual
       
        If cerrar Then
            mAtencion.Estado = BLCemi.eFinalizado
        Else
            mAtencion.Estado = IIf(DatosCompletos, BLCemi.eestadoatencion.eListaParaCerrar, BLCemi.eestadoatencion.ePendiente)
        End If
        
        'mAtencion.HoraLlamada = Time 'Trim(Str(dtpHora.Hour)) + ":" + Trim(Str(dtpHora.Minute)) + ":" + Trim(Str(dtpHora.Second))
        mAtencion.NroIncidente = txtNroIncidente
        mAtencion.nroIncidenteInterno = txtNroInterno
        mAtencion.Observaciones = txtObservaciones
        
        mAtencion.DescripcionPerjuiciosCampo = txtPerjuiciosCampo
        mAtencion.MaterialesCombustibles = txtMaterialesCombustibles
        mAtencion.DescripcionPerjuiciosVivienda = txtPerjuiciosVivienda
        mAtencion.AccesoPor = txtAccesoPor
        mAtencion.Referencias = txtReferencias
        mAtencion.DescripcionMaterial = txtMaterialPrevencionFuego
        mAtencion.EquiposEspeciales = txtEquiposEspeciales
        mAtencion.Poliza = txtPoliza
        
        mAtencion.PoliciaACargo = txtPoliciaACargo
        If IsNumeric(txtPoliciaCantEfectivos) Then
            mAtencion.PoliciaCantidad = txtPoliciaCantEfectivos
        End If
        mAtencion.PoliciaMovil = txtPoliciaNroMovil
        mAtencion.SEMedico = txtMedicoNombre
        mAtencion.SECentroAsistencial = txtCentroAsistenciales
        mAtencion.SEMedicoMP = txtMedicoMP
        mAtencion.SEAmbulancias = IIf(chkAmbulancias.Value = 1, True, False)
        
        If txtAmbientes <> "" Then
            mAtencion.AmbientesAfectadosVivienda = Val(txtAmbientes)
        End If
        If txtHectareas <> "" Then
            mAtencion.HectareasAfectadasCampo = Val(txtHectareas)
        End If
        
        Set mAtencion.InstalacionElectrica = cmbElectricidad.SelectedItem
        Set mAtencion.InstalacionGas = cmbGas.SelectedItem
        
        'ver por el check mAtencion.Aseguradora = txtAseguradora
    
        Set mAtencion.Sintoma = cmbSintoma.SelectedItem
        
        mAtencion.SalidaPreInspeccion = txtSalidaPreInspeccion
        mAtencion.LlegadaPreInspeccion = txtPreInspeccion
        mAtencion.SalidaDotacion = txtSalidaDotacion
        mAtencion.QTH = txtQTH
        mAtencion.VL = txtVL
    
        If mAtencion.id = 0 Then
            Set mAtencion.Involucrados = mInvolucrados 'los tengo q referenciar solo por nuevo, por modificar ya queda cuando cargo el formulario
            mAtencion.fecha = CDate(txtFecha.Text)
            mAtencion.HoraLlamada = txtHora
            mAtencion.Guardar
        Else
            mAtencion.GuardarModificaciones UsuarioActual
        End If
        
        mFrmParent.Refrescar
        Unload Me
    
    End If
End Sub

Private Sub cmdCancelar_Click()
'preguntar si esta seguro cuando cargo datos
If Not mAfiliadoPropio Is Nothing Then
    If MsgBox("Esta seguro que desea descartar los ultimos cambios?", vbQuestion + vbOKCancel) = vbOK Then
        If Not mAtencion Is Nothing Then
            Set mAtencion.Equipos = Nothing 'para evitarme el tema del beginedit etc.
            Set mAtencion.Vehiculos = Nothing
        End If
        Unload Me
    End If
Else
    Unload Me
End If

End Sub

Private Function DatosBasicosCorrectos() As Boolean
    Dim msj As String
    Dim msjDir As String
    
    Dim sint As BLCemi.Sintoma
        
    If Not cmbSintoma.SelectedItem Is Nothing Then
        Set sint = cmbSintoma.SelectedItem
        
        If Not ctlDirEmergencia.DireccionCompleta(msjDir) Then msj = msj + msjDir
                
        If mInvolucrados.GetByTipo(BLCemi.eSolicitante).Count = 0 Then msj = msj + "Debe ingresar los datos de al menos un Solicitante." + vbCrLf
    Else
        msj = msj + "Debe seleccionar un Tipo de Siniestro." + vbCrLf
    End If
       
        
    If msj = "" Then
        DatosBasicosCorrectos = True
    Else
        MsgBox "Faltan los siguientes datos:" + vbCrLf + msj, vbExclamation
        DatosBasicosCorrectos = False
    End If
End Function

Private Function DatosCompletos() As Boolean
    Dim aux As Boolean
    'asumo q los datos estan completos, al operarlos con un and si alguno no esta me cambia a falso.
    aux = True
    
    'principal
    aux = aux And tvw.Coleccion.Count > 0 ' me fijo que haya equipos
    aux = aux And TiemposCorrectos
    
    'datos
        'txtAccesoPor
        'txtReferencias
        'txtObservaciones
        'txtReseña
    aux = aux And ctlDirEmergencia.MiDireccion.Calle <> ""
    
    'personas
    aux = aux And mInvolucrados.GetByTipo(BLCemi.eSolicitante).Count > 0
    
    'bienes
    'txtPerjuiciosCampo
    'txtMaterialesCombustibles
    'txtPerjuiciosVivienda
    'txtPoliza
    'txtAmbientes
    'txtHectareas
    
    'colab
    'txtMaterialPrevencionFuego
    'txtEquiposEspeciales
'    txtPoliciaACargo = mAtencion.PoliciaACargo
'    txtPoliciaCantEfectivos = mAtencion.PoliciaCantidad
'    txtPoliciaNroMovil = mAtencion.PoliciaMovil
'    txtMedicoNombre = mAtencion.SEMedico
'    txtCentroAsistenciales = mAtencion.SECentroAsistencial
'    txtMedicoMP = mAtencion.SEMedicoMP
    
    DatosCompletos = aux
     
End Function

Private Function TiemposCorrectos() As Boolean
    Dim aux As Boolean
    Dim aux2 As Boolean
    'si salio y llego la preinspeccion...
    aux = (txtSalidaPreInspeccion.Text <> "" And txtPreInspeccion <> "")
    'si salio, llego y volvio la dotacion
    aux2 = (txtSalidaDotacion.Text <> "" And txtQTH.Text <> "" And txtVL.Text <> "")
    'por ultimo si salio y llego la preinspeccion o si salio, llego y volvio la dotacion
    TiemposCorrectos = aux Or aux2
End Function

Public Sub RecibirLlamadoTelefono(pFrmParent As frmConsultaAtencion, pNumero As String)
    NuevaAtencion pFrmParent
    'guardar el numero en una variable, par mostrarlo despues en el abm involuc
End Sub

Public Sub NuevaAtencion(pFrmParent As frmConsultaAtencion)
    Set mAtencion = New BLCemi.AtencionB
    Set mFrmParent = pFrmParent
    Set tvw.Coleccion = mAtencion.Equipos
    Set mInvolucrados = New BLCemi.InvolucradoManager
    Set lvwVehiculos.Coleccion = mAtencion.Vehiculos
    Set lvwOtrosCuerpos.Coleccion = mAtencion.ColaboracionBomberos
    Me.Show
    txtHora = Time
    cmdFinalizar.Visible = False
    cmdGuardar.Visible = True
End Sub

Public Sub ModificarAtencion(pAtencion As BLCemi.AtencionB, pFrmParent As frmConsultaAtencion)
    Set mAtencion = pAtencion
    Set mFrmParent = pFrmParent
    Me.Show
    If mAtencion.Estado = BLCemi.eListaParaCerrar Then cmdFinalizar.Visible = True
    
    Set mInvolucrados = mAtencion.Involucrados
    
    Set lvwVehiculos.Coleccion = mAtencion.Vehiculos
    Set lvwOtrosCuerpos.Coleccion = mAtencion.ColaboracionBomberos
    
    ActualizarInvolucrados
            
    Set ctlDirEmergencia.MiDireccion = mAtencion.Direccion
    
    'mAtencion.Despachador empleadoactual, ver
    
    Set tvw.Coleccion = mAtencion.Equipos
    'mAtencion.Estado ver
    txtFecha.Text = mAtencion.fecha
    txtHora.Text = mAtencion.HoraLlamada
    txtNroIncidente = mAtencion.NroIncidente
    txtNroInterno = mAtencion.nroIncidenteInterno
    txtObservaciones = mAtencion.Observaciones
           
    txtPerjuiciosCampo = mAtencion.DescripcionPerjuiciosCampo
    txtMaterialesCombustibles = mAtencion.MaterialesCombustibles
    txtPerjuiciosVivienda = mAtencion.DescripcionPerjuiciosVivienda
    txtAccesoPor = mAtencion.AccesoPor
    txtReferencias = mAtencion.Referencias
    txtMaterialPrevencionFuego = mAtencion.DescripcionMaterial
    txtEquiposEspeciales = mAtencion.EquiposEspeciales
    txtPoliza = mAtencion.Poliza
    txtAmbientes = mAtencion.AmbientesAfectadosVivienda
    txtHectareas = mAtencion.HectareasAfectadasCampo
    Set cmbElectricidad.SelectedItem = mAtencion.InstalacionElectrica
    Set cmbGas.SelectedItem = mAtencion.InstalacionGas
    
    txtPoliciaACargo = mAtencion.PoliciaACargo
    txtPoliciaCantEfectivos = mAtencion.PoliciaCantidad
    txtPoliciaNroMovil = mAtencion.PoliciaMovil
    txtMedicoNombre = mAtencion.SEMedico
    txtCentroAsistenciales = mAtencion.SECentroAsistencial
    txtMedicoMP = mAtencion.SEMedicoMP
    chkAmbulancias.Value = IIf(mAtencion.SEAmbulancias, 1, 0)
    
    'ver por el check txtAseguradora=mAtencion.Aseguradora
       
       
    Set cmbSintoma.SelectedItem = mAtencion.Sintoma
    Set cmbCodigo.SelectedItem = mAtencion.Sintoma.Parent
    'tiempos
    txtSalidaPreInspeccion = mAtencion.SalidaPreInspeccion
    txtPreInspeccion = mAtencion.LlegadaPreInspeccion
    txtSalidaDotacion = mAtencion.SalidaDotacion
    txtVL = mAtencion.VL
    txtQTH = mAtencion.QTH
    
    If txtSalidaPreInspeccion <> "" Then
        cmdSalidaPreInspeccion.Visible = False
        cmdArriboPreinspeccion.Enabled = True
    End If
    
    If txtPreInspeccion <> "" Then
        cmdArriboPreinspeccion.Visible = False
    End If
    
    If txtSalidaDotacion <> "" Then
        cmdSalidaDotacion.Visible = False
        cmdRegistrarArribo.Enabled = True
    End If
    
    If txtQTH <> "" Then
        cmdRegistrarArribo.Visible = False
        cmdRegistrarLiberacion.Enabled = True
    End If
    
    If txtVL <> "" Then cmdRegistrarLiberacion.Visible = False
            
    cmdGuardar.Visible = True
End Sub

Private Sub Form_Load()
    On Local Error GoTo errF6
    
    TERR.Anotar "abag"
    Set cmdBuscarSintoma.Picture = MDI.il16.ListImages("buscar").Picture
    Set cmdAgregarAfectado.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdAgregarCuerpo.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdAgregarSolicitante.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdAgregarVehiculo.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdEditarAfectado.Picture = MDI.il32.ListImages("modificar").Picture
    Set cmdEditarCuerpo.Picture = MDI.il32.ListImages("modificar").Picture
    Set cmdEditarSolicitante.Picture = MDI.il32.ListImages("modificar").Picture
    Set cmdEditarVehiculo.Picture = MDI.il32.ListImages("modificar").Picture
    Set cmdEliminarAfectado.Picture = MDI.il32.ListImages("eliminar").Picture
    Set cmdEliminarCuerpo.Picture = MDI.il32.ListImages("eliminar").Picture
    Set cmdEliminarSolicitante.Picture = MDI.il32.ListImages("eliminar").Picture
    Set cmdEliminarVehiculo.Picture = MDI.il32.ListImages("eliminar").Picture
    TERR.Anotar "abah"
    
    Set cmbCodigo.Coleccion = GBL.CodigoEmergenciaGBL
    Set cmbElectricidad.Coleccion = GBL.InstElectricasGBL
    Set cmbGas.Coleccion = GBL.InstalacionesGasGBL
    TERR.Anotar "abai"
    
    Set ctlDirEmergencia.MiDireccion = New BLCemi.Direccion
    TERR.Anotar "abaj"
    
    InicializarDireccion ctlDirEmergencia
    TERR.Anotar "abak"
    Set cmbSintoma.Coleccion = GBL.SintomasGBL
    txtFecha.Text = Date
    txtHora.Text = Time
    Set Me.Icon = MDI.Icon

    TERR.Anotar "abal"
    Me.Move (MDI.Width - 10650) / 2, 0 'no uso me.width porq lo cambio despues del load
    
    TERR.Anotar "abam"
    AplicarConfiguracion
    TERR.Anotar "aban"
    AplicarPermisos
    
    Exit Sub
errF6:
    TERR.AppendLog "errF6--", TERR.ErrToTXT(Err)
    
End Sub

Private Sub AplicarConfiguracion()
    
'    lblSintoma.Left = lblTipo.Left
'    txtSintoma.Left = lblSintoma.Left + 100 + lblSintoma.Width
'    cmbSintoma.Left = txtSintoma.Left + 100 + txtSintoma.Width
'    cmbSintoma.Width = 5635

End Sub

Private Sub AplicarPermisos()

End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "siniestro"
End Function

Public Sub Refrescar()
    AplicarConfiguracion
    AplicarPermisos
    Set cmbElectricidad.Coleccion = GBL.InstElectricasGBL
    Set cmbGas.Coleccion = GBL.InstalacionesGasGBL
    ActualizarInvolucrados
End Sub

'----------------------colaboracion cuerpos--------------------------

Private Sub cmdAgregarCuerpo_Click()
    Set mFrmColaboracion = New frmColaboracion
    mFrmColaboracion.NuevaColaboracion
End Sub

Private Sub cmdEditarCuerpo_Click()
    If Not lvwOtrosCuerpos.SelectedItem Is Nothing Then
        Set mFrmColaboracion = New frmColaboracion
        mFrmColaboracion.ModificarColaboracion lvwOtrosCuerpos.SelectedItem
    End If
End Sub

Private Sub cmdEliminarCuerpo_Click()
'
End Sub

Private Sub mFrmColaboracion_ColaboracionNueva(pColaboracion As BLCemi.Colaboracion)
    lvwOtrosCuerpos.Coleccion.AddItem pColaboracion
    lvwOtrosCuerpos.Refresh
End Sub

Private Sub mFrmColaboracion_ColaboracionModificada(pColaboracion As BLCemi.Colaboracion)
    lvwOtrosCuerpos.Refresh
    Set lvwOtrosCuerpos.SelectedItem = pColaboracion
End Sub


'----------------------vehiculos--------------------------

Private Sub cmdAgregarVehiculo_Click()
    Set mFrmABMVehiculo = New frmABMVehiculo
    mFrmABMVehiculo.Nuevo lvwVehiculos.Coleccion
End Sub

Private Sub cmdEditarVehiculo_Click()
    If Not lvwVehiculos.SelectedItem Is Nothing Then
        Set mFrmABMVehiculo = New frmABMVehiculo
        mFrmABMVehiculo.Modificar lvwVehiculos.SelectedItem
    End If
End Sub

Private Sub mFrmABMVehiculo_NuevoVehiculo(pVehiculo As BLCemi.Vehiculo)
    lvwVehiculos.Refresh
End Sub

Private Sub mFrmABMVehiculo_VehiculoModificado(pVehiculo As BLCemi.Vehiculo)
    lvwVehiculos.Refresh
End Sub

'----------------------tiempos--------------------------
Private Sub cmdSalidaPreInspeccion_Click()
    txtSalidaPreInspeccion = Time
    cmdSalidaPreInspeccion.Visible = False
    cmdArriboPreinspeccion.Enabled = True
End Sub

Private Sub cmdArriboPreinspeccion_Click()
    txtPreInspeccion = Time
    cmdArriboPreinspeccion.Visible = False
End Sub

Private Sub cmdSalidaDotacion_Click()
    txtSalidaDotacion = Time
    cmdRegistrarArribo.Enabled = True
    cmdSalidaDotacion.Visible = False
End Sub

Private Sub cmdRegistrarArribo_Click()
    txtQTH = Time
    cmdRegistrarArribo.Visible = False
    cmdRegistrarLiberacion.Enabled = True
End Sub

Private Sub cmdRegistrarLiberacion_Click()
    txtVL = Time
    cmdRegistrarLiberacion.Visible = False
End Sub

'------------busqueda de sintomas----------------------------
Private Sub cmdBuscarSintoma_Click()
    Set mFrmBuscarSintoma = New frmSeleccionarSintoma
    mFrmBuscarSintoma.Show
End Sub


Private Sub mFrmBuscarSintoma_SeleccionCancelada()
'
End Sub

Private Sub mFrmBuscarSintoma_SintomaSeleccionado(pSintoma As BLCemi.Sintoma)
    Set cmbSintoma.SelectedItem = pSintoma
End Sub

Private Sub cmdDotacion_Click()
    Set mFrmConsultarDotacion = New frmConsultarDotaciones
    mFrmConsultarDotacion.Consultar GBL.EquiposGBL, etConRetorno
End Sub

'--------------------gestion de involucrados----------------------

Private Sub cmdAgregarAfectado_Click()
    Set mFrmABMInvolucrado = New frmABMInvolucrado
    mFrmABMInvolucrado.Nuevo mInvolucrados, BLCemi.eAfectado
End Sub

Private Sub cmdAgregarSolicitante_Click()
    Set mFrmABMInvolucrado = New frmABMInvolucrado
    mFrmABMInvolucrado.Nuevo mInvolucrados, BLCemi.eSolicitante
End Sub

Private Sub lvwPersonas_ItemClick(Item As Object)
    Set mFrmABMInvolucrado = New frmABMInvolucrado
    mFrmABMInvolucrado.Modificar Item, BLCemi.eAfectado
End Sub

Private Sub cmdEditarAfectado_Click()
    If Not lvwPersonas.SelectedItem Is Nothing Then
        Set mFrmABMInvolucrado = New frmABMInvolucrado
        mFrmABMInvolucrado.Modificar lvwPersonas.SelectedItem, BLCemi.eAfectado
    End If
End Sub

Private Sub lvwSolicitantes_ItemDblClick(Item As Object)
    Set mFrmABMInvolucrado = New frmABMInvolucrado
    mFrmABMInvolucrado.Modificar Item, BLCemi.eSolicitante
End Sub

Private Sub cmdEditarSolicitante_Click()
    If Not lvwSolicitantes.SelectedItem Is Nothing Then
        Set mFrmABMInvolucrado = New frmABMInvolucrado
        mFrmABMInvolucrado.Modificar lvwSolicitantes.SelectedItem, BLCemi.eSolicitante
    End If
End Sub

Private Sub mFrmABMInvolucrado_InvolucradoModificado(pInvolucrado As BLCemi.Involucrado)
    ActualizarInvolucrados
End Sub

Private Sub mFrmABMInvolucrado_NuevoInvolucrado(pInvolucrado As BLCemi.Involucrado)
    ActualizarInvolucrados
End Sub

Private Sub ActualizarInvolucrados()
    Set lvwSolicitantes.Coleccion = mInvolucrados.GetByTipo(BLCemi.eSolicitante)
    Set lvwPersonas.Coleccion = mInvolucrados.GetByTipo(BLCemi.eAfectado)
End Sub

'Private Sub lblDireccion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbLeftButton Then lblDireccion.Drag
'End Sub

Private Sub mFrmConsultarDotacion_EquipoSeleccionado(pEquipo As BLCemi.Equipo)
    mAtencion.Equipos.AddItem pEquipo
    tvw.Refresh
End Sub

Private Sub mFrmConsultarDotacion_EquiposSeleccionados(pEquipos As BLCemi.EquipoManager)
    Set mAtencion.Equipos = pEquipos
    Set tvw.Coleccion = mAtencion.Equipos
    tvw.Refresh
End Sub

Private Sub mFrmConsultarAfiliado_SeleccionCancelada()
Me.SetFocus
End Sub

Private Sub cmbCodigo_ItemSeleccionado(Item As Object)
    Dim codEm As BLCemi.CodigoEmergencia
    Set codEm = Item
    If Not cmbSintoma.SelectedItem Is Nothing Then
        If cmbSintoma.SelectedItem.Parent.id <> codEm.id Then
            Set cmbSintoma.Coleccion = codEm.Sintomas
        End If
    Else
        Set cmbSintoma.Coleccion = cmbCodigo.SelectedItem.Sintomas
    End If
End Sub

Private Sub cmbSintoma_ItemSeleccionado(Item As Object)
    txtSintoma = ""
    txtCodigo = ""
    Set cmbCodigo.SelectedItem = Item.Parent
End Sub

Private Sub txtCodigo_Change()
    Dim cod As BLCemi.CodigoEmergencia
    Set cod = GBL.CodigoEmergenciaGBL(Val(txtCodigo))
    If Not cod Is Nothing Then
        txtSintoma = ""
        Set cmbCodigo.SelectedItem = cod
    End If
End Sub

Private Sub txtPoliciaCantEfectivos_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub

Private Sub txtSintoma_Change()
On Error Resume Next
If txtSintoma <> "" Then
    txtCodigo = ""
    Dim sint As BLCemi.Sintoma
    Set sint = GBL.SintomasGBL.Item(Val(txtSintoma.Text))
    If Not sint Is Nothing Then
        Set cmbSintoma.SelectedItem = sint
    End If
End If
End Sub

Private Sub tvw_ItemKeyDeletePressed(Item As Object)
mAtencion.Equipos.Remove Item.id
tvw.Refresh
End Sub
