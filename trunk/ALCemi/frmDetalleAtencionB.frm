VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmDetalleAtencionB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Emergencia"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   9555
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Principal"
      TabPicture(0)   =   "frmDetalleAtencionB.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tvw"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtReseña"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtObservaciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fraAfiliadoExterno"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "fraComun"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraCodigoEmerg"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Datos del Siniestro"
      TabPicture(1)   =   "frmDetalleAtencionB.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAtencion"
      Tab(1).Control(1)=   "fraSolicitantes"
      Tab(1).Control(2)=   "ctlDirEmergencia"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Personas Afectadas"
      TabPicture(2)   =   "frmDetalleAtencionB.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Bienes Afectados"
      TabPicture(3)   =   "frmDetalleAtencionB.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraComplementarios"
      Tab(3).Control(1)=   "fraCampo"
      Tab(3).Control(2)=   "fraVivienda"
      Tab(3).Control(3)=   "fraVehiculos"
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Colaboracion"
      TabPicture(4)   =   "frmDetalleAtencionB.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame1"
      Tab(4).Control(1)=   "fraPolicia"
      Tab(4).Control(2)=   "fraBomberos"
      Tab(4).ControlCount=   3
      TabCaption(5)   =   "Historial"
      TabPicture(5)   =   "frmDetalleAtencionB.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame15"
      Tab(5).Control(1)=   "Frame23"
      Tab(5).ControlCount=   2
      Begin VB.Frame Frame1 
         Caption         =   "Servicios Emergencias"
         Height          =   1095
         Left            =   -74880
         TabIndex        =   82
         Top             =   4320
         Width           =   9135
         Begin VB.CheckBox chkAmbulancias 
            Caption         =   "Ambulancias."
            Height          =   255
            Left            =   7440
            TabIndex        =   86
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtCentroAsistenciales 
            Height          =   285
            Left            =   1800
            MaxLength       =   255
            TabIndex        =   85
            Top             =   720
            Width           =   5535
         End
         Begin VB.TextBox txtMedicoMP 
            Height          =   285
            Left            =   8280
            MaxLength       =   25
            TabIndex        =   84
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtMedicoNombre 
            Height          =   285
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   83
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label25 
            Caption         =   "Centros Asistenciales:"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblMP 
            Caption         =   "MP:"
            Height          =   255
            Left            =   7440
            TabIndex        =   88
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "Médico:"
            Height          =   255
            Left            =   120
            TabIndex        =   87
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame fraPolicia 
         Caption         =   "Policía"
         Height          =   855
         Left            =   -74880
         TabIndex        =   75
         Top             =   3240
         Width           =   9135
         Begin VB.TextBox txtPoliciaNroMovil 
            Height          =   285
            Left            =   8280
            MaxLength       =   50
            TabIndex        =   78
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox txtPoliciaCantEfectivos 
            Height          =   285
            Left            =   6600
            TabIndex        =   77
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtPoliciaACargo 
            Height          =   285
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   76
            Top             =   360
            Width           =   3495
         End
         Begin VB.Label Label23 
            Caption         =   "Móvil Nº:"
            Height          =   255
            Left            =   7440
            TabIndex        =   81
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label22 
            Caption         =   "Cantidad de Efectivos:"
            Height          =   255
            Left            =   4800
            TabIndex        =   80
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label21 
            Caption         =   "A Cargo De:"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fraBomberos 
         Caption         =   "Colaboración de Otros Cuerpos"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   71
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtEquiposEspeciales 
            Height          =   615
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   72
            Top             =   1800
            Width           =   8895
         End
         Begin ControlesPOO.ListViewConsulta lvwOtrosCuerpos 
            Height          =   1215
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
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
         Begin VB.Label Label20 
            Caption         =   "Equipos especiales:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   1560
            Width           =   2895
         End
      End
      Begin VB.Frame fraComplementarios 
         Caption         =   "Datos Complementarios"
         Height          =   1455
         Left            =   -74880
         TabIndex        =   62
         Top             =   4200
         Width           =   9135
         Begin VB.TextBox txtGas 
            Height          =   285
            Left            =   6840
            TabIndex        =   92
            Top             =   360
            Width           =   2175
         End
         Begin VB.TextBox txtElectricidad 
            Height          =   285
            Left            =   2280
            TabIndex        =   91
            Top             =   360
            Width           =   2895
         End
         Begin VB.TextBox txtMaterialPrevencionFuego 
            Height          =   285
            Left            =   5400
            TabIndex        =   67
            Top             =   1080
            Width           =   3615
         End
         Begin VB.CheckBox chkPrevencionFuego 
            Caption         =   "Material de Prevención y Lucha contra el Fuego.         Descripcion:"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   1080
            Width           =   5055
         End
         Begin VB.TextBox txtPoliza 
            Height          =   285
            Left            =   6840
            TabIndex        =   65
            Top             =   720
            Width           =   2175
         End
         Begin VB.TextBox txtAseguradora 
            Height          =   285
            Left            =   2280
            TabIndex        =   64
            Top             =   720
            Width           =   2895
         End
         Begin VB.CheckBox chkSeguro 
            Caption         =   "Seguro.        Aseguradora:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   720
            Width           =   2175
         End
         Begin VB.Label Label19 
            Caption         =   "Poliza Nº:"
            Height          =   255
            Left            =   5400
            TabIndex        =   70
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label18 
            Caption         =   "Instalación de Gas:"
            Height          =   255
            Left            =   5400
            TabIndex        =   69
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label16 
            Caption         =   "Instalación Eléctrica:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame fraCampo 
         Caption         =   "Campo - Baldío"
         Height          =   2055
         Left            =   -70560
         TabIndex        =   55
         Top             =   2040
         Width           =   4815
         Begin VB.TextBox txtMaterialesCombustibles 
            Height          =   285
            Left            =   1920
            TabIndex        =   58
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtPerjuiciosCampo 
            Height          =   735
            Left            =   120
            TabIndex        =   57
            Top             =   1200
            Width           =   4575
         End
         Begin VB.TextBox txtHectareas 
            Height          =   285
            Left            =   1920
            TabIndex        =   56
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Materiales Combustibles:"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "Descripción daños:"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label8 
            Caption         =   "Hectareas Afectadas:"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame fraVivienda 
         Caption         =   "Vivienda"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   50
         Top             =   2040
         Width           =   4215
         Begin VB.TextBox txtPerjuiciosVivienda 
            Height          =   1095
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   3975
         End
         Begin VB.TextBox txtAmbientes 
            Height          =   285
            Left            =   1920
            TabIndex        =   51
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblPerjuicios 
            Caption         =   "Descripción daños:"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label7 
            Caption         =   "Ambientes Afectados:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.Frame fraVehiculos 
         Caption         =   "Vehiculos"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   48
         Top             =   360
         Width           =   9135
         Begin ControlesPOO.ListViewConsulta lvwVehiculos 
            Height          =   1215
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
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
      End
      Begin VB.Frame Frame2 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   46
         Top             =   480
         Width           =   9135
         Begin ControlesPOO.ListViewConsulta lvwPersonas 
            Height          =   3255
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
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
      End
      Begin VB.Frame fraAtencion 
         Caption         =   "Referencias"
         Height          =   2415
         Left            =   -70200
         TabIndex        =   40
         Top             =   3120
         Width           =   4455
         Begin VB.TextBox txtReferencias 
            Height          =   765
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   4215
         End
         Begin VB.TextBox txtAccesoPor 
            Height          =   645
            Left            =   120
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   41
            Top             =   1680
            Width           =   4215
         End
         Begin VB.Label Label26 
            Caption         =   "Referencia más próxima:"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label Label27 
            Caption         =   "Acceso por:"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   1440
            Width           =   1815
         End
      End
      Begin VB.Frame fraSolicitantes 
         Caption         =   "Datos de los Solicitantes"
         Height          =   2535
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   9135
         Begin ControlesPOO.ListViewConsulta lvwSolicitantes 
            Height          =   2175
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
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
      End
      Begin VB.Frame fraCodigoEmerg 
         Caption         =   "Seleccione el Grado y Tipo del Siniestro"
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtCodigo 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   720
            TabIndex        =   30
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtSintoma 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4680
            TabIndex        =   29
            Top             =   240
            Width           =   375
         End
         Begin ControlesPOO.Combo cmbSintoma 
            Height          =   315
            Left            =   5160
            TabIndex        =   31
            Top             =   240
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            AtributoAMostrar=   "nombrecompuesto"
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbCodigo 
            Height          =   315
            Left            =   1200
            TabIndex        =   32
            Top             =   240
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   556
            AtributoAMostrar=   "nombrecompuesto"
            Enabled         =   -1  'True
         End
         Begin VB.Label lblTipo 
            Caption         =   "Grado:"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblSintoma 
            Caption         =   "Tipo:"
            Height          =   195
            Left            =   4080
            TabIndex        =   33
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame fraComun 
         Caption         =   "Tiempos"
         Height          =   2895
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   3975
         Begin VB.TextBox txtSalidaPreInspeccion 
            Height          =   285
            Left            =   2400
            TabIndex        =   20
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox txtSalidaDotacion 
            Height          =   285
            Left            =   2400
            TabIndex        =   19
            Top             =   1800
            Width           =   1215
         End
         Begin VB.TextBox txtPreInspeccion 
            Height          =   285
            Left            =   2400
            TabIndex        =   18
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtQTH 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   17
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txtVL 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   2520
            Width           =   1215
         End
         Begin VB.TextBox txtHora 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   720
            Width           =   1215
         End
         Begin VB.TextBox txtFecha 
            Height          =   285
            Left            =   2400
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Salida Pre-Inspección:"
            Height          =   255
            Left            =   240
            TabIndex        =   27
            Top             =   1080
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "Salida Dotación:"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "Arribo Pre-Inspección:"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   1440
            Width           =   1695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "QTH:"
            Height          =   195
            Left            =   240
            TabIndex        =   24
            Top             =   2160
            Width           =   390
         End
         Begin VB.Label Label14 
            Caption         =   "VL:"
            Height          =   195
            Left            =   240
            TabIndex        =   23
            Top             =   2520
            Width           =   240
         End
         Begin VB.Label Label12 
            Caption         =   "Hora Llamado:"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame fraAfiliadoExterno 
         Height          =   735
         Left            =   4320
         TabIndex        =   8
         Top             =   1320
         Width           =   4935
         Begin VB.TextBox txtNroInterno 
            Height          =   285
            Left            =   3840
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox txtNroIncidente 
            Height          =   285
            Left            =   2040
            TabIndex        =   9
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "Interno:"
            Height          =   255
            Left            =   3120
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label32 
            Caption         =   "Nro. de Incidente Externo:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   855
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   4560
         Width           =   3975
      End
      Begin VB.TextBox txtReseña 
         Height          =   855
         Left            =   4320
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   4560
         Width           =   4935
      End
      Begin VB.Frame Frame15 
         Caption         =   "Historial de Modificaciones"
         Height          =   4215
         Left            =   -74880
         TabIndex        =   4
         Top             =   1440
         Width           =   9135
         Begin ControlesPOO.ListViewConsulta lvwCambios 
            Height          =   3855
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   6800
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
            NEncabezado0    =   "Fecha"
            MEncabezado0    =   "Fecha"
            AEncabezado0    =   10
            NEncabezado1    =   "Hora"
            MEncabezado1    =   "Hora"
            AEncabezado1    =   10
            NEncabezado2    =   "Empleado"
            MEncabezado2    =   "pgempleado"
            AEncabezado2    =   17
            NEncabezado3    =   "Campo"
            MEncabezado3    =   "Campo"
            AEncabezado3    =   13
            NEncabezado4    =   "Valor Anterior"
            MEncabezado4    =   "pgValorAnterior"
            AEncabezado4    =   25
            NEncabezado5    =   "Valor Nuevo"
            MEncabezado5    =   "pgValorNuevo"
            AEncabezado5    =   25
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
      Begin VB.Frame Frame23 
         Height          =   855
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   9135
         Begin VB.TextBox txtDespachador 
            Height          =   285
            Left            =   2400
            TabIndex        =   2
            Top             =   360
            Width           =   6615
         End
         Begin VB.Label Label16 
            Caption         =   "El siniestro fue registrado por:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   2535
         End
      End
      Begin ControlesPOO.TreeViewConsulta tvw 
         Height          =   1695
         Left            =   4320
         TabIndex        =   35
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
         SoloConsulta    =   -1  'True
         EntrecallesVisible=   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Dotación:"
         Height          =   255
         Left            =   4320
         TabIndex        =   90
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   4320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Reseña de la Actuación:"
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   4320
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmDetalleAtencionB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAtencionB As blcemi.AtencionB

Private Sub MostrarFrames()
    'cmdCancelar.Left = 8760
   ' cmdCancelar.Top = cmdGuardar.Top
    Me.Width = 9645
    Me.Height = 6360
End Sub

Private Sub cmdCancelar_Click()
     
    Unload Me

End Sub

Public Sub VerDetalleAtencionB(pAtencionB As blcemi.AtencionB)
    Set mAtencionB = pAtencionB
    
    Me.Show
    MostrarFrames

    Set ctlDirEmergencia.MiDireccion = mAtencionB.Direccion
    'ver por el check txtAseguradora=mAtencion.Aseguradora
    
    Set lvwSolicitantes.Coleccion = mAtencionB.Involucrados.GetByTipo(blcemi.eSolicitante)
    Set lvwPersonas.Coleccion = mAtencionB.Involucrados.GetByTipo(blcemi.eAfectado)
    
    Set tvw.Coleccion = mAtencionB.Equipos
    Set lvwVehiculos.Coleccion = mAtencionB.Vehiculos
    Set lvwOtrosCuerpos.Coleccion = mAtencionB.ColaboracionBomberos

    txtFecha.Text = mAtencionB.fecha
    txtHora.Text = mAtencionB.HoraLlamada
    txtNroIncidente = mAtencionB.NroIncidente
    txtNroInterno = mAtencionB.nroIncidenteInterno
    txtObservaciones = mAtencionB.Observaciones
    
    txtPerjuiciosCampo = mAtencionB.DescripcionPerjuiciosCampo
    txtMaterialesCombustibles = mAtencionB.MaterialesCombustibles
    txtPerjuiciosVivienda = mAtencionB.DescripcionPerjuiciosVivienda
    txtAccesoPor = mAtencionB.AccesoPor
    txtReferencias = mAtencionB.Referencias
    txtMaterialPrevencionFuego = mAtencionB.DescripcionMaterial
    txtEquiposEspeciales = mAtencionB.EquiposEspeciales
    txtPoliza = mAtencionB.Poliza
    txtAmbientes = mAtencionB.AmbientesAfectadosVivienda
    txtHectareas = mAtencionB.HectareasAfectadasCampo
    If mAtencionB.InstalacionElectrica Is Nothing Then
        txtElectricidad = "-Sin asignar-"
    Else
        txtElectricidad = mAtencionB.InstalacionElectrica.Nombre
    End If
    If mAtencionB.InstalacionGas Is Nothing Then
        txtGas = "-Sin asignar-"
    Else
        txtGas = mAtencionB.InstalacionGas.Nombre
    End If
    txtPoliciaACargo = mAtencionB.PoliciaACargo
    txtPoliciaCantEfectivos = mAtencionB.PoliciaCantidad
    txtPoliciaNroMovil = mAtencionB.PoliciaMovil
    txtMedicoNombre = mAtencionB.SEMedico
    txtCentroAsistenciales = mAtencionB.SECentroAsistencial
    txtMedicoMP = mAtencionB.SEMedicoMP
    chkAmbulancias.Value = IIf(mAtencionB.SEAmbulancias, 1, 0)

    
    Set cmbSintoma.SelectedItem = mAtencionB.Sintoma
    Set cmbCodigo.SelectedItem = mAtencionB.Sintoma.Parent
    txtSalidaPreInspeccion = mAtencionB.SalidaPreInspeccion
    txtPreInspeccion = mAtencionB.LlegadaPreInspeccion
    txtSalidaDotacion = mAtencionB.SalidaDotacion
    txtVL = mAtencionB.VL
    txtQTH = mAtencionB.QTH
           
    Dim detalles As New blcemi.DetalleSegManager
   'warning: revisar esto conrespecto al segiumiento
    detalles.CargarXAtencion mAtencionB.id
    Set lvwCambios.Coleccion = detalles
    
    txtDespachador = mAtencionB.Despachador.NombreCompleto
    On Error Resume Next
    Dim ctl As Control
    For Each ctl In Me.Controls
        If TypeOf ctl Is TextBox Then
            ctl.Locked = True
        End If
    Next
    
End Sub


Private Sub Form_Load()
            
    Set cmbCodigo.Coleccion = GBL.CodigoEmergenciaGBL
    Set ctlDirEmergencia.MiDireccion = New blcemi.Direccion
        
    InicializarDireccion ctlDirEmergencia
    Set cmbSintoma.Coleccion = GBL.SintomasGBL
    'txtFecha.Text = Date
   ' dtpHora.Value = Time
   ' txtHora.Text = Time
   Set Me.Icon = MDI.Icon

    Me.Move (MDI.Width - 10650) / 2, 0 'no uso me.width porq lo cambio despues del load
    AplicarConfiguracion
End Sub

Private Sub AplicarConfiguracion()
    lvwCambios.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
End Sub

Public Sub Refrescar()
    AplicarConfiguracion
End Sub


'Dim mAtencionB As blcemi.AtencionB
'
'Public Sub MostrarDetalles(pAtencionB As blcemi.AtencionB)
'    Set mAtencionB = pAtencionB
''    mAtencionB.Afiliado
''    mAtencionB.AfiliadoExterno
''    mAtencionB.AreaProtegida
''    mAtencionB.Despachador
''    mAtencionB.Diagnostico
''    mAtencionB.Direccion
''    mAtencionB.DireccionDestino
''    mAtencionB.DireccionOrigen
''    mAtencionB.Equipos
''    mAtencionB.Estado
''    mAtencionB.fecha
''    mAtencionB.HoraLlamada
''    mAtencionB.id
''    mAtencionB.NroIncidente
''    mAtencionB.ObraSocial
''    mAtencionB.Observaciones
''    mAtencionB.Operador
''    mAtencionB.QTH
''    mAtencionB.ServicioEmergencia
''    mAtencionB.Sintoma
''    mAtencionB.Telefono
''    mAtencionB.TelefonoAuxilar
''    mAtencionB.VL
'
'    AppendBigTitle "Ficha de AtencionB"
'    AppendLine
'    AppendLine
'    AppendSubTitle "Fecha: "
'    AppendContent "23/23/2009"
'    AppendSubTitle "Hora: "
'    AppendContent "12.12.12 am."
'
'
'
'End Sub
'
'Public Sub AppendBigTitle(pTitle As String)
'    rTxt.SelStart = Len(rTxt.Text)
'    rTxt.SelLength = Len(pTitle)
'    rTxt.SelColor = vbBlack
'    rTxt.SelBold = True
'    rTxt.SelFontSize = 12
'    rTxt.SelUnderline = True
'    rTxt.SelText = pTitle
'
'End Sub
'
'Public Sub AppendTitle(pTitle As String)
'    rTxt.SelStart = Len(rTxt.Text)
'    rTxt.SelLength = Len(pTitle)
'    rTxt.SelColor = vbBlack
'    rTxt.SelBold = True
'    rTxt.SelUnderline = False
'    rTxt.SelFontSize = 10
'    rTxt.SelText = pTitle
'
'End Sub
'Public Sub AppendSubTitle(pTitle As String)
'    rTxt.SelStart = Len(rTxt.Text)
'    rTxt.SelLength = Len(pTitle)
'    rTxt.SelColor = vbBlack
'    rTxt.SelBold = True
'     rTxt.SelUnderline = False
'     rTxt.SelFontSize = 8.5
'    rTxt.SelText = pTitle
'End Sub
'
'Public Sub AppendContent(pContent As String)
'    rTxt.SelStart = Len(rTxt.Text)
'    rTxt.SelLength = Len(pContent)
'    rTxt.SelColor = vbBlack
'    rTxt.SelBold = False
'     rTxt.SelUnderline = False
'     rTxt.SelFontSize = 8.5
'    rTxt.SelText = pContent
'End Sub
'
'Public Sub AppendTabedContent(pContent As String)
'    rTxt.SelStart = Len(rTxt.Text)
'    rTxt.SelLength = Len(pContent) + 1
'    rTxt.SelColor = vbBlack
'    rTxt.SelBold = False
'     rTxt.SelUnderline = False
'     rTxt.SelFontSize = 8.5
'    rTxt.SelText = vbCrLf + vbTab + pContent
'End Sub
'
'Public Sub AppendLine()
'    rTxt.SelStart = Len(rTxt.Text)
'    rTxt.SelLength = 1
'    rTxt.SelColor = vbBlack
'    rTxt.SelText = vbCrLf
'End Sub
'
'Private Sub Form_Load()
'MostrarDetalles Nothing
'End Sub
Private Sub SSTab1_DblClick()

End Sub
