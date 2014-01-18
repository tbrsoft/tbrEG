VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmDetalleAtencion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Atencion"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10785
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Atención"
      TabPicture(0)   =   "frmDetalleAtencion.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAtencion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraTraslado"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraCodigoEmerg"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraComun"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fraAfiliadoExterno"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraAfiliadoPropio"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Historial"
      TabPicture(1)   =   "frmDetalleAtencion.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   -74880
         TabIndex        =   96
         Top             =   480
         Width           =   10335
         Begin VB.TextBox txtDespachador 
            Height          =   285
            Left            =   4440
            TabIndex        =   97
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label Label16 
            Caption         =   "La atencion fue registrada por:"
            Height          =   255
            Left            =   1200
            TabIndex        =   98
            Top             =   360
            Width           =   2295
         End
      End
      Begin VB.Frame fraAfiliadoPropio 
         Caption         =   "Afiliado Propio"
         Height          =   3615
         Left            =   120
         TabIndex        =   76
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Label lblNroAfiliado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   95
            Top             =   360
            Width           =   2520
         End
         Begin VB.Label lblDireccionaPropio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   94
            Top             =   2880
            Width           =   3000
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   "Direccion:"
            Height          =   195
            Left            =   855
            TabIndex        =   93
            ToolTipText     =   "Si esta es la direccion de la emergencia, la puede arrastrar..."
            Top             =   2880
            Width           =   720
         End
         Begin VB.Label lblOcupacion 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   92
            Top             =   2565
            Width           =   3000
         End
         Begin VB.Label lblSexoAPropio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   91
            Top             =   1935
            Width           =   3000
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Sexo:"
            Height          =   195
            Left            =   1170
            TabIndex        =   90
            Top             =   1935
            Width           =   405
         End
         Begin VB.Label lblEstadoCivil 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   89
            Top             =   1620
            Width           =   3000
         End
         Begin VB.Label lblEdadAPropio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   88
            Top             =   1305
            Width           =   3000
         End
         Begin VB.Label lblTipoDoc 
            Alignment       =   1  'Right Justify
            Caption         =   "DNI:"
            Height          =   195
            Left            =   1245
            TabIndex        =   87
            Top             =   990
            Width           =   330
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Edad:"
            Height          =   195
            Left            =   1155
            TabIndex        =   86
            Top             =   1305
            Width           =   420
         End
         Begin VB.Label lblApeNomAfilPropio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   85
            Top             =   675
            Width           =   3000
         End
         Begin VB.Label Label1 
            Caption         =   "Apellido y Nombre:"
            Height          =   195
            Index           =   0
            Left            =   255
            TabIndex        =   84
            Top             =   675
            Width           =   1320
         End
         Begin VB.Label lblNroDocAPropio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   83
            Top             =   990
            Width           =   3000
         End
         Begin VB.Label lblDatosAPropio 
            Alignment       =   1  'Right Justify
            Caption         =   "Nº de afiliado:"
            Height          =   195
            Index           =   0
            Left            =   585
            TabIndex        =   82
            Top             =   360
            Width           =   990
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Ocupacion:"
            Height          =   195
            Left            =   750
            TabIndex        =   81
            Top             =   2565
            Width           =   825
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado Civil:"
            Height          =   195
            Left            =   705
            TabIndex        =   80
            Top             =   1620
            Width           =   870
         End
         Begin VB.Label Label27 
            Caption         =   "Obra Social:"
            Height          =   195
            Left            =   705
            TabIndex        =   79
            Top             =   2250
            Width           =   870
         End
         Begin VB.Label lblObraSocialAPropio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1680
            TabIndex        =   78
            Top             =   2250
            Width           =   3000
         End
         Begin VB.Label lblCiudadAPropio 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1680
            TabIndex        =   77
            Top             =   3240
            Width           =   3135
         End
      End
      Begin VB.Frame fraAfiliadoExterno 
         Height          =   3735
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Frame fraAreaProtegida 
            Caption         =   "Area Protegida"
            Height          =   1935
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Visible         =   0   'False
            Width           =   4815
            Begin VB.Label lblSinDatosAP 
               Caption         =   "No se registran datos."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   63
               Top             =   840
               Width           =   2415
            End
            Begin VB.Label lblCiudadArea 
               Caption         =   "Bº Alto Alberdi, Cordoba, Cordoba."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   62
               Top             =   1600
               Visible         =   0   'False
               Width           =   3135
            End
            Begin VB.Label lblDireccionArea 
               AutoSize        =   -1  'True
               Caption         =   "Sol de Mayo 321."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   61
               Top             =   1320
               Visible         =   0   'False
               Width           =   3180
            End
            Begin VB.Label lblDatosArea 
               Alignment       =   1  'Right Justify
               Caption         =   "Direccion:"
               Height          =   255
               Index           =   1
               Left            =   600
               TabIndex        =   60
               ToolTipText     =   "Si esta es la direccion de la emergencia, la puede arrastrar..."
               Top             =   1320
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblDatosArea 
               Alignment       =   1  'Right Justify
               Caption         =   "DNI:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   59
               Top             =   960
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label lblNombreArea 
               AutoSize        =   -1  'True
               Caption         =   "Colegio Lourdes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   58
               Top             =   240
               Visible         =   0   'False
               Width           =   1380
            End
            Begin VB.Label lblDatosArea 
               Alignment       =   1  'Right Justify
               Caption         =   "Nombre Area:"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   57
               Top             =   240
               Visible         =   0   'False
               Width           =   1320
            End
            Begin VB.Label lblDocRespArea 
               AutoSize        =   -1  'True
               Caption         =   "28.974.150"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   56
               Top             =   960
               Visible         =   0   'False
               Width           =   3075
            End
            Begin VB.Label lblDatosArea 
               Alignment       =   1  'Right Justify
               Caption         =   "Apellido y nombre:"
               Height          =   195
               Index           =   3
               Left            =   120
               TabIndex        =   55
               ToolTipText     =   "Apellido y nombre del responsable del area"
               Top             =   600
               Visible         =   0   'False
               Width           =   1290
            End
            Begin VB.Label lblResponsableArea 
               Caption         =   "Paliza, Martin"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   54
               Top             =   600
               Visible         =   0   'False
               Width           =   3015
            End
         End
         Begin VB.Frame fraServicioEmergencia 
            Caption         =   "Servicio Emergencia"
            Height          =   1935
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Visible         =   0   'False
            Width           =   4815
            Begin VB.Label lblSinDatosSE 
               Caption         =   "No se registran datos."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   52
               Top             =   840
               Width           =   2415
            End
            Begin VB.Label lblNombreSE 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   51
               Top             =   360
               Visible         =   0   'False
               Width           =   3135
            End
            Begin VB.Label lblCiudadSE 
               Caption         =   "Bº Alto Alberdi, Cordoba, Cordoba."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   50
               Top             =   1080
               Visible         =   0   'False
               Width           =   3135
            End
            Begin VB.Label lblDireccionSE 
               AutoSize        =   -1  'True
               Caption         =   "Sol de Mayo 321."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1200
               TabIndex        =   49
               Top             =   720
               Visible         =   0   'False
               Width           =   3180
            End
            Begin VB.Label lblDatosSE 
               Alignment       =   1  'Right Justify
               Caption         =   "Direccion:"
               Height          =   255
               Index           =   1
               Left            =   240
               TabIndex        =   48
               ToolTipText     =   "Si esta es la direccion de la emergencia, la puede arrastrar..."
               Top             =   720
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblDatosSE 
               Caption         =   "Nombre:"
               Height          =   195
               Index           =   0
               Left            =   495
               TabIndex        =   47
               Top             =   360
               Visible         =   0   'False
               Width           =   600
            End
         End
         Begin VB.Frame fraObraSocial 
            Caption         =   "Obra Social"
            Height          =   1935
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Visible         =   0   'False
            Width           =   4815
            Begin VB.Label lblSinDatosOS 
               Caption         =   "No se registran datos."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   45
               Top             =   840
               Width           =   2415
            End
            Begin VB.Label lblCodCubiertosOS 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3720
               TabIndex        =   44
               Top             =   600
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label lblServicioEmergenciaOS 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   43
               Top             =   960
               Visible         =   0   'False
               Width           =   3135
            End
            Begin VB.Label lblCoseguro 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   42
               Top             =   600
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblNombreOS 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   41
               Top             =   240
               Visible         =   0   'False
               Width           =   3135
            End
            Begin VB.Label lblDatosOS 
               Caption         =   "Serv. Emergencia:"
               Height          =   255
               Index           =   3
               Left            =   105
               TabIndex        =   40
               Top             =   960
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.Label lblDatosOS 
               Caption         =   "Cod. Cubiertos:"
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   39
               Top             =   600
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.Label lblDatosOS 
               Caption         =   "Coseguro:"
               Height          =   255
               Index           =   2
               Left            =   705
               TabIndex        =   38
               Top             =   600
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblCiudadOS 
               Caption         =   "Bº Alto Alberdi, Cordoba, Cordoba."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   37
               Top             =   1560
               Visible         =   0   'False
               Width           =   3135
            End
            Begin VB.Label lblDireccionOS 
               AutoSize        =   -1  'True
               Caption         =   "Sol de Mayo 321."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1560
               TabIndex        =   36
               Top             =   1320
               Visible         =   0   'False
               Width           =   3180
            End
            Begin VB.Label lblDatosOS 
               Alignment       =   1  'Right Justify
               Caption         =   "Direccion:"
               Height          =   255
               Index           =   4
               Left            =   585
               TabIndex        =   35
               ToolTipText     =   "Si esta es la direccion de la emergencia, la puede arrastrar..."
               Top             =   1320
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblDatosOS 
               Caption         =   "Nombre:"
               Height          =   195
               Index           =   0
               Left            =   840
               TabIndex        =   34
               Top             =   240
               Visible         =   0   'False
               Width           =   600
            End
         End
         Begin VB.TextBox txtNroIncidente 
            Height          =   285
            Left            =   1800
            TabIndex        =   32
            Top             =   240
            Width           =   3135
         End
         Begin VB.Frame fraPacienteArea 
            Caption         =   "Datos Paciente"
            Height          =   1095
            Left            =   120
            TabIndex        =   24
            Top             =   2530
            Width           =   4815
            Begin VB.Label lblSinDatosAE 
               Caption         =   "No se registran datos."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1200
               TabIndex        =   31
               Top             =   480
               Width           =   2415
            End
            Begin VB.Label lblSexoAExterno 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2880
               TabIndex        =   30
               Top             =   720
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lblDatosPaciente 
               Caption         =   "Sexo:"
               Height          =   255
               Index           =   2
               Left            =   2280
               TabIndex        =   29
               Top             =   720
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label lblEdadAExterno 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   28
               Top             =   720
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.Label lblDatosPaciente 
               Caption         =   "Edad:"
               Height          =   195
               Index           =   1
               Left            =   1020
               TabIndex        =   27
               Top             =   720
               Visible         =   0   'False
               Width           =   420
            End
            Begin VB.Label lblApeNomAExterno 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   26
               Top             =   360
               Visible         =   0   'False
               Width           =   2895
            End
            Begin VB.Label lblDatosPaciente 
               Caption         =   "Apellido y Nombre:"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   25
               Top             =   360
               Visible         =   0   'False
               Width           =   1320
            End
         End
         Begin VB.Label Label32 
            Caption         =   "Nro. de Incidente:"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame fraComun 
         Height          =   3615
         Left            =   5280
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   5175
         Begin VB.TextBox txtHora 
            Height          =   285
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtFecha 
            Height          =   285
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtDiagnostico 
            Height          =   1095
            Left            =   2640
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   975
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   2400
            Width           =   4935
         End
         Begin MSComCtl2.DTPicker dtpHora 
            Height          =   285
            Left            =   3840
            TabIndex        =   16
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "HH:mm"
            Format          =   45809666
            CurrentDate     =   39293
            MinDate         =   -105918
         End
         Begin ControlesPOO.TreeViewConsulta tvw 
            Height          =   1095
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   1931
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
         Begin VB.Label Label21 
            Caption         =   "Observaciones:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label22 
            Caption         =   "Diagnostico final:"
            Height          =   255
            Left            =   2640
            TabIndex        =   21
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label12 
            Caption         =   "Hora Llamado:"
            Height          =   195
            Left            =   2640
            TabIndex        =   20
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label15 
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Dotacion:"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame fraCodigoEmerg 
         Caption         =   "Codigo y Sintoma de la Atencion"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   4200
         Visible         =   0   'False
         Width           =   10335
         Begin ControlesPOO.Combo cmbSintoma 
            Height          =   315
            Left            =   5280
            TabIndex        =   7
            Top             =   240
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   556
            AtributoAMostrar=   "nombrecompuesto"
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbCodigo 
            Height          =   315
            Left            =   1080
            TabIndex        =   8
            Top             =   240
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            AtributoAMostrar=   "nombrecompuesto"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label11 
            Caption         =   "Sintoma:"
            Height          =   195
            Left            =   4440
            TabIndex        =   10
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Codigo:"
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame fraTraslado 
         Caption         =   "Traslado"
         Height          =   2895
         Left            =   120
         TabIndex        =   3
         Top             =   4920
         Visible         =   0   'False
         Width           =   10215
         Begin ALCemi.ctlDireccion dirOrigen 
            Height          =   2565
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   3995
            ProvinciaVisible=   0   'False
            Caption         =   ""
            CanDragDrop     =   0   'False
            SoloConsulta    =   0   'False
            EntrecallesVisible=   -1  'True
         End
         Begin ALCemi.ctlDireccion dirDestino 
            Height          =   2565
            Left            =   5280
            TabIndex        =   5
            Top             =   240
            Width           =   4485
            _ExtentX        =   7911
            _ExtentY        =   3995
            ProvinciaVisible=   0   'False
            Caption         =   ""
            CanDragDrop     =   0   'False
            SoloConsulta    =   0   'False
            EntrecallesVisible=   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Historial de Modificaciones"
         Height          =   5895
         Left            =   -74880
         TabIndex        =   1
         Top             =   1560
         Width           =   10335
         Begin ControlesPOO.ListViewConsulta lvwCambios 
            Height          =   5535
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   9763
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
      Begin VB.Frame fraAtencion 
         Caption         =   "Atencion"
         Height          =   2895
         Left            =   120
         TabIndex        =   65
         Top             =   4920
         Visible         =   0   'False
         Width           =   10335
         Begin VB.TextBox txtTelefono 
            Height          =   285
            Left            =   1200
            TabIndex        =   70
            Top             =   840
            Width           =   3015
         End
         Begin VB.CommandButton cmdLlamar 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4320
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Llamar al telefono"
            Top             =   840
            Width           =   375
         End
         Begin VB.TextBox txtVL 
            Height          =   285
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtQTH 
            Height          =   285
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   67
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtOperador 
            Height          =   285
            Left            =   1200
            TabIndex        =   66
            ToolTipText     =   "Ingrese el nombre de la persona que llama"
            Top             =   360
            Width           =   3495
         End
         Begin ALCemi.ctlDireccion ctlDirEmergencia 
            Height          =   2565
            Left            =   5280
            TabIndex        =   71
            Top             =   240
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   3995
            ProvinciaVisible=   0   'False
            Caption         =   "Direccion de la Emergencia"
            CanDragDrop     =   -1  'True
            SoloConsulta    =   0   'False
            EntrecallesVisible=   -1  'True
         End
         Begin VB.Label Label2 
            Caption         =   "Telefono:"
            Height          =   255
            Left            =   240
            TabIndex        =   75
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "VL:"
            Height          =   195
            Left            =   600
            TabIndex        =   74
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "QTH:"
            Height          =   195
            Left            =   2760
            TabIndex        =   73
            Top             =   1320
            Width           =   390
         End
         Begin VB.Label Label9 
            Caption         =   "Operador:"
            Height          =   195
            Left            =   240
            TabIndex        =   72
            Top             =   360
            Width           =   705
         End
      End
   End
End
Attribute VB_Name = "frmDetalleAtencion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mAtencion As blcemi.Atencion

Private mAfiliadoPropio As blcemi.Afiliado
Private mAfiliadoExterno As blcemi.AfiliadoExterno

Private mObraSocial As blcemi.ObraSocial
Private mAreaProtegida As blcemi.AreaProtegida
Private mServicioEmergencia As blcemi.ServicioEmergencia

'Private mEquipos As blcemi.EquipoManager

Private telsAux As blcemi.TelefonoManager


Private Sub fraAfiliadoPropio_DblClick()
If Not mAfiliadoPropio Is Nothing Then
    Dim frm As New frmABMAfiliado
    frm.VerDatos mAfiliadoPropio
End If
End Sub

Private Sub fraAreaProtegida_DblClick()
If Not mAreaProtegida Is Nothing Then
    Dim frm As New frmABMAreaProtegida
    frm.VerDatos mAreaProtegida
End If
End Sub


Private Sub fraObraSocial_DblClick()
If Not mObraSocial Is Nothing Then
    Dim frm As New frmABMObraSocial
    frm.VerDatos mObraSocial
End If
End Sub

Private Sub fraPacienteArea_DblClick()
If Not mAfiliadoExterno Is Nothing Then
    Dim frm As New frmABMAfiliadoExterno
    frm.VerDatos mAfiliadoExterno
End If
End Sub

Private Sub fraServicioEmergencia_DblClick()
If Not mServicioEmergencia Is Nothing Then
    Dim frm As New frmABMServicioEmergencia
    frm.VerDatos mServicioEmergencia
End If
End Sub

Private Sub MostrarFrames()
    fraCodigoEmerg.Visible = True
    fraAtencion.Visible = True
    fraComun.Visible = True
    'cmdCancelar.Left = 8760
   ' cmdCancelar.Top = cmdGuardar.Top
    Me.Width = 10875
    Me.Height = 8265
End Sub


Private Sub cmdCancelar_Click()
     
    Unload Me

End Sub

Public Sub VerDetalleAtencion(pAtencion As blcemi.Atencion)
    Set mAtencion = pAtencion
    
    Me.Show
    MostrarFrames
    
    Set mAfiliadoPropio = mAtencion.Afiliado
    Set mAfiliadoExterno = mAtencion.AfiliadoExterno
    Set mAreaProtegida = mAtencion.AreaProtegida
    Set mObraSocial = mAtencion.ObraSocial
    Set mServicioEmergencia = mAtencion.ServicioEmergencia
    
    If Not mAfiliadoPropio Is Nothing Then
        fraAfiliadoPropio.Visible = True
        LlenarCamposAfiliadoPropio
    End If
    
    If Not mAreaProtegida Is Nothing Then
        fraAfiliadoExterno.Visible = True
        fraAreaProtegida.Visible = True
        LlenarCamposAreaProtegida
    End If
    
    If Not mObraSocial Is Nothing Then
        fraAfiliadoExterno.Visible = True
        fraObraSocial.Visible = True
        LlenarCamposObraSocial
    End If
    
    If Not mServicioEmergencia Is Nothing Then
        fraAfiliadoExterno.Visible = True
        fraServicioEmergencia.Visible = True
        LlenarCamposServicioEmergencia
    End If
    
    'si hay AP, OS o SE, cargo el afiliadoexterno (si hay...)
    If fraAreaProtegida.Visible Or fraObraSocial.Visible Or fraServicioEmergencia.Visible Then
        If Not mAfiliadoExterno Is Nothing Then LlenarCamposAfiliadoExterno
    End If
    
    'si no esta visible alguno de los frames, quiere decir q no selecciono nada, le muestro el fraseleccion
    If Not (fraAreaProtegida.Visible Or fraObraSocial.Visible Or fraServicioEmergencia.Visible Or fraAfiliadoPropio.Visible) Then
        'fraSeleccion.Visible = True
    End If
    
    Set dirDestino.MiDireccion = mAtencion.DireccionDestino
    Set dirOrigen.MiDireccion = mAtencion.DireccionOrigen
    Set ctlDirEmergencia.MiDireccion = mAtencion.Direccion
    
    'mAtencion.Despachador empleadoactual, ver
    txtDiagnostico = mAtencion.Diagnostico
    Set tvw.Coleccion = mAtencion.Equipos
    'mAtencion.Estado ver
    txtFecha.Text = mAtencion.fecha
    dtpHora.Value = mAtencion.HoraLlamada
    txtHora.Text = mAtencion.HoraLlamada
    txtNroIncidente = mAtencion.NroIncidente
    txtObservaciones = mAtencion.Observaciones
    txtOperador = mAtencion.Operador
   
    If mAtencion.Sintoma.Parent.id <> 4 Then 'si no es un traslado
        'si hasta ahora no cargue un telefono muestro el auxiliar
         txtTelefono.Text = mAtencion.TelefonoAuxilar
    End If
   
    Set cmbSintoma.SelectedItem = mAtencion.Sintoma
    Set cmbCodigo.SelectedItem = mAtencion.Sintoma.Parent
    txtVL = mAtencion.VL
    txtQTH = mAtencion.QTH
    
    Dim detalles As New blcemi.DetalleSegManager
    detalles.CargarXAtencion (mAtencion.id)
    Set lvwCambios.Coleccion = detalles
    
    txtDespachador = mAtencion.Despachador.NombreCompleto
    
End Sub


Private Sub Form_Load()
   
    Set cmdLlamar.Picture = MDI.il16.ListImages("llamar").Picture
      
    Set cmbCodigo.Coleccion = GBL.CodigoEmergenciaGBL
    Set ctlDirEmergencia.MiDireccion = New blcemi.Direccion
    Set dirOrigen.MiDireccion = New blcemi.Direccion
    Set dirDestino.MiDireccion = New blcemi.Direccion
    InicializarDireccion dirOrigen
    InicializarDireccion dirDestino
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
    
    'refresco los datos del que correponda
    If Not mAfiliadoPropio Is Nothing Then
        fraAfiliadoPropio.Visible = True
        LlenarCamposAfiliadoPropio
    End If
    
    If Not mAreaProtegida Is Nothing Then
        fraAfiliadoExterno.Visible = True
        fraAreaProtegida.Visible = True
        LlenarCamposAreaProtegida
    End If
    
    If Not mObraSocial Is Nothing Then
        fraAfiliadoExterno.Visible = True
        fraObraSocial.Visible = True
        LlenarCamposObraSocial
    End If
    
    If Not mServicioEmergencia Is Nothing Then
        fraAfiliadoExterno.Visible = True
        fraServicioEmergencia.Visible = True
        LlenarCamposServicioEmergencia
    End If
End Sub


Private Sub LlenarCamposAfiliadoPropio()
    
    lblApeNomAfilPropio = mAfiliadoPropio.NombreCompleto
    lblNroDocAPropio = mAfiliadoPropio.NroDoc
    lblTipoDoc = mAfiliadoPropio.TipoDoc.Nombre
    lblDireccionaPropio = mAfiliadoPropio.Direccion.Calle + " " + mAfiliadoPropio.Direccion.Nro
    lblCiudadAPropio = mAfiliadoPropio.Direccion.GetBarrioCiudadProvincia
    lblNroAfiliado = mAfiliadoPropio.IdCompleto
    lblEdadAPropio = mAfiliadoPropio.Edad
    lblOcupacion = mAfiliadoPropio.Ocupacion.Nombre
    lblSexoAPropio = IIf(mAfiliadoPropio.Sexo = 1, "Masculino", "Femenino")
    lblEstadoCivil = mAfiliadoPropio.EstadoCivil.Nombre
    
    lblObraSocialAPropio = mAfiliadoPropio.ObraSocial.Nombre
    
    If Not ctlDirEmergencia.DireccionCompleta("") Then 'si la direccion esta vacia...
        Set ctlDirEmergencia.MiDireccion = mAfiliadoPropio.Direccion
    End If
    
   ' MostrarTelefonos mAfiliadoPropio.Telefonos
    
End Sub

Private Sub LlenarCamposObraSocial()
   lblNombreOS.Visible = True
   lblCoseguro.Visible = True
   lblCodCubiertosOS.Visible = True
   lblDireccionOS.Visible = True
   lblCiudadOS.Visible = True
   lblServicioEmergenciaOS.Visible = True
   
   lblDatosOS(0).Visible = True
   lblDatosOS(1).Visible = True
   lblDatosOS(2).Visible = True
   lblDatosOS(3).Visible = True
   lblDatosOS(4).Visible = True
   
   lblSinDatosOS.Visible = False
   
   lblNombreOS = mObraSocial.Nombre
   lblCoseguro = Str(mObraSocial.Coseguro)
   lblCodCubiertosOS = mObraSocial.CodigosCubiertos.GetCadenaCodigos
   lblDireccionOS = mObraSocial.Direccion.Calle + " " + mObraSocial.Direccion.Nro
   lblCiudadOS = mObraSocial.Direccion.GetBarrioCiudadProvincia
   lblServicioEmergenciaOS = mObraSocial.ServicioEmergencia.Nombre
   
  ' MostrarTelefonos mObraSocial.Telefonos
   
   
End Sub

Private Sub LlenarCamposServicioEmergencia()
    lblNombreSE.Visible = True
    lblDireccionSE.Visible = True
    lblCiudadSE.Visible = True
    
    lblDatosSE(0).Visible = True
    lblDatosSE(1).Visible = True
    
    lblSinDatosSE.Visible = False
    
    lblNombreSE = mServicioEmergencia.Nombre
    lblDireccionSE = mServicioEmergencia.Direccion.Calle + " " + mServicioEmergencia.Direccion.Nro
    lblCiudadSE = mServicioEmergencia.Direccion.GetBarrioCiudadProvincia
    
    'MostrarTelefonos mServicioEmergencia.Telefonos
End Sub

Private Sub LlenarCamposAreaProtegida()
    lblNombreArea.Visible = True
    lblResponsableArea.Visible = True
    lblDocRespArea.Visible = True
    lblDireccionArea.Visible = True
    lblCiudadArea.Visible = True
    
    lblDatosArea(0).Visible = True
    lblDatosArea(1).Visible = True
    lblDatosArea(2).Visible = True
    lblDatosArea(3).Visible = True
        
    lblSinDatosAP.Visible = False
    
    lblNombreArea = mAreaProtegida.NombreArea
    lblResponsableArea = mAreaProtegida.NombreCompleto
    lblDocRespArea = mAreaProtegida.NroDocResp
    lblDireccionArea = mAreaProtegida.Direccion.Calle + " " + mAreaProtegida.Direccion.Nro
    lblCiudadArea = mAreaProtegida.Direccion.GetBarrioCiudadProvincia
    
    If Not ctlDirEmergencia.DireccionCompleta("") Then
        Set ctlDirEmergencia.MiDireccion = mAreaProtegida.Direccion
    End If
    
   ' MostrarTelefonos mAreaProtegida.Telefonos
End Sub

Private Sub LlenarCamposAfiliadoExterno()
    lblSinDatosAE.Visible = False
    lblApeNomAExterno.Visible = True
    lblEdadAExterno.Visible = True
    lblSexoAExterno.Visible = True
    lblDatosPaciente(0).Visible = True
    lblDatosPaciente(1).Visible = True
    lblDatosPaciente(2).Visible = True
       
    lblApeNomAExterno = mAfiliadoExterno.NombreCompleto
    lblEdadAExterno = mAfiliadoExterno.Edad
    lblSexoAExterno = IIf(mAfiliadoExterno.Sexo = 1, "Masculino", "Femenino")
    
End Sub

'Dim mAtencion As blcemi.Atencion
'
'Public Sub MostrarDetalles(pAtencion As blcemi.Atencion)
'    Set mAtencion = pAtencion
''    mAtencion.Afiliado
''    mAtencion.AfiliadoExterno
''    mAtencion.AreaProtegida
''    mAtencion.Despachador
''    mAtencion.Diagnostico
''    mAtencion.Direccion
''    mAtencion.DireccionDestino
''    mAtencion.DireccionOrigen
''    mAtencion.Equipos
''    mAtencion.Estado
''    mAtencion.fecha
''    mAtencion.HoraLlamada
''    mAtencion.id
''    mAtencion.NroIncidente
''    mAtencion.ObraSocial
''    mAtencion.Observaciones
''    mAtencion.Operador
''    mAtencion.QTH
''    mAtencion.ServicioEmergencia
''    mAtencion.Sintoma
''    mAtencion.Telefono
''    mAtencion.TelefonoAuxilar
''    mAtencion.VL
'
'    AppendBigTitle "Ficha de Atencion"
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
Private Sub Label22_Click()

End Sub


