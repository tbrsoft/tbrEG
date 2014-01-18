VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConfiguracion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   23
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   22
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Frame fraBotones 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      Begin VB.Label lblBoton 
         AutoSize        =   -1  'True
         Caption         =   "Campos Opcionales"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   240
         TabIndex        =   55
         Top             =   3120
         Width           =   1710
      End
      Begin VB.Label lblBoton 
         AutoSize        =   -1  'True
         Caption         =   "Predeterminados"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   240
         TabIndex        =   45
         Top             =   2640
         Width           =   1515
      End
      Begin VB.Label lblBoton 
         AutoSize        =   -1  'True
         Caption         =   "Codigos Emergencia"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   240
         TabIndex        =   38
         Top             =   2160
         Width           =   1800
      End
      Begin VB.Label lblBoton 
         AutoSize        =   -1  'True
         Caption         =   "Comportamiento"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label lblBoton 
         AutoSize        =   -1  'True
         Caption         =   "Base de Datos"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label lblBoton 
         AutoSize        =   -1  'True
         Caption         =   "Red"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   345
      End
      Begin VB.Label lblBoton 
         AutoSize        =   -1  'True
         Caption         =   "Apariencia"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame fraComportamiento 
      Height          =   4575
      Left            =   2400
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chkEnviarErrores 
         Caption         =   "Enviar informes de errores sin preguntar."
         Height          =   255
         Left            =   120
         TabIndex        =   76
         Top             =   3120
         Width           =   3615
      End
      Begin VB.CheckBox chkSepararAtenciones 
         Caption         =   "Separar atenciones asignadas y sin asignar"
         Height          =   195
         Left            =   120
         TabIndex        =   44
         ToolTipText     =   "En el formulario Consulta de Atenciones se muestran todas las atenciones juntas o separadas segun si estan asignadas o no. "
         Top             =   240
         Width           =   3855
      End
      Begin VB.CheckBox chkExportToCalc 
         Caption         =   "Permitir Exportar Listados a OO Calc"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   1680
         Width           =   3495
      End
      Begin VB.CheckBox chkExportToWrite 
         Caption         =   "Permitir Exportar Listados a OO Write"
         Height          =   195
         Left            =   120
         TabIndex        =   36
         Top             =   1320
         Width           =   3255
      End
      Begin VB.CheckBox chkAtencionesPendientes 
         Caption         =   "Mostrar Aviso de Atenciones Pendientes"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   2760
         Width           =   3495
      End
      Begin VB.CheckBox chkSugerencias 
         Caption         =   "Mostrar sugerencias si faltan datos no obligatorios"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   3855
      End
      Begin VB.CheckBox chkExportToExcel 
         Caption         =   "Permitir Exportar Listados a MS Excel"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3135
      End
      Begin VB.CheckBox chkExportToWord 
         Caption         =   "Permitir Exportar Listados a MS Word"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox chkBarraMenues 
         Caption         =   "Mostrar Barra de Menues"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   2400
         Width           =   3735
      End
   End
   Begin VB.Frame fraBaseDatos 
      Height          =   4575
      Left            =   2400
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdPathBD 
         Caption         =   "..."
         Height          =   285
         Left            =   4800
         TabIndex        =   14
         Top             =   840
         Width           =   285
      End
      Begin VB.TextBox txtPathBD 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Ruta de la Base de Datos:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame fraRed 
      Height          =   4575
      Left            =   2400
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   2160
         TabIndex        =   24
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtDirIp 
         Height          =   285
         Left            =   2160
         TabIndex        =   18
         Text            =   "192.128.255.255"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton optCliente 
         Caption         =   "Modo Cliente"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton optServer 
         Caption         =   "Modo Servidor"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Nombre 
         Caption         =   "Nombre Pc:"
         Height          =   255
         Left            =   960
         TabIndex        =   25
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Dir IP Servidor:"
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame fraApariencia 
      Height          =   4575
      Left            =   2400
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdWallpaper 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   35
         Top             =   2400
         Width           =   255
      End
      Begin VB.TextBox txtWallpaper 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   2400
         Width           =   3375
      End
      Begin VB.CommandButton cmdFuenteContenidos 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   31
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton cmdFuenteTitulo 
         Caption         =   "..."
         Height          =   255
         Left            =   4200
         TabIndex        =   28
         Top             =   1320
         Width           =   255
      End
      Begin VB.CheckBox chkGridLinesOtros 
         Caption         =   "Mostrar Cuadricula en las demas listas"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   3735
      End
      Begin VB.CheckBox chkGridLinesConsulta 
         Caption         =   "Mostrar Cuadriculas en formularios de consulta"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label5 
         Caption         =   "Fondo:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label lblMuestraContenido 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Muestra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         TabIndex        =   30
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblMuestraTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Muestra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2400
         TabIndex        =   29
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Fuente contenido listados."
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Fuente titulo listados."
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1320
         Width           =   1575
      End
   End
   Begin VB.Frame fraCamposOpcionales 
      Height          =   4575
      Left            =   2400
      TabIndex        =   56
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CheckBox chkTopeAP 
         Caption         =   "Utilizar Tope de Atenciones."
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox chkDNIAP 
         Caption         =   "Exigir DNI."
         Height          =   255
         Left            =   360
         TabIndex        =   62
         Top             =   1560
         Width           =   3615
      End
      Begin VB.CheckBox chkTopeArea 
         Caption         =   "Utilizar Tope de Atenciones."
         Height          =   255
         Left            =   360
         TabIndex        =   60
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CheckBox chkDNIArea 
         Caption         =   "Exigir DNI del Responsable."
         Height          =   255
         Left            =   360
         TabIndex        =   59
         Top             =   2640
         Width           =   2775
      End
      Begin VB.CheckBox chkTopeAE 
         Caption         =   "Utilizar Tope de Atenciones."
         Height          =   255
         Left            =   360
         TabIndex        =   58
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox chkDNIAE 
         Caption         =   "Exigir DNI."
         Height          =   255
         Left            =   360
         TabIndex        =   57
         Top             =   480
         Width           =   3615
      End
      Begin VB.Label Label11 
         Caption         =   "Area Protegida"
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
         Left            =   240
         TabIndex        =   65
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Afiliado Propio"
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
         Left            =   240
         TabIndex        =   64
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Afiliado Externo"
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
         Left            =   240
         TabIndex        =   61
         Top             =   240
         Width           =   2895
      End
   End
   Begin VB.Frame fraCodigo 
      Height          =   4575
      Left            =   2400
      TabIndex        =   39
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Frame Frame2 
         Caption         =   "Codigos"
         Height          =   3135
         Left            =   120
         TabIndex        =   66
         Top             =   1320
         Width           =   4935
         Begin VB.CheckBox chkColor 
            Caption         =   "Usar codigo como color de fuente."
            Height          =   255
            Left            =   120
            TabIndex        =   73
            Top             =   840
            Width           =   2895
         End
         Begin VB.CheckBox chkNegrita 
            Caption         =   "Fuente en negrita."
            Height          =   255
            Left            =   3120
            TabIndex        =   72
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtVencimiento 
            Height          =   285
            Left            =   1200
            TabIndex        =   71
            Top             =   480
            Width           =   495
         End
         Begin ControlesPOO.ListViewConsulta lvwCodigos 
            Height          =   1815
            Left            =   120
            TabIndex        =   67
            Top             =   1200
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   3201
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   -1  'True
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            AllowModify     =   -1  'True
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Vence"
            MEncabezado0    =   "vencimiento"
            AEncabezado0    =   20
            NEncabezado1    =   "Codigo"
            MEncabezado1    =   "nombrecompuesto"
            AEncabezado1    =   40
            NEncabezado2    =   "Negrita"
            MEncabezado2    =   "pgBold"
            AEncabezado2    =   20
            NEncabezado3    =   "ColorFuente"
            MEncabezado3    =   "pgcolorfuente"
            AEncabezado3    =   20
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
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
         Begin VB.Label Label12 
            Caption         =   "Vencimiento:"
            Height          =   255
            Left            =   120
            TabIndex        =   70
            Top             =   480
            Width           =   975
         End
         Begin VB.Label lblCodigo 
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
            TabIndex        =   69
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label8 
            Caption         =   "(Minutos)"
            Height          =   255
            Left            =   1800
            TabIndex        =   68
            Top             =   480
            Width           =   735
         End
      End
      Begin ALCemi.GraphicButton cmdModificar 
         Height          =   375
         Left            =   3600
         TabIndex        =   54
         Top             =   2040
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.CheckBox chkVencimiento 
         Caption         =   "Habilitar vencimiento de los Codigos."
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   960
         Width           =   3615
      End
      Begin VB.CheckBox chkCoseguro 
         Caption         =   "Utilizar distintos Coseguros por Codigo."
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   3615
      End
      Begin VB.CheckBox chkTipos 
         Caption         =   "Utilizar Tipos de Codigos (Diurno, Nocturno,etc.)"
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   480
         Width           =   3735
      End
      Begin VB.CheckBox chkExigirCodigos 
         Caption         =   "Exigir Codigos Cubiertos."
         Height          =   255
         Left            =   240
         TabIndex        =   40
         Top             =   240
         Width           =   3735
      End
   End
   Begin VB.Frame fraDefaults 
      Height          =   4575
      Left            =   2400
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
      Begin VB.Frame Frame1 
         Caption         =   "Direcciones"
         Height          =   2055
         Left            =   120
         TabIndex        =   47
         Top             =   480
         Width           =   4935
         Begin ControlesPOO.Combo cmbPais 
            Height          =   315
            Left            =   960
            TabIndex        =   74
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbBarrio 
            Height          =   315
            Left            =   960
            TabIndex        =   48
            Top             =   1440
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            NuevoEnabled    =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbCiudad 
            Height          =   315
            Left            =   960
            TabIndex        =   49
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            NuevoEnabled    =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbProvincia 
            Height          =   315
            Left            =   960
            TabIndex        =   50
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            NuevoEnabled    =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.Label Label13 
            Caption         =   "Pais:"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblProvincia 
            Caption         =   "Provincia:"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lblCiudad 
            Caption         =   "Ciudad:"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblBarrio 
            Caption         =   "Barrio:"
            Height          =   255
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCodigos As blcemi.CodigoEmergenciaManager
Dim mCodigoAnterior As blcemi.CodigoEmergencia

Private Sub cmdAceptar_Click()
    If DatosCorrectos Then
        'apariencia
        CCFFGG.Configuracion.Apariencia.GridLinesConsultas = IIf(chkGridLinesConsulta.Value = 1, True, False)
        CCFFGG.Configuracion.Apariencia.GridLinesOtros = IIf(chkGridLinesOtros.Value = 1, True, False)
        Set CCFFGG.Configuracion.Apariencia.ContentsFont = lblMuestraContenido.Font
        Set CCFFGG.Configuracion.Apariencia.TitleFont = lblMuestraTitulo.Font
        CCFFGG.Configuracion.Apariencia.PathFondo = txtWallpaper
        
        'ccffgg.configuracion.Apariencia.ToolBarStyleFlat
        'comportamiento
        CCFFGG.Configuracion.Comportamiento.SepararAtenciones = IIf(chkSepararAtenciones.Value = 0, False, True)
        CCFFGG.Configuracion.Comportamiento.AllowExportToExcel = IIf(chkExportToExcel.Value = 0 Or chkExportToExcel.Value = 2, False, True)
        CCFFGG.Configuracion.Comportamiento.AllowExportToWord = IIf(chkExportToWord.Value = 0 Or chkExportToExcel.Value = 2, False, True)
        CCFFGG.Configuracion.Comportamiento.AllowExportToWrite = IIf(chkExportToWrite.Value = 0 Or chkExportToWrite.Value = 2, False, True)
        CCFFGG.Configuracion.Comportamiento.AllowExportToCalc = IIf(chkExportToCalc.Value = 0 Or chkExportToCalc.Value = 2, False, True)
        CCFFGG.Configuracion.Comportamiento.MostrarBarraMenu = IIf(chkBarraMenues.Value = 1, True, False)
        CCFFGG.Configuracion.Comportamiento.MostrarSugerenciasDatosFaltantes = IIf(chkSugerencias.Value = 1, True, False)
        CCFFGG.Configuracion.Comportamiento.MostrarAvisoAtencionesPendientes = IIf(chkAtencionesPendientes.Value = 1, True, False)
        CCFFGG.Configuracion.Comportamiento.EnviarErrores = chkEnviarErrores.Value
        'Bd
        CCFFGG.Configuracion.DBLayer.PathDB = txtPathBD
        'red
        CCFFGG.Configuracion.Red.ModoServer = optServer.Value
        CCFFGG.Configuracion.Red.Nombre = txtNombre
        CCFFGG.Configuracion.Red.DirIPRemota = txtDirIp
        
        'codigos
        CCFFGG.Configuracion.Codigo.CosegurosPorCodigo = IIf(chkCoseguro.Value = 1, True, False)
        CCFFGG.Configuracion.Codigo.ExigirCodigos = IIf(chkExigirCodigos.Value = 1, True, False)
        CCFFGG.Configuracion.Codigo.HabilitarVencimiento = IIf(chkVencimiento.Value = 1, True, False)
        CCFFGG.Configuracion.Codigo.UtilizarTipos = IIf(chkTipos.Value = 1, True, False)
        Dim c As blcemi.CodigoEmergencia
        For Each c In lvwCodigos.Coleccion
            GBL.CodigoEmergenciaGBL.Item(c.id).Vencimiento = c.Vencimiento
            GBL.CodigoEmergenciaGBL.Item(c.id).Bold = c.Bold
            GBL.CodigoEmergenciaGBL.Item(c.id).ColorFuente = c.ColorFuente
            GBL.CodigoEmergenciaGBL.Item(c.id).SaveChanges
        Next
        'defaults
        CCFFGG.Configuracion.Defaults.Pais = cmbPais.SelectedItem.Nombre
        CCFFGG.Configuracion.Defaults.Barrio = cmbBarrio.SelectedItem.Nombre
        CCFFGG.Configuracion.Defaults.Ciudad = cmbCiudad.SelectedItem.Nombre
        CCFFGG.Configuracion.Defaults.Provincia = cmbProvincia.SelectedItem.Nombre
        
        'requeridos
        CCFFGG.Configuracion.Requeridos.ExigirDNIAE = IIf(chkDNIAE.Value = 0, False, True)
        CCFFGG.Configuracion.Requeridos.ExigirDNIAP = IIf(chkDNIAP.Value = 0, False, True)
        CCFFGG.Configuracion.Requeridos.ExigirDNIRespArea = IIf(chkDNIArea.Value = 0, False, True)
        CCFFGG.Configuracion.Requeridos.UsarTopeAtencAE = IIf(chkTopeAE.Value = 0, False, True)
        CCFFGG.Configuracion.Requeridos.UsarTopeAtencAP = IIf(chkTopeAP.Value = 0, False, True)
        CCFFGG.Configuracion.Requeridos.UsarTopeAtencArea = IIf(chkTopeArea.Value = 0, False, True)
        
        'escondo el formulario porq cuando llama a actualizar toma este como activeform
        Me.Hide
        'aviso q cambie la configuracion
        CCFFGG.Configuracion.ConfiguracionModificada
        'ver si hay q guardar la config
        Unload Me
    End If
End Sub

Private Function DatosCorrectos() As Boolean
    Dim msg As String
    If optCliente.Value Then
        If Not TextBoxValidado(txtDirIp, eDireccionIP) Then msg = "La direccion IP ingresada no es correcta."
        If Not TextBoxValidado(txtNombre, eString) Then msg = msg + "Ingrese el nombre de la PC."
    End If
    
    'defaults
    If cmbProvincia.SelectedItem Is Nothing Then
        msg = "Seleccione una Provincia."
    End If
    
    If cmbCiudad.SelectedItem Is Nothing Then
        msg = "Seleccione una Ciudad."
    End If
    
    If cmbBarrio.SelectedItem Is Nothing Then
        msg = "Seleccione un Barrio."
    End If
    
    If msg = "" Then
        DatosCorrectos = True
    Else
        MsgBox "Falta la siguiente informacion: " + vbLf + msg, vbExclamation
        DatosCorrectos = False
    End If
End Function

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdFuenteContenidos_Click()
Dim cd As New CommonDialog
With lblMuestraContenido.Font
    cd.FontName = .Name
    cd.Bold = .Bold
    cd.Italic = .Italic
    cd.FontSize = .size
    cd.StrikeThru = .Strikethrough
    cd.Underline = .Underline
End With
cd.flags = cdlCFBoth
cd.ShowFont
'si cancela quedan los mismos valores...
With lblMuestraContenido.Font
    .Name = cd.FontName
    .Bold = cd.Bold
    .Italic = cd.Italic
    .size = cd.FontSize
    .Strikethrough = cd.StrikeThru
    .Underline = cd.Underline
End With
End Sub

Private Sub cmdFuenteTitulo_Click()
Dim cd As New CommonDialog
With lblMuestraTitulo.Font
    cd.FontName = .Name
    cd.Bold = .Bold
    cd.Italic = .Italic
    cd.FontSize = .size
    cd.StrikeThru = .Strikethrough
    cd.Underline = .Underline
End With
cd.flags = cdlCFBoth
cd.ShowFont
'si cancela quedan los mismos valores...
With lblMuestraTitulo.Font
    .Name = cd.FontName
    .Bold = cd.Bold
    .Italic = cd.Italic
    .size = cd.FontSize
    .Strikethrough = cd.StrikeThru
    .Underline = cd.Underline
End With

End Sub

Private Sub cmdModificar_Click()
lvwCodigos.SetFocus
lvwCodigos.Editar
End Sub

Private Sub cmdPathBD_Click()
Dim cd As New CommonDialog
cd.Filter = "Archivo de Base de Datos (*.mdb)|*.mdb"
cd.ShowOpen
If cd.FileName <> "" Then 'si no eligio cancelar...
    If GBL.VerificarRutaBaseDatos(cd.FileName) Then
        txtPathBD = cd.FileName
    Else
        MsgBox "El archivo seleccionado no es valido!", vbExclamation
    End If
End If
End Sub

Private Sub cmdWallpaper_Click()
Dim cd As New CommonDialog
cd.Filter = "Archivos JPEG (*.jpg)|*.jpg|Archivos bitmap (*.bmp)|*.bmp"
cd.ShowOpen
If cd.FileName <> "" Then 'si no eligio cancelar...
    txtWallpaper = cd.FileName
End If

End Sub

Private Sub Form_Load()
    TERR.Anotar "arha"
    
    On Local Error GoTo errCGGD
    
    Set Me.Icon = MDI.Icon

    lblBoton(0).FontBold = True
    fraApariencia.Visible = True
    'Apariencia
    TERR.Anotar "arhb"
    chkGridLinesConsulta.Value = IIf(CCFFGG.Configuracion.Apariencia.GridLinesConsultas, 1, 0)
    chkGridLinesOtros.Value = IIf(CCFFGG.Configuracion.Apariencia.GridLinesOtros, 1, 0)
    
    TERR.Anotar "arhc"
    'ccffgg.configuracion.Apariencia.ToolBarStyleFlat
    'si hago el set irectamente me toma la ref y si cambio la fuente y pongo cancelar me la cambia igual
    With CCFFGG.Configuracion.Apariencia.ContentsFont
        TERR.Anotar "arhd"
        lblMuestraContenido.FontName = .Name
        lblMuestraContenido.FontSize = .size
        lblMuestraContenido.FontBold = .Bold
        lblMuestraContenido.FontItalic = .Italic
        lblMuestraContenido.FontUnderline = .Underline
        lblMuestraContenido.FontStrikethru = .Strikethrough
    End With
    
    With CCFFGG.Configuracion.Apariencia.TitleFont
        TERR.Anotar "arhe"
        lblMuestraTitulo.FontName = .Name
        lblMuestraTitulo.FontSize = .size
        lblMuestraTitulo.FontBold = .Bold
        lblMuestraTitulo.FontItalic = .Italic
        lblMuestraTitulo.FontUnderline = .Underline
        lblMuestraTitulo.FontStrikethru = .Strikethrough
    End With
    
    TERR.Anotar "arhf"
    txtWallpaper = CCFFGG.Configuracion.Apariencia.PathFondo
    
    TERR.Anotar "arhg"
    'comportamiento
    If ApplicationInstalled(eExcel) Then
        TERR.Anotar "arhh"
        chkExportToExcel.Value = IIf(CCFFGG.Configuracion.Comportamiento.AllowExportToExcel, 1, 0)
    Else
        TERR.Anotar "arhi"
        chkExportToExcel.Value = 2
        chkExportToExcel.Enabled = False
    End If
    
    If ApplicationInstalled(eWord) Then
        TERR.Anotar "arhj"
        chkExportToWord.Value = IIf(CCFFGG.Configuracion.Comportamiento.AllowExportToWord, 1, 0)
    Else
        TERR.Anotar "arhk"
        chkExportToWord.Value = 2
        chkExportToWord.Enabled = False
    End If
    'verifico si esta instalado el OpenOffice
    If ApplicationInstalled(eOpenOffice) Then
        TERR.Anotar "arhl"
        chkExportToWrite.Value = IIf(CCFFGG.Configuracion.Comportamiento.AllowExportToWrite, 1, 0)
        chkExportToCalc.Value = IIf(CCFFGG.Configuracion.Comportamiento.AllowExportToCalc, 1, 0)
    Else
        TERR.Anotar "arhm"
        chkExportToWrite.Value = 2
        chkExportToCalc.Value = 2
        chkExportToCalc.Enabled = False
        chkExportToWrite.Enabled = False
    End If
  
    TERR.Anotar "arhn"
    chkSepararAtenciones.Value = IIf(CCFFGG.Configuracion.Comportamiento.SepararAtenciones, 1, 0)
    chkBarraMenues.Value = IIf(CCFFGG.Configuracion.Comportamiento.MostrarBarraMenu, 1, 0)
    chkSugerencias.Value = IIf(CCFFGG.Configuracion.Comportamiento.MostrarSugerenciasDatosFaltantes, 1, 0)
    TERR.Anotar "arho"
    chkAtencionesPendientes.Value = IIf(CCFFGG.Configuracion.Comportamiento.MostrarAvisoAtencionesPendientes, 1, 0)
    chkEnviarErrores.Value = CCFFGG.Configuracion.Comportamiento.EnviarErrores
    
    TERR.Anotar "arhp"
    'Bd
    txtPathBD = CCFFGG.Configuracion.DBLayer.PathDB
    TERR.Anotar "arhq", txtPathBD
    'red
    optServer = CCFFGG.Configuracion.Red.ModoServer
    optCliente = Not CCFFGG.Configuracion.Red.ModoServer
    txtNombre = CCFFGG.Configuracion.Red.Nombre
    txtDirIp = CCFFGG.Configuracion.Red.DirIPRemota
    
    TERR.Anotar "arhr", txtNombre
    'codigos
    chkCoseguro.Value = IIf(CCFFGG.Configuracion.Codigo.CosegurosPorCodigo, 1, 0)
    chkExigirCodigos.Value = IIf(CCFFGG.Configuracion.Codigo.ExigirCodigos, 1, 0)
    chkVencimiento.Value = IIf(CCFFGG.Configuracion.Codigo.HabilitarVencimiento, 1, 0)
    chkTipos.Value = IIf(CCFFGG.Configuracion.Codigo.UtilizarTipos, 1, 0)
    
    TERR.Anotar "arhs"
    Dim c As blcemi.CodigoEmergencia
    Set mCodigos = New blcemi.CodigoEmergenciaManager
    For Each c In GBL.CodigoEmergenciaGBL
        TERR.Anotar "arht", c.Nombre
        mCodigos.AddItem c.Clone
    Next
    
    TERR.Anotar "arhu"
    Set lvwCodigos.Coleccion = mCodigos
    TERR.Anotar "arhu21"
    Set cmdModificar.Picture = MDI.il32.ListImages("modificar").Picture
    'defaults
    TERR.Anotar "arhu22"
    Set cmbPais.Coleccion = GBL.PaisesGBL
    
    Dim mPais As blcemi.Pais
    TERR.Anotar "arhu24"
    Set mPais = GBL.PaisesGBL.ItemByName(CCFFGG.Configuracion.Defaults.Pais)
    TERR.Anotar "arhu23"
    Set cmbPais.SelectedItem = mPais
    TERR.Anotar "arhu25"
    Set cmbProvincia.Coleccion = mPais.Provincias
    TERR.Anotar "arhu26"
    Dim mProv As blcemi.Provincia
    TERR.Anotar "arhu27"
    Set mProv = mPais.Provincias.ItemByName(CCFFGG.Configuracion.Defaults.Provincia)
    TERR.Anotar "arhu28"
    Set cmbProvincia.SelectedItem = mProv
    TERR.Anotar "arhu29"
    Set cmbCiudad.Coleccion = mProv.Ciudades
    TERR.Anotar "arhu30"
    Dim mCiudad As blcemi.Ciudad
    Set mCiudad = mProv.Ciudades.ItemByName(CCFFGG.Configuracion.Defaults.Ciudad)
    TERR.Anotar "arhu31"
    Set cmbCiudad.SelectedItem = mCiudad
    TERR.Anotar "arhu32"
    Set cmbBarrio.Coleccion = mCiudad.Barrios
    TERR.Anotar "arhu33"
    'si no hay ciudad no habra barrios !
    Dim mBarrio As blcemi.Barrio
    TERR.Anotar "arhu34"
    Set mBarrio = mCiudad.Barrios.ItemByName(CCFFGG.Configuracion.Defaults.Barrio)
    TERR.Anotar "arhu34"
    Set cmbBarrio.SelectedItem = mBarrio
    
    TERR.Anotar "arhv"
    'requeridos
    chkDNIAE.Value = IIf(CCFFGG.Configuracion.Requeridos.ExigirDNIAE, 1, 0)
    chkDNIAP.Value = IIf(CCFFGG.Configuracion.Requeridos.ExigirDNIAP, 1, 0)
    chkDNIArea.Value = IIf(CCFFGG.Configuracion.Requeridos.ExigirDNIRespArea, 1, 0)
    chkTopeAE.Value = IIf(CCFFGG.Configuracion.Requeridos.UsarTopeAtencAE, 1, 0)
    chkTopeAP.Value = IIf(CCFFGG.Configuracion.Requeridos.UsarTopeAtencAP, 1, 0)
    chkTopeArea.Value = IIf(CCFFGG.Configuracion.Requeridos.UsarTopeAtencArea, 1, 0)
    
    TERR.Anotar "arhw"
    'aplicar permisos
    lblBoton(0).Enabled = UsuarioActual.Permisos.Can(blcemi.ConfigurarApariencia)
    lblBoton(1).Enabled = UsuarioActual.Permisos.Can(blcemi.ConfigurarRed)
    lblBoton(2).Enabled = UsuarioActual.Permisos.Can(blcemi.ConfigurarBaseDatos)
    lblBoton(3).Enabled = UsuarioActual.Permisos.Can(blcemi.ConfigurarComportamiento)
    lblBoton(4).Enabled = UsuarioActual.Permisos.Can(blcemi.ConfigurarCodigo)
    lblBoton(5).Enabled = UsuarioActual.Permisos.Can(blcemi.ConfigurarDefaults)
    'lblBoton(6).Enabled = UsuarioActual.Permisos.Can(ConfigurarDefaults)
    
    TERR.Anotar "arhx"
    Exit Sub
errCGGD:
    TERR.AppendLog "errCGGD", TERR.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub fraBotones_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim l As Label
    For Each l In lblBoton
        l.ForeColor = vbBlack
    Next
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "configuracion"
End Function

Private Sub lblBoton_Click(Index As Integer)
    Dim l As Label
    For Each l In lblBoton
        l.FontBold = False
    Next
    'lblBoton(Index).BackColor = RGB(255, 255, 170)
    lblBoton(Index).FontBold = True
    
    fraApariencia.Visible = False
    fraRed.Visible = False
    fraBaseDatos.Visible = False
    fraComportamiento.Visible = False
    fraCodigo.Visible = False
    fraDefaults.Visible = False
    fraCamposOpcionales.Visible = False
    
    Select Case Index
        Case 0 'apariencia
            fraApariencia.Visible = True
        Case 1 'red
            fraRed.Visible = True
        Case 2 'bd
            fraBaseDatos.Visible = True
        Case 3 'comport
            fraComportamiento.Visible = True
        Case 4 'codigo
            fraCodigo.Visible = True
        Case 5 'defaults
            fraDefaults.Visible = True
        Case 6
            fraCamposOpcionales.Visible = True
    End Select
    
End Sub

Private Sub lblBoton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim l As Label
    For Each l In lblBoton
        l.ForeColor = vbBlack
    Next
    lblBoton(Index).ForeColor = vbRed
End Sub

Private Sub lvwCodigos_ItemClick(Item As Object)
    If Not mCodigoAnterior Is Nothing Then
        mCodigoAnterior.Vencimiento = IIf(txtVencimiento = "", 0, txtVencimiento)
        mCodigoAnterior.Bold = IIf(chkNegrita.Value = Checked, True, False)
        mCodigoAnterior.ColorFuente = GetColor
    End If
    Dim c As blcemi.CodigoEmergencia
    Set c = Item
    lblCodigo = c.Nombre
    txtVencimiento = c.Vencimiento
    chkNegrita.Value = IIf(c.Bold, 1, 0)
    chkColor.Value = IIf(c.ColorFuente <> 0, 1, 0)
    Set mCodigoAnterior = c
    lvwCodigos.Refresh
    Set lvwCodigos.SelectedItem = c
End Sub

Private Function GetColor() As Long
    If chkColor.Value = 1 Then
        Select Case mCodigoAnterior.Nombre
            Case "Rojo"
                GetColor = vbRed
            Case "Amarillo"
                GetColor = RGB(200, 200, 0)
            Case "Verde"
                GetColor = vbGreen
            Case "Azul - Traslado"
                GetColor = vbBlue
            Case "Celeste"
                GetColor = vbBlue
            Case "Enfermeria"
                GetColor = vbBlack
            Case "Emergencia"
                GetColor = vbBlack
        End Select
    Else
        GetColor = 0
    End If
End Function


Private Sub lvwCodigos_ItemEdited(Item As Object, pNewValue As String, pCancel As Boolean)
    Dim c As blcemi.CodigoEmergencia
    If IsNumeric(pNewValue) Then
        Set c = Item
        c.Vencimiento = CInt(pNewValue)
    Else
        pCancel = True
    End If
End Sub

Private Sub txtWallpaper_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If MsgBox("¿Desea quitar el fondo de pantalla?", vbYesNo + vbQuestion) = vbYes Then
        txtWallpaper = "(Sin fondo de pantalla)"
    End If
End If
End Sub

'-------------------------auxiliares de defaults.direccion-------------------
Private Sub cmbBarrio_NuevoSeleccionado()

If Not cmbCiudad.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Barrio")
    If Nombre <> "" Then
        If cmbBarrio.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim b As blcemi.Barrio
            Set b = cmbBarrio.Coleccion.Nuevobarrio(Nombre, cmbCiudad.SelectedItem)
            cmbBarrio.Refresh
            Set cmbBarrio.SelectedItem = b
        End If
    End If
Else
    'debe seleccionar una ciudad primero
End If

End Sub

Private Sub cmbCiudad_ItemSeleccionado(Item As Object)
    Dim c As blcemi.Ciudad
    Set c = Item
    Set cmbBarrio.Coleccion = c.Barrios
End Sub

Private Sub cmbCiudad_NuevoSeleccionado()
If Not cmbProvincia.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Ciudad")
    If Nombre <> "" Then
        If cmbCiudad.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim c As blcemi.Ciudad
            Set c = cmbCiudad.Coleccion.Nuevaciudad(Nombre, cmbProvincia.SelectedItem)
            cmbCiudad.Refresh
            Set cmbCiudad.SelectedItem = c
        End If
    End If
Else
    'debe seleccionar una provincia antes
End If

End Sub

Private Sub cmbProvincia_ItemSeleccionado(Item As Object)
    Dim p As blcemi.Provincia
    Set p = Item
    Set cmbCiudad.Coleccion = p.Ciudades
    Set cmbBarrio.Coleccion = Nothing
End Sub

Private Sub cmbProvincia_NuevoSeleccionado()
If Not cmbPais.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Provincia")
    If Nombre <> "" Then
        If cmbProvincia.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim c As blcemi.Provincia
            Set c = cmbProvincia.Coleccion.NuevaProvincia(Nombre, cmbPais.SelectedItem)
            cmbProvincia.Refresh
            Set cmbProvincia.SelectedItem = c
        End If
    End If
Else
    'debe seleccionar un pais antes
End If

End Sub

Private Sub cmbPais_ItemSeleccionado(Item As Object)
    Dim p As blcemi.Pais
    Set p = Item
    If p.PrimerOrden <> "" Then lblProvincia.Caption = p.PrimerOrden + ":"
    If p.SegundoOrden <> "" Then lblCiudad.Caption = p.SegundoOrden + ":"
    If p.TercerOrden <> "" Then lblBarrio.Caption = p.TercerOrden + ":"

    Set cmbProvincia.Coleccion = p.Provincias
    Set cmbCiudad.Coleccion = Nothing
    Set cmbBarrio.Coleccion = Nothing
End Sub
