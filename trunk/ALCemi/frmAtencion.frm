VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmAtencion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Operador"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9915
   ScaleWidth      =   10635
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar"
      Height          =   495
      Left            =   5280
      TabIndex        =   131
      Top             =   8640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame fraComun 
      Height          =   3855
      Left            =   5160
      TabIndex        =   86
      Top             =   4680
      Visible         =   0   'False
      Width           =   5295
      Begin TabDlg.SSTab sTabDetalles 
         Height          =   1815
         Left            =   120
         TabIndex        =   117
         Top             =   1920
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3201
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         TabCaption(0)   =   "Observaciones"
         TabPicture(0)   =   "frmAtencion.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtObservaciones"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Diagnostico Final"
         TabPicture(1)   =   "frmAtencion.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "txtDiagnostico"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Informacion Contable"
         TabPicture(2)   =   "frmAtencion.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmbIva"
         Tab(2).Control(1)=   "txtServicio"
         Tab(2).Control(2)=   "txtSaldo"
         Tab(2).Control(3)=   "txtAbonado"
         Tab(2).Control(4)=   "txtCopago"
         Tab(2).Control(5)=   "Label19"
         Tab(2).Control(6)=   "Label18"
         Tab(2).Control(7)=   "Label16"
         Tab(2).Control(8)=   "lblSaldo"
         Tab(2).Control(9)=   "Label11"
         Tab(2).Control(10)=   "Label6"
         Tab(2).ControlCount=   11
         Begin VB.ComboBox cmbIva 
            Height          =   315
            ItemData        =   "frmAtencion.frx":0054
            Left            =   -74880
            List            =   "frmAtencion.frx":0061
            Style           =   2  'Dropdown List
            TabIndex        =   133
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox txtServicio 
            Height          =   285
            Left            =   -72840
            TabIndex        =   129
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSaldo 
            Height          =   285
            Left            =   -70680
            Locked          =   -1  'True
            TabIndex        =   125
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtAbonado 
            Height          =   285
            Left            =   -71760
            TabIndex        =   123
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtCopago 
            Height          =   285
            Left            =   -71760
            TabIndex        =   121
            Top             =   1200
            Width           =   615
         End
         Begin VB.TextBox txtDiagnostico 
            Height          =   1215
            Left            =   -74880
            MultiLine       =   -1  'True
            TabIndex        =   119
            Top             =   480
            Width           =   4815
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   1215
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   118
            Top             =   480
            Width           =   4815
         End
         Begin VB.Label Label19 
            Caption         =   "IVA:"
            Height          =   255
            Left            =   -74880
            TabIndex        =   134
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label18 
            Caption         =   "Servicio"
            Height          =   255
            Left            =   -72840
            TabIndex        =   130
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label16 
            Caption         =   "-                ="
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72000
            TabIndex        =   126
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblSaldo 
            Alignment       =   2  'Center
            Caption         =   "Saldo"
            Height          =   255
            Left            =   -70680
            TabIndex        =   124
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "Paciente abonó"
            Height          =   255
            Left            =   -72000
            TabIndex        =   122
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Copago:"
            Height          =   255
            Left            =   -72840
            TabIndex        =   120
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtFecha 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdDotacion 
         Caption         =   "Dotacion"
         Height          =   375
         Left            =   120
         TabIndex        =   89
         Top             =   240
         Width           =   2535
      End
      Begin ControlesPOO.TreeViewConsulta tvw 
         Height          =   1215
         Left            =   120
         TabIndex        =   97
         Top             =   600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2143
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
      Begin VB.Label Label12 
         Caption         =   "Hora Llamado:"
         Height          =   195
         Left            =   2880
         TabIndex        =   88
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label15 
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   3360
         TabIndex        =   87
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   6960
      TabIndex        =   28
      Top             =   8640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame fraCodigoEmerg 
      Caption         =   "Seleccione el codigo y el Sintoma"
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   10335
      Begin ALCemi.GraphicButton cmdBuscarSintoma 
         Height          =   375
         Left            =   9840
         TabIndex        =   132
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin ControlesPOO.Combo cmbTipo 
         Height          =   315
         Left            =   3600
         TabIndex        =   116
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   720
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdVerFichaPreArribo 
         Caption         =   "Ver ficha Pre-Arribo..."
         Height          =   375
         Left            =   8520
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin ControlesPOO.Combo cmbSintoma 
         Height          =   315
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         AtributoAMostrar=   "nombrecompuesto"
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtSintoma 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5880
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin ControlesPOO.Combo cmbCodigo 
         Height          =   315
         Left            =   1200
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         AtributoAMostrar=   "nombrecompuesto"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   3120
         TabIndex        =   115
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblSintoma 
         Caption         =   "Sintoma:"
         Height          =   195
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   8760
      TabIndex        =   107
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Frame fraAfiliadoExterno 
      Height          =   3735
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   10335
      Begin VB.TextBox txtNroInterno 
         Height          =   285
         Left            =   3840
         TabIndex        =   128
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtNroIncidente 
         Height          =   285
         Left            =   2040
         TabIndex        =   61
         Top             =   240
         Width           =   975
      End
      Begin VB.Frame fraPacienteArea 
         Caption         =   "Datos Paciente"
         Height          =   1095
         Left            =   120
         TabIndex        =   23
         Top             =   2530
         Width           =   4815
         Begin ALCemi.GraphicButton cmdConsultarAExterno 
            Height          =   375
            Left            =   4440
            TabIndex        =   106
            ToolTipText     =   "Consultar Afiliados Externos"
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
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
            Left            =   1230
            TabIndex        =   92
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
            TabIndex        =   33
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lblDatosPaciente 
            Caption         =   "Sexo:"
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   720
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label lblDatosPaciente 
            Caption         =   "Edad:"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   360
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label lblDatosPaciente 
            Caption         =   "Apellido y Nombre:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   360
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin TabDlg.SSTab sTabArea 
         Height          =   3400
         Left            =   5040
         TabIndex        =   21
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   6006
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Historia Clinica"
         TabPicture(0)   =   "frmAtencion.frx":0086
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "rTxtHCAExterno"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Atenciones"
         TabPicture(1)   =   "frmAtencion.frx":00A2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lvwAtencionesAE"
         Tab(1).ControlCount=   1
         Begin ControlesPOO.ListViewConsulta lvwAtencionesAE 
            Height          =   2895
            Left            =   -74880
            TabIndex        =   90
            ToolTipText     =   "Doble click en una atencion para ver los detalles"
            Top             =   360
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   5106
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   -1  'True
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            AllowModify     =   0   'False
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Fecha"
            MEncabezado0    =   "fecha"
            AEncabezado0    =   15
            NEncabezado1    =   "Afiliado"
            MEncabezado1    =   "afiliado"
            AEncabezado1    =   30
            NEncabezado2    =   "Sintoma"
            MEncabezado2    =   "sintoma"
            AEncabezado2    =   25
            NEncabezado3    =   "Diagnostico"
            MEncabezado3    =   "Diagnostico"
            AEncabezado3    =   30
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
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
         Begin RichTextLib.RichTextBox rTxtHCAExterno 
            Height          =   3015
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   4935
            _ExtentX        =   8705
            _ExtentY        =   5318
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmAtencion.frx":00BE
         End
      End
      Begin VB.Frame fraServicioEmergencia 
         Caption         =   "Servicio Emergencia"
         Height          =   1935
         Left            =   120
         TabIndex        =   60
         Top             =   600
         Visible         =   0   'False
         Width           =   4815
         Begin ALCemi.GraphicButton cmdConsultarSE 
            Height          =   375
            Left            =   4440
            TabIndex        =   109
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
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
            TabIndex        =   94
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
            TabIndex        =   67
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
            TabIndex        =   66
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
            TabIndex        =   65
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
            TabIndex        =   64
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
            TabIndex        =   63
            Top             =   360
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.Frame fraObraSocial 
         Caption         =   "Obra Social"
         Height          =   1935
         Left            =   120
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   4815
         Begin ALCemi.GraphicButton cmdConsultarOS 
            Height          =   375
            Left            =   4440
            TabIndex        =   105
            ToolTipText     =   "Permite seleccionar otra obra social."
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
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
            TabIndex        =   95
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
            TabIndex        =   59
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
            TabIndex        =   58
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   240
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label lblDatosOS 
            Caption         =   "Serv. Emergencia:"
            Height          =   255
            Index           =   3
            Left            =   105
            TabIndex        =   55
            Top             =   960
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblDatosOS 
            Caption         =   "Cod. Cubiertos:"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   54
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lblDatosOS 
            Caption         =   "Coseguro:"
            Height          =   255
            Index           =   2
            Left            =   705
            TabIndex        =   53
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
            TabIndex        =   52
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
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   240
            Visible         =   0   'False
            Width           =   600
         End
      End
      Begin VB.Frame fraAreaProtegida 
         Caption         =   "Area Protegida"
         Height          =   1935
         Left            =   120
         TabIndex        =   34
         Top             =   600
         Visible         =   0   'False
         Width           =   4815
         Begin ALCemi.GraphicButton cmdConsultarAP 
            Height          =   375
            Left            =   4440
            TabIndex        =   108
            ToolTipText     =   "Consultar Areas Protegidas"
            Top             =   120
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
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
            TabIndex        =   93
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
            TabIndex        =   43
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
         End
      End
      Begin VB.Label Label17 
         Caption         =   "Interno:"
         Height          =   255
         Left            =   3120
         TabIndex        =   127
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label32 
         Caption         =   "Nro. de Incidente Externo:"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame fraAtencion 
      Caption         =   "Atencion"
      Height          =   4575
      Left            =   120
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   4935
      Begin VB.TextBox txtTelefono 
         Height          =   285
         Left            =   1200
         TabIndex        =   102
         Top             =   720
         Width           =   2535
      End
      Begin VB.CommandButton cmdAsignar 
         Height          =   285
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   101
         ToolTipText     =   "Asignar el telefono a a persona que llamo"
         Top             =   720
         Width           =   375
      End
      Begin ControlesPOO.Combo cmbTelefono 
         Height          =   315
         Left            =   1200
         TabIndex        =   100
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         NuevoEnabled    =   -1  'True
         AtributoAMostrar=   "numero"
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdLlamar 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Llamar al telefono"
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdRegistrarLiberacion 
         Caption         =   "Registrar Liberacion del Movil (VL)"
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   19
         Top             =   1560
         Width           =   3495
      End
      Begin VB.CommandButton cmdRegistrarArribo 
         Caption         =   "Registrar Arribo del Movil al Lugar (QTH)"
         Height          =   285
         Left            =   1200
         TabIndex        =   18
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtVL 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox txtQTH 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
      End
      Begin ALCemi.ctlDireccion ctlDirEmergencia 
         Height          =   2565
         Left            =   120
         TabIndex        =   15
         Top             =   1920
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   3995
         ProvinciaVisible=   0   'False
         Caption         =   "Direccion de la Emergencia"
         CanDragDrop     =   -1  'True
         SoloConsulta    =   0   'False
         EntrecallesVisible=   -1  'True
      End
      Begin VB.TextBox txtOperador 
         Height          =   285
         Left            =   1200
         TabIndex        =   14
         ToolTipText     =   "Ingrese el nombre de la persona que llama"
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Telefono:"
         Height          =   255
         Left            =   240
         TabIndex        =   98
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "VL:"
         Height          =   195
         Left            =   3000
         TabIndex        =   17
         Top             =   1680
         Width           =   240
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "QTH:"
         Height          =   195
         Left            =   2760
         TabIndex        =   16
         Top             =   1200
         Width           =   390
      End
      Begin VB.Label Label9 
         Caption         =   "Operador:"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame fraSeleccion 
      Caption         =   "Seleccione el tipo de Afiliado"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      Begin VB.CommandButton cmdAfiliadoPropio 
         Caption         =   "Afiliado Propio"
         Height          =   615
         Left            =   360
         TabIndex        =   103
         Top             =   360
         Width           =   2055
      End
      Begin VB.PictureBox picLogo 
         BorderStyle     =   0  'None
         Height          =   3180
         Left            =   2640
         Picture         =   "frmAtencion.frx":0140
         ScaleHeight     =   212
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   505
         TabIndex        =   96
         Top             =   360
         Width           =   7575
      End
      Begin VB.CommandButton cmdServicioEmergencia 
         Caption         =   "Servicio de Emergencia"
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton cmdObraSocial 
         Caption         =   "Obra Social"
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   2040
         Width           =   2055
      End
      Begin VB.CommandButton cmdAreaProtegida 
         Caption         =   "Area Protegida"
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   2055
      End
   End
   Begin VB.Frame fraAfiliadoPropio 
      Caption         =   "Afiliado Propio"
      Height          =   3615
      Left            =   120
      TabIndex        =   44
      Top             =   120
      Visible         =   0   'False
      Width           =   10335
      Begin ALCemi.GraphicButton cmdConsultarAPropio 
         Height          =   375
         Left            =   4320
         TabIndex        =   104
         ToolTipText     =   "Permite seleccionar otro afiliado."
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin TabDlg.SSTab sTab 
         Height          =   3255
         Left            =   4800
         TabIndex        =   45
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5741
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Historia Clinica"
         TabPicture(0)   =   "frmAtencion.frx":93A5
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "rTxtHCAPropio"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Atenciones"
         TabPicture(1)   =   "frmAtencion.frx":93C1
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lvwAtencionesAP"
         Tab(1).ControlCount=   1
         Begin ControlesPOO.ListViewConsulta lvwAtencionesAP 
            Height          =   2775
            Left            =   -74880
            TabIndex        =   91
            ToolTipText     =   "Doble click en una atencion para ver los detalles"
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4895
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   -1  'True
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            AllowModify     =   0   'False
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Fecha"
            MEncabezado0    =   "fecha"
            AEncabezado0    =   20
            NEncabezado1    =   "Sintoma"
            MEncabezado1    =   "sintoma"
            AEncabezado1    =   30
            NEncabezado2    =   "Diagnostico Final"
            MEncabezado2    =   "diagnostico"
            AEncabezado2    =   50
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
            NEncabezado0    =   ""
            MEncabezado0    =   ""
            AEncabezado0    =   0
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
         Begin RichTextLib.RichTextBox rTxtHCAPropio 
            Height          =   2775
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   4895
            _Version        =   393217
            ScrollBars      =   2
            TextRTF         =   $"frmAtencion.frx":93DD
         End
      End
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
         TabIndex        =   85
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
         TabIndex        =   84
         Top             =   2880
         Width           =   3000
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   855
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
         Top             =   1935
         Width           =   3000
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Sexo:"
         Height          =   195
         Left            =   1170
         TabIndex        =   80
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
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   1305
         Width           =   3000
      End
      Begin VB.Label lblTipoDoc 
         Alignment       =   1  'Right Justify
         Caption         =   "DNI:"
         Height          =   195
         Left            =   1245
         TabIndex        =   77
         Top             =   990
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Edad:"
         Height          =   195
         Left            =   1155
         TabIndex        =   76
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
         TabIndex        =   75
         Top             =   675
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido y Nombre:"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   74
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
         TabIndex        =   73
         Top             =   990
         Width           =   3000
      End
      Begin VB.Label lblDatosAPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº de afiliado:"
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   72
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ocupacion:"
         Height          =   195
         Left            =   750
         TabIndex        =   71
         Top             =   2565
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado Civil:"
         Height          =   195
         Left            =   705
         TabIndex        =   70
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label27 
         Caption         =   "Obra Social:"
         Height          =   195
         Left            =   705
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   47
         Top             =   3240
         Width           =   3135
      End
   End
   Begin VB.Frame fraTraslado 
      Caption         =   "Traslado"
      Height          =   4455
      Left            =   120
      TabIndex        =   27
      Top             =   4680
      Visible         =   0   'False
      Width           =   4935
      Begin TabDlg.SSTab SSTab1 
         Height          =   3015
         Left            =   120
         TabIndex        =   112
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5318
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Direccion Origen"
         TabPicture(0)   =   "frmAtencion.frx":945F
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "dirOrigen"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Direccion Destino"
         TabPicture(1)   =   "frmAtencion.frx":947B
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dirDestino"
         Tab(1).ControlCount=   1
         Begin ALCemi.ctlDireccion dirOrigen 
            Height          =   2565
            Left            =   120
            TabIndex        =   113
            Top             =   360
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
            Left            =   -74880
            TabIndex        =   114
            Top             =   360
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
   End
End
Attribute VB_Name = "frmAtencion"
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

Private WithEvents mFrmConsultarAfiliado As frmConsultarAfiliado
Attribute mFrmConsultarAfiliado.VB_VarHelpID = -1
Private WithEvents mFrmConsultarAfiliadoExterno As frmConsultarAfiliadoExterno
Attribute mFrmConsultarAfiliadoExterno.VB_VarHelpID = -1

Private WithEvents mFrmConsultarObraSocial As frmConsultarObraSocial
Attribute mFrmConsultarObraSocial.VB_VarHelpID = -1
Private WithEvents mFrmConsultarAreaProtegida As frmConsultarAreaProtegida
Attribute mFrmConsultarAreaProtegida.VB_VarHelpID = -1
Private WithEvents mFrmConsultarServicioEmergencia As frmConsultarServiciosEmergencia
Attribute mFrmConsultarServicioEmergencia.VB_VarHelpID = -1

Private WithEvents mFrmConsultarDotacion As frmConsultarDotaciones
Attribute mFrmConsultarDotacion.VB_VarHelpID = -1

Private mFrmParent As frmConsultaAtencion 'para enviarle eventos porq en gral se van a abrir varias frmAtencion

Private WithEvents mFrmBuscarSintoma As frmSeleccionarSintoma
Attribute mFrmBuscarSintoma.VB_VarHelpID = -1

Private WithEvents mFrmABMT As frmABMTelefono
Attribute mFrmABMT.VB_VarHelpID = -1
Private telsAux As blcemi.TelefonoManager


Private Sub cmbTelefono_ItemSeleccionado(Item As Object)
    On Error GoTo errman:
    Set mAtencion.Telefono = Item
    Exit Sub
errman:
    GBL.PrintToErrorLog "frmatencion", "cmbtelefono_itemSeleccionado", Err.Description
End Sub

Private Sub cmdAsignar_Click()

    If Not mAfiliadoPropio Is Nothing Then
        If MsgBox("Desea asignar el telefono al afiliado?", vbYesNo) = vbYes Then
            Set telsAux = mAfiliadoPropio.Telefonos
            AsignarTel telsAux
        End If
    ElseIf Not mAreaProtegida Is Nothing Then
        If MsgBox("Desea asignar al Area Protegida?", vbYesNo) = vbYes Then
            Set telsAux = mAreaProtegida.Telefonos
            AsignarTel telsAux
        End If
    ElseIf Not mObraSocial Is Nothing Then
        If MsgBox("Desea asignar el telefono a la Obra Social?", vbYesNo) = vbYes Then
           Set telsAux = mObraSocial.Telefonos
           AsignarTel telsAux
        End If
    ElseIf Not mServicioEmergencia Is Nothing Then
        If MsgBox("Desea asignar el telefono al Servicio de Emergencias?", vbYesNo) = vbYes Then
           Set telsAux = mServicioEmergencia.Telefonos
           AsignarTel telsAux
        End If
    Else
        MsgBox "Para poder asignar el telefono tiene que seleccionar un Afiliado, una Obra Social, un Servicio de Emergencias o un Area Protegida primero.", vbInformation
    End If
End Sub

Private Sub AsignarTel(tels As blcemi.TelefonoManager)
    Set mFrmABMT = New frmABMTelefono
    mFrmABMT.Nuevo tels
    mFrmABMT.txtNumero = txtTelefono
    cmdAsignar.Enabled = False
End Sub

Private Sub cmbTelefono_NuevoSeleccionado()
    If Not mAfiliadoPropio Is Nothing Then
        AsignarTel mAfiliadoPropio.Telefonos
    ElseIf Not mAreaProtegida Is Nothing Then
        AsignarTel mAreaProtegida.Telefonos
    ElseIf Not mObraSocial Is Nothing Then
        AsignarTel mObraSocial.Telefonos
    ElseIf Not mServicioEmergencia Is Nothing Then
        AsignarTel mServicioEmergencia.Telefonos
    End If
End Sub

Private Sub lvwAtencionesAE_ItemClick(Item As Object)
    Dim frmDA As New frmDetalleAtencion
    frmDA.VerDetalleAtencion Item
End Sub

Private Sub lvwAtencionesAP_ItemDblClick(Item As Object)
    Dim frmDA As New frmDetalleAtencion
    frmDA.VerDetalleAtencion Item
End Sub

Private Sub mFrmABMT_Cancelado()
    cmdAsignar.Enabled = True
End Sub

Private Sub mFrmABMT_Nuevo(pTelefono As blcemi.Telefono)
    On Error Resume Next
    Select Case telsAux.OwnerType
        Case blcemi.eOwnerType.eOTAfiliado
            mAfiliadoPropio.GuardarModificaciones
        Case blcemi.eOwnerType.eOTAreaProtegida
            mAreaProtegida.GuardarModificaciones
        Case blcemi.eOwnerType.eOTObraSocial
            mObraSocial.GuardarModificaciones
        Case blcemi.eOwnerType.eOTServicioEmergencia
            mServicioEmergencia.GuardarModificaciones
    End Select
    Set mAtencion.Telefono = pTelefono
End Sub

'muestra un combo con los telefonos si el txtTelefono esta vacio
'y si la coleccion tiene telefonos
Private Sub MostrarTelefonos(tels As blcemi.TelefonoManager)
    If mAtencion.Telefono Is Nothing Then 'si no hay un telefono
        If Not tels Is Nothing Then
            If tels.Count <> 0 Then ' si tiene telefonos, los muestro a menos q ya haya cargado algo en el txttel
                If txtTelefono = "" Then
                    cmbTelefono.Visible = True
                    txtTelefono.Visible = False
                    cmdAsignar.Enabled = False
                    Set cmbTelefono.Coleccion = tels
                Else 'pero si el txt tiene algo, me fijo a ver si lo tengo guardado
                    If Not tels.ItemByTelNumber(txtTelefono) Is Nothing Then
                        cmdAsignar.Enabled = False 'ya esta guardado
                    End If
                End If
            Else 'no tiene telefonos, lo dejamos ingresar uno
               cmdAsignar.Enabled = True
               cmbTelefono.Visible = False
               txtTelefono.Visible = True
            End If
        End If
    Else 'si ya esta guardado, lo muestro
        txtTelefono.Visible = True
        txtTelefono = mAtencion.Telefono.numero
        cmdAsignar.Enabled = False
    End If
End Sub

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

'funciona
'Private Sub fraAtencion_DragDrop(Source As Control, X As Single, Y As Single)
'If TypeOf Source Is ctlTelefonos Then
'    Dim ctlTel As ctlTelefonos
'    Set ctlTel = Source
'    txtTelefono = ctlTel.TelefonoDragged.Numero
'End If
'End Sub

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

Private Sub cmdConsultarAExterno_Click()

    If Not mAreaProtegida Is Nothing Then
        Set mFrmConsultarAfiliadoExterno = New frmConsultarAfiliadoExterno
        mFrmConsultarAfiliadoExterno.Consultar mAreaProtegida.Afiliados, etConRetorno
    ElseIf Not mObraSocial Is Nothing Then
        Set mFrmConsultarAfiliadoExterno = New frmConsultarAfiliadoExterno
        mFrmConsultarAfiliadoExterno.Consultar mObraSocial.Afiliados, etConRetorno
    ElseIf Not mServicioEmergencia Is Nothing Then
        Set mFrmConsultarAfiliadoExterno = New frmConsultarAfiliadoExterno
        mFrmConsultarAfiliadoExterno.Consultar mServicioEmergencia.Afiliados, etConRetorno
    End If

End Sub

Private Sub cmdDotacion_Click()
    Set mFrmConsultarDotacion = New frmConsultarDotaciones
    mFrmConsultarDotacion.Consultar GBL.EquiposGBL, etConRetorno
End Sub

Private Sub cmdConsultarAPropio_Click()
    Set mFrmConsultarAfiliado = New frmConsultarAfiliado
    'muestro todos los afiliados, titulares y a cargo
    mFrmConsultarAfiliado.Consultar GBL.AfiliadosGBL, etConRetorno
End Sub

Private Sub cmdAfiliadoPropio_Click()
    cmdConsultarAPropio_Click
    fraSeleccion.Visible = False
    fraAfiliadoPropio.Visible = True
    MostrarFrames
End Sub

Private Sub cmdConsultarAP_Click()
    Set mFrmConsultarAreaProtegida = New frmConsultarAreaProtegida
    mFrmConsultarAreaProtegida.Consultar GBL.AreasProtegidasGBL, etConRetorno
End Sub

Private Sub cmdAreaProtegida_Click()
    cmdConsultarAP_Click
    fraSeleccion.Visible = False
    fraAfiliadoExterno.Visible = True
    fraAreaProtegida.Visible = True
    MostrarFrames
End Sub

Private Sub cmdConsultarOS_Click()
    Set mFrmConsultarObraSocial = New frmConsultarObraSocial
    mFrmConsultarObraSocial.Consultar GBL.ObrasSocialesGBL, etConRetorno
End Sub

Private Sub cmdObraSocial_Click()
    cmdConsultarOS_Click
    fraAfiliadoExterno.Visible = True
    fraObraSocial.Visible = True
    MostrarFrames
End Sub

Private Sub cmdConsultarSE_Click()
    Set mFrmConsultarServicioEmergencia = New frmConsultarServiciosEmergencia
    mFrmConsultarServicioEmergencia.Consultar GBL.ServiciosEmergenciaGBL, etConRetorno
End Sub

Private Sub cmdServicioEmergencia_Click()
    cmdConsultarSE_Click
    fraAfiliadoExterno.Visible = True
    fraServicioEmergencia.Visible = True
    MostrarFrames
End Sub

Private Sub MostrarFrames()
    fraSeleccion.Visible = False
    fraCodigoEmerg.Visible = True
    fraAtencion.Visible = True
    fraComun.Visible = True
    cmdGuardar.Visible = True
    cmdCancelar.Left = 8760
    cmdCancelar.Top = cmdGuardar.Top
    Me.Width = 10650
    Me.Height = 9620
End Sub

Private Sub cmdFinalizar_Click()
    If MsgBox("Esta seguro que desea finalizar la Atención? " + vbCrLf + "(Una vez cerrada la misma no se podrán realizar modificaciones).", vbQuestion + vbYesNo, "tbrEmergencyGroup") = vbYes Then
        Guardar True
    End If
End Sub

Private Sub cmdGuardar_Click()
    Guardar False
End Sub

Private Sub Guardar(cerrar As Boolean)

    TERR.Anotar "gret", cerrar
    On Local Error GoTo errSave

    'seteo todos, alguno debe ser, si estan todos vacios despues hay q preguntar
    'lo unico q necesito si o si es el sintoma
    If DatosBasicosCorrectos Then
        TERR.Anotar "grev"
        'los equipos no los asigno porq los tengo referenciados en el tvw.
        Set mAtencion.Afiliado = mAfiliadoPropio
        TERR.Anotar "grew"
        Set mAtencion.AfiliadoExterno = mAfiliadoExterno
        TERR.Anotar "grex"
        Set mAtencion.AreaProtegida = mAreaProtegida
        TERR.Anotar "grey"
        Set mAtencion.ObraSocial = mObraSocial
        TERR.Anotar "grez"
        Set mAtencion.ServicioEmergencia = mServicioEmergencia
        TERR.Anotar "grfa"
        If Not cmbCodigo.SelectedItem Is Nothing Then
            If cmbCodigo.SelectedItem.id = 100 Then 'traslado
                TERR.Anotar "grfb"
                Set mAtencion.DireccionDestino = dirDestino.MiDireccion
                Set mAtencion.DireccionOrigen = dirOrigen.MiDireccion
            Else
                TERR.Anotar "grfc"
                If mAtencion.Telefono Is Nothing Then mAtencion.TelefonoAuxilar = txtTelefono.Text
                Set mAtencion.Direccion = ctlDirEmergencia.MiDireccion
            End If
            TERR.Anotar "grfd"
        End If
        TERR.Anotar "grfe"
        Set mAtencion.despachador = UsuarioActual
        TERR.Anotar "grff", txtDiagnostico
        mAtencion.Diagnostico = txtDiagnostico
        TERR.Anotar "grfg", cerrar
        If cerrar Then
            TERR.Anotar "grfh"
            mAtencion.Estado = blcemi.eFinalizado
        Else
            TERR.Anotar "grfi"
            mAtencion.Estado = IIf(DatosCompletos, blcemi.eestadoatencion.eListaParaCerrar, blcemi.eestadoatencion.ePendiente)
        End If
        TERR.Anotar "grfj", txtFecha.Text
        mAtencion.fecha = CDate(txtFecha.Text)
        ' Trim(Str(dtpHora.Hour)) + ":" + Trim(Str(dtpHora.Minute)) + ":" + Trim(Str(dtpHora.Second))
        TERR.Anotar "grfk", txtNroIncidente
        mAtencion.NroIncidente = txtNroIncidente
        TERR.Anotar "grfl", txtNroInterno
        mAtencion.nroIncidenteInterno = txtNroInterno
        TERR.Anotar "grfm", txtObservaciones
        mAtencion.Observaciones = txtObservaciones
        TERR.Anotar "grfo", txtQTH, txtOperador
        mAtencion.Operador = txtOperador
        mAtencion.QTH = txtQTH
        TERR.Anotar "grfp"
        Set mAtencion.Sintoma = cmbSintoma.SelectedItem
        TERR.Anotar "grfq", txtVL
        If CCFFGG.Configuracion.Codigo.UtilizarTipos = True Then
            TERR.Anotar "grfq2"
            Set mAtencion.TipoCodigo = cmbTipo.SelectedItem
        End If
        mAtencion.VL = txtVL
        TERR.Anotar "grfr", txtCopago, txtAbonado, txtServicio
        'ver al condicion de IVA si hace falta
        'If Trim(txtServicio) <> "" Then mAtencion.InfoContable.CondicionIVA
        If Trim(txtCopago) <> "" Then mAtencion.InfoContable.Coseguro = CCurrency(txtCopago)
        If Trim(txtAbonado) <> "" Then mAtencion.InfoContable.MontoAbonado = CCurrency(txtAbonado)
        If Trim(txtServicio) <> "" Then mAtencion.InfoContable.Servicio = CCurrency(txtServicio)
        TERR.Anotar "grfs", cmbIva.ListIndex
        If cmbIva.ListIndex <> -1 Then
            TERR.Anotar "grft"
            mAtencion.InfoContable.CondicionIVA = Choose(cmbIva.ListIndex + 1, blcemi.eCondicionIVA.eCINoInformado, blcemi.eCondicionIVA.eCIGravado, blcemi.eCondicionIVA.eCIExento)
        End If
        
        TERR.Anotar "grfu", mAtencion.id
        If mAtencion.id = 0 Then
            mAtencion.HoraLlamada = Time
            mAtencion.Guardar
        Else
            TERR.Anotar "grfv"
            mAtencion.GuardarModificaciones UsuarioActual
        End If
        TERR.Anotar "grfw"
        mFrmParent.Refrescar
        TERR.Anotar "grfx"
        Unload Me
    
    End If
    
    Exit Sub
errSave:
    TERR.AppendLog "greu", "error al grabar atencion" + vbCrLf + TERR.ErrToTXT(Err)
    MsgBox "Ocurrieron errores al grabar la atencion, envie infoirme de errores a info@tbrSoft"
    frmInformarError2.ForzarEnvio "Se estaba grabando la atencion"
    
End Sub

Private Sub cmdCancelar_Click()
'preguntar si esta seguro cuando cargo datos
If Not mAreaProtegida Is Nothing Or Not mObraSocial Is Nothing Or Not mServicioEmergencia Is Nothing Or Not mAfiliadoPropio Is Nothing Then
    If MsgBox("Esta seguro que desea cancelar la atencion?", vbQuestion + vbOKCancel) = vbOK Then
        If Not mAtencion Is Nothing Then Set mAtencion.Equipos = Nothing 'para evitarme el tema del beginedit etc.
        Unload Me
    End If
Else
    Unload Me
End If

End Sub

Private Function DatosBasicosCorrectos() As Boolean
    Dim msj As String
    Dim msjDirOrigen As String
    Dim msjDirDestino As String
    Dim msjDir As String
    
    Dim sint As blcemi.Sintoma
    
    If Not cmbSintoma.SelectedItem Is Nothing Then
        Set sint = cmbSintoma.SelectedItem
        If sint.Parent.id <> 100 Then
            If Not ctlDirEmergencia.DireccionCompleta(msjDir) Then msj = msj + msjDir
            'verifico el telefono, solo si es atencion, habria que ver si lo tengo q agregar en traslado tmb
            If mAtencion.Telefono Is Nothing Then 'si no tiene un telefono...
                If txtTelefono.Text = "" Then
                    msj = "Debe ingresar el telefono desde el cual se realizo la llamada." + vbCrLf
                End If
            End If
        
        Else
            If Not dirDestino.DireccionCompleta(msjDirDestino) Then msj = msj + msjDirDestino
            If Not dirOrigen.DireccionCompleta(msjDirOrigen) Then msj = msj + msjDirOrigen
        End If
    Else
        msj = "Debe seleccionar un Sintoma para la Emergencia." + vbCrLf
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
    If fraTraslado.Visible Then 'es un traslado
        aux = aux And dirOrigen.MiDireccion.Calle <> ""
        aux = aux And dirDestino.MiDireccion.Calle <> ""
    Else 'es una atencion
        aux = aux And ctlDirEmergencia.MiDireccion.Calle <> ""
        aux = aux And txtQTH.Text <> ""
        aux = aux And txtVL <> ""
        aux = aux And txtOperador <> ""
    End If
    
    aux = aux And txtDiagnostico <> ""
    'si utiliza tipos, me fijo q tenga alguno seleccionado
    If CCFFGG.Configuracion.Codigo.UtilizarTipos Then
        aux = aux And Not (cmbTipo.SelectedItem Is Nothing)
    End If
    'me fijo si alguno de los destinos de la atencion esta seteado
    aux = aux And (Not mAfiliadoPropio Is Nothing Or ((Not mAreaProtegida Is Nothing Or Not mServicioEmergencia Is Nothing Or Not mObraSocial Is Nothing) And Not mAfiliadoExterno Is Nothing))
         
    'tiene que seleccionar un tipo de iva
    aux = aux And (cmbIva.ListIndex = 1 Or cmbIva.ListIndex = 2)
         
    'ver equipos e infocontable
    DatosCompletos = aux
     
End Function


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

Private Sub mFrmBuscarSintoma_SintomaSeleccionado(pSintoma As blcemi.Sintoma)
    Set cmbSintoma.SelectedItem = pSintoma
End Sub

'pnumero es si no conozco el tel
Public Sub RecibirLlamadoTelefono(pFrmParent As frmConsultaAtencion, pTelefono As blcemi.Telefono, Optional pNumero As String)
    
'TERMINAR!!!
    
    Set mAtencion = New blcemi.Atencion
    Set mFrmParent = pFrmParent
    Set tvw.Coleccion = mAtencion.Equipos
    MostrarFrames
    
    If Not pTelefono Is Nothing Then
        txtTelefono = pTelefono.numero
        'desabilito el boton porq el tel puede estar guardado, excepto q ownertype no este seteado
        'cmdAsignar.Enabled = False ver
        
        'me fijo quien llama
        Select Case pTelefono.OwnerType
            Case blcemi.eOwnerType.eOTAfiliado
                Set mAfiliadoPropio = GBL.AfiliadosGBL.Item(pTelefono.OwnerId)
                If Not mAfiliadoPropio Is Nothing Then
                    fraAfiliadoPropio.Visible = True
                    LlenarCamposAfiliadoPropio
                End If
            Case blcemi.eOwnerType.eOTAreaProtegida
                Set mAreaProtegida = GBL.AreasProtegidasGBL.Item(pTelefono.OwnerId)
                If Not mAreaProtegida Is Nothing Then
                    fraAfiliadoExterno.Visible = True
                    fraAreaProtegida.Visible = True
                    LlenarCamposAreaProtegida
                End If
            Case blcemi.eOwnerType.eOTObraSocial
                Set mObraSocial = GBL.ObrasSocialesGBL.Item(pTelefono.OwnerId)
                If Not mObraSocial Is Nothing Then
                    fraAfiliadoExterno.Visible = True
                    fraObraSocial.Visible = True
                    LlenarCamposObraSocial
                End If
            Case blcemi.eOwnerType.eOTServicioEmergencia
                Set mServicioEmergencia = GBL.ServiciosEmergenciaGBL.Item(pTelefono.OwnerId)
                If Not mServicioEmergencia Is Nothing Then
                    fraAfiliadoExterno.Visible = True
                    fraServicioEmergencia.Visible = True
                    LlenarCamposServicioEmergencia
                End If
            Case blcemi.eOwnerType.eOTAfiliadoExterno
                Dim afs As blcemi.AfiliadoExternoManager
                Set afs = New blcemi.AfiliadoExternoManager
                Set mAfiliadoExterno = afs.LoadById(pTelefono.OwnerId)
                If TypeOf mAfiliadoExterno.Parent Is blcemi.ObraSocial Then
                    Set mObraSocial = mAfiliadoExterno.Parent
                    fraObraSocial.Visible = True
                    LlenarCamposObraSocial
                ElseIf TypeOf mAfiliadoExterno.Parent Is blcemi.ServicioEmergencia Then
                    Set mServicioEmergencia = mAfiliadoExterno.Parent
                    fraServicioEmergencia.Visible = True
                    LlenarCamposServicioEmergencia
                End If
                fraAfiliadoExterno.Visible = True
                LlenarCamposAfiliadoExterno
            Case Else
                fraSeleccion.Visible = True
                
        End Select
    Else
        fraSeleccion.Visible = True
        txtTelefono = pNumero
    End If
    Me.Show
End Sub

Public Sub NuevaAtencion(pFrmParent As frmConsultaAtencion)
    Set mAtencion = New blcemi.Atencion
    Set mFrmParent = pFrmParent
    Set tvw.Coleccion = mAtencion.Equipos
    Me.Show
    cmdFinalizar.Visible = False
End Sub

Public Sub ModificarAtencion(pAtencion As blcemi.Atencion, pFrmParent As frmConsultaAtencion)
    Set mAtencion = pAtencion
    Set mFrmParent = pFrmParent
    Me.Show
    If mAtencion.Estado = blcemi.eListaParaCerrar Then cmdFinalizar.Visible = True
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
        fraSeleccion.Visible = True
    End If
    
    Set dirDestino.MiDireccion = mAtencion.DireccionDestino
    Set dirOrigen.MiDireccion = mAtencion.DireccionOrigen
    Set ctlDirEmergencia.MiDireccion = mAtencion.Direccion
    
    'mAtencion.Despachador empleadoactual, ver
    txtDiagnostico = mAtencion.Diagnostico
    Set tvw.Coleccion = mAtencion.Equipos
    'mAtencion.Estado ver
    txtFecha.Text = mAtencion.fecha
    txtHora.Text = mAtencion.HoraLlamada
    txtNroIncidente = mAtencion.NroIncidente
    txtNroInterno = mAtencion.nroIncidenteInterno
    txtObservaciones = mAtencion.Observaciones
    txtOperador = mAtencion.Operador
   
    If mAtencion.Sintoma.Parent.id <> 100 Then 'si no es un traslado
        'si hasta ahora no cargue un telefono muestro el auxiliar
         txtTelefono.Text = mAtencion.TelefonoAuxilar
    End If
    
    LlenarInfoContable
    
    Set cmbSintoma.SelectedItem = mAtencion.Sintoma
    Set cmbCodigo.SelectedItem = mAtencion.Sintoma.Parent
    If CCFFGG.Configuracion.Codigo.UtilizarTipos = True Then Set cmbTipo.SelectedItem = mAtencion.TipoCodigo
    txtVL = mAtencion.VL
    txtQTH = mAtencion.QTH
    If txtVL <> "" Then cmdRegistrarLiberacion.Visible = False
    If txtQTH <> "" Then
        cmdRegistrarArribo.Visible = False
        cmdRegistrarLiberacion.Enabled = True
    End If
End Sub

Public Sub LlenarInfoContable()
    txtCopago = IIf(mAtencion.InfoContable.Coseguro <> -1, mAtencion.InfoContable.Coseguro, "")
    txtAbonado = IIf(mAtencion.InfoContable.MontoAbonado <> -1, mAtencion.InfoContable.MontoAbonado, "")
    txtServicio = IIf(mAtencion.InfoContable.Servicio <> -1, mAtencion.InfoContable.Servicio, "")
    cmbIva.ListIndex = mAtencion.InfoContable.CondicionIVA
End Sub

'Private Sub ctlDirEmergencia_DireccionDragDrop(Source As Control, X As Single, Y As Single)
'If Source.Name = "lblDireccion" Then Set ctlDirEmergencia.MiDireccion = mAfiliadoPropio.Direccion
'
'End Sub

Private Sub Form_Load()

    On Local Error GoTo ErrATC

    TERR.Anotar "Atnc.load"
    Set cmdConsultarAExterno.Picture = MDI.il16.ListImages("buscar").Picture
    Set cmdConsultarAP.Picture = MDI.il16.ListImages("buscar").Picture
    Set cmdConsultarAPropio.Picture = MDI.il16.ListImages("buscar").Picture
    Set cmdConsultarOS.Picture = MDI.il16.ListImages("buscar").Picture
    Set cmdConsultarSE.Picture = MDI.il16.ListImages("buscar").Picture
    Set cmdLlamar.Picture = MDI.il16.ListImages("llamar").Picture
    Set cmdAsignar.Picture = MDI.il16.ListImages("guardar").Picture
    Set cmdBuscarSintoma.Picture = MDI.il16.ListImages("buscar").Picture
       
    TERR.Anotar "Atnc.load.2"
    Set cmbTipo.Coleccion = GBL.TiposCodigoGBL
    TERR.Anotar "Atnc.load.3"
    Set cmbCodigo.Coleccion = GBL.CodigoEmergenciaGBL
    TERR.Anotar "Atnc.load.4"
    Set ctlDirEmergencia.MiDireccion = New blcemi.Direccion
    TERR.Anotar "Atnc.load.5"
    Set dirOrigen.MiDireccion = New blcemi.Direccion
    TERR.Anotar "Atnc.load.6"
    Set dirDestino.MiDireccion = New blcemi.Direccion
    TERR.Anotar "Atnc.load.7"
    InicializarDireccion ctlDirEmergencia
    TERR.Anotar "Atnc.load.8"
    InicializarDireccion dirOrigen
    TERR.Anotar "Atnc.load.9"
    InicializarDireccion dirDestino
    TERR.Anotar "Atnc.load.10"
    Set cmbSintoma.Coleccion = GBL.SintomasGBL
    TERR.Anotar "Atnc.load.11"
    txtFecha.Text = Date
    txtHora.Text = Time
    Set Me.Icon = MDI.Icon

    Me.Move (MDI.Width - 10650) / 2, 0 'no uso me.width porq lo cambio despues del load
    TERR.Anotar "Atnc.load.12"
    AplicarConfiguracion
    TERR.Anotar "Atnc.load.13"
    AplicarPermisos
    TERR.Anotar "Atnc.load.14"
    
    Exit Sub
    
ErrATC:
    TERR.AppendLog "Atnc.load.ERROR", TERR.ErrToTXT(Err)
    Resume Next 'TODO esto no deberia ser asi!
End Sub

Private Sub AplicarConfiguracion()

    TERR.Anotar "Atnc.AplicCFG.1"

    If Not CCFFGG.Configuracion.Codigo.UtilizarTipos Then
        TERR.Anotar "Atnc.AplicCFG.2"
        cmbTipo.Visible = False
        lblTipo.Visible = False
        lblSintoma.Left = lblTipo.Left
        TERR.Anotar "Atnc.AplicCFG.4"
        txtSintoma.Left = lblSintoma.Left + 100 + lblSintoma.Width
        cmbSintoma.Left = txtSintoma.Left + 100 + txtSintoma.Width
        cmbSintoma.Width = 5635
    Else
        TERR.Anotar "Atnc.AplicCFG.3"
        cmbTipo.Visible = True
        lblTipo.Visible = True
        lblSintoma.Left = 5160
        txtSintoma.Left = 5880
        cmbSintoma.Left = 6360
        cmbSintoma.Width = 3435
    End If
    TERR.Anotar "Atnc.AplicCFG.5"
    lvwAtencionesAE.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    TERR.Anotar "Atnc.AplicCFG.6"
    lvwAtencionesAP.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    TERR.Anotar "Atnc.AplicCFG.7"

End Sub

Private Sub AplicarPermisos()
    TERR.Anotar "Atnc.AplicPER.1"
    sTabDetalles.TabVisible(2) = UsuarioActual.Permisos.Can(blcemi.VerInformacionContableAtencion)
    TERR.Anotar "Atnc.AplicCFG.2"
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "nuevaatencion"
End Function

Public Sub Refrescar()
    AplicarConfiguracion
    AplicarPermisos
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
    cmbTipo.Refresh
End Sub

'Private Sub lblDireccion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    If Button = vbLeftButton Then lblDireccion.Drag
'End Sub

Private Sub mFrmConsultarAfiliado_AfiliadoSeleccionado(pAfiliado As blcemi.Afiliado)
    Set mAfiliadoPropio = pAfiliado
    LlenarCamposAfiliadoPropio
    Me.SetFocus
End Sub

Private Sub mFrmConsultarAfiliadoExterno_AfiliadoExternoSeleccionado(pAfiliadoExterno As blcemi.AfiliadoExterno)
    Set mAfiliadoExterno = pAfiliadoExterno
    LlenarCamposAfiliadoExterno
    Me.SetFocus
End Sub

Private Sub mFrmConsultarAreaProtegida_AreaProtegidaSeleccionada(pAreaProtegida As blcemi.AreaProtegida)
    Set mAreaProtegida = pAreaProtegida
    LlenarCamposAreaProtegida
    Set mFrmConsultarAfiliadoExterno = New frmConsultarAfiliadoExterno
    mFrmConsultarAfiliadoExterno.Consultar pAreaProtegida.Afiliados, etConRetorno
End Sub

Private Sub mFrmConsultarDotacion_EquipoSeleccionado(pEquipo As blcemi.Equipo)
    mAtencion.Equipos.AddItem pEquipo
    tvw.Refresh
End Sub

Private Sub mFrmConsultarDotacion_EquiposSeleccionados(pEquipos As blcemi.EquipoManager)
    Set mAtencion.Equipos = pEquipos
    Set tvw.Coleccion = mAtencion.Equipos
    tvw.Refresh
End Sub

Private Sub mFrmConsultarObraSocial_ObraSocialSeleccionada(pObraSocial As blcemi.ObraSocial)
    Set mObraSocial = pObraSocial
    LlenarCamposObraSocial
    Set mFrmConsultarAfiliadoExterno = New frmConsultarAfiliadoExterno
    mFrmConsultarAfiliadoExterno.Consultar pObraSocial.Afiliados, etConRetorno
End Sub

Private Sub mFrmConsultarServicioEmergencia_ServicioEmergenciaSeleccionado(pServicioEmergencia As blcemi.ServicioEmergencia)
    Set mServicioEmergencia = pServicioEmergencia
    LlenarCamposServicioEmergencia
    Set mFrmConsultarAfiliadoExterno = New frmConsultarAfiliadoExterno
    mFrmConsultarAfiliadoExterno.Consultar pServicioEmergencia.Afiliados, etConRetorno
End Sub

Private Sub mFrmConsultarAfiliado_SeleccionCancelada()
Me.SetFocus
End Sub

Private Sub mFrmConsultarAfiliadoExterno_SeleccionCancelada()
Me.SetFocus
End Sub

Private Sub mFrmConsultarAreaProtegida_SeleccionCancelada()
Me.SetFocus
End Sub

Private Sub mFrmConsultarObraSocial_SeleccionCancelada()
Me.SetFocus
End Sub

Private Sub mFrmConsultarServicioEmergencia_SeleccionCancelada()
Me.SetFocus
End Sub

Private Sub cmbCodigo_ItemSeleccionado(Item As Object)
    Dim codEm As blcemi.CodigoEmergencia
    Set codEm = Item
    If Not cmbSintoma.SelectedItem Is Nothing Then
        If cmbSintoma.SelectedItem.Parent.id <> codEm.id Then
            Set cmbSintoma.Coleccion = codEm.Sintomas
        End If
    Else
        Set cmbSintoma.Coleccion = cmbCodigo.SelectedItem.Sintomas
    End If
    
    If codEm.id = 100 Then 'es un traslado
        fraTraslado.Visible = True
        fraAtencion.Visible = False
    Else
        fraTraslado.Visible = False
        fraAtencion.Visible = True
    End If
    MostrarCoseguro
End Sub

Private Sub cmbSintoma_ItemSeleccionado(Item As Object)
    txtSintoma = ""
    txtCodigo = ""
    Set cmbCodigo.SelectedItem = Item.Parent
End Sub

Private Sub txtCodigo_Change()
    Dim cod As blcemi.CodigoEmergencia
    Set cod = GBL.CodigoEmergenciaGBL(Val(txtCodigo))
    If Not cod Is Nothing Then
        txtSintoma = ""
        Set cmbCodigo.SelectedItem = cod
    End If
End Sub


Private Sub txtSintoma_Change()
On Error Resume Next
If txtSintoma <> "" Then
    txtCodigo = ""
    Dim sint As blcemi.Sintoma
    Set sint = GBL.SintomasGBL.Item(Val(txtSintoma.Text))
    If Not sint Is Nothing Then
        Set cmbSintoma.SelectedItem = sint
    End If
End If
End Sub

Private Sub cmbTipo_ItemSeleccionado(Item As Object)
    MostrarCoseguro
End Sub

Private Sub MostrarCoseguro()
    If Not cmbCodigo.SelectedItem Is Nothing Then
        Dim cods As blcemi.CodigoCubiertoManager
        Dim codCub As blcemi.CodigoCubierto
        If Not mServicioEmergencia Is Nothing Then
            Set cods = mServicioEmergencia.CodigosCubiertos
        ElseIf Not mObraSocial Is Nothing Then
            Set cods = mObraSocial.CodigosCubiertos
        End If
        If Not cods Is Nothing Then
            If CCFFGG.Configuracion.Codigo.UtilizarTipos And Not cmbTipo.SelectedItem Is Nothing Then
                Set codCub = cods.Item(cmbCodigo.SelectedItem.id, cmbTipo.SelectedItem.id)
            Else
                Set codCub = cods.Item(cmbCodigo.SelectedItem.id, 0)
            End If
            If Not codCub Is Nothing Then
                txtCopago = codCub.Coseguro
                txtServicio = codCub.Servicio
            End If
        End If
    End If
End Sub

'--------info contable--------------
Private Sub txtAbonado_Change()
CalcularSaldo
End Sub

Private Sub txtAbonado_KeyPress(KeyAscii As Integer)
        SoloNumeros KeyAscii, True
End Sub

Private Sub txtCopago_Change()
CalcularSaldo
End Sub

Private Sub txtCopago_KeyPress(KeyAscii As Integer)
        SoloNumeros KeyAscii, True
End Sub

Private Sub txtServicio_Change()
CalcularSaldo
End Sub

Private Sub txtServicio_KeyPress(KeyAscii As Integer)
        SoloNumeros KeyAscii, True
End Sub

Private Sub CalcularSaldo()
On Error GoTo errman
txtSaldo = CCurrency(txtServicio) - CCurrency(txtAbonado)
Exit Sub
errman:
txtSaldo = "-"
End Sub

Private Sub tvw_ItemKeyDeletePressed(Item As Object)
mAtencion.Equipos.Remove Item.id
tvw.Refresh
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
    
    rTxtHCAPropio.Text = GetResumenHC(mAfiliadoPropio.HistoriaClinica)
    
    lblObraSocialAPropio = mAfiliadoPropio.ObraSocial.Nombre
    
    If Not ctlDirEmergencia.DireccionCompleta("") Then 'si la direccion esta vacia...
        Set ctlDirEmergencia.MiDireccion = mAfiliadoPropio.Direccion
    End If
    
    Set lvwAtencionesAP.Coleccion = mAfiliadoPropio.Atenciones
    MostrarTelefonos mAfiliadoPropio.Telefonos
'    If mTipoAfiliado = eTitular Then
'        'mAfiliado.Vehiculo
'        Set lvwACargo.Coleccion = mAfiliado.PersonasACargo
'        Set cmbCobrador.SelectedItem = mAfiliado.Cobrador
'        txtTope = mAfiliado.TopeAtenciones
'        'mAfiliado.Pagos
'    Else
'        'no deberia estar, lo dejo para acordarme...
'        'mAfiliado.Parent
'        Set cmbParentezco.SelectedItem = mAfiliadoPropio.Parentezco
'    End If
    
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
   
   MostrarTelefonos mObraSocial.Telefonos
      
   'no porq nunca va a ser desde la obra social la emergencia
   'Set ctlDirEmergencia.MiDireccion = mObraSocial.Direccion
   
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
    
    MostrarTelefonos mServicioEmergencia.Telefonos
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
    
    MostrarTelefonos mAreaProtegida.Telefonos
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
    Set lvwAtencionesAE.Coleccion = mAfiliadoExterno.Atenciones
    
    If Not ctlDirEmergencia.DireccionCompleta("") Then
        Set ctlDirEmergencia.MiDireccion = mAfiliadoExterno.Direccion
    End If
    
End Sub

Private Function GetResumenHC(pHC As blcemi.HistoriaClinica) As String
    Dim aux As String
    Dim a As blcemi.Alergia
    Dim e As blcemi.Enfermedad
    Dim m As blcemi.Medicamento
    aux = aux + "Alergias:" + vbCrLf
    
    If pHC.Alergias.Count = 0 Then aux = aux + vbTab + "No tiene alergias registradas." + vbCrLf
    For Each a In pHC.Alergias
        aux = aux + vbTab + a.Nombre + vbCrLf
    Next
     aux = aux + "Enfermedades:" + vbCrLf
    If pHC.Enfermedades.Count = 0 Then aux = aux + vbTab + "No tiene enfermedades registradas." + vbCrLf
    For Each e In pHC.Enfermedades
        aux = aux + vbTab + e.Nombre + vbCrLf
    Next
    aux = aux + "Medicamentos:" + vbCrLf
    If pHC.Alergias.Count = 0 Then aux = aux + vbTab + "No tiene medicamentos registrados." + vbCrLf
    For Each m In pHC.Medicamentos
        aux = aux + vbTab + m.Nombre + vbCrLf
    Next
    GetResumenHC = aux
End Function
