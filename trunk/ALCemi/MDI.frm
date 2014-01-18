VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "TBR Emergency Group"
   ClientHeight    =   6540
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11235
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList il32 
      Left            =   120
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   13027014
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   28
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":08A6
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1558
            Key             =   "salidap"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":21AA
            Key             =   "llegadap"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":2DFC
            Key             =   "salidad"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3A4E
            Key             =   "recibosanulados"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":4AA2
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":56F6
            Key             =   "liquidacion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":634A
            Key             =   "registrarcobro"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6F9E
            Key             =   "writer"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7C50
            Key             =   "qth"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":88A4
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":94F8
            Key             =   "restaurar"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A14C
            Key             =   "vl"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":ADA0
            Key             =   "papelerallena"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B9F4
            Key             =   "papeleravacia"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":C648
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":D29C
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":DEF0
            Key             =   "registraratencion"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":EB44
            Key             =   "word"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":F798
            Key             =   "excel"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":103EC
            Key             =   "agregar"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":11040
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":11C94
            Key             =   "detalles"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":128E8
            Key             =   "guardia"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1353C
            Key             =   "afiliados"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":14190
            Key             =   "listado"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":14DE4
            Key             =   "aceptar"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":15A38
            Key             =   "configurar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilBarra 
      Left            =   1560
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   13027014
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1668A
            Key             =   "alerta"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":172DC
            Key             =   "areasprotegidas"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":17F2E
            Key             =   "moviles"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":18B80
            Key             =   "atencionesalerta"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":197D4
            Key             =   "empleados"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1A426
            Key             =   "afiliados"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1B078
            Key             =   "dotaciones"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1BCCA
            Key             =   "iniciosesion"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1C91C
            Key             =   "consultaratenciones"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1D570
            Key             =   "cerrarsesion"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1E1C2
            Key             =   "obrasocial"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1EE14
            Key             =   "registraratencion"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":1FA68
            Key             =   "servicioemergencia"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLlamada 
      Align           =   1  'Align Top
      BackColor       =   &H00C0C0C0&
      Height          =   915
      Left            =   0
      Picture         =   "MDI.frx":206BA
      ScaleHeight     =   855
      ScaleWidth      =   11175
      TabIndex        =   3
      Top             =   330
      Visible         =   0   'False
      Width           =   11235
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Llamada Entrante!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label lblTelefono 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "(03453) - 15515245"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label lblTipo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Servicio de Emergencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   7
         Top             =   120
         Width           =   4455
      End
      Begin VB.Label lblNombre 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Algun servicio de emergencias"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   480
         Width           =   4455
      End
      Begin VB.Label lblRegistrarAtencion 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "Registrar Atencion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6840
         MouseIcon       =   "MDI.frx":2D5CD
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label lblCancelar 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   6840
         MouseIcon       =   "MDI.frx":2D8D7
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6285
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il16 
      Left            =   840
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":2DBE1
            Key             =   "aceptar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3C783
            Key             =   "atencion"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":3D3D5
            Key             =   "agregar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":43C37
            Key             =   "empleados"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":44889
            Key             =   "afiliados"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":454DB
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":4612D
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":4951F
            Key             =   "buscar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":4A173
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":51675
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":522C7
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":58B29
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":5EDC3
            Key             =   "papeleravacia"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6505D
            Key             =   "llamar"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":653AF
            Key             =   "guardar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il32Vieja 
      Left            =   120
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   26
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":65701
            Key             =   "excel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6BF63
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6CC15
            Key             =   "writer"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6D8C7
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6E519
            Key             =   "registrarcobro"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":6F52B
            Key             =   "word"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":75D8D
            Key             =   "restaurar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":769DF
            Key             =   "agregar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7D241
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":7DE93
            Key             =   "aceptar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8CA35
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":8D687
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":90A79
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":97F7B
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":9E7DD
            Key             =   "qth"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":9F42F
            Key             =   "vl"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A0081
            Key             =   "recibosanulados"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A0CD3
            Key             =   "papeleravacia"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A1925
            Key             =   "papelerallena"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A2577
            Key             =   "detalles"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A8DD9
            Key             =   "registraratencion"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":A9A2B
            Key             =   "moviles"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":AA67D
            Key             =   "liquidacion"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":AB2CF
            Key             =   "guardia"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":AC005
            Key             =   "afiliados"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":ACC57
            Key             =   "listado"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      DragIcon        =   "MDI.frx":AD8A9
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _ExtentX        =   19817
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "iniciosesion"
            Object.ToolTipText     =   "Iniciar Sesion"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "afiliados"
            Object.ToolTipText     =   "Consultar Afiliados"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "areasprotegidas"
            Object.ToolTipText     =   "Consultar Areas Protegidas"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep1"
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "obrasocial"
            Object.ToolTipText     =   "Consultar Obras Sociales"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "servicioemergencia"
            Object.ToolTipText     =   "Consultar Servicios de Emergencias"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "sep2"
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "registraratencion"
            Object.ToolTipText     =   "Registrar Atencion"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultaratenciones"
            Object.ToolTipText     =   "Consultar Listado de Atenciones Pendientes"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "empleados"
            Object.ToolTipText     =   "Consultar Empleados"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "moviles"
            Object.ToolTipText     =   "Consultar Moviles"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "dotaciones"
            Object.ToolTipText     =   "Consultar Dotaciones"
         EndProperty
      EndProperty
      Begin VB.Timer timerLlamada 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   8280
         Top             =   0
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   7800
         Top             =   0
      End
      Begin VB.PictureBox picAlerta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   8760
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   2
         ToolTipText     =   "Hay atenciones pendientes!"
         Top             =   0
         Width           =   615
      End
   End
   Begin MSComctlLib.ImageList ilBarraB 
      Left            =   2160
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   13027014
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":ADBB3
            Key             =   "alerta"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":AE805
            Key             =   "areasprotegidas"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":AF457
            Key             =   "atencionesalerta"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B00AB
            Key             =   "afiliados"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B0CFD
            Key             =   "moviles"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B194F
            Key             =   "iniciosesion"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B25A1
            Key             =   "empleados"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B31F3
            Key             =   "consultaratenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B3E47
            Key             =   "dotaciones"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B4A99
            Key             =   "cerrarsesion"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B56EB
            Key             =   "obrasocial"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B633D
            Key             =   "registraratencion"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDI.frx":B6F91
            Key             =   "servicioemergencia"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuInicioSesion 
         Caption         =   "Iniciar Sesion..."
      End
      Begin VB.Menu mnuCambioPass 
         Caption         =   "Cambiar Contraseña..."
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAfiliados 
         Caption         =   "Afiliados"
      End
      Begin VB.Menu mnuAreasProtegidas 
         Caption         =   "Areas Protegidas"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObrasSociales 
         Caption         =   "Obras Sociales"
      End
      Begin VB.Menu mnuServiciosEmergencia 
         Caption         =   "Servicios de Emergencia"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovil 
         Caption         =   "Moviles"
      End
      Begin VB.Menu mnuDotaciones 
         Caption         =   "Dotaciones"
      End
      Begin VB.Menu mnusep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnCreateReg 
         Caption         =   "Reporte de errores"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuActualizar 
         Caption         =   "Actualizar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuVerMenu 
         Caption         =   "Barra de Menu"
      End
   End
   Begin VB.Menu mnuAtencion 
      Caption         =   "Atencion"
      Begin VB.Menu mnuRegistrarAtencion 
         Caption         =   "Registrar Atencion"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuConsultarAtencionesPendientes 
         Caption         =   "Listado de Atenciones Pendientes"
      End
      Begin VB.Menu mnusep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuListadoAtenciones 
         Caption         =   "Listado de Atenciones"
      End
   End
   Begin VB.Menu mnuAdministracion 
      Caption         =   "Administracion"
      Begin VB.Menu mnuListadoCuotasACobrar 
         Caption         =   "Emitir Listado de Cuotas a Cobrar"
      End
      Begin VB.Menu mnuRegistrarRecibosAnulados 
         Caption         =   "Registrar Devolucion de Recibos Anulados"
      End
      Begin VB.Menu mnuSep10 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLiqEmpleados 
         Caption         =   "Consultar Liquidaciones a Empleados"
      End
      Begin VB.Menu mnuLiqEmpresas 
         Caption         =   "Consultar Liquidaciones a Empresas"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnuInformes 
         Caption         =   "Informes"
      End
   End
   Begin VB.Menu mnuMantenimiento 
      Caption         =   "Mantenimiento"
      Begin VB.Menu mnuCargos 
         Caption         =   "Cargos"
      End
      Begin VB.Menu mnuOcupaciones 
         Caption         =   "Ocupaciones"
      End
      Begin VB.Menu mnuParentezcos 
         Caption         =   "Parentezcos"
      End
      Begin VB.Menu mnusep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAlergias 
         Caption         =   "Alergias"
      End
      Begin VB.Menu mnuEnfermedades 
         Caption         =   "Enfermedades"
      End
      Begin VB.Menu mnuMedicamentos 
         Caption         =   "Medicamentos"
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTiposCodigo 
         Caption         =   "Tipos de Codigo de Emergencia"
      End
      Begin VB.Menu mnuTipoTelefono 
         Caption         =   "Tipos de Telefono"
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCuerpos 
         Caption         =   "Cuerpos de Bomberos"
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInstGas 
         Caption         =   "Tipos de Instalacion de Gas"
      End
      Begin VB.Menu mnuInstElectrica 
         Caption         =   "Tipos de Instalacion Electrica"
      End
   End
   Begin VB.Menu mnuHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu mnuConfiguracion 
         Caption         =   "Configuracion"
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCamara 
         Caption         =   "Ver Camara"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuImportar 
         Caption         =   "Importar datos externos"
      End
      Begin VB.Menu mnuEstadoRed 
         Caption         =   "Estado de la red"
      End
      Begin VB.Menu mnuSetearUltimoRecibo 
         Caption         =   "Setear Ultimo Recibo"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnLicencia 
         Caption         =   "Licencia"
         Begin VB.Menu mnGenerarLic 
            Caption         =   "Generar archivo para validar equipo"
         End
         Begin VB.Menu mnInsertLic 
            Caption         =   "Insertar licencia recibida"
         End
      End
   End
   Begin VB.Menu mnuVentana 
      Caption         =   "Ventana"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ayuda"
      Begin VB.Menu mnuContenido 
         Caption         =   "Contenido"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuManual 
         Caption         =   "Manual de Usuario"
         Visible         =   0   'False
      End
      Begin VB.Menu mnusep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca de..."
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAX_CAMPOS = 2
Public Enum HH_COMMAND
    HH_DISPLAY_TOPIC = &H0
    HH_HELP_FINDER = &H0 ' WinHelp equivalent
    HH_DISPLAY_TOC = &H1 ' not currently implemented
    HH_DISPLAY_INDEX = &H2 ' not currently implemented
    HH_DISPLAY_SEARCH = &H3 ' not currently implemented
    HH_SET_WIN_TYPE = &H4
    HH_GET_WIN_TYPE = &H5
    HH_GET_WIN_HANDLE = &H6
    HH_GET_INFO_TYPES = &H7 ' not currently implemented
    HH_SET_INFO_TYPES = &H8 ' not currently implemented
    HH_SYNC = &H9
    HH_ADD_NAV_UI = &HA ' not currently implemented
    HH_ADD_BUTTON = &HB ' not currently implemented
    HH_GETBROWSER_APP = &HC ' not currently implemented
    HH_KEYWORD_LOOKUP = &HD
    HH_DISPLAY_TEXT_POPUP = &HE ' display string resource id or text in a popup window
    HH_HELP_CONTEXT = &HF ' display mapped numeric value in dwData
    HH_TP_HELP_CONTEXTMENU ' Text pop-up help,
    ' similar to WinHelp's HELP_CONTEXTMENU
    HH_TP_HELP_WM_HELP = &H11 ' text pop-up help, similar to WinHelp's HELP_WM_HELP.
    HH_CLOSE_ALL = &H12 ' close all windows opened directly or indirectly by the caller
    HH_ALINK_LOOKUP = &H13 ' ALink version of HH_KEYWORD_LOOKUP
End Enum

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As HH_COMMAND, ByVal dwData As Long) As Long

'Dim Formularios As New FormManager
Dim WithEvents mFrmInicioSesion As frmInicioSesion
Attribute mFrmInicioSesion.VB_VarHelpID = -1
Dim WithEvents miConfig As tbrConfig.clsConfiguracion
Attribute miConfig.VB_VarHelpID = -1
Dim WithEvents mDBMonitor As blcemi.DBMonitor
Attribute mDBMonitor.VB_VarHelpID = -1
Dim WithEvents mErrHandler As blcemi.ErrHandler
Attribute mErrHandler.VB_VarHelpID = -1
Dim WithEvents mFrmAtencionesPendientes As frmConsultaAtencion
Attribute mFrmAtencionesPendientes.VB_VarHelpID = -1
Dim WithEvents mNetMonitor As blcemi.NetMonitor 'para q me avise el estado de la red
Attribute mNetMonitor.VB_VarHelpID = -1

Public WithEvents clienteTel As blcemi.tbrEGClient
Attribute clienteTel.VB_VarHelpID = -1
Dim tel As blcemi.Telefono 'lo uso para las llamadas entrantes
Dim telNumber As String 'por si el telefono es desconocido
Dim mAtencion As blcemi.Atencion 'si llama por tel alguien q tiene una atencion pendiente

Public Property Get FrmAtencionesPendientes() As frmConsultaAtencion
    If mFrmAtencionesPendientes Is Nothing Then Set mFrmAtencionesPendientes = New frmConsultaAtencion
    mFrmAtencionesPendientes.Inicializar
    Set FrmAtencionesPendientes = mFrmAtencionesPendientes
End Property


'-------------Llamadas telefonicas----------------------------
Private Sub clienteTel_LlamadaEntrando(numero As String)
RecibirLlamado numero
End Sub

Private Sub ConectarConGrabadorLlamadas()
    Set clienteTel = New blcemi.tbrEGClient
    On Error GoTo errman
    Dim cli As Object
    Set cli = clienteTel
    Set llamadasApp = CreateObject("tbrCalls.Application")
    llamadasApp.NotifyMe cli
    Exit Sub
errman:
    'gbl.PrintToErrorLog "No se pudo conectar a llamadas: " + Err.Description
End Sub

Private Sub lblCancelar_Click()
    OcultarPicLlamada True
End Sub

Private Sub lblRegistrarAtencion_Click()
     Select Case modoSoftware
         Case eModoFuncionamiento.eMFBomberos:
            Dim frmB As New frmAtencionBomberos
            If lblRegistrarAtencion.Caption = "Registrar Atencion" Then
                frmB.RecibirLlamadoTelefono FrmAtencionesPendientes, lblTelefono.Caption
                OcultarPicLlamada False
            Else 'modificar
                'warning: aca hay q mandarle un atencionB
                'frmB.ModificarAtencion mAtencion, FrmAtencionesPendientes
                OcultarPicLlamada False
            End If
         Case eModoFuncionamiento.eMFEmergencia:
            Dim frm As New frmAtencion
            If lblRegistrarAtencion.Caption = "Registrar Atencion" Then
                frm.RecibirLlamadoTelefono FrmAtencionesPendientes, tel, lblTelefono.Caption
                OcultarPicLlamada False
            Else 'modificar
                frm.ModificarAtencion mAtencion, FrmAtencionesPendientes
                OcultarPicLlamada False
            End If
    End Select
End Sub

Private Sub OcultarPicLlamada(pAnimar As Boolean)
Dim altura As Single
altura = picLlamada.Height - 1

'no hace falta un timer, con el retraso propio del for queda bien el efecto
If pAnimar Then
    For I = 1 To altura
        picLlamada.Height = picLlamada.Height - 1
    Next
End If
picLlamada.Visible = False
End Sub

Private Sub MostrarPicLlamada()
picLlamada.Visible = True
picLlamada.Height = 1
For I = 1 To 915
    picLlamada.Height = picLlamada.Height + 1
Next
End Sub

'enganchar esto donde me llama el programa de llamadas
Private Sub RecibirLlamado(pTelNumber As String)
    Dim a As blcemi.Atencion
    
    MostrarPicLlamada
    Set tel = GBL.TelefonosGBL.ItemByTelNumber(pTelNumber)
    If Not tel Is Nothing Then
        lblTelefono = tel.numero
        Select Case tel.OwnerType
            Case eOwnerType.eOTAfiliado:
                lblTipo = "Afiliado"
                lblNombre = GBL.AfiliadosGBL.Item(tel.OwnerId).NombreCompleto
                
                'me fijo si tengo una atencion pendiente de este afiliado
                For Each a In GBL.AtencionesGBL.GetByEstado(blcemi.ePendiente)
                    If Not a.Afiliado Is Nothing Then
                        If a.Afiliado.id = tel.OwnerId Then
                            lblRegistrarAtencion.Caption = "Modificar Atencion"
                            Set mAtencion = a
                            Exit For
                        End If
                    End If
                Next
            Case eOwnerType.eOTAreaProtegida:
                lblTipo = "Area Protegida"
                lblNombre = GBL.AreasProtegidasGBL.Item(tel.OwnerId).NombreArea
                For Each a In GBL.AtencionesGBL.GetByEstado(blcemi.ePendiente)
                If Not a.AreaProtegida Is Nothing Then
                        If a.AreaProtegida.id = tel.OwnerId Then
                            lblRegistrarAtencion.Caption = "Modificar Atencion"
                            Set mAtencion = a
                            Exit For
                        End If
                    End If
                Next
            Case eOwnerType.eOTObraSocial:
                lblTipo = "Obra Social"
                lblNombre = GBL.ObrasSocialesGBL.Item(tel.OwnerId).Nombre
                For Each a In GBL.AtencionesGBL.GetByEstado(blcemi.ePendiente)
                    If Not a.ObraSocial Is Nothing Then
                        If a.ObraSocial.id = tel.OwnerId Then
                            lblRegistrarAtencion.Caption = "Modificar Atencion"
                            Set mAtencion = a
                            Exit For
                        End If
                    End If
                Next
            Case eOwnerType.eOTServicioEmergencia:
                lblTipo = "Servicio de Emergencias"
                lblNombre = GBL.ServiciosEmergenciaGBL.Item(tel.OwnerId).Nombre
                For Each a In GBL.AtencionesGBL.GetByEstado(blcemi.ePendiente)
                    If Not a.ServicioEmergencia Is Nothing Then
                        If a.ServicioEmergencia.id = tel.OwnerId Then
                            lblRegistrarAtencion.Caption = "Modificar Atencion"
                            Set mAtencion = a
                            Exit For
                        End If
                    End If
                Next
            Case eOwnerType.eOTEmpleado
                'ver
            Case eOwnerType.eOTAfiliadoExterno
                lblTipo = "Afiliado Externo"
                Dim afs As blcemi.AfiliadoExternoManager
                Set afs = New blcemi.AfiliadoExternoManager
                lblNombre = afs.LoadById(tel.OwnerId).NombreCompleto
                For Each a In GBL.AtencionesGBL.GetByEstado(blcemi.ePendiente)
                    If Not a.AfiliadoExterno Is Nothing Then
                        If a.AfiliadoExterno.id = tel.OwnerId Then
                            lblRegistrarAtencion.Caption = "Modificar Atencion"
                            Set mAtencion = a
                            Exit For
                        End If
                    End If
                Next
        End Select
    Else
        'numero desconocido
        lblTelefono = pTelNumber
        lblTipo = "Desconocido"
        lblNombre = "Desconocido"
        lblRegistrarAtencion.Caption = "Registrar Atencion"
    End If
        
End Sub

Private Sub BlanquearBotones()
    lblCancelar.ForeColor = vbBlue
    lblRegistrarAtencion.ForeColor = vbBlue
End Sub

Private Sub lblRegistrarAtencion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblRegistrarAtencion.ForeColor = vbBlue + 200
End Sub

Private Sub lblCancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCancelar.ForeColor = vbBlue + 200
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BlanquearBotones
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'QueryUnload method
    'vbFormCode         1   Unload method invoked from code.
    'vbAppWindows       2   Current Windows session ending.
    'vbFormMDIForm      4   MDI child form is closing because the MDI form is closing.
    'vbFormControlMenu  0   User has chosen Close command from the Control-menu box on a form.
    'vbAppTaskManager   3   Windows Task Manager is closing the application.

    'no permitir cerrar desde la X
    If UnloadMode = vbFormControlMenu Then
        Cancel = True
    End If

End Sub

Public Sub mnuSalir_Click()
    SuperSalir
End Sub

Private Sub SuperSalir(Optional NotMe As Boolean = False)

    If MsgBox("Esta seguro que desea salir del Sistema?", vbOKCancel + vbQuestion) = vbOK Then
    
        Unload frmAyuda 'estaba en mdiFrm_unload
        
        Set mFrmInicioSesion = Nothing
        Set miConfig = Nothing
        Set mDBMonitor = Nothing
        Set mErrHandler = Nothing
        Set mFrmAtencionesPendientes = Nothing
        Set mNetMonitor = Nothing
        Set clienteTel = Nothing
        Set tel = Nothing
        Set mAtencion = Nothing
        Set GBL = Nothing
        Set CCFFGG = Nothing
    
        If Forms.Count <> 1 Then '1 porq esta el formconsultaratencionespendientes en segundo plano
            'andres marz 2010
            'descargar todos los formularios
            Dim f As Form
            For Each f In Forms
                If f.hwnd <> Me.hwnd Then Unload f
            Next
        End If
    Else
        Cancel = True
        Exit Sub
    End If
    
    If NotMe = False Then Unload Me
    End
End Sub

Private Sub mErrHandler_Error(informe As String)
    
    'version que no entiendo de martin
    'Dim frmIE As New frmInformarError
    'frmIE.InformarError informe
    
    frmInformarError2.Show 1
    
End Sub

Private Sub mnCreteReg_Click()
    frmInformarError2.Show 1
End Sub

Private Sub mnCreateReg_Click()
    frmInformarError2.Show 1
End Sub

Private Sub mNetMonitor_AtencionesChanged()
    On Error Resume Next
    GBL.ResetearColecciones
    MDI.ActiveForm.Refrescar
End Sub

Private Sub mnGenerarLic_Click()
    'VERIFICACION DE LICENCIA
    Dim cm As New CommonDialog
    cm.InitDir = APh
    
    cm.ShowFolder
    
    Dim f As String
    f = cm.SelectedDir
    
    If f = "" Then Exit Sub
    If Right(f, 1) <> "\" Then f = f + "\"
    
    Dim TD As New tbrDATA.clsTODO, j As Long
    TD.SetLog APh + "loglic.txt"
    TD.SetSF "tbrEG_v2"
    
    j = TD.DoNow(f + "Licencia_TbrEG.LIC")
    
    MsgBox "Archivo generado.", vbInformation
End Sub

Private Sub mnInsertLic_Click()
    Dim cm As New CommonDialog
    cm.ShowOpen
    
    Dim f As String
    f = cm.FileName
    
    If f = "" Then Exit Sub
    Dim FS As New Scripting.FileSystemObject
    'borro la licencia que hay si la hay ...
    If FS.FileExists(APh + "lic.insertada") Then FS.DeleteFile APh + "lic.insertada", False
    'y grabo la nueva
    FS.CopyFile f, APh + "lic.insertada", True
    VerificarLicencia True
End Sub
  
Private Sub picLlamada_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BlanquearBotones
End Sub

'-------------hasta aca lo de las llamadas-------------------------

'recibo eventos desde la db
Private Sub mDBMonitor_Error(eError As blcemi.eDBErrors, pDescription As String)
Select Case eError
    Case eDBErrors.eDBCantFindDB
        If CCFFGG.Configuracion.Red.ModoServer Then
            MsgBox "No se puede encontrar la Base de Datos. Ejecute el reparador."
            Unload Me
        Else
            MsgBox "No se puede encontrar la Base de Datos. Comuniquese con el responsable del sistema."
            Unload Me
        End If
        
    Case eDBErrors.eDBConnectionClosed
        frmWait.Show vbModal
    
    Case Else
        MsgBox "Ocurrio un error desconocido."
        
End Select
End Sub

Private Sub MDIForm_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If X < Me.Width And X > Me.Width - 500 Then
        tBar.Align = 4
        Exit Sub
    End If
    
    If X > 0 And X < 500 Then
        tBar.Align = 3
        Exit Sub
    End If
    
    If Y > 0 And Y < 500 Then
        tBar.Align = 1
        Exit Sub
    End If
    
    If Y < Me.Height And Y > Me.ScaleHeight - 500 Then
        tBar.Align = 2
        Exit Sub
    End If

End Sub


Private Sub MDIForm_Resize()
If Me.WindowState = vbMaximized Or Me.WindowState = vbNormal Then
'    If Me.Width < 9000 Then Me.Width = 9000 'si dejo esto queda en un ciclo infinito
'    If Me.Height < 7000 Then Me.Width = 7000
    lblRegistrarAtencion.Left = Me.Width - lblRegistrarAtencion.Width
    lblCancelar.Left = Me.Width - lblCancelar.Width
    lblNombre.Left = (Me.Width - lblNombre.Width) / 2
    lblTipo.Left = lblNombre.Left
    picAlerta.Left = Me.Width - picAlerta.Width - 100
End If
End Sub

Private Sub mFrmAtencionesPendientes_AtencionesModificadas(pCantidadAtencionesPendientes As Integer)
'aca me avisa si se agrego o modifico una atencion,
If pCantidadAtencionesPendientes <> 0 Then
    sBar.Panels(2).Text = "Hay " + Trim(Str(pCantidadAtencionesPendientes)) + " atencion" + IIf(pCantidadAtencionesPendientes = 1, " pendiente.", "es pendientes.")
    tBar.Buttons("consultaratenciones").Image = "atencionesalerta"
    Timer1.Enabled = True
Else
    sBar.Panels(2).Text = ""
    tBar.Buttons("consultaratenciones").Image = "consultaratenciones"
    Timer1.Enabled = False
    Set picAlerta.Picture = LoadPicture("")
End If

End Sub

Private Sub mGlobal_Error(pNumber As Long, pDescription As String)
    On Error Resume Next
    TERR.AppendLog "globalError", pDescription
    MsgBox "descripcion: " + pDescription
End Sub

Private Sub miConfig_ConfigChanged()
    mnuActualizar_Click
    If CCFFGG.Configuracion.Comportamiento.MostrarBarraMenu Then
        frmBarraMenu.Show
    Else
        frmBarraMenu.Hide
    End If
    On Error GoTo errman
    Set MDI.Picture = LoadPicture(CCFFGG.Configuracion.Apariencia.PathFondo)
    Exit Sub
errman:
    Set MDI.Picture = LoadPicture()
End Sub

Private Sub mNetMonitor_NetStatusChanged(pState As String)
    sBar.Panels(4).Text = "Red - Modo: " + IIf(mNetMonitor.MiRedLocal.modo = 1, "Cliente", "Servidor") + " - Estado: " + pState
End Sub

Private Sub mNetMonitor_PedirNombreUsuario(pNombreUsuario As String)
    If Not UsuarioActual Is Nothing Then
        pNombreUsuario = UsuarioActual.NombreCompleto
    Else
        pNombreUsuario = "-"
    End If
End Sub

Public Sub mnuActualizar_Click()
    On Error Resume Next
    Dim id As Long
    id = UsuarioActual.id
    GBL.ResetearColecciones
    Set UsuarioActual = GBL.EmpleadosGBL.Item(id)
    MDI.ActiveForm.Refrescar
End Sub

Private Sub mnuConfiguracion_Click()
    frmConfiguracion.Show
End Sub

Private Sub mnuEstadoRed_Click()
    GBL.MostrarEstadoRed
End Sub

Private Sub mnuSetearUltimoRecibo_Click()
    MsgBox "Debe tener mucho cuidado al utilizar esta funcion ya que se pueden dar inconsistencias en la base de datos.", vbInformation + vbOKOnly
    Dim actual As Long
    Dim res As String
    actual = GBL.CuotasByEstadoGBL(blcemi.eAnulado).GetNumeroDeReciboActual
    res = InputBox("Ingrese el numero del ultimo recibo emitido.", , actual)
    If IsNumeric(res) Then
        GBL.CuotasByEstadoGBL(blcemi.eAnulado).SetNumeroDeReciboActual CLng(res)
        MsgBox "Se guardo " + Trim(Str(res)) + " como el ultimo numero de recibo.", vbInformation + vbOKOnly
    End If
End Sub

Private Sub mnuVerMenu_Click()
    frmBarraMenu.Show
End Sub

Private Sub tBar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static xx As Single
    Static yy As Single
    If Button = vbLeftButton Then
        If Abs(xx - X) > 15 And Abs(yy - Y) > 15 Then tBar.Drag
    End If
    xx = X
    yy = Y
    BlanquearBotones 'pone los botones del picllamada en el color default
End Sub

Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "afiliados"
            mnuAfiliados_Click
        Case "areasprotegidas"
            mnuAreasProtegidas_Click
        Case "empleados"
            mnuEmpleados_Click
        Case "registraratencion"
            mnuRegistrarAtencion_Click
        Case "iniciosesion"
           mnuInicioSesion_Click
           ' mFrmInicioSesion_SesionIniciada UsuarioActual 'sacar esta linea
        Case "consultaratenciones"
            mnuConsultarAtencionesPendientes_Click
        Case "moviles"
            mnuMovil_Click
        Case "dotaciones"
            mnuDotaciones_Click
        Case "obrasocial"
            mnuObrasSociales_Click
        Case "servicioemergencia"
            mnuServiciosEmergencia_Click
    End Select
End Sub

Private Sub DesabilitarMenues()

    On Error Resume Next
    
    Dim c As Control
    
    For Each c In Me.Controls
        If TypeOf c Is Menu Then If c.Caption <> "-" Then c.Enabled = False
    Next
    
    Dim b As Button
    For Each b In tBar.Buttons
        If b.Style = tbrDefault Then b.Enabled = False
    Next
    
    'el boton de registro de errores siempre activo
    mnCreateReg.Enabled = True
    
    tBar.Buttons("iniciosesion").Enabled = True
        mnuArchivo.Enabled = True
        mnuInicioSesion.Enabled = True
        mnuSalir.Enabled = True
        mnuVentana.Enabled = True
        mnuHerramientas.Enabled = True
        mnLicencia.Enabled = True
        mnGenerarLic.Enabled = True
        mnInsertLic.Enabled = True
        mnuAyuda.Enabled = True
        mnuAcercaDe.Enabled = True
        mnuManual.Enabled = True
        
End Sub

Private Sub MDIForm_Load()
    
    On Local Error GoTo errMDI
    TERR.Anotar "argd"
    
    Me.Caption = "tbrSoft Emergency group v" + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    
    'Set mDBMonitor = New BLCemi.DBMonitor
    Set mDBMonitor = GBL.DBMonitorGBL 'imagino que quiere este que es contolado y no uno nuevo
    
    TERR.Anotar "arge"
    Set mErrHandler = GBL.ErrHandlerGBL
    
    TERR.Anotar "argf"
    frmSplash.Mensaje = "Conectando con grabador de llamadas..."
    
    TERR.Anotar "argg"
    ConectarConGrabadorLlamadas
    
    TERR.Anotar "argh", modoSoftware
    On Error Resume Next
    Select Case modoSoftware
         Case eModoFuncionamiento.eMFBomberos:
            Set tBar.ImageList = ilBarraB
         Case eModoFuncionamiento.eMFEmergencia:
            Set tBar.ImageList = ilBarra
    End Select
    
    TERR.Anotar "argi"
    Dim b As MSComctlLib.Button
    
    TERR.Anotar "argj"
    For Each b In tBar.Buttons
        TERR.Anotar "argk", b.Index
        If b.Style <> MSComctlLib.tbrSeparator Then b.Image = b.Key
    Next
    
    TERR.Anotar "argl"
    DesabilitarMenues
    mnuInicioSesion.Tag = "iniciar"
    
    TERR.Anotar "argm", modoSoftware
    Set miConfig = Configuracion
    'fijarse si esto es lo mejor
    Select Case modoSoftware
        Case eModoFuncionamiento.eMFBomberos:
            mFrmAtencionesPendientes_AtencionesModificadas (GBL.AtencionesBGBL.GetByEstado(blcemi.ePendiente).Count)
        Case eModoFuncionamiento.eMFEmergencia:
            mFrmAtencionesPendientes_AtencionesModificadas (GBL.AtencionesGBL.GetByEstado(blcemi.ePendiente).Count)
    End Select
    
    TERR.Anotar "argn"
    frmSplash.Mensaje = "Iniciando red..."
    Set mNetMonitor = GBL.NetMonitorGBL 'para q me avise el estado de la red
    If CCFFGG.Configuracion.Comportamiento.MostrarBarraMenu Then
        TERR.Anotar "argo"
        frmBarraMenu.Show
    Else
        TERR.Anotar "argp"
        frmBarraMenu.Hide
    End If
    
    TERR.Anotar "argq"
    picAlerta.BackColor = vbButtonFace
    
    Dim cantAt As Long
    Dim cantC As Long
    TERR.Anotar "argr"
    cantAt = GBL.GetCantidadRegistros("atencion")
    cantC = GBL.GetCantidadRegistros("cuota")
    TERR.Anotar "args", cantAt, cantC, modo
    If modo = eModoDemo Then
        If (cantAt > 700 Or cantC > 400) Then
            TERR.Anotar "argt"
            frmAbout.Show
        End If
    End If
    
    TERR.Anotar "argu"
    frmSplash.Mensaje = "Cargando fondo..."
    
    TERR.Anotar "argv"
    Set MDI.Picture = LoadPicture(CCFFGG.Configuracion.Apariencia.PathFondo)
    frmSplash.Mensaje = "Conectando..."
    'por si no se conecto...
    TERR.Anotar "argw"
    mNetMonitor.ForzarConexion
    frmSplash.Mensaje = "Iniciando el programa..."
    'necesario para la ayuda
    
    TERR.Anotar "argx"
    Load frmAyuda
    
    TERR.Anotar "argy"
    Exit Sub
    
errMDI:
    TERR.AppendLog "errMDI", TERR.ErrToTXT(Err)
    Resume Next
End Sub

Private Sub mFrmInicioSesion_SesionIniciada(pUsuarioLogueado As blcemi.Empleado)
    Set UsuarioActual = pUsuarioLogueado
    AplicarPermisos
    mnuInicioSesion.Caption = "Cerrar Sesion"
    mnuInicioSesion.Tag = "cerrar"
    tBar.Buttons("iniciosesion").Image = "cerrarsesion"
    sBar.Panels(1).Text = "Usuario=" + UsuarioActual.Login
End Sub

Public Sub mnuInicioSesion_Click()
    If mnuInicioSesion.Tag = "iniciar" Then
        MostrarIniciarSesion
    Else
        Set UsuarioActual = Nothing
        DesabilitarMenues
        ActualizarFrmBarra
        
        mnuInicioSesion.Caption = "Iniciar Sesion"
        mnuInicioSesion.Tag = "iniciar"
        tBar.Buttons("iniciosesion").Image = "iniciosesion"
        sBar.Panels(1).Text = "No se ha iniciado sesion."
    End If
End Sub

Private Sub ActualizarFrmBarra()
    'me fijo si esta cargado
    For X = 0 To Forms.Count - 1
        If LCase(Forms(X).Name) = LCase("frmBarraMenu") Then
            frmBarraMenu.Actualizar
            Exit For
        End If
    Next
End Sub

Public Sub MostrarIniciarSesion()
    Set mFrmInicioSesion = New frmInicioSesion
    On Error Resume Next 'esto es porq en el primer uso me larga un error cuando trato de descargarlo al iniciosesion
    If modo = eVersionRegistrada Or modo = eModoDemo Then mFrmInicioSesion.Show
End Sub

Private Sub AplicarPermisos()
    With UsuarioActual.Permisos
         
        Select Case modoSoftware
         Case eModoFuncionamiento.eMFBomberos:
            'oculto los menues que no voy a usar
            mnuAfiliados.Visible = False
            mnuAreasProtegidas.Visible = False
            mnuObrasSociales.Visible = False
            mnuServiciosEmergencia.Visible = False
            mnuSep2.Visible = False  'para q quede mejor
            mnusep5.Visible = False  'para q quede mejor
            mnusep7.Visible = False  'para q quede mejor
            
            mnuOcupaciones.Visible = False
            mnuParentezcos.Visible = False
            mnuAlergias.Visible = False
            mnuEnfermedades.Visible = False
            mnuMedicamentos.Visible = False
            mnuTiposCodigo.Visible = False
                
            tBar.Buttons("afiliados").Visible = False
            tBar.Buttons("areasprotegidas").Visible = False
            tBar.Buttons("obrasocial").Visible = False
            tBar.Buttons("servicioemergencia").Visible = False
            tBar.Buttons("sep1").Visible = False 'para q quede mejor
            tBar.Buttons("sep2").Visible = False 'para q quede mejor
            
            mnuLiqEmpresas.Visible = False
            mnuAdministracion.Visible = False
            mnuSetearUltimoRecibo.Visible = False
                        
            'habilito los menues propios de los bomberos
            'adapto los nombre de los menues
            mnuAtencion.Caption = "Siniestro"
            mnuRegistrarAtencion.Caption = "Registrar Siniestro"
            mnuConsultarAtencionesPendientes.Caption = "Consultar Siniestros Pendientes"
            mnuListadoAtenciones.Caption = "Listado de Siniestros"
            
            tBar.Buttons("registraratencion").ToolTipText = "Registrar Siniestro"
            tBar.Buttons("consultaratenciones").ToolTipText = "Consultar Siniestros Pendientes"
            mnuInstElectrica.Enabled = True
            mnuInstGas.Enabled = True
            'ver permisos
            mnuCuerpos.Enabled = True
         
         Case eModoFuncionamiento.eMFEmergencia:
            'oculto los menues que no voy a usar
            mnuSep14.Visible = False
            mnuSep15.Visible = False
            mnuInstElectrica.Visible = False
            mnuInstGas.Visible = False
            mnuCuerpos.Visible = False
            
            'habilito los menues propios de las centrales de emergencias
            mnuAfiliados.Enabled = .Can(blcemi.ConsultarAfiliados)
            mnuAreasProtegidas.Enabled = .Can(blcemi.ConsultarAreaProtegida)
            mnuObrasSociales.Enabled = .Can(blcemi.ConsultarObraSocial)
            mnuServiciosEmergencia.Enabled = .Can(blcemi.ConsultarServicioEmergencia)
            
            'habilito las consultas siempre
            mnuOcupaciones.Enabled = True
            mnuParentezcos.Enabled = True
            mnuAlergias.Enabled = True
            mnuEnfermedades.Enabled = True
            mnuMedicamentos.Enabled = True
            mnuTiposCodigo.Enabled = True
                
            tBar.Buttons("areasprotegidas").Enabled = .Can(blcemi.ConsultarAreaProtegida)
            tBar.Buttons("obrasocial").Enabled = .Can(blcemi.ConsultarObraSocial)
            tBar.Buttons("servicioemergencia").Enabled = .Can(blcemi.ConsultarServicioEmergencia)
            mnuLiqEmpresas.Enabled = .Can(blcemi.ConsultarLiquidacionEmpresas)
        
        End Select
                
        'habilito todos los comunes a cualquier empresa
        
        'mnuArchivo.Enabled = True
        ActualizarFrmBarra
        
        mnuEmpleados.Enabled = .Can(blcemi.ConsultarEmpleado)
        mnuMovil.Enabled = .Can(blcemi.ConsultarMovil)
        mnuDotaciones.Enabled = .Can(blcemi.ConsultarEquipo)
        
        mnuAtencion.Enabled = True
              
        lblRegistrarAtencion.Enabled = .Can(blcemi.AltaAtencion) 'boton en picllamadas
               
        mnuRegistrarAtencion.Enabled = .Can(blcemi.AltaAtencion)
        
        mnuConsultarAtencionesPendientes.Enabled = .Can(blcemi.ConsultarAtencion)
        mnuListadoAtenciones.Enabled = .Can(blcemi.ConsultarAtencion)
        
        mnuAdministracion.Enabled = True
               
        mnuListadoCuotasACobrar.Enabled = .Can(blcemi.EmitirListadoPagos)
        mnuRegistrarRecibosAnulados.Enabled = .Can(blcemi.RegistrarDevolucionRecibosAnulados)
               
        'barra de accesos directos
        tBar.Buttons("afiliados").Enabled = .Can(blcemi.ConsultarAfiliados)
        tBar.Buttons("empleados").Enabled = .Can(blcemi.ConsultarEmpleado)
        tBar.Buttons("registraratencion").Enabled = .Can(blcemi.AltaAtencion)
        tBar.Buttons("consultaratenciones").Enabled = .Can(blcemi.ConsultarAtencion)
        tBar.Buttons("moviles").Enabled = .Can(blcemi.ConsultarMovil)
        tBar.Buttons("dotaciones").Enabled = .Can(blcemi.ConsultarEquipo)
                
        mnuLiqEmpleados.Enabled = .Can(blcemi.ConsultarLiquidacionEmpleado)
        
        
        '--------VER PERMISOS!!!-------
        mnuReportes.Enabled = True
        mnuInformes.Enabled = True
        
        mnuMantenimiento.Enabled = True
        'habilito estas consultas siempre
        mnuCargos.Enabled = True
       
        mnuTipoTelefono.Enabled = True
        
        
        mnuVer.Enabled = True
        mnuVerMenu.Enabled = True
        mnuActualizar.Enabled = True
        
        mnuHerramientas.Enabled = True
        mnuCamara.Enabled = True
        mnuImportar.Enabled = True 'ver, estan verificados los permisos pero por las dudas...
        mnuEstadoRed.Enabled = True 'VER, de todas formas no hay mucho q se pueda hacer aca
        mnLicencia.Enabled = True
        mnGenerarLic.Enabled = True
        mnInsertLic.Enabled = True
        
        mnuConfiguracion.Enabled = True
        mnuCambioPass.Enabled = True
        mnuSetearUltimoRecibo.Enabled = .Can(blcemi.SetearNumeroRecibo)
        
        mnuContenido.Enabled = True
    End With
End Sub

Public Sub SetStatusBarText(texto As String)
    sBar.Panels(3).Text = texto
End Sub

Private Sub Timer1_Timer()
'hacer titilar el aviso de atenciones pendientes
    Static b As Boolean
    If CCFFGG.Configuracion.Comportamiento.MostrarAvisoAtencionesPendientes Then
        If b Then
            Set picAlerta.Picture = ilBarra.ListImages("alerta").ExtractIcon
        Else
            Set picAlerta.Picture = LoadPicture("")
        End If
    Else
        Set picAlerta.Picture = LoadPicture("")
    End If
    b = Not b
End Sub

'------------------------MENUES------------------------------
Public Sub mnuContenido_Click()
    On Error GoTo errman
    Dim context As String
    context = MDI.ActiveForm.GetHelpContext
    Dim h As Long
    h = HtmlHelp(hWndAyudaHTML, APh + "Ayuda.chm" + "::/" + context + ".htm", HH_DISPLAY_TOPIC, 0&)
    Exit Sub
errman:
'si hay algun problema muestro la basica...
    h = HtmlHelp(hWndAyudaHTML, APh + "Ayuda.chm", HH_DISPLAY_TOPIC, 0&)
End Sub

'con el sistema de ayuda quedo obsoleto...
Private Sub mnuManual_Click()
   On Error GoTo errman
    Ruta = Chr(32) + APh + "..\manual\manual de usuario.doc" + Chr(32)
    Dim a
    Set a = CreateObject("word.application")
    a.Visible = True
    a.Documents.open Ruta, , True
    Set a = Nothing
    Exit Sub
errman:
    MsgBox "No se encuentra el manual de usuario. Intente buscando en la carpeta manual dentro del directorio principal de la aplicacion.", vbInformation
End Sub

Private Sub mnuImportar_Click()
    frmImportar.Show
End Sub

Private Sub mnuCambioPass_Click()
    frmCambioPass.Show
End Sub

Private Sub mnuAcercaDe_Click()
    frmAbout.Show
End Sub

Public Sub mnuDotaciones_Click()
    frmConsultarDotaciones.Consultar GBL.EquiposGBL, etSinRetorno
End Sub

Public Sub mnuListadoAtenciones_Click()
     Select Case modoSoftware
         Case eModoFuncionamiento.eMFBomberos:
            frmFiltroAtencionesB.Show
         Case eModoFuncionamiento.eMFEmergencia:
            frmFiltroAtenciones.Show
    End Select
End Sub

Private Sub mnuListadoCuotasACobrar_Click()
    frmEmitirListadoCobros.Show
End Sub

Private Sub mnuRegistrarRecibosAnulados_Click()
    frmRegistrarRecibosAnulados.MostrarListadoRecibosAnulados Nothing
End Sub

Public Sub mnuAfiliados_Click()
    frmConsultarAfiliado.Consultar GBL.AfiliadosGBL.GetAfiliadosTitulares, etSinRetorno
End Sub

Public Sub mnuAreasProtegidas_Click()
    frmConsultarAreaProtegida.Consultar GBL.AreasProtegidasGBL, etSinRetorno
End Sub

Public Sub mnuConsultarAtencionesPendientes_Click()
    FrmAtencionesPendientes.Mostrar
End Sub

Public Sub mnuEmpleados_Click()
    frmConsultarEmpleado.Consultar GBL.EmpleadosGBL
End Sub

Public Sub mnuMovil_Click()
    frmConsultarMovil.Consultar GBL.MovilesGBL
End Sub

Public Sub mnuObrasSociales_Click()
    frmConsultarObraSocial.Consultar GBL.ObrasSocialesGBL
End Sub

Public Sub mnuRegistrarAtencion_Click()
        
    On Local Error GoTo ErrF6B
    TERR.Anotar "abaa", modoSoftware
    Select Case modoSoftware
         Case eModoFuncionamiento.eMFBomberos:
            TERR.Anotar "abab"
            Dim frmB As New frmAtencionBomberos
            TERR.Anotar "abad"
            frmB.NuevaAtencion FrmAtencionesPendientes
         Case eModoFuncionamiento.eMFEmergencia:
            TERR.Anotar "abac"
            Dim frm As New frmAtencion
            TERR.Anotar "abae"
            frm.NuevaAtencion FrmAtencionesPendientes
    End Select
    
    Exit Sub
    
ErrF6B:
    TERR.AppendLog "ErrF6b.a-", TERR.ErrToTXT(Err)
End Sub

Public Sub mnuServiciosEmergencia_Click()
    frmConsultarServiciosEmergencia.Consultar GBL.ServiciosEmergenciaGBL
End Sub

'-----------------------------menues de mantenimiento-----------------------
Private Sub mnuCuerpos_Click()
    Dim frmCuerpos As New frmConsultaCuerpos
    frmConsultaCuerpos.Consultar GBL.CuerposDeBomberosGBL, etSinRetorno
End Sub

Private Sub mnuInstElectrica_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.InstElectricasGBL, "Consulta de Tipos de Instalaciones Electricas", "Nuevo Tipo de Instalacion Electrica", True, True, True, , etSinRetorno  'UsuarioActual.Permisos.Can(AltaCargo), UsuarioActual.Permisos.Can(ModificacionCargo), UsuarioActual.Permisos.Can(BajaCargo), , etSinRetorno
End Sub

Private Sub mnuInstGas_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.InstalacionesGasGBL, "Consulta de Tipos de Instalaciones de Gas", "Nuevo Tipo de Instalacion de Gas", True, True, True, , etSinRetorno 'UsuarioActual.Permisos.Can(AltaCargo), UsuarioActual.Permisos.Can(ModificacionCargo), UsuarioActual.Permisos.Can(BajaCargo), , etSinRetorno
End Sub

Public Sub mnuCargos_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.CargosGBL, "Consulta de Cargos", "Nuevo Cargo", UsuarioActual.Permisos.Can(blcemi.AltaCargo), UsuarioActual.Permisos.Can(blcemi.ModificacionCargo), UsuarioActual.Permisos.Can(blcemi.BajaCargo), , etSinRetorno
End Sub

Public Sub mnuMedicamentos_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.MedicamentosGBL, "Consulta de Medicamentos", "Nuevo Medicamento", UsuarioActual.Permisos.Can(blcemi.AltaMedicamento), UsuarioActual.Permisos.Can(blcemi.ModificacionMedicamento), UsuarioActual.Permisos.Can(blcemi.BajaMedicamento), , etSinRetorno
End Sub

Public Sub mnuEnfermedades_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.EnfermedadesGBL, "Consulta de Enfermedades", "Nueva Enfermedad", UsuarioActual.Permisos.Can(blcemi.AltaEnfermedad), UsuarioActual.Permisos.Can(blcemi.ModificacionEnfermedad), UsuarioActual.Permisos.Can(blcemi.BajaEnfermedad), , etSinRetorno
End Sub

Public Sub mnuAlergias_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.AlergiasGBL, "Consulta de Alergias", "Nueva Alergia", UsuarioActual.Permisos.Can(blcemi.AltaAlergia), UsuarioActual.Permisos.Can(blcemi.ModificacionAlergia), UsuarioActual.Permisos.Can(blcemi.BajaAlergia), , etSinRetorno
End Sub

Public Sub mnuParentezcos_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.ParentezcosGBL, "Consulta de Parentezcos", "Nuevo Parentezco", UsuarioActual.Permisos.Can(blcemi.AltaParentezco), UsuarioActual.Permisos.Can(blcemi.ModificacionParentezco), UsuarioActual.Permisos.Can(blcemi.BajaParentezco), , etSinRetorno
End Sub

Public Sub mnuOcupaciones_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.OcupacionesGBL, "Consulta de Ocupaciones", "Nueva Ocupacion", UsuarioActual.Permisos.Can(blcemi.AltaOcupacion), UsuarioActual.Permisos.Can(blcemi.ModificacionOcupacion), UsuarioActual.Permisos.Can(blcemi.BajaOcupacion), , etSinRetorno
End Sub

Public Sub mnuTipoTelefono_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.TiposTelefonoGBL, "Consulta de Tipos de Telefono", "Nuevo Tipo de Telefono", UsuarioActual.Permisos.Can(blcemi.AltaTipoTelefono), UsuarioActual.Permisos.Can(blcemi.ModificacionTipoTelefono), UsuarioActual.Permisos.Can(blcemi.BajaTipoTelefono), , etSinRetorno
End Sub

Private Sub mnuTiposCodigo_Click()
    Dim frm As New frmConsultaGenerico
    frm.Consultar GBL.TiposCodigoGBL, "Consulta de Tipos de Codigo de Emergencia", "Nuevo Tipo de Codigo", UsuarioActual.Permisos.Can(blcemi.AltaTipoCodigo), UsuarioActual.Permisos.Can(blcemi.ModificarTipoCodigo), UsuarioActual.Permisos.Can(blcemi.BajaTipoCodigo), , etSinRetorno
End Sub

Private Sub mnuLiqEmpleados_Click()
    frmConsultarLiqEmpleado.Show
End Sub

Private Sub mnuLiqEmpresas_Click()
    frmConsultarLiqEmpresa.Show
End Sub

Private Sub mnuInformes_Click()
    frmConsultaListado.Show
End Sub

Private Sub mnuCamara_Click()
    frmCamaraIP.Show
End Sub

