VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMCuerpoBomberos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   5700
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5535
      Begin VB.Frame Frame3 
         Caption         =   "Unidades"
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   2640
         Width           =   5295
         Begin ControlesPOO.ListViewConsulta lvwUnidades 
            Height          =   1455
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2566
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   0   'False
            FullRowSelection=   -1  'True
            AutoDistribuirColumnas=   -1  'True
            AllowModify     =   0   'False
            ShowCheckBoxes  =   0   'False
            MultiSelect     =   0   'False
            CampoImage      =   ""
            NEncabezado0    =   "Nombre"
            MEncabezado0    =   "nombre"
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
         Begin ALCemi.GraphicButton cmdEliminarUnidad 
            Height          =   375
            Left            =   4800
            TabIndex        =   12
            ToolTipText     =   "Eliminar los datos del vehiculo seleccionado."
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdEditarUnidad 
            Height          =   375
            Left            =   4800
            TabIndex        =   13
            ToolTipText     =   "Modificar los datos del vehiculo afectado."
            Top             =   660
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdAgregarUnidad 
            Height          =   375
            Left            =   4800
            TabIndex        =   14
            ToolTipText     =   "Agregar un vehiculo afectado por el siniestro."
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Responsables"
         Height          =   1815
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   5295
         Begin ControlesPOO.ListViewConsulta lvwResponsables 
            Height          =   1455
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2566
            HideSelection   =   0   'False
            HideEncabezados =   0   'False
            GridLines       =   0   'False
            FullRowSelection=   -1  'True
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
         Begin ALCemi.GraphicButton cmdEliminarResponsable 
            Height          =   375
            Left            =   4800
            TabIndex        =   9
            ToolTipText     =   "Eliminar los datos del vehiculo seleccionado."
            Top             =   1080
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdEditarResponsable 
            Height          =   375
            Left            =   4800
            TabIndex        =   10
            ToolTipText     =   "Modificar los datos del vehiculo afectado."
            Top             =   660
            Visible         =   0   'False
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
         Begin ALCemi.GraphicButton cmdAgregarResponsable 
            Height          =   375
            Left            =   4800
            TabIndex        =   11
            ToolTipText     =   "Agregar un vehiculo afectado por el siniestro."
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
         End
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   4800
      Width           =   1935
   End
End
Attribute VB_Name = "frmABMCuerpoBomberos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar

Public Event NuevoCuerpo(pCuerpo As blcemi.CuerpoBomberos)
Public Event CuerpoModificado(pCuerpo As blcemi.CuerpoBomberos)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mCuerpo As blcemi.CuerpoBomberos
Private mCuerpos As blcemi.CuerpoBomberosManager

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Set mCuerpo = mCuerpos.Nuevo(txtNombre, lvwResponsables.Coleccion, lvwUnidades.Coleccion)
                RaiseEvent NuevoCuerpo(mCuerpo)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            mCuerpo.Nombre = txtNombre
            mCuerpo.GuardarModificaciones
            RaiseEvent CuerpoModificado(mCuerpo)
            Unload Me
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean
DatosCorrectos = True
End Function

Private Sub cmdAgregarResponsable_Click()
    Dim frmGen As frmABMGenerico
    Set frmGen = New frmABMGenerico
    Dim Nombre As String
    Nombre = frmGen.Nuevo("Agregar Responsable")
    'validar que no exista un item con el mismo nombre
    If Nombre <> "" Then
        Dim resp As New blcemi.ResponsableCuerpo
        resp.Nombre = Nombre
        lvwResponsables.Coleccion.AddItem resp
        lvwResponsables.Refresh
    End If
End Sub

Private Sub cmdAgregarUnidad_Click()
    Dim frmGen As frmABMGenerico
    Set frmGen = New frmABMGenerico
    Dim Nombre As String
    Nombre = frmGen.Nuevo("Agregar Unidad")
    'validar que no exista un item con el mismo nombre
    If Nombre <> "" Then
        Dim resp As New blcemi.UnidadCuerpo
        resp.Nombre = Nombre
        lvwUnidades.Coleccion.AddItem resp
        lvwUnidades.Refresh
    End If
End Sub

Private Sub cmdCancelar_Click()
'implementar
If Tipo = etMODIFICACION Then
    Set mCuerpo.Responsables = Nothing
    Set mCuerpo.Unidades = Nothing
End If
Unload Me
End Sub

Public Sub Nuevo(pCuerpos As blcemi.CuerpoBomberosManager)
'implementar
Tipo = etALTA
Set mCuerpos = pCuerpos
Set lvwResponsables.Coleccion = New blcemi.ResponsableCuerpoManager
Set lvwUnidades.Coleccion = New blcemi.UnidadCuerpoManager
Me.Show
Me.Caption = "Nuevo Cuerpo"
End Sub

Public Sub Modificar(pCuerpo As blcemi.CuerpoBomberos)
'implementar
Tipo = etMODIFICACION
Me.Show
Set mCuerpo = pCuerpo
Me.Caption = "Modificar Cuerpo"
txtNombre = mCuerpo.Nombre
Set lvwResponsables.Coleccion = mCuerpo.Responsables
Set lvwUnidades.Coleccion = mCuerpo.Unidades
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
Tipo = etBAJA
Me.Show
End Sub

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
    If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
    Set Me.Icon = MDI.Icon
    
    Set cmdAgregarResponsable.Picture = MDI.il32.ListImages("agregar").Picture
    Set cmdAgregarUnidad.Picture = MDI.il32.ListImages("agregar").Picture
    
    Set cmdEditarResponsable.Picture = MDI.il32.ListImages("modificar").Picture
    Set cmdEditarUnidad.Picture = MDI.il32.ListImages("modificar").Picture
    
    Set cmdEliminarResponsable.Picture = MDI.il32.ListImages("eliminar").Picture
    Set cmdEliminarUnidad.Picture = MDI.il32.ListImages("eliminar").Picture
    
End Sub

Public Function GetHelpContext() As String
    'habilitar cuando agregue estas paginas a la ayuda
    GetHelpContext = "" '"abmcuerpobomberos"
End Function

Public Sub Refrescar()

End Sub

