VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmRegistrarCobro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Cobro"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   5355
   Begin VB.Frame fraTotal 
      Height          =   495
      Left            =   120
      TabIndex        =   33
      Top             =   5760
      Width           =   5175
      Begin VB.Label lblTotal 
         Alignment       =   2  'Center
         Caption         =   "0"
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
         Left            =   3795
         TabIndex        =   35
         Top             =   170
         Width           =   1305
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
         Height          =   195
         Left            =   3240
         TabIndex        =   34
         Top             =   170
         Width           =   405
      End
   End
   Begin VB.Frame fraAreaProtegida 
      Caption         =   "Area Protegida"
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
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
         TabIndex        =   12
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblDatosArea 
         Alignment       =   1  'Right Justify
         Caption         =   "Apellido y nombre:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         ToolTipText     =   "Apellido y nombre del responsable del area"
         Top             =   600
         Width           =   1290
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
         TabIndex        =   10
         Top             =   960
         Width           =   3075
      End
      Begin VB.Label lblDatosArea 
         Alignment       =   1  'Right Justify
         Caption         =   "Nombre Area:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1320
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
         TabIndex        =   8
         Top             =   240
         Width           =   1380
      End
      Begin VB.Label lblDatosArea 
         Alignment       =   1  'Right Justify
         Caption         =   "DNI:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblDatosArea 
         Alignment       =   1  'Right Justify
         Caption         =   "Direccion:"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   6
         ToolTipText     =   "Si esta es la direccion de la emergencia, la puede arrastrar..."
         Top             =   1320
         Width           =   855
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
         TabIndex        =   5
         Top             =   1320
         Width           =   3180
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
         TabIndex        =   4
         Top             =   1600
         Width           =   3135
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   435
      Left            =   2880
      TabIndex        =   2
      Top             =   6360
      Width           =   2415
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Registrar Cobro"
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   6360
      Width           =   2415
   End
   Begin ControlesPOO.ListViewConsulta lvwCuotas 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   3960
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3201
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   -1  'True
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Mes"
      MEncabezado0    =   "Mes"
      AEncabezado0    =   20
      NEncabezado1    =   "Año"
      MEncabezado1    =   "aYear"
      AEncabezado1    =   20
      NEncabezado2    =   "Nro Recibo"
      MEncabezado2    =   "nrorecibo"
      AEncabezado2    =   30
      NEncabezado3    =   "Importe"
      MEncabezado3    =   "monto"
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
   Begin VB.Frame fraAfiliadoPropio 
      Caption         =   "Afiliado Propio"
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
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
         TabIndex        =   32
         Top             =   3120
         Width           =   3135
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
         TabIndex        =   31
         Top             =   2250
         Width           =   3000
      End
      Begin VB.Label Label27 
         Caption         =   "Obra Social:"
         Height          =   195
         Left            =   705
         TabIndex        =   30
         Top             =   2250
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado Civil:"
         Height          =   195
         Left            =   705
         TabIndex        =   29
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ocupacion:"
         Height          =   195
         Left            =   750
         TabIndex        =   28
         Top             =   2565
         Width           =   825
      End
      Begin VB.Label lblDatosAPropio 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº de afiliado:"
         Height          =   195
         Index           =   0
         Left            =   585
         TabIndex        =   27
         Top             =   360
         Width           =   990
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
         TabIndex        =   26
         Top             =   990
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "Apellido y Nombre:"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   25
         Top             =   675
         Width           =   1320
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
         TabIndex        =   24
         Top             =   675
         Width           =   3000
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Edad:"
         Height          =   195
         Left            =   1155
         TabIndex        =   23
         Top             =   1305
         Width           =   420
      End
      Begin VB.Label lblTipoDoc 
         Alignment       =   1  'Right Justify
         Caption         =   "DNI:"
         Height          =   195
         Left            =   1245
         TabIndex        =   22
         Top             =   990
         Width           =   330
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
         TabIndex        =   21
         Top             =   1305
         Width           =   3000
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
         TabIndex        =   20
         Top             =   1620
         Width           =   3000
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Sexo:"
         Height          =   195
         Left            =   1170
         TabIndex        =   19
         Top             =   1935
         Width           =   405
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
         TabIndex        =   18
         Top             =   1935
         Width           =   3000
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
         TabIndex        =   17
         Top             =   2565
         Width           =   3000
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   855
         TabIndex        =   16
         ToolTipText     =   "Si esta es la direccion de la emergencia, la puede arrastrar..."
         Top             =   2880
         Width           =   720
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
         TabIndex        =   15
         Top             =   2880
         Width           =   3000
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
         TabIndex        =   14
         Top             =   360
         Width           =   2520
      End
   End
End
Attribute VB_Name = "frmRegistrarCobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event CobroRegistrado()

Public Enum eTipoCliente
    eTCAfiliado = 1
    eTCAreaProtegida = 2
End Enum

Private mTipo As eTipoCliente
Private mAfiliado As blcemi.Afiliado
Private mAreaProtegida As blcemi.AreaProtegida
Private WithEvents mCuotas As blcemi.CuotaManager
Attribute mCuotas.VB_VarHelpID = -1

Private Sub cmdAceptar_Click()
    
    If MsgBox("Esta seguro que desea anular el recibo y registrar el pago?", vbQuestion + vbOKCancel) = vbOK Then
        Dim cuotasaux As New blcemi.CuotaManager
        Dim c As blcemi.Cuota
        Dim mNroRecibo As Long
        'les asigno el mismo nro recibo a todas porq las voy a imprimir en el mismo
        mNroRecibo = cuotasaux.GetUltimoNroRecibo + 1
        For Each c In lvwCuotas.CheckedItems
            c.Anular UsuarioActual, blcemi.ePedirRecibo
            c.RegistrarCobro UsuarioActual, mNroRecibo
            cuotasaux.AddItem c
        Next
        
        If mTipo = eTCAfiliado Then
            ImprimirReciboAfiliado cuotasaux, True
        Else
            ImprimirReciboAreaProtegida cuotasaux, True
        End If
        
        RaiseEvent CobroRegistrado
        Unload Me
    End If
    
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Public Sub RegistrarPagoAfiliado(pAfiliado As blcemi.Afiliado)
    
    mTipo = eTCAfiliado
    Set mAfiliado = pAfiliado
    fraAfiliadoPropio.Visible = True
    'Me.Show se muestra solo
    lblApeNomAfilPropio = mAfiliado.NombreCompleto
    lblNroDocAPropio = mAfiliado.NroDoc
    lblTipoDoc = mAfiliado.TipoDoc.Nombre
    lblDireccionaPropio = mAfiliado.Direccion.Calle + " " + mAfiliado.Direccion.Nro
    lblCiudadAPropio = mAfiliado.Direccion.GetBarrioCiudadProvincia
    lblNroAfiliado = mAfiliado.IdCompleto
    lblEdadAPropio = mAfiliado.Edad
    lblOcupacion = mAfiliado.Ocupacion.Nombre
    lblSexoAPropio = IIf(mAfiliado.Sexo = 1, "Masculino", "Femenino")
    lblEstadoCivil = mAfiliado.EstadoCivil.Nombre
         
    lblObraSocialAPropio = mAfiliado.ObraSocial.Nombre
    Ajustar
    Set mCuotas = mAfiliado.Cuotas.GetCuotasByEstado(blcemi.eImpaga)
    Set lvwCuotas.Coleccion = mCuotas
    cmdAceptar.Enabled = False

End Sub

Public Sub RegistrarPagoAreaProtegida(pAreaProtegida As blcemi.AreaProtegida)
    mTipo = eTCAreaProtegida
    Set mAreaProtegida = pAreaProtegida
    fraAreaProtegida.Visible = True
    'Me.Show se muestra solo
     
    lblNombreArea = mAreaProtegida.NombreArea
    lblResponsableArea = mAreaProtegida.NombreCompleto
    lblDocRespArea = mAreaProtegida.NroDocResp
    lblDireccionArea = mAreaProtegida.Direccion.Calle + " " + mAreaProtegida.Direccion.Nro
    lblCiudadArea = mAreaProtegida.Direccion.GetBarrioCiudadProvincia
    Ajustar
    Set mCuotas = mAreaProtegida.Cuotas.GetCuotasByEstado(blcemi.eImpaga)
    Set lvwCuotas.Coleccion = mCuotas
    cmdAceptar.Enabled = False
    
End Sub

Private Sub Ajustar()
If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
    If fraAfiliadoPropio.Visible Then
        lvwCuotas.Top = fraAfiliadoPropio.Height + fraAfiliadoPropio.Top + 100
    Else
        lvwCuotas.Top = fraAreaProtegida.Height + fraAreaProtegida.Top + 100
    End If
    fraTotal.Top = lvwCuotas.Top + lvwCuotas.Height + 100
    cmdAceptar.Top = fraTotal.Top + fraTotal.Height + 100
    cmdCancelar.Top = cmdAceptar.Top
    Me.Height = cmdAceptar.Top + cmdAceptar.Height + 400
End If
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-cobradores"
End Function

Public Sub Refrescar()
    If Not mAfiliado Is Nothing Then
        Set lvwCuotas.Coleccion = mAfiliado.Cuotas.GetCuotasByEstado(blcemi.eImpaga)
    ElseIf Not mAreaProtegida Is Nothing Then
        Set lvwCuotas.Coleccion = mAreaProtegida.Cuotas.GetCuotasByEstado(blcemi.eImpaga)
    End If
End Sub

Private Sub Form_Load()
    Dim cantAt As Long
    Dim cantC As Long
    cantAt = GBL.GetCantidadRegistros("atencion")
    cantC = GBL.GetCantidadRegistros("cuota")
    
    If modo = eModoDemo Then
        If (cantAt > 800 Or cantC > 500) Then
            MsgBox "La version demo ha caducado. Consulte el manual de usuario para registrar el presente software.", vbExclamation
            Unload Me
        End If
    End If
    Set Me.Icon = MDI.Icon

End Sub

Private Sub lvwCuotas_ItemCheck(Item As Object)
    cmdAceptar.Enabled = (lvwCuotas.CheckedItems.Count <> 0)
    Dim c As blcemi.Cuota
    Dim montoaux As Currency
    For Each c In lvwCuotas.CheckedItems
        montoaux = montoaux + c.Monto
    Next
    lblTotal = "$" + Trim(Str(montoaux))

End Sub

Private Sub mCuotas_HasChanged()
    Refrescar
End Sub
