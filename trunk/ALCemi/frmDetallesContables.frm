VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmDetallesContables 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informacion Contable"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   5880
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fórmula de cálculo de los servicios"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      Begin VB.OptionButton opt1 
         Caption         =   "(Servicio - Copago) x IVA"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         ToolTipText     =   "A los servicios se les resta el copago y sobre ese valor se calcula el IVA"
         Top             =   360
         Width           =   2535
      End
      Begin VB.OptionButton opt2 
         Caption         =   "Servicio x IVA"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         ToolTipText     =   "El IVA se calcula sobre el total del servicio."
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.TextBox txtIva 
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Text            =   "10,5"
      Top             =   240
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tarifas por Codigo de Emergencia"
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   5655
      Begin ControlesPOO.ListViewConsulta lvwCodigos 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   4048
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
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
         NEncabezado0    =   ""
         MEncabezado0    =   ""
         AEncabezado0    =   0
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
      Begin ALCemi.GraphicButton cmdQuitarCodigo 
         Height          =   495
         Left            =   5040
         TabIndex        =   7
         ToolTipText     =   "Eliminar el cargo seleccionado"
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
      Begin ALCemi.GraphicButton cmdAgregarCodigo 
         Height          =   495
         Left            =   5040
         TabIndex        =   8
         ToolTipText     =   "Agregar un codigo cubierto"
         Top             =   360
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   873
      End
   End
   Begin VB.Label Label3 
      Caption         =   "IVA:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmDetallesContables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event DatosModificados(pCodigos As blcemi.CodigoCubiertoManager)
Private WithEvents frmCodigoC As frmCodigoCubierto
Attribute frmCodigoC.VB_VarHelpID = -1
Private mInfoContable As blcemi.InfoContableEmp
Private mCodigos As blcemi.CodigoCubiertoManager


Public Sub MostrarDetalles(pInfoContable As blcemi.InfoContableEmp, pCodigos As blcemi.CodigoCubiertoManager, pTipo As eTipoAMB)
    Set mCodigos = New blcemi.CodigoCubiertoManager
    Dim cc As blcemi.CodigoCubierto
    For Each cc In pCodigos
        mCodigos.AddItem cc.Clone
    Next
    Set lvwCodigos.Coleccion = mCodigos
    Set mInfoContable = pInfoContable
    txtIva = mInfoContable.IVA
    opt1.Value = IIf(mInfoContable.Formula = eServicioMenosCopagoPorIVA, True, False)
    opt2.Value = Not opt1.Value
    If pTipo = etCONSULTA Then
        cmdAceptar.Visible = False
        cmdCancelar.Caption = "Cerrar"
        cmdAgregarCodigo.Enabled = False
        cmdQuitarCodigo.Enabled = False
        opt1.Enabled = False
        opt2.Enabled = False
    End If
    Me.Show
End Sub

Private Sub cmdAceptar_Click()
    mInfoContable.IVA = CCurrency(txtIva)
    mInfoContable.Formula = IIf(opt1.Value, blcemi.eFormula.eServicioMenosCopagoPorIVA, blcemi.eFormula.eServicioPorIVA)
    RaiseEvent DatosModificados(mCodigos)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon

Set cmdAgregarCodigo.Picture = MDI.il32.ListImages("agregar").Picture
Set cmdQuitarCodigo.Picture = MDI.il32.ListImages("eliminar").Picture
AplicarConfiguracion
End Sub

Private Sub cmdQuitarCodigo_Click()
    If Not lvwCodigos.SelectedItem Is Nothing Then
        lvwCodigos.Coleccion.Remove lvwCodigos.SelectedItem
        lvwCodigos.Refresh
    End If
End Sub

Private Sub cmdAgregarCodigo_Click()
    Set frmCodigoC = New frmCodigoCubierto
    'los -1 vienen de una funcion q no esta mas
    frmCodigoC.NuevoCodigoCubierto -1, -1
End Sub

Private Sub frmCodigoC_ItemSeleccionado(pCodigoCubierto As blcemi.CodigoCubierto)
    lvwCodigos.Coleccion.AddItem pCodigoCubierto
    lvwCodigos.Refresh
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-empresas"
End Function

Private Sub AplicarConfiguracion()
    lvwCodigos.Encabezados.Clear
    If CCFFGG.Configuracion.Codigo.CosegurosPorCodigo Then
        If CCFFGG.Configuracion.Codigo.UtilizarTipos Then
            lvwCodigos.Encabezados.Add "Codigo", "nombrecompuesto", 30
            lvwCodigos.Encabezados.Add "Tipo", "tipo", 30
            lvwCodigos.Encabezados.Add "Servicio", "servicio", 20
            lvwCodigos.Encabezados.Add "Coseguro", "coseguro", 20
        Else
            lvwCodigos.Encabezados.Add "Codigo", "nombrecompuesto", 50
            lvwCodigos.Encabezados.Add "Servicio", "servicio", 25
            lvwCodigos.Encabezados.Add "Coseguro", "coseguro", 25
        End If
    Else
        lvwCodigos.Encabezados.Add "Codigo", "nombrecompuesto", 75
        lvwCodigos.Encabezados.Add "Servicio", "servicio", 25
    End If
    lvwCodigos.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    lvwCodigos.Refresh
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, True
End Sub

'Private Sub txtCopago_Change()
'    If chkCopago.Value = Checked Then
'        If TextBoxValidado(txtCopago, eMoneda) Then
'            Dim cc As blcemi.CodigoCubierto
'            'HACER: Si tiene valores diferentes avisarle q se van a cambiar
'            For Each cc In lvwCodigos.Coleccion
'                cc.Coseguro = CCurrency(txtCopago)
'            Next
'            lvwCodigos.Refresh
'        End If
'    End If
'End Sub
'
'Private Sub txtServicio_Change()
'    If chkServicio.Value = Checked Then
'        If TextBoxValidado(txtServicio, eMoneda) Then
'            Dim cc As blcemi.CodigoCubierto
'            'HACER: Si tiene valores diferentes avisarle q se van a cambiar
'            For Each cc In lvwCodigos.Coleccion
'                cc.Servicio = CCurrency(txtServicio)
'            Next
'            lvwCodigos.Refresh
'        End If
'    End If
'End Sub
'Private Sub chkCopago_Click()
'    If chkCopago.Value = Checked Then
'        txtCopago_Change
'    End If
'    txtCopago.Enabled = (chkCopago.Value = Checked)
'End Sub
'
'Private Sub chkServicio_Click()
'    If chkServicio.Value = Checked Then
'        txtServicio_Change
'    End If
'    txtServicio.Enabled = (chkServicio.Value = Checked)
'End Sub
