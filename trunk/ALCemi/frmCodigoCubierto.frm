VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmCodigoCubierto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nuevo Codigo Cubierto"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4110
   Begin VB.TextBox txtServicio 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   960
      Width           =   735
   End
   Begin ControlesPOO.Combo cmbTipo 
      Height          =   315
      Left            =   1200
      TabIndex        =   7
      Top             =   600
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtCopago 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin ControlesPOO.Combo cmbCodigo 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      Enabled         =   -1  'True
   End
   Begin VB.Label lblServicio 
      Caption         =   "Servicio:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblTipo 
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblCopago 
      Caption         =   "Copago:"
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmCodigoCubierto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event ItemSeleccionado(pCodigoCubierto As blcemi.CodigoCubierto)

Private Sub cmdAceptar_Click()
    If cmbCodigo.SelectedItem Is Nothing Then
        MsgBox "Seleccione un Codigo de Emergencia.", vbExclamation
        Exit Sub
    End If
    If cmbTipo.Visible Then
        If cmbTipo.SelectedItem Is Nothing Then
            MsgBox "Seleccione un Tipo de Codigo de Emergencia.", vbExclamation
            Exit Sub
        End If
    End If
    If Not TextBoxValidado(txtCopago, eMoneda) Then
        MsgBox "Ingrese un valor en Copago.", vbExclamation
        Exit Sub
    End If
    If Not TextBoxValidado(txtServicio, eMoneda) Then
        MsgBox "Ingrese un valor en Servicio.", vbExclamation
        Exit Sub
    End If
    Dim cod As blcemi.CodigoCubierto
    Set cod = New blcemi.CodigoCubierto
    Set cod.TipoCodigo = cmbTipo.SelectedItem
    Set cod.Codigo = cmbCodigo.SelectedItem
    cod.Coseguro = CCur(Replace(txtCopago, ".", ","))
    cod.Servicio = CCur(Replace(txtServicio, ".", ","))
    RaiseEvent ItemSeleccionado(cod)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set cmbCodigo.Coleccion = GBL.CodigoEmergenciaGBL
    Set cmbTipo.Coleccion = GBL.TiposCodigoGBL
    Set Me.Icon = MDI.Icon
    If Not CCFFGG.Configuracion.Codigo.UtilizarTipos Then
        cmbTipo.Visible = False
        lblTipo.Visible = False
        txtCopago.Top = cmbTipo.Top
        lblCopago.Top = lblTipo.Top
        txtServicio.Top = cmbTipo.Top
        lblServicio.Top = cmbTipo.Top
    End If
End Sub

Public Sub NuevoCodigoCubierto(Optional pServicio As Currency = -1, Optional pCoseguro As Currency = -1)
    'lo quite lo de servicio y coseguro
'    If pServicio <> -1 Then
'        txtServicio = pServicio
'        txtServicio.Locked = True
'        txtServicio.ToolTipText = "Para poder modificar este campo destilde la opcion 'Utilizar valor fijo para los servicios'."
'    End If
'    If pCoseguro <> -1 Then
'        txtCopago = pCoseguro
'        txtCopago.Locked = True
'        txtCopago.ToolTipText = "Para poder modificar este campo destilde la opcion 'Utilizar valor fijo para los coseguros'."
'    End If
    Me.Show
End Sub

Private Sub txtCopago_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, True
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-empresas"
End Function

Public Sub Refrescar()
    cmbCodigo.Refresh
    cmbTipo.Refresh
End Sub

Private Sub txtServicio_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, True
End Sub
