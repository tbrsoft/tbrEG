VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistrarGuardia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Guardia"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5880
   Begin VB.TextBox txtObservaciones 
      Height          =   975
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtPlus 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "0"
      ToolTipText     =   "Premios, extras, etc..."
      Top             =   1680
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   503
      _Version        =   393216
      Format          =   47120385
      CurrentDate     =   39850
   End
   Begin VB.TextBox txtAdelanto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Text            =   "0"
      ToolTipText     =   "Es el dinero que ya se le pago en concepto de coseguros, etc.."
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Valor de la guardia."
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Observaciones:"
      Height          =   255
      Left            =   3120
      TabIndex        =   13
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   1560
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblEmpleado 
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
      Left            =   240
      TabIndex        =   11
      Top             =   120
      Width           =   2655
   End
   Begin VB.Line Line1 
      X1              =   1560
      X2              =   2880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label5 
      Caption         =   "Total:"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Plus:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Adelanto:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Monto:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
End
Attribute VB_Name = "frmRegistrarGuardia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mEmpleado As blcemi.Empleado

Private Sub cmdAceptar_Click()
    If DatosCompletos Then
        mEmpleado.Guardias.Nueva mEmpleado.id, CCur(Replace(txtMonto, ".", ",")), CCur(Replace(txtAdelanto, ".", ",")), dtpFecha.Value, CCur(Replace(txtPlus, ".", ",")), txtObservaciones
        Unload Me
    End If
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "empleados"
End Function

Private Function DatosCompletos() As Boolean
Dim msj As String

If Not TextBoxValidado(txtAdelanto, eMoneda) Then msj = msj + "Ingrese un valor en Adelanto" + vbCrLf
If Not TextBoxValidado(txtMonto, eMoneda) Then msj = msj + "Ingrese un Monto" + vbCrLf
If Not TextBoxValidado(txtPlus, eMoneda) Then msj = msj + "Ingrese un valor en Plus." + vbCrLf

If msj = "" Then
    DatosCompletos = True
Else
    MsgBox "Faltan los siguientes datos:" + vbCrLf + msj, vbExclamation
    DatosCompletos = False
End If
End Function

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Public Sub NuevaGuardia(pEmpleado As blcemi.Empleado)
    Set mEmpleado = pEmpleado
    lblEmpleado = mEmpleado.NombreCompleto
    dtpFecha.Value = Date
End Sub

Private Sub MostrarSaldo()
    On Error GoTo errman
    lblTotal = CCur(Replace(txtMonto, ".", ",")) - CCur(Replace(txtAdelanto, ".", ",")) + CCur(Replace(txtPlus, ".", ","))
    Exit Sub
errman:
    lblTotal = "-"
End Sub

Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
End Sub

Private Sub txtAdelanto_Change()
    MostrarSaldo
End Sub

Private Sub txtMonto_Change()
    MostrarSaldo
End Sub

Private Sub txtPlus_Change()
    MostrarSaldo
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub

Private Sub txtAdelanto_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub

Private Sub txtPlus_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub


