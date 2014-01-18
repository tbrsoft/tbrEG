VERSION 5.00
Begin VB.Form frmABMVehiculo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4305
   Begin VB.Frame fraDatos 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   960
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtPerjuicios 
         Height          =   885
         Left            =   960
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtTipo 
         Height          =   285
         Left            =   960
         MaxLength       =   30
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtPatente 
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtMarca 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label5 
         Caption         =   "Color:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Modelo:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Daños"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label N 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Patente:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Marca:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2235
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
End
Attribute VB_Name = "frmABMVehiculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Event NuevoVehiculo(pVehiculo As blcemi.Vehiculo)
Public Event VehiculoModificado(pVehiculo As blcemi.Vehiculo)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mVehiculo As blcemi.Vehiculo
Private mVehiculos As blcemi.VehiculoManager

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Set mVehiculo = mVehiculos.Nuevo(txtTipo, txtMarca, txtModelo, txtPatente, txtPerjuicios, txtColor)
                RaiseEvent NuevoVehiculo(mVehiculo)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            mVehiculo.Color = txtColor
            mVehiculo.Marca = txtMarca
            mVehiculo.Modelo = txtModelo
            mVehiculo.Perjuicios = txtPerjuicios
            mVehiculo.Patente = txtPatente
            mVehiculo.Tipo = txtTipo
            'mVehiculo.Update se guarda en aceptar de siniestro?? ver...
            RaiseEvent VehiculoModificado(mVehiculo)
            Unload Me
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean
DatosCorrectos = True
End Function

Private Sub cmdCancelar_Click()
'implementar
Unload Me
End Sub

Public Sub Nuevo(pVehiculos As blcemi.VehiculoManager)
'implementar
Tipo = etALTA
Set mVehiculos = pVehiculos
Me.Show
Me.Caption = "Nuevo Vehiculo Damnificado"
End Sub

Public Sub Modificar(pVehiculo As blcemi.Vehiculo)
'implementar
Tipo = etMODIFICACION
Me.Show
Set mVehiculo = pVehiculo
Me.Caption = "Modificar Datos de Vehiculo"
txtModelo = mVehiculo.Modelo
txtColor = mVehiculo.Color
txtPatente = mVehiculo.Patente
txtMarca = mVehiculo.Marca
txtPerjuicios = mVehiculo.Perjuicios
txtTipo = mVehiculo.Tipo

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

End Sub
Public Sub Refrescar()

End Sub

Private Sub txtPatente_KeyPress(KeyAscii As Integer)
    'hacer mas comprobaciones
    If KeyAscii = Asc(" ") Then KeyAscii = 0
End Sub

