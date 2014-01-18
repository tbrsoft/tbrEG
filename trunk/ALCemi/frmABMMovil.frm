VERSION 5.00
Begin VB.Form frmABMMovil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4185
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4185
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Frame fraDatos 
      Height          =   1575
      Left            =   50
      TabIndex        =   5
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtEstado 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "1"
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtPatente 
         Height          =   285
         Left            =   960
         MaxLength       =   6
         TabIndex        =   1
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Estado:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Patente:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label N 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmABMMovil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar

Public Event NuevoMovil(pMovil As blcemi.Movil)
Public Event MovilModificado(pMovil As blcemi.Movil)

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mMovil As blcemi.Movil
Private mMoviles As blcemi.MovilManager

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Set mMovil = mMoviles.Nuevo(txtNombre, txtPatente, blcemi.eDisponible)
                RaiseEvent NuevoMovil(mMovil)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
            mMovil.Estado = txtEstado
            mMovil.Nombre = txtNombre
            mMovil.Patente = txtPatente
            mMovil.Update
            RaiseEvent MovilModificado(mMovil)
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

Public Sub Nuevo(pMoviles As blcemi.MovilManager)
'implementar
Tipo = etALTA
Set mMoviles = pMoviles
Me.Show
Me.Caption = "Nuevo Movil"
End Sub

Public Sub Modificar(pMovil As blcemi.Movil)
'implementar
Tipo = etMODIFICACION
Me.Show
Set mMovil = pMovil
Me.Caption = "Modificar Movil"
txtEstado = mMovil.Estado
txtNombre = mMovil.Nombre
txtPatente = mMovil.Patente
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
