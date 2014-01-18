VERSION 5.00
Begin VB.Form frmABMGenerico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4305
   Begin VB.Frame fraDatos 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4095
      Begin VB.TextBox txtNombre 
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "frmABMGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Me.Hide
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Public Function Nuevo(TituloForm As String) As String
    Me.Caption = TituloForm
    Set Me.Icon = MDI.Icon
    Me.Show vbModal
    Nuevo = txtNombre.Text
    Unload Me
End Function

Public Sub Refrescar()

End Sub

