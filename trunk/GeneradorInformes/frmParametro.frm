VERSION 5.00
Begin VB.Form frmParametro 
   Caption         =   "Nuevo Parametro"
   ClientHeight    =   2010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form1"
   ScaleHeight     =   2010
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescripcion 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   1080
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtTipo 
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   720
      Width           =   3495
   End
   Begin VB.TextBox txtNombre 
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Nomn 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event NuevoParametro(par As LParameter)

Private Sub Command1_Click()
Dim lp As New LParameter
lp.Descripcion = txtDescripcion
lp.Nombre = txtNombre
lp.Tipo = txtTipo
RaiseEvent NuevoParametro(lp)
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
