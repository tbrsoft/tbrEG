VERSION 5.00
Begin VB.Form frmCambioPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambiar la contraseña"
   ClientHeight    =   1785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.TextBox txtPass2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Confirme la nueva contraseña:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese su nueva contraseña:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmCambioPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
If txtPass <> "" Then
    If txtPass.Text = txtPass2.Text Then
        UsuarioActual.Pass = txtPass
        UsuarioActual.GuardarModificaciones 'se agrego el 8/10/2010 1?!!?!?!? y antes funcionaba !?!?!?!
        MsgBox "Su contraseña se ha cambiado con exito", vbInformation + vbOKOnly
        Unload Me
    Else
        MsgBox "Ingrese la misma contraseña en ambos campos!"
    End If
Else
    MsgBox "Ingrese una contraseña!"
End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon

End Sub
