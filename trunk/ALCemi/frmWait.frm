VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Espere por favor..."
   ClientHeight    =   690
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   690
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Se esta intentando conectar con la base de datos."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelar_Click()
    End
End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon

End Sub

Private Sub Timer1_Timer()
    If GBL.ConectarBaseDatos Then Unload Me
End Sub
