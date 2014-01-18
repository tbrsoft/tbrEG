VERSION 5.00
Begin VB.Form frmLlamada 
   Caption         =   "Llamada Telefonica"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7740
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origen"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "(03543) - 489426"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Numero:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmLlamada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mApp As Application

Private Sub Command1_Click()
'mApp.NotificarLlamadas txtNumero
'mapp.TelephoneCalls.NewTelephoneCall ()'probar
End Sub

Public Sub setApp(pApp As Application)
Set mApp = pApp
End Sub
