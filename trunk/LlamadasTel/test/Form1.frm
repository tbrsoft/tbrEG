VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   3015
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents tel As Cliente
Attribute tel.VB_VarHelpID = -1

Private Sub Form_Load()
Set a = CreateObject("tbrllamadastel.application")
Set tel = New Cliente
Dim aux As ICliente
Set aux = tel
a.NotificameLlamadas aux
End Sub

Private Sub tel_LlamadaEntrando(pTelNumber As String)
Text1 = pTelNumber
End Sub

