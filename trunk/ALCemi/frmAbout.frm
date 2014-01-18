VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de ""tbr Emergency Group"""
   ClientHeight    =   6435
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5775
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4441.553
   ScaleMode       =   0  'User
   ScaleWidth      =   5423.023
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4800
      Top             =   3120
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4200
      TabIndex        =   0
      Top             =   4920
      Width           =   1500
   End
   Begin VB.Label lblUltimaVersion 
      Caption         =   "Label3"
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
      Left            =   2160
      TabIndex        =   7
      Top             =   3120
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Ultima Versión disponible:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "MS Word y MS Excel son marcas registradas de Microsoft Corp."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4560
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   2400
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5805
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   3344.106
      Y2              =   3344.106
   End
   Begin VB.Label lblDescription 
      Caption         =   "Software para la gestion de Centrales de Emergencias y Cuarteles de Bomberos."
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   3480
      Width           =   5805
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   3354.46
      Y2              =   3354.46
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión 1.0"
      Height          =   465
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   5685
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":CE8D
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   3870
   End
   Begin VB.Label lblModo 
      Caption         =   "Label2"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents ws As tbrSoftwareVersion
Attribute ws.VB_VarHelpID = -1
Dim Index As Integer
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
    Timer1.Enabled = True
    Select Case modo
        Case eModo.eModoDemo
            lblModo = "Modo: Demo. Puede registrar " + Str(500 - GBL.GetCantidadRegistros("cuota")) + " pagos y " + Str(800 - GBL.GetCantidadRegistros("atencion")) + " atenciones. Para registrar el presente software consulte el manual de usuario."
        Case eModo.eNoRegistrada
            lblModo = "Copia NO REGISTRADA. Consultar manual de usuario para registrar el presente software."
        Case eModo.eVersionRegistrada
            lblModo = "Copia registrada."
    End Select
    
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & "Version Base de Datos: " & GBL.GetDatabaseVersion
   On Error GoTo errman:
    Set ws = New tbrSoftwareVersion
    ws.GetLastVersion "tbrEG"
    Exit Sub
errman:
    lblUltimaVersion.Caption = "Sin datos."
End Sub

Private Sub ws_Done(pLastVersion As String)
    Timer1.Enabled = False
    lblUltimaVersion.Caption = pLastVersion
End Sub

Private Sub ws_Error(pMessage As String)
    Timer1.Enabled = False
    lblUltimaVersion.Caption = "No disponible: " + pMessage
End Sub

Private Sub Timer1_Timer()
    Index = Index + 1
    If Index = 5 Then Index = 1
    lblUltimaVersion = "Obteniendo información " + Choose(Index, ".", "..", "...", "....")
End Sub
