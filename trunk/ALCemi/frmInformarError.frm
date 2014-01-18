VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmInformarError 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informe de Error"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkEnviarSiempre 
      Caption         =   "Enviar siempre, no preguntarme nuevamente."
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   6000
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "Enviar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   7095
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   6600
         Top             =   360
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblMensaje 
         Height          =   855
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   7095
      Begin VB.TextBox txtContexto 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   4680
         Width           =   6855
      End
      Begin VB.TextBox txtError 
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         Height          =   2295
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   1800
         Width           =   6855
      End
      Begin VB.Label Label3 
         Caption         =   $"frmInformarError.frx":0000
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   6735
      End
      Begin VB.Label Label2 
         Caption         =   $"frmInformarError.frx":00C4
         Height          =   855
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6735
      End
      Begin VB.Label Label1 
         Caption         =   "Contenido del Informe:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmInformarError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents mEService As tbrErrorService
Attribute mEService.VB_VarHelpID = -1

Dim enviarSinPreguntar As Boolean

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdEnviar_Click()
    Frame2.Visible = True
    Frame1.Visible = False
    cmdEnviar.Visible = False
    cmdCancelar.Caption = "Cerrar"
    lblMensaje.Caption = "Enviando informe..."
    Enviar
    Timer1.Enabled = True
    CCFFGG.Configuracion.Comportamiento.EnviarErrores = chkEnviarSiempre.Value
    CCFFGG.Configuracion.Save
End Sub

Private Sub Enviar()
    Set mEService = New tbrErrorService
    mEService.NotificateError Replace(txtError.Text + IIf(txtContexto.Text <> "", vbCrLf + vbCrLf + "Contexto del Error:" + vbCrLf + txtContexto.Text, ""), vbCrLf, "***")
End Sub

Private Sub Form_Load()
    chkEnviarSiempre.Value = CCFFGG.Configuracion.Comportamiento.EnviarErrores
End Sub

Private Sub mEService_Done(pLastVersion As String)
    If enviarSinPreguntar Then
        Unload Me
    Else
        lblMensaje.Caption = "El informe ha sido recibido. Muchas gracias por su colaboración."
        Timer1.Enabled = False
        cmdCancelar.Enabled = True
        pb.Visible = False
    End If
End Sub

Private Sub mEService_Error(pMessage As String)
    If enviarSinPreguntar Then
        Unload Me
    Else
        lblMensaje.Caption = "Ocurrio un error al intentar enviar el informe. Puede encontrar una copia del mismo en la carpeta del programa con el nombre ""errorlog.txt""."
        Timer1.Enabled = False
        cmdCancelar.Enabled = True
        pb.Visible = False
    End If
End Sub

Private Sub Timer1_Timer()
pb.Value = pb.Value + 1
If pb.Value = 20 Then pb.Value = 5
End Sub

Public Sub InformarError(pError As String)
    txtError.Text = pError
    If CCFFGG.Configuracion.Comportamiento.EnviarErrores = 1 Then
        enviarSinPreguntar = True
        Enviar
    Else
        Me.Show
    End If
End Sub
