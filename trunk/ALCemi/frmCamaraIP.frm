VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmCamaraIP 
   Caption         =   "Visor de Cámara"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   8970
   Begin VB.Frame fraConfig 
      Caption         =   "Configuracion"
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   8895
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   7680
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtUrl 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "URL Camara:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdConfiguracion 
      Caption         =   "Configuracion"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   2040
      Width           =   8415
      ExtentX         =   14843
      ExtentY         =   7435
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmCamaraIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    fraConfig.Visible = False
    SaveSetting "TbrEmergencyGroup", "ComplementoCamara", "URL", txtUrl.Text
    wb.Navigate txtUrl.Text
    Form_Resize
End Sub

Private Sub cmdCancelar_Click()
    fraConfig.Visible = False
    Form_Resize
End Sub

Private Sub cmdConfiguracion_Click()
    fraConfig.Visible = True
    Form_Resize
End Sub

Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
    txtUrl.Text = GetSetting("TbrEmergencyGroup", "ComplementoCamara", "URL", "http://87.25.182.3/view/view.shtml")
   
    wb.Navigate txtUrl.Text
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        
        If fraConfig.Visible Then
            fraConfig.Width = Me.Width - 100
            cmdCancelar.Left = fraConfig.Width - 100 - cmdCancelar.Width
            cmdAceptar.Left = cmdCancelar.Left - cmdAceptar.Width - 100
            txtUrl.Width = cmdAceptar.Left - txtUrl.Left - 100
            wb.Height = Me.ScaleHeight - fraConfig.Height - 100
            wb.Top = fraConfig.Height + fraConfig.Top + 100
        Else
            wb.Top = cmdConfiguracion.Top + cmdConfiguracion.Height + 100
            wb.Height = Me.ScaleHeight - wb.Top - 100
        End If
        
        wb.Width = Me.Width - 100
    
    End If
End Sub
