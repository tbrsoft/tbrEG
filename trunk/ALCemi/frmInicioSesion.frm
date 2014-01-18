VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmInicioSesion 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Emergency Group"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1350
      Left            =   2640
      Picture         =   "frmInicioSesion.frx":0000
      ScaleHeight     =   90
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   225
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   3370
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2280
      Width           =   1575
   End
   Begin VB.TextBox txtPass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1800
      Width           =   3375
   End
   Begin ControlesPOO.ListViewConsulta lvwUsuarios 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   4471
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Usuarios"
      MEncabezado0    =   "login"
      AEncabezado0    =   100
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
      NEncabezado0    =   ""
      MEncabezado0    =   ""
      AEncabezado0    =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Password:"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frmInicioSesion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event SesionIniciada(pUsuarioLogueado As blcemi.Empleado)

Dim mUsuarioActual As blcemi.Empleado

Private Sub cmdAceptar_Click()
    Set mUsuarioActual = lvwUsuarios.SelectedItem
    If mUsuarioActual.Pass = txtPass Then
        RaiseEvent SesionIniciada(mUsuarioActual)
        Unload Me
    Else
        MsgBox "La contraseña introducida no es correcta!"
        txtPass = ""
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        MDI.mnuContenido_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Move (MDI.Width - Me.Width) / 2, (MDI.ScaleHeight / 2) - Me.Height
    'separo los usuarios q tienen login
              
    Dim em As blcemi.EmpleadoManager
    Set em = New blcemi.EmpleadoManager
    Dim e As blcemi.Empleado
    
    For Each e In GBL.EmpleadosGBL
        If e.Login <> "" Then em.AddItem e
    Next
    
    Set lvwUsuarios.Coleccion = em
'    If em.Count = 0 Then
'        MsgBox "En el siguiente formulario, ingrese los datos del usuario a cargo del sistema.", vbInformation + vbOKOnly, "Preparando el sistema para el primer uso..."
'        frmABMEmpleado.Nuevo gbl.EmpleadosGBL
'        Unload Me 'esto provoca un error en mdi mostrariniciosesion, no se por q
'    End If
    Set Me.Icon = MDI.Icon

End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "seguridad"
End Function

