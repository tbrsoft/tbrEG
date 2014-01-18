VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmColaboracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Colaboración"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   5385
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtObservaciones 
         Height          =   1095
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtEfectivos 
         Height          =   285
         Left            =   1320
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin ControlesPOO.Combo cmbCuerpo 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin ControlesPOO.Combo cmbResponsable 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin ControlesPOO.Combo cmbUnidad 
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Top             =   960
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin ALCemi.GraphicButton cmdBuscar 
         Height          =   315
         Left            =   4680
         TabIndex        =   11
         Top             =   240
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
      End
      Begin VB.Label Observaciones 
         Caption         =   "Observaciones:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Unidad:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Cant. Efectivos:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Responsable:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cuerpo:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   3120
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "frmColaboracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event ColaboracionModificada(pColaboracion As blcemi.Colaboracion)
Public Event ColaboracionNueva(pColaboracion As blcemi.Colaboracion)

Dim WithEvents frmCuerpos As frmConsultaCuerpos
Attribute frmCuerpos.VB_VarHelpID = -1

Dim mColaboracion As blcemi.Colaboracion

Private Sub cmbCuerpo_ItemSeleccionado(Item As Object)
    Set cmbResponsable.Coleccion = Item.Responsables
    Set cmbUnidad.Coleccion = Item.Unidades
End Sub

Private Sub cmdAceptar_Click()
    Dim esNuevo As Boolean
    esNuevo = False
    If DatosCorrectos Then
        If mColaboracion Is Nothing Then
            esNuevo = True
            Set mColaboracion = New blcemi.Colaboracion
        End If
        
        Set mColaboracion.Cuerpo = cmbCuerpo.SelectedItem
        Set mColaboracion.Responsable = cmbResponsable.SelectedItem
        Set mColaboracion.Unidad = cmbUnidad.SelectedItem
        If IsNumeric(txtEfectivos.Text) Then
            mColaboracion.CantidadEfectivos = CInt(txtEfectivos.Text)
        End If
        mColaboracion.Observaciones = txtObservaciones.Text
        
        If esNuevo Then
            RaiseEvent ColaboracionNueva(mColaboracion)
        Else
            RaiseEvent ColaboracionModificada(mColaboracion)
        End If
        
        Unload Me
    End If
End Sub

Private Function DatosCorrectos() As Boolean
    Dim msj As String
    
    If Not TextBoxValidado(txtEfectivos, eInteger) Then msj = msj + "Ingrese la cantidad de efectivos." + vbCrLf
            
    If cmbCuerpo.SelectedItem Is Nothing Then msj = msj + "Seleccione un Cuerpo de Bomberos" + vbCrLf
        
    If msj = "" Then
        DatosCorrectos = True
    Else
        MsgBox "Faltan los siguientes datos:" + vbCrLf + msj, vbExclamation
        DatosCorrectos = False
    End If
    
End Function

Private Sub cmdBuscar_Click()
    Set frmCuerpos = New frmConsultaCuerpos
    frmCuerpos.Consultar GBL.CuerposDeBomberosGBL, etConRetorno
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
    Set cmdBuscar.Picture = MDI.il16.ListImages("buscar").Picture
    Set cmbCuerpo.Coleccion = GBL.CuerposDeBomberosGBL
End Sub

Private Sub frmCuerpos_CuerpoSeleccionado(pCuerpo As blcemi.CuerpoBomberos)
    Set cmbCuerpo.Coleccion = GBL.CuerposDeBomberosGBL
    Set cmbCuerpo.SelectedItem = pCuerpo
End Sub

Public Sub NuevaColaboracion()
    Me.Show
End Sub

Public Sub ModificarColaboracion(pColaboracion As blcemi.Colaboracion)
    Me.Show
    
    Dim mCuerpo As blcemi.CuerpoBomberos
    Set mColaboracion = pColaboracion
    Set mCuerpo = mColaboracion.Cuerpo
    Set cmbCuerpo.SelectedItem = mCuerpo
    Set cmbResponsable.Coleccion = mCuerpo.Responsables
    Set cmbUnidad.Coleccion = mCuerpo.Unidades
    Set cmbResponsable.SelectedItem = mColaboracion.Responsable
    Set cmbUnidad.SelectedItem = mColaboracion.Unidad
    txtEfectivos = mColaboracion.CantidadEfectivos
    txtObservaciones = mColaboracion.Observaciones
    
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "siniestro"
End Function

Public Sub Refrescar()
    Dim c As blcemi.CuerpoBomberos
    Dim r As blcemi.ResponsableCuerpo
    Dim u As blcemi.UnidadCuerpo
    
    Set c = cmbCuerpo.SelectedItem
    Set r = cmbResponsable.SelectedItem
    Set u = cmbUnidad.SelectedItem
    Set cmbCuerpo.Coleccion = GBL.CuerposDeBomberosGBL
    Set cmbResponsable.Coleccion = c.Responsables
    Set cmbUnidad.Coleccion = c.Unidades
    
    Set cmbCuerpo.SelectedItem = c
    Set cmbUnidad.SelectedItem = u
    Set cmbResponsable.SelectedItem = r
End Sub

Private Sub txtEfectivos_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii, False
End Sub
