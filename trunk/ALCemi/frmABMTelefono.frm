VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmABMTelefono 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4410
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Frame fraDatos 
      Height          =   2055
      Left            =   50
      TabIndex        =   2
      Top             =   0
      Width           =   4335
      Begin VB.TextBox txtObservaciones 
         Height          =   735
         Left            =   1320
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1200
         Width           =   2895
      End
      Begin ControlesPOO.Combo cmbTipo 
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         NuevoEnabled    =   -1  'True
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNumero 
         Height          =   315
         Left            =   1320
         TabIndex        =   0
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Observaciones:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "Numero:"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   240
         Width           =   600
      End
   End
End
Attribute VB_Name = "frmABMTelefono"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar
Public Event Nuevo(pTelefono As blcemi.Telefono)
Public Event Modificado(pTelefono As blcemi.Telefono)
Public Event Eliminado(pTelefono As blcemi.Telefono)
Public Event Cancelado()

Private Tipo As eTipoAMB 'enumeracion definida en Modulo

Private mTelefono As blcemi.Telefono
Private mTelefonos As blcemi.TelefonoManager

Private Sub cmbTipo_NuevoSeleccionado()
    Dim aux As String
    
    aux = frmABMGenerico.Nuevo("Ingrese el nuevo Tipo de Telefono:")
    
    If aux <> "" Then
        Dim ocAux As blcemi.TipoTelefono
        Dim newOc As blcemi.TipoTelefono
        Set ocAux = GBL.TiposTelefonoGBL.ItemByName(aux)
        If ocAux Is Nothing Then
            Set newOc = GBL.TiposTelefonoGBL.Nuevo(aux)
            cmbTipo.Refresh
            Set cmbTipo.SelectedItem = newOc
        End If
    End If

End Sub

Private Sub cmdAceptar_Click()
    
    Select Case Tipo
        Case etALTA
            If DatosCorrectos Then
                Set mTelefono = mTelefonos.Nuevo(txtNumero, cmbTipo.SelectedItem, txtObservaciones)
                
                RaiseEvent Nuevo(mTelefono)
                Unload Me
            End If
        Case etBAJA
            'implementar
        Case etMODIFICACION
            If DatosCorrectos Then
                mTelefono.numero = txtNumero
                Set mTelefono.Tipo = cmbTipo.SelectedItem
                mTelefono.Observaciones = txtObservaciones
                RaiseEvent Modificado(mTelefono)
                Unload Me
            End If
        Case etCONSULTA
            'implementar
    End Select
    
End Sub

Private Function DatosCorrectos() As Boolean

Dim msj As String

If Not TextBoxValidado(txtNumero, eString) Then msj = msj + "Ingrese el numero de telefono." + vbCrLf
If cmbTipo.SelectedItem Is Nothing Then msj = msj + "Seleccione un tipo de telefono." + vbCrLf

If msj = "" Then
    DatosCorrectos = True
Else
    MsgBox "Faltan los siguientes datos:" + vbCrLf + msj, vbExclamation
    DatosCorrectos = False
End If

End Function
Private Sub cmdCancelar_Click()
RaiseEvent Cancelado
Unload Me
End Sub

Public Sub Nuevo(pTelefonos As blcemi.TelefonoManager)  'mandar la coleccion como parametro

Set mTelefonos = pTelefonos
'Set mTelefono = New Telefono
Tipo = etALTA
Me.Show
Me.Caption = "Nuevo Telefono"
End Sub

Public Sub Modificar(pTelefono As blcemi.Telefono)  'mandar como parametro el elemento a modificar
Set mTelefono = pTelefono
Tipo = etMODIFICACION
Me.Show
Me.Caption = "Modificar Telefono"
txtNumero = mTelefono.numero
Set cmbTipo.SelectedItem = mTelefono.Tipo
txtObservaciones = mTelefono.Observaciones

End Sub

'Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
''implementar
'tipo = etBAJA
'Me.Show
'Me.Caption = "Eliminar Telefono"
'End Sub

'Public Sub VerDatos() 'mandar como parametro el elemento del q queremos ver los datos
''implementar
'tipo = etCONSULTA
'Me.Show
'End Sub

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If Tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
Set Me.Icon = MDI.Icon

Set cmbTipo.Coleccion = GBL.TiposTelefonoGBL
On Error Resume Next
cmbTipo.Enabled = UsuarioActual.Permisos.Can(blcemi.AltaTipoTelefono)

End Sub

Public Sub Refrescar()

End Sub

