VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.UserControl ctlDireccion 
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   ScaleHeight     =   2565
   ScaleWidth      =   4125
   Begin VB.Frame fra 
      Caption         =   "Direccion"
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4000
      Begin VB.TextBox txtTorre 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   18
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtEntreCalles 
         Height          =   285
         Left            =   1080
         TabIndex        =   16
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox txtDepto 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   14
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtPiso 
         Height          =   285
         Left            =   2280
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtNumero 
         Height          =   315
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtCalle 
         Height          =   315
         Left            =   1080
         MaxLength       =   254
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin ControlesPOO.Combo cmbBarrio 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   1080
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         NuevoEnabled    =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ControlesPOO.Combo cmbCiudad 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         NuevoEnabled    =   -1  'True
         Enabled         =   -1  'True
      End
      Begin ControlesPOO.Combo cmbProvincia 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   360
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTorre 
         Caption         =   "Torre:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label lblEntrecalles 
         Caption         =   "EntreCalles:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblDepto 
         Caption         =   "Depto:"
         Height          =   255
         Left            =   2880
         TabIndex        =   13
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblPiso 
         Caption         =   "Piso:"
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label lblNro 
         Caption         =   "Nro:"
         Height          =   195
         Left            =   2880
         TabIndex        =   8
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label3 
         Caption         =   "Calle:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label lblBarrio 
         Caption         =   "Barrio:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblCiudad 
         Caption         =   "Ciudad:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblProvincia 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "ctlDireccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event DireccionDragDrop(Source As Control, X As Single, Y As Single)

Private mProvinciaVisible As Boolean
Private mCanDragDrop As Boolean
Private mSoloConsulta As Boolean
Private mEntrecallesVisible As Boolean
Private mDireccion As BLCemi.Direccion

Public Property Get MinHeight() As Single
    'depende de provincia visible
    'implementar
End Property

Public Property Get MinWidth() As Single
    MinWidth = 4000
End Property

Public Property Get ProvinciaVisible() As Boolean
Attribute ProvinciaVisible.VB_Description = "Indica si el Combo Provincia es visible o no."
    ProvinciaVisible = mProvinciaVisible
End Property

Public Property Get SoloConsulta() As Boolean
    SoloConsulta = mSoloConsulta
End Property

Public Property Let SoloConsulta(pValor As Boolean)
    mSoloConsulta = pValor
    PropertyChanged
    cmbBarrio.Enabled = Not pValor
    cmbProvincia.Enabled = Not pValor
    cmbCiudad.Enabled = Not pValor
    txtCalle.Locked = pValor
    txtNumero.Locked = pValor
    txtDepto.Locked = pValor
    txtPiso.Locked = pValor
    txtEntreCalles.Locked = pValor
    txtTorre.Locked = pValor
End Property

Public Property Get CanDragDrop() As Boolean
    CanDragDrop = mCanDragDrop
End Property

Public Property Let CanDragDrop(pValue As Boolean)
    mCanDragDrop = pValue
End Property

Public Property Let ProvinciaVisible(pValue As Boolean)
    mProvinciaVisible = pValue
    PropertyChanged
End Property

Public Property Get EntrecallesVisible() As Boolean
    EntrecallesVisible = mEntrecallesVisible
End Property

Public Property Let EntrecallesVisible(pValue As Boolean)
    mEntrecallesVisible = pValue
    txtEntreCalles.Visible = pValue
    lblEntrecalles.Visible = pValue
    PropertyChanged
End Property

Public Property Get MiDireccion() As BLCemi.Direccion
    'actualizamos los campos antes de devolver la direccion
    mDireccion.Calle = txtCalle.Text
    mDireccion.Nro = txtNumero.Text
    Set mDireccion.Barrio = cmbBarrio.SelectedItem
    mDireccion.Piso = txtPiso.Text
    mDireccion.Depto = txtDepto.Text
    mDireccion.Torre = txtTorre.Text
    mDireccion.Entrecalles = txtEntreCalles.Text
    Set MiDireccion = mDireccion
End Property

Public Property Set MiDireccion(pValue As BLCemi.Direccion)
    Set mDireccion = pValue
    If Not mDireccion Is Nothing Then
        If Not mDireccion.Provincia Is Nothing Then
            Set cmbProvincia.SelectedItem = mDireccion.Provincia
            Set cmbCiudad.SelectedItem = mDireccion.Ciudad
            Set cmbBarrio.SelectedItem = mDireccion.Barrio
            txtCalle = mDireccion.Calle
            txtNumero = mDireccion.Nro
            txtDepto = mDireccion.Depto
            txtPiso = mDireccion.Piso
            txtEntreCalles = mDireccion.Entrecalles
            txtTorre = mDireccion.Torre
        End If
    End If
End Property

Public Property Get Caption() As String
    Caption = fra.Caption
End Property

Public Property Let Caption(pValor As String)
    fra.Caption = pValor
    PropertyChanged
End Property

Public Sub Inicializar( _
    pPais As BLCemi.Pais, _
    pProvincias As BLCemi.ProvinciaManager, _
    Optional pProvinciaSeleccionada As BLCemi.Provincia, _
    Optional pCiudadSeleccionada As BLCemi.Ciudad, _
    Optional pBarrioSeleccionado As BLCemi.Barrio)
    
    On Local Error GoTo ErrInitDIRCTL
    
    TERR.Anotar "abat6"
    
    On Error Resume Next
    'le adapto los nombres segun corresponda
    If pPais.PrimerOrden <> "" Then lblProvincia.Caption = pPais.PrimerOrden + ":"
    If pPais.SegundoOrden <> "" Then lblCiudad.Caption = pPais.SegundoOrden + ":"
    If pPais.TercerOrden <> "" Then lblBarrio.Caption = pPais.TercerOrden + ":"
    
    TERR.Anotar "abau6"
    
    Set cmbProvincia.Coleccion = pProvincias
    TERR.Anotar "abav"
    If Not pProvinciaSeleccionada Is Nothing Then
        TERR.Anotar "abaw"
        Set cmbProvincia.SelectedItem = pProvinciaSeleccionada
        If Not pCiudadSeleccionada Is Nothing Then
            TERR.Anotar "abax"
            Set cmbCiudad.SelectedItem = pCiudadSeleccionada
            If Not pBarrioSeleccionado Is Nothing Then
                TERR.Anotar "abay"
                Set cmbBarrio.SelectedItem = pBarrioSeleccionado
            End If
        End If
    End If
    
    TERR.Anotar "abaz"
    Exit Sub
ErrInitDIRCTL:
    TERR.AppendLog "ErrInitDIRCTL", TERR.ErrToTXT(Err)
End Sub

Public Sub Refresh()
'no funciona, hacer 'TODO
'    cmbProvincia.Refresh
'    cmbCiudad.Refresh
'    cmbBarrio.Refresh
End Sub

Public Function DireccionCompleta(outMensajeError As String) As Boolean
    Dim msj As String
    
    If Not TextBoxValidado(txtCalle, eString) Then msj = msj + "Ingrese el nombre de la calle." + vbCrLf
    If Not TextBoxValidado(txtNumero, eString) Then msj = msj + "Ingrese la altura de la calle." + vbCrLf
    If cmbProvincia.SelectedItem Is Nothing Then msj = msj + "Seleccione una Provincia." + vbCrLf
    If cmbCiudad.SelectedItem Is Nothing Then msj = msj + "Seleccione una Ciudad." + vbCrLf
    If cmbBarrio.SelectedItem Is Nothing Then msj = msj + "Seleccione un Barrio" + vbCrLf
    
    'ver de hacer propiedad para exigir entrecalles
    
    If msj = "" Then
        DireccionCompleta = True
    Else
        outMensajeError = msj
        DireccionCompleta = False
    End If

End Function

Private Sub cmbBarrio_NuevoSeleccionado()

If Not cmbCiudad.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Barrio")
    If Nombre <> "" Then
        If cmbBarrio.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim b As BLCemi.Barrio
            Set b = cmbBarrio.Coleccion.Nuevobarrio(Nombre, cmbCiudad.SelectedItem)
            cmbBarrio.Refresh
            Set cmbBarrio.SelectedItem = b
        End If
    End If
Else
    'debe seleccionar una ciudad primero
End If

End Sub

Private Sub cmbCiudad_ItemSeleccionado(Item As Object)
    Dim c As BLCemi.Ciudad
    Set c = Item
    Set cmbBarrio.Coleccion = c.Barrios
End Sub

Private Sub cmbCiudad_NuevoSeleccionado()
If Not cmbProvincia.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Ciudad")
    If Nombre <> "" Then
        If cmbCiudad.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim c As BLCemi.Ciudad
            Set c = cmbCiudad.Coleccion.Nuevaciudad(Nombre, cmbProvincia.SelectedItem)
            cmbCiudad.Refresh
            Set cmbCiudad.SelectedItem = c
        End If
    End If
Else
    'debe seleccionar una provincia antes
End If

End Sub

Private Sub cmbProvincia_ItemSeleccionado(Item As Object)
    Dim p As BLCemi.Provincia
    Set p = Item
    Set cmbCiudad.Coleccion = p.Ciudades
    Set cmbBarrio.Coleccion = Nothing
End Sub

Private Sub fra_DragDrop(Source As Control, X As Single, Y As Single)
    If mCanDragDrop Then RaiseEvent DireccionDragDrop(Source, X, Y)
End Sub


'------------------eventos-------------------------------

Private Sub UserControl_Resize()
If UserControl.Width < 4000 Then UserControl.Width = 4000
'el alto depende de provincia visible, cambiar despues
If mEntrecallesVisible Then
    If UserControl.Height <> 2565 Then UserControl.Height = 2565
Else
    If UserControl.Height <> 2265 Then UserControl.Height = 2265
End If

fra.Width = UserControl.Width
fra.Height = UserControl.Height
txtNumero.Left = fra.Width - 100 - txtNumero.Width
txtDepto.Left = fra.Width - 100 - txtDepto.Width
lblNro.Left = txtNumero.Left - 100 - lblNro.Width
lblDepto.Left = txtDepto.Left - 100 - lblDepto.Width
txtCalle.Width = lblDepto.Left - 100 - txtCalle.Left

txtPiso.Left = txtCalle.Left + txtCalle.Width - txtPiso.Width
lblPiso.Left = txtPiso.Left - 100 - lblPiso.Width

cmbBarrio.Width = fra.Width - 100 - cmbBarrio.Left
cmbCiudad.Width = cmbBarrio.Width
cmbProvincia.Width = cmbBarrio.Width
txtEntreCalles.Width = cmbBarrio.Width
End Sub

'-----------------------------------mantener estado-----------------------------

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   mProvinciaVisible = PropBag.ReadProperty("ProvinciaVisible", True)
   fra.Caption = PropBag.ReadProperty("Caption", "Direccion")
   mCanDragDrop = PropBag.ReadProperty("CanDragDrop", False)
   SoloConsulta = PropBag.ReadProperty("SoloConsulta", False)
   EntrecallesVisible = PropBag.ReadProperty("EntrecallesVisible", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ProvinciaVisible", ProvinciaVisible
    PropBag.WriteProperty "Caption", fra.Caption
    PropBag.WriteProperty "CanDragDrop", CanDragDrop
    PropBag.WriteProperty "SoloConsulta", SoloConsulta
    PropBag.WriteProperty "EntrecallesVisible", EntrecallesVisible
End Sub
