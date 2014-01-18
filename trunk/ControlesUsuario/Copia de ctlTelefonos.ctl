VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "*\AListViewConsultaCtl.vbp"
Begin VB.UserControl ctlTelefonos 
   ClientHeight    =   2850
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4695
   ScaleHeight     =   2850
   ScaleWidth      =   4695
   Begin VB.Frame fra 
      Caption         =   "Telefonos"
      Height          =   2055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4000
      Begin ControlesPOO.ListViewConsulta lvw 
         Height          =   1695
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         HideSelection   =   0   'False
         HideEncabezados =   0   'False
         GridLines       =   -1  'True
         FullRowSelection=   -1  'True
         AutoDistribuirColumnas=   -1  'True
         AllowModify     =   0   'False
         ShowCheckBoxes  =   0   'False
         MultiSelect     =   0   'False
         NEncabezado0    =   "Tipo"
         MEncabezado0    =   "tipo"
         AEncabezado0    =   20
         NEncabezado1    =   "Numero"
         MEncabezado1    =   "Numero"
         AEncabezado1    =   30
         NEncabezado2    =   "Observaciones"
         MEncabezado2    =   "observaciones"
         AEncabezado2    =   50
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
      Begin MSForms.CommandButton cmdEliminar 
         Height          =   495
         Left            =   3360
         TabIndex        =   4
         ToolTipText     =   "Eliminar el telefono seleccionado"
         Top             =   1440
         Width           =   495
         Size            =   "873;873"
         Picture         =   "ctlTelefonos.ctx":0000
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEditar 
         Height          =   495
         Left            =   3360
         TabIndex        =   3
         ToolTipText     =   "Modificar el telefono seleccionado"
         Top             =   840
         Width           =   495
         Size            =   "873;873"
         Picture         =   "ctlTelefonos.ctx":1709A
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdNuevo 
         Height          =   495
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "Agregar un telefono"
         Top             =   240
         Width           =   495
         Size            =   "873;873"
         Picture         =   "ctlTelefonos.ctx":1E59C
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
End
Attribute VB_Name = "ctlTelefonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private WithEvents frm As frmABMTelefono
Attribute frm.VB_VarHelpID = -1

Public Event BeginDrag(pTelefono As Telefono)

Private xx As Single 'para ver el incremento para saber si arrastro
Private yy As Single

Private mTelefonos As TelefonoManager
Private mSoloConsulta As Boolean
Private mTelefonoDragged As Telefono

Public Property Get Telefonos() As TelefonoManager
    Set Telefonos = mTelefonos
End Property

Public Property Set Telefonos(pTelefonos As TelefonoManager)
    Set mTelefonos = pTelefonos
    Set lvw.Coleccion = mTelefonos
End Property

Public Property Get MinHeight() As Single
    MinHeight = 2000
End Property

Public Property Get MinWidth() As Single
    MinWidth = 4000
End Property

Public Property Get SoloConsulta() As Boolean
Attribute SoloConsulta.VB_Description = "Devuelve o establece si el control muestra los botones agregar, editar y eliminar."
    SoloConsulta = mSoloConsulta
End Property

Public Property Let SoloConsulta(pValor As Boolean)
    mSoloConsulta = pValor
    PropertyChanged
    UserControl_Resize
End Property

Public Property Get Caption() As String
    Caption = fra.Caption
End Property

Public Property Let Caption(pValor As String)
    fra.Caption = pValor
    PropertyChanged
End Property

Public Property Get TelefonoDragged() As Telefono
    Set TelefonoDragged = mTelefonoDragged
End Property

Private Sub cmdEditar_Click()
If Not lvw.SelectedItem Is Nothing Then
    Set frm = New frmABMTelefono
    frm.Modificar lvw.SelectedItem
End If
End Sub

Private Sub cmdEliminar_Click()
    If Not lvw.SelectedItem Is Nothing Then
        mTelefonos.Remove lvw.SelectedItem.id
        lvw.Refresh
    End If
End Sub

Private Sub cmdNuevo_Click()
    Set frm = New frmABMTelefono
    frm.Nuevo mTelefonos
End Sub

Private Sub frm_Eliminado(pTelefono As BLCemi.Telefono)
    lvw.Refresh
End Sub

Private Sub frm_Modificado(pTelefono As BLCemi.Telefono)
    lvw.Refresh
End Sub

Private Sub frm_Nuevo(pTelefono As BLCemi.Telefono)
    lvw.Refresh
End Sub

Private Sub lvw_MouseMove(Item As Object, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If (Abs(xx - X) > 50 Or Abs(yy - Y) > 50) And Button = vbLeftButton Then
        Set mTelefonoDragged = Item
        RaiseEvent BeginDrag(Item)
    End If
    xx = X
    yy = Y
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    fra.Caption = PropBag.ReadProperty("Caption")
    mSoloConsulta = PropBag.ReadProperty("SoloConsulta")
    UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Caption", fra.Caption
    PropBag.WriteProperty "SoloConsulta", mSoloConsulta
End Sub

Private Sub UserControl_Resize()
If UserControl.Width < MinWidth Then UserControl.Width = MinWidth
If UserControl.Height < MinHeight Then UserControl.Height = MinHeight

fra.Width = UserControl.Width
fra.Height = UserControl.Height

If mSoloConsulta Then
    cmdNuevo.Visible = False
    cmdEditar.Visible = False
    cmdEliminar.Visible = False
    lvw.Width = fra.Width - 100 - lvw.Left
    lvw.Height = fra.Height - 100 - lvw.Top
Else
    cmdNuevo.Visible = True
    cmdEditar.Visible = True
    cmdEliminar.Visible = True
    cmdNuevo.Left = fra.Width - 100 - cmdNuevo.Width
    cmdEditar.Left = cmdNuevo.Left
    cmdEliminar.Left = cmdNuevo.Left
    lvw.Width = cmdNuevo.Left - 100 - lvw.Left
    lvw.Height = fra.Height - 100 - lvw.Top
End If

End Sub

