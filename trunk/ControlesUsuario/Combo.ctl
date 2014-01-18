VERSION 5.00
Begin VB.UserControl Combo 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   525
   ScaleWidth      =   4800
   ToolboxBitmap   =   "Combo.ctx":0000
   Begin VB.ComboBox cmb 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Combo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event NuevoSeleccionado()
Attribute NuevoSeleccionado.VB_Description = "Sucede cuando el usuario hace click en la opcion Nuevo dentro del combo. Para que se muestre la ocion Nuevo modifique la propiedad NuevoEnabled."
Public Event ItemSeleccionado(Item As Object)
Attribute ItemSeleccionado.VB_Description = "Sucede cuando el usuaio selecciona un item de la lista. Item es del mismo tipo de objeto que la coleccion."
Private mNuevoEnabled As Boolean
Private vColeccion As Object
Private vAtributoAMostrar As String
Private vCampoKey As String

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
   Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
   UserControl.Enabled = NewValue
   PropertyChanged "Enabled"
End Property

Public Property Get CampoKey() As String
Attribute CampoKey.VB_Description = "El nombre del atributo clave de la clase."
    CampoKey = vCampoKey
End Property

Public Property Let CampoKey(ByVal vNewValue As String)
    vCampoKey = vNewValue
    PropertyChanged
End Property

Public Property Get AtributoAMostrar() As String
Attribute AtributoAMostrar.VB_Description = "El nombre del atributo de la clase que se quiere mostrar en el combo."
    AtributoAMostrar = vAtributoAMostrar
End Property

Public Property Let AtributoAMostrar(ByVal vNewValue As String)
    vAtributoAMostrar = vNewValue
    PropertyChanged
End Property

Public Property Get Coleccion() As Object
Attribute Coleccion.VB_Description = "Elementos que se van a mostrar. La coleccion debe tener el metodo enumerator para funcionar."
    Set Coleccion = vColeccion
End Property

Public Property Set Coleccion(ByRef vNewValue As Object)
    Set vColeccion = vNewValue
    Refresh
End Property

Public Property Get ListIndex() As Integer
ListIndex = cmb.ListIndex
End Property

Public Property Let ListIndex(vData As Integer)
On Error Resume Next
cmb.ListIndex = vData
End Property

Public Property Get NuevoEnabled() As Boolean
Attribute NuevoEnabled.VB_Description = "Valor boolean que indica si se muestra la opcion Nuevo en el combo."
    NuevoEnabled = mNuevoEnabled
End Property

Public Property Let NuevoEnabled(vData As Boolean)
    mNuevoEnabled = vData
    cmb.Clear
    If NuevoEnabled Then cmb.AddItem "<Nuevo>"
           
    Refresh
End Property

Private Sub cmb_Click()
    If cmb.Text = "<Nuevo>" Then
        RaiseEvent NuevoSeleccionado
    Else
        Dim obj As Object
        indice = cmb.ItemData(cmb.ListIndex)
        For Each obj In vColeccion
            If Trim(CallByName(obj, vCampoKey, VbGet)) = Trim(Str(indice)) Then
            Exit For
            End If
        Next
        
        RaiseEvent ItemSeleccionado(obj)
    End If
End Sub

Private Sub UserControl_Initialize()
    cmb.Clear
    If NuevoEnabled Then cmb.AddItem "<Nuevo>"
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    NuevoEnabled = PropBag.ReadProperty("NuevoEnabled", False)
    CampoKey = PropBag.ReadProperty("CampoKey", "id")
    AtributoAMostrar = PropBag.ReadProperty("AtributoAMostrar", "nombre")
    Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

Private Sub UserControl_Resize()
UserControl.Height = cmb.Height
cmb.Width = UserControl.Width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "NuevoEnabled", NuevoEnabled, False
    PropBag.WriteProperty "CampoKey", vCampoKey, "id"
    PropBag.WriteProperty "AtributoAMostrar", vAtributoAMostrar, "nombre"
    PropBag.WriteProperty "Enabled", UserControl.Enabled
End Sub

'Public Sub addItem(id, valor As String)
'    cmb.addItem valor
'    cmb.ItemData(cmb.NewIndex) = id
'End Sub
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refresca el contenido del combo. Llamar despues de modificar la coleccion incluida el combo (agregar o quitar elemento/s)."
    On Error GoTo salir
    Dim v
    Dim sel
    Set sel = Me.SelectedItem
    cmb.Clear
    For Each v In vColeccion
        cmb.AddItem CallByName(v, vAtributoAMostrar, VbGet)
        cmb.ItemData(cmb.NewIndex) = CallByName(v, vCampoKey, VbGet)
    Next
        If NuevoEnabled Then cmb.AddItem "<Nuevo>"
    Set Me.SelectedItem = sel
salir:
End Sub

Public Sub Clear()
    cmb.Clear
End Sub

'Public Function getId()
'    If cmb.ListIndex <> -1 Then getId = cmb.ItemData(cmb.ListIndex) Else getId = -1
'End Function

'Public Sub seleccionarItem(id)
'    For i = 0 To cmb.ListCount - 1
'        If cmb.ItemData(i) = id Then
'            cmb.ListIndex = i
'            Exit For
'        End If
'    Next
'End Sub

Public Property Get SelectedItem() As Object
Attribute SelectedItem.VB_Description = "Se usa para obtener o setear el item que se muestra en el combo. En caso del set asegurarse que el item sea parte de la coleccion del combo."
    On Error GoTo errman
    Dim obj As Object
    If cmb.ListIndex <> -1 Then
        indice = cmb.ItemData(cmb.ListIndex)
        For Each obj In vColeccion
            If Trim(CallByName(obj, vCampoKey, VbGet)) = Trim(Str(indice)) Then
            Exit For
            End If
        Next
        Set SelectedItem = obj
    End If
Exit Property
errman:

End Property

Public Property Set SelectedItem(value As Object)
    If Not value Is Nothing Then
        For i = 0 To cmb.ListCount - 1
            If cmb.ItemData(i) = CallByName(value, vCampoKey, VbGet) Then
                cmb.ListIndex = i
                Exit For
            End If
        Next
    End If
End Property

