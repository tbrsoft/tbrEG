VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl TreeViewConsulta 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   PropertyPages   =   "TreeViewConsulta.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSComctlLib.TreeView tvw 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   2778
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "TreeViewConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
'Property Variables:
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer

Dim WithEvents mNodes As TVNodeManager
Attribute mNodes.VB_VarHelpID = -1
Dim mFirstNode As TVNode
Dim xx As Single
Dim yy As Single

'Event Declarations:
''siempre se levanta el evento click (aunque haga dobleclick)
Public Event ItemClick(Item As Object)
Public Event ItemDblClick(Item As Object)
Public Event ItemKeyEnterPressed(Item As Object)
Public Event ItemKeyDeletePressed(Item As Object)
Public Event ItemGotFocus(Item As Object)
Public Event ItemCheck(Item As Object, pCancel As Boolean)
'Public Event ItemEdited(Item As Object, pNewValue As String, pCancel As Boolean)

Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Ocurre cuando el usuario libera una tecla mientras un objeto tiene el enfoque."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."


''para agregar una propiedad no olvidarse:
''procedimientos set/let/get
''read y writeProperty
''pagina de propiedades
''no olvidarse el PropertyChanged para q se guarde el estado

'
'Private WithEvents vEncabezados As LVCEncabezadoManager
'
'Private vHideSelection As Boolean
'Private vHideEncabezados As Boolean
'Private vFullRowSelection As Boolean
'Private vGridLines As Boolean
'Private vAutoDistribuirColumnas As Boolean
'Private vCampoKey As String
'Private vAllowModify As Boolean
'Private vMultiSelect As Boolean
'Private vCheckBoxes As Boolean
'
'Private vColeccion As Object
'
'Private vMenu As Menu
'

'
''--------------------propiedades-------------------------
'Public Property Get MultiSelect() As Boolean
'    MultiSelect = vMultiSelect
'End Property
'
'Public Property Let MultiSelect(var As Boolean)
'     PropertyChanged
'     vMultiSelect = var
'     lv.MultiSelect = var
'End Property
'
'Public Property Get AllowModify() As Boolean
'    AllowModify = vAllowModify
'End Property
'
'Public Property Let AllowModify(pvalor As Boolean)
'    vAllowModify = pvalor
'    PropertyChanged
'End Property
'
'Public Property Get Picture() As StdPicture
'    Set Picture = lv.Picture
'End Property
'
'Public Property Set Picture(pPicture As StdPicture)
'    Set lv.Picture = pPicture
'    PropertyChanged
'End Property
'
'Public Property Set CheckedItems(col As Collection)
'    Dim li As ListItem
'    For Each li In lv.ListItems
'        For Each obj In col
'            If CallByName(li.Tag, vCampoKey, VbGet) = CallByName(obj, vCampoKey, VbGet) Then
'                li.Checked = True
'            End If
'        Next
'    Next
'End Property
'
'Public Property Get SelectedItems() As Collection
'    If vMultiSelect Then
'        Dim c As New Collection
'        Dim li As ListItem
'        For Each li In lv.ListItems
'            If li.Selected Then c.Add li.Tag
'        Next
'        Set SelectedItems = c
''    Else
''        MsgBox "No esta habilitada la opcion Multiselect"
'    End If
'End Property
'
'
'Public Property Get MenuPopUp() As Object
'    Set MenuPopUp = vMenu
'End Property
'
'Public Property Set MenuPopUp(pMenu As Object)
'    If TypeOf pMenu Is Menu Then
'        Set vMenu = pMenu
'    Else
'        Err.Raise 13 'no coinciden los tipos
'    End If
'End Property
'
''------------------------funciones publicas---------------------------
'Public Sub Editar()
'    If vAllowModify Then
'        If Not SelectedItem Is Nothing Then
'            lv.StartLabelEdit
'        End If
'    End If
'End Sub
'
'
'Public Sub filtrar(cadena As String)
'
'Dim v As Object
'Dim first As Boolean
'Dim li As ListItem
'Dim enc As LVCEncabezado
'first = True
'lv.ListItems.Clear
'
'For Each v In Coleccion
'    If InStr(1, getCadenaAtributos(v), cadena, vbTextCompare) <> 0 Then
'
'        For Each enc In Encabezados
'
'            If first Then
'100             Set li = lv.ListItems.Add(, , getValue(v, enc))
'                Set li.Tag = v
'                first = False
'            Else
'101             aux = getValue(v, enc)
'                li.ListSubItems.Add , , aux
'            End If
'        Next
'    End If
'    first = True
'Next
'If lv.ListItems.Count <> 0 Then RaiseEvent ItemGotFocus(lv.SelectedItem.Tag)
'Exit Sub
'e:
'If Erl = 100 Or Erl = 101 Then
'    Err.Raise 2008, , "Debe Implementar la interfaz IlvwConsulta" 'ya no la uso, ver q errores se presentan y adaptar
'End If
'
'End Sub
'

'
''------------------------funciones privadas------------------------------------
'
'
'Private Function getCadenaAtributos(v As Object) As String
'Dim cad As String
'Dim enc As LVCEncabezado
'For Each enc In vEncabezados
'    If enc.filtrar Then
'        mTipoLlamado = detectarMetodoLlamado(v, enc)
'        cad = cad + getValue(v, enc)
'    End If
'Next
'getCadenaAtributos = cad
'End Function
'
''funciones para exportar el contenido

'------------------------Propiedades--------------------------------------


'ver un campo para asignarle iconos a los nodos

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get Appearance() As AppearanceConstants
Attribute Appearance.VB_Description = "Devuelve o establece si los controles, formularios y formularios MDI se dibujan en tiempo de ejecución con efectos 3D."
    Appearance = tvw.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As AppearanceConstants)
    tvw.Appearance() = New_Appearance
    PropertyChanged "Appearance"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indica si un control Label o el color de fondo de un control Shape es transparente u opaco."
Attribute BackStyle.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As BorderStyleConstants
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    BorderStyle = tvw.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleConstants)
    tvw.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indica si los elementos de un control se ordenan automáticamente de forma alfabética."
Attribute Sorted.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Sorted = tvw.Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    tvw.Sorted = New_Sorted
    PropertyChanged "Sorted"
End Property

Public Property Get Style() As TreeStyleConstants
Attribute Style.VB_Description = "Muestra una lista jerárquica de objetos Node, cada uno con una etiqueta y un mapa de bits opcional."
    Style = tvw.Style
End Property

Public Property Let Style(ByVal New_Style As TreeStyleConstants)
    tvw.Style() = New_Style
    PropertyChanged "Style"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Devuelve o establece el texto mostrado cuando el mouse se sitúa sobre un control."
Attribute ToolTipText.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    ToolTipText = tvw.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    tvw.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Public Property Get FullRowSelect() As Boolean
Attribute FullRowSelect.VB_Description = "Devuelve o establece un valor que determina si se resalta la fila entera del elemento seleccionado y si al hacer clic en cualquier lugar de la fila se selecciona el elemento."
    FullRowSelect = tvw.FullRowSelect
End Property

Public Property Let FullRowSelect(ByVal New_FullRowSelect As Boolean)
    tvw.FullRowSelect() = New_FullRowSelect
    PropertyChanged "FullRowSelect"
End Property

Public Property Get HideSelection() As Boolean
Attribute HideSelection.VB_Description = "Determina si el elemento seleccionado se mostrará como seleccionado cuando TreeView pierda el enfoque"
Attribute HideSelection.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    HideSelection = tvw.HideSelection
End Property

Public Property Let HideSelection(ByVal New_HideSelection As Boolean)
    tvw.HideSelection() = New_HideSelection
    PropertyChanged "HideSelection"
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_Description = "Devuelve o establece un valor que determina si los elementos se resaltan cuando el puntero del mouse pasa sobre ellos."
Attribute HotTracking.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    HotTracking = tvw.HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    tvw.HotTracking() = New_HotTracking
    PropertyChanged "HotTracking"
End Property

Public Property Get Indentation() As Single
Attribute Indentation.VB_Description = "Devuelve o establece el ancho de la sangría de un control TreeView."
Attribute Indentation.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Indentation = tvw.Indentation
End Property

Public Property Let Indentation(ByVal New_Indentation As Single)
    tvw.Indentation() = New_Indentation
    PropertyChanged "Indentation"
End Property

Public Property Get LabelEdit() As LabelEditConstants
Attribute LabelEdit.VB_Description = "Devuelve o establece un valor que determina si un usuario puede modificar la etiqueta de un objeto ListItem o Node."
    LabelEdit = tvw.LabelEdit
End Property

Public Property Let LabelEdit(ByVal New_LabelEdit As LabelEditConstants)
    tvw.LabelEdit() = New_LabelEdit
    PropertyChanged "LabelEdit"
End Property

Public Property Get LineStyle() As TreeLineStyleConstants
Attribute LineStyle.VB_Description = "Devuelve o establece el estilo de las líneas mostradas entre objetos Node."
    LineStyle = tvw.LineStyle
End Property

Public Property Let LineStyle(ByVal New_LineStyle As TreeLineStyleConstants)
    tvw.LineStyle() = New_LineStyle
    PropertyChanged "LineStyle"
End Property

Public Property Get SingleSel() As Boolean
Attribute SingleSel.VB_Description = "Devuelve o establece un valor que determina si al seleccionar un nuevo elemento de la estructura se expande ese elemento y se contraen los elementos seleccionados anteriormente."
Attribute SingleSel.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    SingleSel = tvw.SingleSel
End Property

Public Property Let SingleSel(ByVal New_SingleSel As Boolean)
    tvw.SingleSel = New_SingleSel
    PropertyChanged "SingleSel"
End Property

Public Property Get ShowCheckBoxes() As Boolean
    ShowCheckBoxes = tvw.Checkboxes
End Property

Public Property Let ShowCheckBoxes(var As Boolean)
     tvw.Checkboxes = var
     PropertyChanged "ShowCheckBoxes"
End Property

'--------propiedades funcionamiento---------------------

Public Property Get SelectedItem() As Object
    If Not tvw.SelectedItem Is Nothing Then
        Set SelectedItem = tvw.SelectedItem.Tag
    End If
End Property

Public Property Set SelectedItem(pItem As Object)
'    Dim nod As Node
'    For Each nod In tvw.Nodes
'
'        If getValue(li.Tag, vCampoKey) = CallByName(pItem, vCampoKey, VbGet) Then
'            li.Selected = True
'            Exit For
'        End If
'    Next
End Property

Public Property Get CheckedItems() As Collection
    
    Dim c As New Collection
    Dim nod As Node
    For Each nod In tvw.Nodes
        If nod.Checked Then c.Add nod.Tag
    Next
    Set CheckedItems = c
    
End Property

Public Property Get Nodos() As TVNodeManager
    Set Nodos = mNodes
End Property

Public Property Get Coleccion() As Object
   Set Coleccion = mFirstNode.Collection
End Property

Public Property Set Coleccion(var As Object)
    
    If Nodos.Count <> 0 Then
        If mFirstNode Is Nothing Then
            Set mFirstNode = Nodos.Item("k 0")
        End If
        Set mFirstNode.Collection = var
        Refresh
    Else
        MsgBox "No se seteo ningun nodo."
    End If
    
End Property

'----------------------funciones publicas------------------------
Public Sub Refresh()
    On Error GoTo e:
    
    If Not Coleccion Is Nothing Then
        Dim v As Object
        Dim first As Boolean
        Dim nod As TVNode
        first = True
        tvw.Nodes.Clear
        
        LlenarRama mFirstNode, Nothing
        If tvw.Nodes.Count <> 0 Then
            Set tvw.SelectedItem = tvw.Nodes(1)
            RaiseEvent ItemGotFocus(tvw.SelectedItem.Tag)
        End If
    Else
        tvw.Nodes.Clear
    End If
Exit Sub
e:
     MsgBox Err.Description
End Sub

'----------------Funciones Privadas-------------------------
Private Function LlenarRama(tvNod As TVNode, nodAux As Node)
    On Error GoTo e
    Static id As Long  'prueba
    Dim v As Object
    Dim nod As Node
100 For Each v In tvNod.Collection
        id = id + 1
        If Not nodAux Is Nothing Then
            Set nod = tvw.Nodes.Add(nodAux.Key, tvwChild, "k" + Str(id), getValue(v, tvNod.TextField))
        Else
            Set nod = tvw.Nodes.Add(, , "k" + Str(id), getValue(v, tvNod.TextField))
        End If
        'seteamos propiedades
        nod.Expanded = tvNod.Expanded
        nod.ForeColor = tvNod.ForeColor
        nod.Bold = tvNod.Bold
        nod.BackColor = tvNod.BackColor
                
        Set nod.Tag = v
        'si tiene un nodo child cargo esa rama tambien
        If Not tvNod.Child Is Nothing Then
            Set tvNod.Child.Collection = CallByName(v, tvNod.ChildCollectionField, VbGet)
            LlenarRama tvNod.Child, nod
        End If
    Next
Exit Function
e:
     If Erl() = 100 Then
        MsgBox "La coleccion no tiene la propiedad IDprocedimiento=-4"
     Else
        MsgBox Err.Description
     End If
End Function

Private Sub mNodes_NodeAdded()
    PropertyChanged
End Sub

Private Sub UserControl_Resize()
    tvw.Height = UserControl.Height
    tvw.Width = UserControl.Width
End Sub

''-----------------eventos del control de usuario-----------------
'
'Private Sub lv_AfterLabelEdit(Cancel As Integer, NewString As String)
'    Dim pvalor As Boolean
'    RaiseEvent ItemEdited(SelectedItem, NewString, pvalor)
'    If pvalor Then Cancel = 1
'End Sub

Private Sub tvw_Click()
    Dim nod As Node
    Set nod = tvw.HitTest(xx, yy)
    If Not nod Is Nothing Then
        RaiseEvent ItemClick(nod.Tag)
    End If
End Sub

Private Sub tvw_DblClick()
    Dim nod As Node
    Set nod = tvw.HitTest(xx, yy)
    If Not nod Is Nothing Then
        RaiseEvent ItemDblClick(nod.Tag)
    End If
End Sub

Private Sub tvw_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    xx = X
    yy = y
End Sub

Private Sub tvw_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim aux As Boolean
    RaiseEvent ItemCheck(Node.Tag, aux)
    'no funciona, ver q pasa
    If aux Then tvw.Nodes(Node.Key).Checked = False
    
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    RaiseEvent ItemGotFocus(Node.Tag)
End Sub

Private Sub tvw_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not tvw.SelectedItem Is Nothing Then
                RaiseEvent ItemKeyEnterPressed(tvw.SelectedItem.Tag)
            End If
            KeyAscii = 0
        Case vbKeyDelete
            If Not tvw.SelectedItem Is Nothing Then
                RaiseEvent ItemKeyDeletePressed(tvw.SelectedItem.Tag)
            End If
            KeyAscii = 0
        Case Else
            RaiseEvent KeyDown(KeyCode, Shift)
    End Select
End Sub
'
'Private Sub lv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = vbRightButton Then
'        If Not vMenu Is Nothing Then UserControl.PopupMenu vMenu
'    End If
'End Sub
'


'----------------Mantener el estado----------------------------------

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    Set mNodes = New TVNodeManager
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set mNodes = New TVNodeManager
    
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    tvw.Appearance = PropBag.ReadProperty("Appearance", 1)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    tvw.Checkboxes = PropBag.ReadProperty("ShowCheckBoxes", False)
    tvw.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    tvw.Sorted = PropBag.ReadProperty("Sorted", False)
    tvw.Style = PropBag.ReadProperty("Style", 7)
    tvw.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    tvw.FullRowSelect = PropBag.ReadProperty("FullRowSelect", False)
    tvw.HideSelection = PropBag.ReadProperty("HideSelection", verdadero)
    tvw.HotTracking = PropBag.ReadProperty("HotTracking", False)
    tvw.Indentation = PropBag.ReadProperty("Indentation", "300")
    tvw.LabelEdit = PropBag.ReadProperty("LabelEdit", 0)
    tvw.LineStyle = PropBag.ReadProperty("LineStyle", 1)
    tvw.SingleSel = PropBag.ReadProperty("SingleSel", False)
    
    Dim nAux As TVNode
    Dim nParent As TVNode
    For i = 0 To 19
        Set nAux = New TVNode
        nAux.id = i
        nAux.BackColor = CLng(PropBag.ReadProperty("Nodo.BackColor" + Trim(Str(i)), Str(vbWhite)))
        nAux.Bold = IIf(PropBag.ReadProperty("Nodo.Bold" + Trim(Str(i)), "False") = "True", True, False)
        nAux.ChildCollectionField = PropBag.ReadProperty("Nodo.ChildCollectionField" + Trim(Str(i)), "")
        nAux.Expanded = IIf(PropBag.ReadProperty("Nodo.Expanded" + Trim(Str(i)), "False") = "True", True, False)
        nAux.ForeColor = CLng(PropBag.ReadProperty("Nodo.ForeColor" + Trim(Str(i)), vbBlack))
        nAux.IdField = PropBag.ReadProperty("Nodo.IdField" + Trim(Str(i)), "")
        nAux.TextField = PropBag.ReadProperty("Nodo.TextField" + Trim(Str(i)), "")
        
        If nAux.TextField <> "" Then
            Nodos.AddItem nAux
            If Not nParent Is Nothing Then
                Set nParent.Child = nAux
            End If
            Set nParent = nAux
        End If
    Next
    
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Appearance", tvw.Appearance, 1)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("ShowCheckBoxes", tvw.Checkboxes, False)
    Call PropBag.WriteProperty("BorderStyle", tvw.BorderStyle, 0)
    Call PropBag.WriteProperty("Sorted", tvw.Sorted, False)
    Call PropBag.WriteProperty("Style", tvw.Style, 7)
    Call PropBag.WriteProperty("ToolTipText", tvw.ToolTipText, "")
    Call PropBag.WriteProperty("FullRowSelect", tvw.FullRowSelect, False)
    Call PropBag.WriteProperty("HideSelection", tvw.HideSelection, verdadero)
    Call PropBag.WriteProperty("HotTracking", tvw.HotTracking, False)
    Call PropBag.WriteProperty("Indentation", tvw.Indentation, "600")
    Call PropBag.WriteProperty("LabelEdit", tvw.LabelEdit, 0)
    Call PropBag.WriteProperty("LineStyle", tvw.LineStyle, 0)
    Call PropBag.WriteProperty("SingleSel", tvw.SingleSel, False)
    
    Dim nod As TVNode
    Dim i As Integer
    For Each nod In Nodos

        PropBag.WriteProperty "Nodo.BackColor" + Trim(Str(i)), Str(nod.BackColor)
        PropBag.WriteProperty "Nodo.Bold" + Trim(Str(i)), IIf(nod.Bold, "True", "False")
        PropBag.WriteProperty "Nodo.ChildCollectionField" + Trim(Str(i)), nod.ChildCollectionField
        PropBag.WriteProperty "Nodo.Expanded" + Trim(Str(i)), IIf(nod.Expanded, "True", "False")
        PropBag.WriteProperty "Nodo.ForeColor" + Trim(Str(i)), Str(nod.ForeColor)
        PropBag.WriteProperty "Nodo.IdField" + Trim(Str(i)), nod.IdField
        PropBag.WriteProperty "Nodo.TextField" + Trim(Str(i)), nod.TextField
        
        i = i + 1
    Next
    For i = j To 19
        PropBag.WriteProperty "Nodo.BackColor" + Trim(Str(i)), "0"
        PropBag.WriteProperty "Nodo.Bold" + Trim(Str(i)), "False"
        PropBag.WriteProperty "Nodo.ChildCollectionField" + Trim(Str(i)), ""
        PropBag.WriteProperty "Nodo.Expanded" + Trim(Str(i)), "False"
        PropBag.WriteProperty "Nodo.ForeColor" + Trim(Str(i)), "0"
        PropBag.WriteProperty "Nodo.IdField" + Trim(Str(i)), ""
        PropBag.WriteProperty "Nodo.TextField" + Trim(Str(i)), ""
    Next
End Sub

