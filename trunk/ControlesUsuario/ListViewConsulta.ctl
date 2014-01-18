VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ListViewConsulta 
   ClientHeight    =   2865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4110
   PropertyPages   =   "ListViewConsulta.ctx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   4110
   Begin MSComctlLib.ImageList il 
      Left            =   360
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListViewConsulta.ctx":0019
            Key             =   "uv"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListViewConsulta.ctx":340B
            Key             =   "u"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListViewConsulta.ctx":3565
            Key             =   "d"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ListViewConsulta.ctx":36BF
            Key             =   "dv"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "il"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblResumen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   ""
      Begin VB.Menu mnJ 
         Caption         =   "prueba"
      End
   End
End
Attribute VB_Name = "ListViewConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'para agregar una propiedad no olvidarse:
'procedimientos set/let/get
'read y writeProperty
'pagina de propiedades
'no olvidarse el PropertyChanged para q se guarde el estado

'fijarse si no hacen falta mas eventos(keypress, etc...)
'siempre se levanta el evento click (aunque haga dobleclick)
Public Event ItemClick(Item As Object)
Attribute ItemClick.VB_Description = "Devuelve un item al que se le hizo click. El tipo de item es el mismo que el de los elementos de la coleccion incluida."
Public Event ItemDblClick(Item As Object)
Attribute ItemDblClick.VB_Description = "Devuelve un item al que se le hizo doble click. El tipo de item es el mismo que el de los elementos de la coleccion incluida."
Public Event ItemKeyEnterPressed(Item As Object)
Public Event ItemGotFocus(Item As Object)
Public Event ItemCheck(Item As Object)
Public Event ItemEdited(Item As Object, pNewValue As String, pCancel As Boolean)
Attribute ItemEdited.VB_Description = "Ocurre cuando se edito un elemento. Aqui se debe validar el nuevo valor, asignarlo a la propiedad que corresponda. Si los datos no son correctos, establezca pCancel=True."
Public Event MouseMove(Item As Object, Button As Integer, Shift As Integer, X As Single, y As Single)
Public Event ItemDataBound(Item As Object, listItem As listItem)
Public Enum eOrientacionPagina
    eHorizontal = 1
    eVertical = 2
End Enum

Private WithEvents vEncabezados As LVCEncabezadoManager
Attribute vEncabezados.VB_VarHelpID = -1

Private vHideSelection As Boolean
Private vHideEncabezados As Boolean
Private vFullRowSelection As Boolean
Private vGridLines As Boolean
Private vAutoDistribuirColumnas As Boolean
Private vCampoKey As String
Private vAllowModify As Boolean
Private vMultiSelect As Boolean
Private vCheckBoxes As Boolean
'Private vShowFooter As Boolean hacer!

Private vColeccion As Object

Private vMenu As Menu
Private vImages As ImageList
Private vCampoImages As String 'nombre del campo cuyo valor coincide con los keys del list image

Dim xx As Single
Dim yy As Single

'--------------------propiedades-------------------------
Public Property Get ListImage() As Object
    Set ListImage = vImages
End Property

Public Property Set ListImage(pValue As Object)
    If Not pValue Is Nothing Then
        If TypeOf pValue Is ImageList Then
            Set vImages = pValue
        Else
            Set vImages = Nothing
        End If
    Else
        Set vImages = Nothing
    End If
    Set lv.SmallIcons = vImages
End Property

Public Property Get CampoKey() As String
Attribute CampoKey.VB_Description = "Devuelve o establece el nombre del campo identificador unico de la clase contenida en el ListView."
    CampoKey = vCampoKey
End Property

Public Property Let CampoKey(ByVal vNewValue As String)
    vCampoKey = vNewValue
    PropertyChanged
End Property

Public Property Get CampoImage() As String
    CampoImage = vCampoImages
End Property

Public Property Let CampoImage(ByVal vNewValue As String)
    vCampoImages = vNewValue
    PropertyChanged
End Property

Public Property Get HideSelect() As Boolean
Attribute HideSelect.VB_Description = "Idem a HideSelection de un listview."
    HideSelect = vHideSelection
End Property

Public Property Let HideSelect(var As Boolean)
     PropertyChanged
     vHideSelection = var
     lv.HideSelection = var
End Property

Public Property Get MultiSelect() As Boolean
    MultiSelect = vMultiSelect
End Property

Public Property Let MultiSelect(var As Boolean)
     PropertyChanged
     vMultiSelect = var
     lv.MultiSelect = var
End Property

Public Property Get ShowCheckBoxes() As Boolean
    ShowCheckBoxes = vCheckBoxes
End Property

Public Property Let ShowCheckBoxes(var As Boolean)
     PropertyChanged
     vCheckBoxes = var
     lv.Checkboxes = var
End Property

Public Property Get AutoDistribuirColumnas() As Boolean
    AutoDistribuirColumnas = vAutoDistribuirColumnas
End Property

Public Property Let AutoDistribuirColumnas(var As Boolean)
     PropertyChanged
     vAutoDistribuirColumnas = var
     If AutoDistribuirColumnas Then DistribuirColumnas
End Property

'Public Property Get ShowFooter() As Boolean
'    ShowFooter = vShowFooter
'End Property
'
'Public Property Let ShowFooter(var As Boolean)
'     PropertyChanged
'     vShowFooter = var
'     If ShowFooter Then MostrarPie
'End Property

Public Property Get Coleccion() As Object
Attribute Coleccion.VB_Description = "Esta es la lista de elementos que se mostraran en el listview."
   Set Coleccion = vColeccion
End Property

Public Property Set Coleccion(var As Object)
    Set vColeccion = var
    Refresh
End Property

Public Property Get Encabezados() As LVCEncabezadoManager
Attribute Encabezados.VB_Description = "Coleccion de los encabezados que se incluyen en el listview."
   Set Encabezados = vEncabezados
End Property

Public Property Set Encabezados(pValue As LVCEncabezadoManager)
   Set vEncabezados = pValue
   refrescarHeaders
End Property

Public Property Get HideEncabezados() As Boolean
    HideEncabezados = vHideEncabezados
End Property

Public Property Let HideEncabezados(ByVal vNewValue As Boolean)
    vHideEncabezados = vNewValue
    lv.HideColumnHeaders = vNewValue
End Property

Public Property Get FullRowSelection() As Boolean
    FullRowSelection = vFullRowSelection
End Property

Public Property Let FullRowSelection(ByVal vNewValue As Boolean)
    PropertyChanged
    vFullRowSelection = vNewValue
    lv.FullRowSelect = vNewValue
End Property

Public Property Get AllowModify() As Boolean
    AllowModify = vAllowModify
End Property

Public Property Let AllowModify(pvalor As Boolean)
    vAllowModify = pvalor
    PropertyChanged
End Property

Public Property Get GridLines() As Boolean
    GridLines = vGridLines
End Property

Public Property Let GridLines(ByVal vNewValue As Boolean)
    PropertyChanged
    vGridLines = vNewValue
    lv.GridLines = vNewValue
End Property

Public Property Get Picture() As StdPicture
    Set Picture = lv.Picture
End Property

Public Property Set Picture(pPicture As StdPicture)
    Set lv.Picture = pPicture
    PropertyChanged
End Property

Public Property Get SelectedItem() As Object
Attribute SelectedItem.VB_Description = "Devuelve una referencia al item seleccionado. El tipo de SelectedItem es el mismo que el de los objetos de  la coleccion incluida en el listview."
    If Not lv.SelectedItem Is Nothing Then
        Set SelectedItem = lv.SelectedItem.Tag
    End If
End Property

Public Property Set SelectedItem(pItem As Object)
    If Not pItem Is Nothing Then
        Dim li As listItem
        For Each li In lv.ListItems
            li.Selected = False
            If CallByName(li.Tag, vCampoKey, VbGet) = CallByName(pItem, vCampoKey, VbGet) Then
                li.Selected = True
                RaiseEvent ItemGotFocus(li.Tag)
                Exit For
            End If
        Next
    End If
End Property

Public Property Get CheckedItems() As Collection
    If vCheckBoxes Then
        Dim c As New Collection
        Dim li As listItem
        For Each li In lv.ListItems
            If li.Checked Then c.Add li.Tag
        Next
        Set CheckedItems = c
    Else
        MsgBox "No esta habilitada la opcion ShowCheckBoxes"
    End If
End Property

Public Property Set CheckedItems(col As Collection)
    Dim li As listItem
    For Each li In lv.ListItems
        For Each obj In col
            If CallByName(li.Tag, vCampoKey, VbGet) = CallByName(obj, vCampoKey, VbGet) Then
                li.Checked = True
            End If
        Next
    Next
End Property

Public Property Get SelectedItems() As Collection
    If vMultiSelect Then
        Dim c As New Collection
        Dim li As listItem
        For Each li In lv.ListItems
            If li.Selected Then c.Add li.Tag
        Next
        Set SelectedItems = c
'    Else
'        MsgBox "No esta habilitada la opcion Multiselect"
    End If
End Property


Public Property Get MenuPopUp() As Object
    Set MenuPopUp = vMenu
End Property

Public Property Set MenuPopUp(pMenu As Object)
    If TypeOf pMenu Is Menu Then
        Set vMenu = pMenu
    Else
        Err.Raise 13 'no coinciden los tipos
    End If
End Property

'------------------------funciones publicas---------------------------
Public Sub CheckAll()
    If vCheckBoxes Then
        Dim li As listItem
        For Each li In lv.ListItems
            li.Checked = True
        Next
    End If
End Sub

Public Sub CheckNone()
    If vCheckBoxes Then
        Dim li As listItem
        For Each li In lv.ListItems
            li.Checked = False
        Next
    End If
End Sub

Public Sub Editar()
    If vAllowModify Then
        If Not SelectedItem Is Nothing Then
            lv.StartLabelEdit
        End If
    End If
End Sub

Public Sub DistribuirColumnas()
    If Not Me.Encabezados Is Nothing Then
        ancho = lv.Width - 100 'el 100 es para q no aparezca la barra de abajo
        Dim enc As LVCEncabezado
        Dim i As Integer
        
        For Each enc In Me.Encabezados
            i = i + 1
            lv.ColumnHeaders(i).Width = ancho * enc.ancho / 100
        Next
    End If
'    For i = 0 To UBound(anchos)
'        lv.ColumnHeaders(i + 1).Width = ancho * anchos(i) / 100
'    Next

End Sub

Public Sub ActualizarAnchos()
'aca guardo si el usuario cambio los anchos
On Error GoTo errman:
Dim ch As ColumnHeader
ancho = lv.Width - 100
For Each ch In lv.ColumnHeaders
    ch.Tag.ancho = ch.Width * 100 / ancho
Next
Exit Sub
errman:

End Sub

Public Sub filtrar(cadena As String)
Attribute filtrar.VB_Description = "Filtra el contenido del listview. Para seleccionar que campos deben incluirse utilice la propiedad filtrar de cada encabezado."
    
Dim v As Object
Dim first As Boolean
Dim li As listItem
Dim enc As LVCEncabezado
first = True
lv.ListItems.Clear

For Each v In Coleccion
    If InStr(1, getCadenaAtributos(v), cadena, vbTextCompare) <> 0 Then

        For Each enc In Encabezados
             
            If first Then
100             Set li = lv.ListItems.Add(, , getValue(v, enc.miembro))
                Set li.Tag = v
                first = False
            Else
101             aux = getValue(v, enc.miembro)
                li.ListSubItems.Add , , aux
            End If
        Next
        RaiseEvent ItemDataBound(v, li)
    End If
    first = True
Next
If lv.ListItems.Count <> 0 Then RaiseEvent ItemGotFocus(lv.SelectedItem.Tag)
Exit Sub
e:
If Erl = 100 Or Erl = 101 Then
    Err.Raise 2008, , "Debe Implementar la interfaz IlvwConsulta" 'ya no la uso, ver q errores se presentan y adaptar
End If

End Sub

Public Sub Refresh()
On Error GoTo e:
'ver como hacer para no tener q eliminar todos los items porq pueden perderse referencias
If Not Coleccion Is Nothing Then
    Dim v As Object
    Dim first As Boolean
    Dim li As listItem
    Dim enc As LVCEncabezado
    first = True
    lv.ListItems.Clear
    
100 For Each v In Coleccion
    
        For Each enc In Encabezados
    
            If first Then
                Set li = lv.ListItems.Add(, , getValue(v, enc.miembro))
                Set li.Tag = v
                first = False
            Else
                aux = getValue(v, enc.miembro)
                li.ListSubItems.Add , , aux
            End If
            
        Next
        RaiseEvent ItemDataBound(v, li)
        first = True
        'muestro el listimage q corresponda
        If Not lv.SmallIcons Is Nothing And vCampoImages <> "" Then
            On Error Resume Next
            li.SmallIcon = getValue(v, vCampoImages)
            On Error GoTo e:
        End If
    Next
    If lv.ListItems.Count <> 0 Then RaiseEvent ItemGotFocus(lv.SelectedItem.Tag)
Else
    lv.ListItems.Clear
End If
Exit Sub
e:
     If el = 100 Then MsgBox "La coleccion no tiene la propiedad IDprocedimiento=-4"
     MsgBox Err.Description
End Sub

'------------------------funciones privadas------------------------------------

'Private Function getValue(var As Object, enc As LVCEncabezado) As String
'    On Error GoTo e
'    Dim mTipoLlamado As eTipoLlamado
'    mTipoLlamado = detectarMetodoLlamado(var, enc)
'    Select Case mTipoLlamado
'        Case eGetPropertyImplementado
'            getValue = var.GetProperty(enc.miembro)
'        Case eMethod
'            getValue = CallByName(var, enc.miembro, VbMethod)
'        Case ePropertyGet
'            getValue = CallByName(var, enc.miembro, VbGet)
'    End Select
'    Exit Function
'e: 'teoricamente nunca deberia llegar aca
'    Err.Raise 2010, , enc.miembro + "No se encuentra"
'End Function

'Private Function detectarMetodoLlamado(var As Object, enc As LVCEncabezado) As eTipoLlamado
''detecta de que tipo es el miembro del encabezado, si es property get, un metodo o la funcion getPROPERTY
'On Error Resume Next
'    'limpio si hay errores
'    Err.Clear
'    a = CallByName(var, enc.miembro, VbGet)
'    If Err.Number = 0 Then
'        detectarMetodoLlamado = ePropertyGet
'        Exit Function
'    End If
'
'    Err.Clear
'    a = CallByName(var, enc.miembro, VbMethod)
'    If Err.Number = 0 Then
'        detectarMetodoLlamado = eMethod
'        Exit Function
'    End If
'
'    Err.Clear
'    a = var.GetProperty(enc.miembro)
'
'    If Err.Number = 0 Then
'        detectarMetodoLlamado = eGetPropertyImplementado
'        Exit Function
'    End If
'
'    'si encuentra uno de los tres metodos no deberia llegar nunca a este punto
'    'y si llego significa q no encontro la propiedad
'    On Error GoTo 0
'    Err.Clear
'
'    Err.Raise 438 'no se encontro la propiedad o el metodo
'
'End Function

Private Function getCadenaAtributos(v As Object) As String
Dim cad As String
Dim enc As LVCEncabezado
For Each enc In vEncabezados
    If enc.filtrar Then
        'mTipoLlamado = detectarMetodoLlamado(v, enc)
        cad = cad + getValue(v, enc.miembro)
    End If
Next
getCadenaAtributos = cad
End Function

'-----------------eventos del control de usuario-----------------

Private Sub lv_AfterLabelEdit(Cancel As Integer, NewString As String)
    Dim pvalor As Boolean
    RaiseEvent ItemEdited(SelectedItem, NewString, pvalor)
    If pvalor Then Cancel = 1
End Sub

Private Sub lv_ColumnClick(ByVal ch As MSComctlLib.ColumnHeader)
    Dim chaux As ColumnHeader
    For Each chaux In lv.ColumnHeaders
        chaux.Icon = 0
    Next
    If lv.SortKey = ch.Index - 1 Then
        lv.SortOrder = IIf(lv.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        ch.Icon = IIf(lv.SortOrder = lvwAscending, "d", "u")
    Else
        ch.Icon = IIf(lv.SortOrder = lvwAscending, "d", "u")

        lv.SortKey = ch.Index - 1
    End If
    
    lv.Sorted = True
    lv.Sorted = False

End Sub

Private Sub lv_Click()

Dim li As listItem
    Set li = lv.HitTest(xx, yy)
    If Not li Is Nothing Then
        RaiseEvent ItemClick(li.Tag)
    End If
End Sub

Private Sub lv_DblClick()
    Dim li As listItem
    Set li = lv.HitTest(xx, yy)
    If Not li Is Nothing Then
        RaiseEvent ItemDblClick(li.Tag)
    End If
End Sub

Private Sub lv_ItemCheck(ByVal Item As MSComctlLib.listItem)
    RaiseEvent ItemCheck(Item.Tag)
End Sub

'levantarlo tambien despues de un refresh
Private Sub lv_ItemClick(ByVal Item As MSComctlLib.listItem)
    RaiseEvent ItemGotFocus(Item.Tag)
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not lv.SelectedItem Is Nothing Then
        RaiseEvent ItemKeyEnterPressed(lv.SelectedItem.Tag)
    End If
End If
End Sub

Private Sub lv_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        If Not vMenu Is Nothing Then UserControl.PopupMenu vMenu
    End If
End Sub

Private Sub lv_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
'If Button = vbRightButton Then UserControl.PopupMenu mnuPopUp
End Sub

Private Sub vEncabezados_EncabezadoAgregado()
    refrescarHeaders
End Sub

Private Sub refrescarHeaders()

    lv.ColumnHeaders.Clear
    
    Dim enc As LVCEncabezado
    For Each enc In Encabezados
        Set lv.ColumnHeaders.Add(, enc.miembro, enc.nombre).Tag = enc
    Next
    'If AutoDistribuirColumnas Then DistribuirColumnas
    PropertyChanged
End Sub

Private Sub lv_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    xx = X
    yy = y
    Dim li As listItem
    Set li = lv.HitTest(X, y)
    If Not li Is Nothing Then RaiseEvent MouseMove(li.Tag, Button, Shift, X, y)
End Sub

Private Sub UserControl_InitProperties()
    Set vEncabezados = New LVCEncabezadoManager
End Sub

Private Sub UserControl_Resize()
    lv.Height = UserControl.Height
    lv.Width = UserControl.Width
    
    If Me.AutoDistribuirColumnas Then DistribuirColumnas
End Sub

'--------------------Funciones para mantener el estado--------------------

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set vEncabezados = New LVCEncabezadoManager
    
    HideSelect = PropBag.ReadProperty("HideSelection", True)
    HideEncabezados = PropBag.ReadProperty("HideEncabezados", False)
    GridLines = PropBag.ReadProperty("GridLines", False)
    FullRowSelection = PropBag.ReadProperty("FullRowSelection", False)
    AutoDistribuirColumnas = PropBag.ReadProperty("AutoDistribuirColumnas", False)
    CampoKey = PropBag.ReadProperty("CampoKey", "id")
    AllowModify = PropBag.ReadProperty("AllowModify", False)
    ShowCheckBoxes = PropBag.ReadProperty("ShowCheckBoxes", False)
    MultiSelect = PropBag.ReadProperty("MultiSelect", False)
    Set Picture = PropBag.ReadProperty("Picture", lv.Picture)
    CampoImage = PropBag.ReadProperty("CampoImage", "")
    
    Dim nombre As String
    Dim miembro As String
    Dim ancho As Integer
    For i = 0 To 19
        nombre = PropBag.ReadProperty("NEncabezado" + Trim(Str(i)), "")
        miembro = PropBag.ReadProperty("MEncabezado" + Trim(Str(i)), "")
        ancho = CInt(PropBag.ReadProperty("AEncabezado" + Trim(Str(i)), 0))
        If nombre <> "" Then
            Encabezados.Add nombre, miembro, ancho
        End If
    Next
    
    refrescarHeaders
    If Me.AutoDistribuirColumnas Then DistribuirColumnas
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HideSelection", HideSelect
    PropBag.WriteProperty "HideEncabezados", HideEncabezados
    PropBag.WriteProperty "GridLines", GridLines
    PropBag.WriteProperty "FullRowSelection", FullRowSelection
    PropBag.WriteProperty "AutoDistribuirColumnas", AutoDistribuirColumnas
    PropBag.WriteProperty "CampoKey", vCampoKey, "id"
    PropBag.WriteProperty "AllowModify", vAllowModify
    PropBag.WriteProperty "ShowCheckBoxes", vCheckBoxes
    PropBag.WriteProperty "MultiSelect", vMultiSelect
    PropBag.WriteProperty "Picture", lv.Picture
    PropBag.WriteProperty "CampoImage", vCampoImages
    
    Dim enc As LVCEncabezado
    Dim i As Integer
    For Each enc In Encabezados
        
        PropBag.WriteProperty "NEncabezado" + Trim(Str(i)), enc.nombre
        PropBag.WriteProperty "MEncabezado" + Trim(Str(i)), enc.miembro
        PropBag.WriteProperty "AEncabezado" + Trim(Str(i)), enc.ancho
        
        i = i + 1
    Next
    For i = j To 19
        PropBag.WriteProperty "NEncabezado" + Trim(Str(j)), ""
        PropBag.WriteProperty "MEncabezado" + Trim(Str(j)), ""
        PropBag.WriteProperty "AEncabezado" + Trim(Str(j)), 0
    Next
        
End Sub


'funciones para exportar el contenido

'esta funcion exporta la coleccion, incluyendo los items q estan filtrados y desordenados
'Public Sub ExportToWord(titulo As String, Optional pOrientacion As eOrientacionPagina = eVertical)
'If titulo = "" Then titulo = "Informe"
'
'    Dim wrd As Object 'As Word.Application
'    Dim doc As Object 'As Word.Document
'
'    Set wrd = GetApplication(eWord)
'    If Not wrd Is Nothing Then
'        Set doc = wrd.Documents.Add
'        'aca agrego la tabla
'        Dim filas As Integer
'        Dim columnas As Integer
'        If pOrientacion = eHorizontal Then doc.PageSetup.Orientation = 1
'        doc.Tables.Add doc.Range(), 1, 1
'        'manejarse con parametros
'        doc.Tables(1).Cell(1, 1).Range.InsertBefore (titulo)
'        doc.Tables(1).Cell(1, 1).Range.Font.Size = 16
'        doc.Tables(1).Cell(1, 1).Range.Font.Bold = True
'        doc.Tables(1).Cell(1, 1).Range.Paragraphs(1).Alignment = 1 ' wdAlignParagraphCenter
'        doc.Tables(1).Cell(1, 1).Range.Font.Underline = 2 ' wdUnderlineWords
'        doc.Tables(1).Cell(1, 1).Height = 25
'
'
'        filas = Coleccion.Count
'        columnas = Encabezados.Count
'        'filas +1 porq la primera es nombre de campos
'        doc.Tables.Add doc.Range(Len(titulo) + 2), filas + 1, columnas
'    '    Dim bor As Border
'    '    For Each bor In doc.Tables(1).Borders
'    '        If Not bor.Inside Then
'    '            bor.LineStyle = wdLineStyleSingle
'    '        End If
'    '    Next
'        j = 2  '1=titulo, 2=subtitulos
'
'        Dim obj As Object
'        Dim enc As LVCEncabezado
'        Dim i As Integer
'        i = 0
'
'        'calculo el ancho de la tabla
'        Dim ancho As Single
'        ancho = 0
'        For Each enc In Encabezados
'            i = i + 1
'            ancho = ancho + doc.Tables(1).Cell(j, i).Width
'            'doc.Tables(1).Cell(j, i).Width = ancho * enc.ancho / 100
'        Next
'        i = 0
'        For Each enc In Encabezados
'            i = i + 1
'            doc.Tables(1).Cell(j, i).Range.InsertAfter (enc.nombre)
'            doc.Tables(1).Cell(j, i).Range.Bold = 1
'            doc.Tables(1).Cell(j, i).Width = ancho * enc.ancho / 100
'        Next
'        i = 0
'        For Each obj In Coleccion
'            j = j + 1
'            For Each enc In Encabezados
'                i = i + 1
'                doc.Tables(1).Cell(j, i).Range.InsertAfter (getValue(obj, enc))
'                doc.Tables(1).Cell(j, i).Width = ancho * enc.ancho / 100
'            Next
'            i = 0
'        Next
'        wrd.Visible = True
'        wrd.Activate
'        Set wrd = Nothing
'        Set doc = Nothing
'    End If
'End Sub

'ver tema de los anchos!!! en los redondeos se deforman las columnas
'devuelve el documento generado
Public Function ExportToWord(titulo As String, Optional pOrientacion As eOrientacionPagina = eVertical, Optional pContentsFont As StdFont, Optional pTitleFont As StdFont) As Object
Attribute ExportToWord.VB_Description = "Exporta el contenido del ListViewConsulta a MS Word. Devuelve el documento generado."
If titulo = "" Then titulo = "Informe"
On Error GoTo errman
    Dim wrd As Object 'As Word.Application
    Dim doc As Object 'As Word.Document

    Set wrd = GetApplication(eWord)
    If Not wrd Is Nothing Then
        Set doc = wrd.Documents.Add
        'aca agrego la tabla
        Dim filas As Integer
        Dim columnas As Integer
        If pOrientacion = eHorizontal Then doc.PageSetup.Orientation = 1
        doc.Tables.Add doc.Range(), 1, 1
        Dim titleRange As Object
        Set titleRange = doc.Tables(1).Cell(1, 1).Range
        titleRange.InsertBefore (titulo)
        
        titleRange.Paragraphs(1).Alignment = 1 ' wdAlignParagraphCenter
        doc.Tables(1).Cell(1, 1).Height = 25


        filas = lv.ListItems.Count
        columnas = lv.ColumnHeaders.Count
        'filas +1 porq la primera es nombre de campos
        Dim t As Object
        Set t = doc.Tables.Add(doc.Range(Len(titulo) + 2), filas + 1, columnas)
        
        If Not pContentsFont Is Nothing Then
            t.Range.Font.Name = pContentsFont.Name
            t.Range.Font.Size = pContentsFont.Size
            t.Range.Font.Bold = pContentsFont.Bold
            t.Range.Font.Italic = pContentsFont.Italic
            t.Range.Font.Underline = pContentsFont.Underline
        End If
        'seteo las propiedades de la fuente del titulo
        If pTitleFont Is Nothing Then
            titleRange.Font.Name = "Times New Roman"
            titleRange.Font.Size = 16
            titleRange.Font.Bold = True
            titleRange.Font.Italic = False
            titleRange.Font.Underline = True
        Else
            titleRange.Font.Name = pTitleFont.Name
            titleRange.Font.Size = pTitleFont.Size
            titleRange.Font.Bold = pTitleFont.Bold
            titleRange.Font.Italic = pTitleFont.Italic
            titleRange.Font.Underline = pTitleFont.Underline
        End If
        
    '    Dim bor As Border
    '    For Each bor In doc.Tables(1).Borders
    '        If Not bor.Inside Then
    '            bor.LineStyle = wdLineStyleSingle
    '        End If
    '    Next
        j = 2  '1=titulo, 2=subtitulos

        Dim obj As listItem
        Dim enc As ColumnHeader
        Dim i As Integer
        i = 0

        'calculo el ancho de la tabla
        Dim ancho As Single
        Dim anchoLv As Single
       
        ancho = doc.Tables(1).Cell(1, 1).Width
        anchoLv = lv.Width
        
        i = 0
        For Each enc In lv.ColumnHeaders
            i = i + 1
            doc.Tables(1).Cell(j, i).Range.InsertAfter (enc.Text)
            doc.Tables(1).Cell(j, i).Range.Bold = 1
            'doc.Tables(1).Cell(j, i).Width = Round(ancho * (100 * enc.Width / anchoLv) / 100, 1)
        Next
        i = 0
        For Each obj In lv.ListItems
            j = j + 1
            doc.Tables(1).Cell(j, 1).Range.InsertAfter (obj.Text)
            ''doc.Tables(1).Cell(j, 1).Width = Round(ancho * (100 * lv.ColumnHeaders(1).Width / anchoLv) / 100, 1)
            
            For i = 1 To columnas - 1
                doc.Tables(1).Cell(j, i + 1).Range.InsertAfter (obj.ListSubItems(i).Text)
                'doc.Tables(1).Cell(j, i + 1).Width = Round(ancho * (100 * lv.ColumnHeaders(i).Width / anchoLv) / 100, 1)
            Next
            
        Next
        wrd.Visible = True
        wrd.Activate
        Set wrd = Nothing
        Set ExportToWord = doc
        Set doc = Nothing
    End If
    Exit Function
errman:
    MsgBox "Ocurrio un error intentando exportar el listado."
End Function

'esta exporta toda la coleccion
'Public Sub ExportToExcel(titulo As String)
'    Dim Obj_Excel As Object
'    Dim Obj_Libro As Object
'    Dim Obj_Hoja As Object
'    On Error GoTo ErrSub
'    filas = Coleccion.Count
'    columnas = Encabezados.Count
'
'    Dim i As Integer, j As Integer
'
'    Set Obj_Excel = GetApplication(eExcel)
'    If Not Obj_Excel Is Nothing Then
'        Set Obj_Libro = Obj_Excel.Workbooks.Add()
'
'        'Ponemos la aplicación excel visible
'        Obj_Excel.Visible = True
'
'        Dim enc As LVCEncabezado
'        'Hoja activa
'        Set Obj_Hoja = Obj_Excel.ActiveSheet
'
'        ' poner los caption
'        iCol = 0
'        For Each enc In Encabezados
'            iCol = iCol + 1
'            Obj_Hoja.Cells(1, iCol) = enc.nombre
'        Next
'        'Obj_Hoja.Cells(1, iCol) = enc.nombre
'        j = 1
'
'        Dim obj As Object
'        For Each obj In Coleccion
'        iCol = 0
'        j = j + 1
'            For Each enc In Encabezados
'                    iCol = iCol + 1
'                    Obj_Hoja.Cells(j, iCol) = getValue(obj, enc)
'            Next
'        Next
'        'Opcional : colocamos en negrita y de color rojo los enbezados en la hoja
'        Obj_Hoja.Rows(1).Font.Bold = True
'        'Obj_Hoja.Rows(1).Font.Color = vbRed
'
'        'Autoajustamos
'        Obj_Hoja.Columns("A:Z").AutoFit
'
'        'Eliminamos las variables de objeto excel
'        Set Obj_Hoja = Nothing
'        Set Obj_Libro = Nothing
'        Set Obj_Excel = Nothing
'    End If
'Exit Sub
'
''Error
'ErrSub:
'
'    MsgBox Err.Description, vbCritical
'    On Error Resume Next
'
'    Set Obj_Hoja = Nothing
'    Set Obj_Libro = Nothing
'    Set Obj_Excel = Nothing
'
'End Sub

'retorna la hoja
Public Function ExportToExcel(titulo As String, Optional rowOffset As Integer = 0) As Object
    Dim Obj_Excel As Object
    Dim Obj_Libro As Object
    Dim Obj_Hoja As Object
    On Error GoTo ErrSub
    filas = lv.ListItems.Count
    columnas = lv.ColumnHeaders.Count

    Dim i As Integer, j As Integer

    'para q empieze a escribir mas abajo
    j = rowOffset
    
    Set Obj_Excel = GetApplication(eExcel)
    If Not Obj_Excel Is Nothing Then
        Set Obj_Libro = Obj_Excel.Workbooks.Add()

        'Ponemos la aplicación excel visible
        Obj_Excel.Visible = True

        Dim enc As ColumnHeader
        'Hoja activa
        Set Obj_Hoja = Obj_Excel.ActiveSheet
        
        j = j + 1
        ' poner los caption
        For i = 1 To columnas
            Obj_Hoja.Cells(j, i) = lv.ColumnHeaders(i).Text
            Next
        'Obj_Hoja.Cells(1, iCol) = enc.nombre
       

        Dim obj As listItem
        For Each obj In lv.ListItems
            j = j + 1
            Obj_Hoja.Cells(j, 1) = obj.Text
            For i = 1 To columnas - 1
                Obj_Hoja.Cells(j, i + 1) = IIf(IsNumeric(obj.ListSubItems(i).Text), Replace(obj.ListSubItems(i).Text, ",", "."), obj.ListSubItems(i).Text)
            Next
        Next
        'Opcional : colocamos en negrita y de color rojo los enbezados en la hoja
        Obj_Hoja.Rows(rowOffset + 1).Font.Bold = True
        'Obj_Hoja.Rows(1).Font.Color = vbRed

        'Autoajustamos
        Obj_Hoja.Columns("A:Z").AutoFit

        'Eliminamos las variables de objeto excel
        Set ExportToExcel = Obj_Hoja
        Set Obj_Hoja = Nothing
        Set Obj_Libro = Nothing
        Set Obj_Excel = Nothing
    End If
Exit Function

'Error
ErrSub:

    MsgBox Err.Description, vbCritical
    On Error Resume Next

    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing

End Function

Public Function ExportToHtml(titulo As String) As String
    Dim res As String
    On Error GoTo ErrSub
    filas = lv.ListItems.Count
    columnas = lv.ColumnHeaders.Count
    res = "<html><body><table style='width: 100%'><tr><td colspan='" + Str(columnas) + "' class ='formheader'>" + titulo + "</td></tr>"
    
    Dim i As Integer, j As Integer
        
    Dim enc As ColumnHeader
            
    j = j + 1
    ' poner los caption
    res = res + "<tr>"
    For i = 1 To columnas
        res = res + "<td class='enc'>" + lv.ColumnHeaders(i).Text + "</td>"
    Next
    res = res + "</tr>"
    
    Dim obj As listItem
    For Each obj In lv.ListItems
        res = res + "<tr class='contents'>"
        j = j + 1
        res = res + "<td>" + obj.Text + "</td>"
        For i = 1 To columnas - 1
           res = res + "<td>" + IIf(IsNumeric(obj.ListSubItems(i).Text), Replace(obj.ListSubItems(i).Text, ",", "."), obj.ListSubItems(i).Text) + "</td>"
        Next
        res = res + "</tr>"
    Next

    res = res + "</table></body></html>"
    ExportToHtml = res
           
Exit Function

'Error
ErrSub:

    MsgBox Err.Description, vbCritical
    On Error Resume Next

End Function
    
'Public Function MakePropertyValue(cName, uValue) As Object
'Dim oStruct, oServiceManager As Object
'    Set oServiceManager = CreateObject("com.sun.star.ServiceManager")
'    Set oStruct = oServiceManager.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
'    oStruct.Name = cName
'    oStruct.Value = uValue
'    Set MakePropertyValue = oStruct
'End Function

Public Function ExportToOOWriter(titulo As String, Optional pContentsFont As StdFont, Optional pTitleFont As StdFont) As Object
If titulo = "" Then titulo = "Informe"
On Error GoTo errman
    
    Dim oSM                   'Root object for accessing OpenOffice from VB
    Dim oDesk, oDoc As Object 'First objects from the API
    Dim arg()                 'Ignore it for the moment !
        
    'Instanciate OOo : this line is mandatory with VB for OOo API
    Set oSM = GetApplication(eOpenOffice)
    If Not oSM Is Nothing Then
        
        'Create the first and most important service
        Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
      
        'Create a new doc
        Set oDoc = oDesk.loadComponentFromURL("private:factory/swriter", "_blank", 0, arg())
        Dim xTable As Object
        Set xTable = oDoc.createInstance("com.sun.star.text.TextTable")
        xTable.Initialize 1, 1
        xTable.HoriOrient = 0 'com::sun::star::text::HoriOrientation::NONE
        xTable.LeftMargin = 10
        xTable.RightMargin = 10
         
        'ver la orientacion dela pagina
         
        'aca inserto la tabla para el titulo
        Dim xTC As Object
        Set xTC = oDoc.GetText().createTextCursor()
        oDoc.GetText().insertTextContent xTC, xTable, False
              
        insertIntoCell "A1", titulo, xTable, pTitleFont
       
        'esta es la tabla del contenido
        Dim filas As Integer
        Dim columnas As Integer
        
        filas = lv.ListItems.Count
        columnas = lv.ColumnHeaders.Count
        
        Set xTable = oDoc.createInstance("com.sun.star.text.TextTable")
        xTable.Initialize filas + 1, columnas 'filas +1 porq la primera es nombre de campos
        xTable.HoriOrient = 0 'com::sun::star::text::HoriOrientation::NONE
        xTable.LeftMargin = 10
        xTable.RightMargin = 10
        oDoc.GetText().insertTextContent xTC, xTable, False
        
        Dim j As Integer
        j = 1  '1=subtitulos

        Dim obj As listItem
        Dim enc As ColumnHeader
        Dim i As Integer
        i = 0

'        'calculo el ancho de la tabla
'        Dim ancho As Single
'        Dim anchoLv As Single
'
'        ancho = doc.Tables(1).Cell(1, 1).Width
'        anchoLv = lv.Width
        
        i = 0
        For Each enc In lv.ColumnHeaders
            i = i + 1
            insertIntoCell ToNamedCell(i, j), enc.Text, xTable, pContentsFont, True
            'doc.Tables(1).Cell(j, i).Range.Bold = 1
            'doc.Tables(1).Cell(j, i).Width = Round(ancho * (100 * enc.Width / anchoLv) / 100, 1)
        Next
        i = 0
        For Each obj In lv.ListItems
            j = j + 1
            insertIntoCell ToNamedCell(1, j), obj.Text, xTable, pContentsFont
            ''doc.Tables(1).Cell(j, 1).Width = Round(ancho * (100 * lv.ColumnHeaders(1).Width / anchoLv) / 100, 1)
            
            For i = 1 To columnas - 1
                insertIntoCell ToNamedCell(i + 1, j), obj.ListSubItems(i).Text, xTable, pContentsFont
                'doc.Tables(1).Cell(j, i + 1).Width = Round(ancho * (100 * lv.ColumnHeaders(i).Width / anchoLv) / 100, 1)
            Next
            
        Next
             
        Set ExportToOOWriter = oDoc
        'Set oDoc = Nothing
    End If
    Exit Function
errman:
    MsgBox "Ocurrio un error intentando exportar el listado."
End Function


Private Sub insertIntoCell(sCellName As String, sText As String, xTable As Object, pFont As StdFont, Optional pBold As Boolean = False)
' Access the XText interface of the cell referred to by sCellName
    Dim xCellText
    Set xCellText = xTable.getCellByName(sCellName)
       
    Dim xCellCursor
    Set xCellCursor = xCellText.createTextCursor()
 
    'propiedades de la fuente
    xCellCursor.charfontname = pFont.Name '"Tahoma"
    xCellCursor.charheight = pFont.Size '16
    'si me pasan true fuerzo negrita
    If pBold Then
        xCellCursor.charweight = 150
    Else
        xCellCursor.charweight = IIf(pFont.Bold, 150, 100) '100=normal, 150=bold
    End If
    xCellCursor.charunderline = IIf(pFont.Underline, 1, 0) '0=normal, 1=subrayado
    
    xCellText.setString sText
   
End Sub
    
Private Function ToNamedCell(pCol As Integer, pRow As Integer) As String
    Dim mCol As String
    mCol = Choose(pCol, "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
    ToNamedCell = mCol + Trim(Str(pRow))
End Function


Public Sub ExportToOOCalc(titulo As String)
    On Error GoTo ErrSub
    filas = lv.ListItems.Count
    columnas = lv.ColumnHeaders.Count
    
    Dim oSM                   'Root object for accessing OpenOffice from VB
    Dim oDesk As Object, oDoc As Object 'First objects from the API
    Dim oSheet As Object
    Dim arg()
    Dim i As Integer, j As Integer

    Set oSM = GetApplication(eOpenOffice)
    If Not oSM Is Nothing Then
        Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
      
        'Create a new doc
        Set oDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg())
                
        Set oSheet = oDoc.getSheets().getByIndex(0)
        
        Dim enc As ColumnHeader
        'Hoja activa
        
        ' poner los caption
        For i = 1 To columnas
            oSheet.getcellbyposition(i, 1).setFormula (lv.ColumnHeaders(i).Text)
            oSheet.getcellbyposition(i, 1).SetPropertyValue "HoriJustify", 2
            oSheet.getcellbyposition(i, 1).SetPropertyValue "VertJustify", 2
            'no funciona oSheet.getcellbyposition(1, 0).charweight = 150
        Next
        j = 1

        Dim obj As listItem
        For Each obj In lv.ListItems
            j = j + 1
            oSheet.getcellbyposition(1, j).setFormula (obj.Text)
            oSheet.getcellbyposition(1, j).SetPropertyValue "HoriJustify", 2
            oSheet.getcellbyposition(1, j).SetPropertyValue "VertJustify", 2
            For i = 1 To columnas - 1
                oSheet.getcellbyposition(i + 1, j).setFormula (obj.ListSubItems(i))
                oSheet.getcellbyposition(i + 1, j).SetPropertyValue "HoriJustify", 2
                oSheet.getcellbyposition(i + 1, j).SetPropertyValue "VertJustify", 2
            Next
        Next
        
        Set oSM = Nothing
        Set oDesk = Nothing
        Set oDoc = Nothing
    End If
Exit Sub

'Error
ErrSub:

    MsgBox Err.Description, vbCritical
    On Error Resume Next

    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing

End Sub

Public Sub ExportToCSV(Optional pPath As String = "")
    
    On Error GoTo ErrSub
    Dim contenido As String
    Dim i As Integer, j As Integer
       
    Dim enc As ColumnHeader
    filas = lv.ListItems.Count
    columnas = lv.ColumnHeaders.Count
    ' poner los caption
    For i = 1 To columnas
        contenido = contenido + lv.ColumnHeaders(i).Text + ";"
    Next
    'le saco la ultima coma
    contenido = Mid(contenido, 1, Len(contenido) - 1)
    contenido = contenido + vbCrLf
    
    j = 1

    Dim obj As listItem
    For Each obj In lv.ListItems
        j = j + 1
        contenido = contenido + obj.Text + ";"
        For i = 1 To columnas - 1
           contenido = contenido + obj.ListSubItems(i) + ";"
        Next
        contenido = Mid(contenido, 1, Len(contenido) - 1)
        contenido = contenido + vbCrLf
    Next
    If pPath = "" Then
        Dim cd As New CommonDialog
        cd.Filter = "Delimitado por comas (*.csv)|*.csv"
        cd.DefaultExt = "csv"
        cd.ShowSave
        If cd.FileName <> "" Then
            EscribirArchivo cd.FileName, contenido
        End If
    End If
Exit Sub

'Error
ErrSub:

    MsgBox Err.Description, vbCritical
    
End Sub

