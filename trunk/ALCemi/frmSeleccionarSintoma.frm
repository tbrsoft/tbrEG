VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmSeleccionarSintoma 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ingrese las primeras letras del item buscado"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   5835
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtFiltro 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5775
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   9763
      HideSelection   =   0   'False
      HideEncabezados =   -1  'True
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   " "
      MEncabezado0    =   "Codigo"
      AEncabezado0    =   30
      NEncabezado1    =   " "
      MEncabezado1    =   "NombreCompuesto"
      AEncabezado1    =   70
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
End
Attribute VB_Name = "frmSeleccionarSintoma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event SintomaSeleccionado(pSintoma As BLCemi.Sintoma)
Public Event SeleccionCancelada()

Private Sub Form_Load()
    Set lvw.Coleccion = GBL.SintomasGBL
    lvw.Encabezados.Item("nombrecompuesto").filtrar = True
    lvw.Encabezados.Item("codigo").filtrar = True
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        RaiseEvent SeleccionCancelada
        Unload Me
    Case vbKeyA To vbKeyZ
        If ActiveControl.Name <> "txtFiltro" Then
            txtFiltro = txtFiltro + Chr$(KeyCode)
            txtFiltro.SelStart = Len(txtFiltro)
            txtFiltro.SetFocus
        End If
    Case 13
        If Not lvw.SelectedItem Is Nothing Then
            RaiseEvent SintomaSeleccionado(lvw.SelectedItem)
            Unload Me
        End If
    End Select
End Sub

Private Sub txtFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    lvw.SetFocus
End If
End Sub

Private Sub txtFiltro_Change()
    lvw.filtrar txtFiltro
End Sub

Private Sub lvw_ItemDblClick(Item As Object)
    RaiseEvent SintomaSeleccionado(Item)
    Unload Me
End Sub

Private Sub lvw_ItemKeyEnterPressed(Item As Object)
    RaiseEvent SintomaSeleccionado(Item)
    Unload Me
End Sub
