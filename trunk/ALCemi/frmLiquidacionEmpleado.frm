VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmLiquidacionEmpleado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidacion Empleados"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   5310
   Begin VB.CheckBox chkEmitir 
      Caption         =   "Emitir informe."
      Height          =   255
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Emite un formulario de excel con los datos de la liquidacion"
      Top             =   8160
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Liquidacion"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   5055
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   615
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   7
         ToolTipText     =   "Valor de la guardia."
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtAdelanto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Text            =   "0"
         ToolTipText     =   "Es el dinero que ya se le pago en concepto de coseguros, etc.."
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txtPlus 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2520
         TabIndex        =   5
         Text            =   "0"
         ToolTipText     =   "Premios, extras, etc..."
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label9 
         Caption         =   "Año:"
         Height          =   255
         Left            =   3360
         TabIndex        =   21
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Observaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   8
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   " -                +                 ="
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Adelanto"
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Plus"
         Height          =   255
         Left            =   2520
         TabIndex        =   10
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Total:"
         Height          =   255
         Left            =   3840
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdLiquidar 
      Caption         =   "Emitir liquidacion"
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   8040
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3600
      TabIndex        =   3
      Top             =   8040
      Width           =   1575
   End
   Begin ControlesPOO.ListViewConsulta lvwDetalle 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8070
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Fecha"
      MEncabezado0    =   "fecha"
      AEncabezado0    =   20
      NEncabezado1    =   "Valor Guardia"
      MEncabezado1    =   "monto"
      AEncabezado1    =   24
      NEncabezado2    =   "Adelanto"
      MEncabezado2    =   "adelanto"
      AEncabezado2    =   18
      NEncabezado3    =   "Plus"
      MEncabezado3    =   "plus"
      AEncabezado3    =   18
      NEncabezado4    =   "Subtotal"
      MEncabezado4    =   "GetSaldo"
      AEncabezado4    =   20
      NEncabezado5    =   "Observaciones"
      MEncabezado5    =   "Observaciones"
      AEncabezado5    =   50
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
   Begin VB.Label Label8 
      Caption         =   "Empleado: "
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblEmpleado 
      Caption         =   "Paliza, Martin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label lblDetalle 
      Caption         =   "Detalle guardias:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "frmLiquidacionEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mEmpleado As blcemi.Empleado

Public Sub NuevaLiquidacion(pEmpleado As blcemi.Empleado)
    lblEmpleado = pEmpleado.NombreCompleto
    Set mEmpleado = pEmpleado
    Set lvwDetalle.Coleccion = pEmpleado.Guardias
    Dim g As blcemi.guardia
    
    'calculo los totales del mes
    Dim mAdelTotal As Currency
    Dim mPlusTotal As Currency
    Dim mMontoTotal As Currency
    
    For Each g In mEmpleado.Guardias
        mAdelTotal = mAdelTotal + g.Adelanto
        mMontoTotal = mMontoTotal + g.Monto
        mPlusTotal = mPlusTotal + g.Plus
    Next
    txtAdelanto = mAdelTotal
    txtMonto = mMontoTotal
    txtPlus = mPlusTotal
    Me.Show
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdLiquidar_Click()
    Dim liqs As New blcemi.LiqEmpleadoManager
    Dim mMes As Integer
    Dim mYear As Integer
    mMes = cmbMes.ListIndex + 1
    mYear = CInt(cmbYear.Text)
    
    liqs.Nueva mEmpleado.id, CCur(Replace(txtMonto, ".", ",")), CCur(Replace(txtAdelanto, ".", ",")), Date, CCur(Replace(txtPlus, ".", ",")), txtObservaciones, lvwDetalle.Coleccion, mMes, mYear
    If chkEmitir.Value = 1 Then ExportarLiquidacion
    
    'lo seteo en nothign para q las descargue
    Set mEmpleado.Guardias = Nothing
    Unload Me
End Sub

Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
    Me.Top = MDI.tBar.Top + 100
    
    'cargo los combos de fecha
    For I = 1 To 12
        cmbMes.AddItem MonthName(I)
        cmbMes.ItemData(cmbMes.NewIndex) = I
    Next
    cmbMes.ListIndex = Month(Date) - 1
    
    For j = 2000 To 2050
        cmbYear.AddItem j
    Next
    
    For k = 0 To 50
        If cmbYear.List(k) = Year(Date) Then
            cmbYear.ListIndex = k
            Exit For
        End If
    Next
    
End Sub

Private Sub lvwEmpleados_ItemClick(Item As Object)
    Dim e As blcemi.Empleado
    Set e = Item
    Set lvwDetalle.Coleccion = e.Guardias
End Sub

Private Sub cmdNinguno_Click()
    lvwEmpleados.CheckNone
End Sub

Private Sub cmdTodos_Click()
    lvwEmpleados.CheckAll
End Sub

Private Sub MostrarSaldo()
    On Error GoTo errman
    lblTotal = CCur(Replace(txtMonto, ".", ",")) - CCur(Replace(txtAdelanto, ".", ",")) + CCur(Replace(txtPlus, ".", ","))
    Exit Sub
errman:
    lblTotal = "-"
End Sub

Private Sub txtAdelanto_Change()
    MostrarSaldo
End Sub

Private Sub txtMonto_Change()
    MostrarSaldo
End Sub

Private Sub txtPlus_Change()
    MostrarSaldo
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub

Private Sub txtAdelanto_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub

Private Sub txtPlus_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub

Private Sub ExportarLiquidacion()
    Dim Obj_Excel As Object
    Dim Obj_Libro As Object
    Dim Obj_Hoja As Object
    Dim obj_HojaResumen As Object
    
    Set Obj_Excel = CreateObject("excel.application")
    If Not Obj_Excel Is Nothing Then
        Set Obj_Libro = Obj_Excel.Workbooks.Add()
        
        Set Obj_Hoja = Obj_Excel.activesheet
               
        LlenarHoja Obj_Hoja, mEmpleado, mEmpleado.Guardias
        
        'Ponemos la aplicación excel visible
        Obj_Excel.Visible = True
        
        'Eliminamos las variables de objeto excel
        Set Obj_Hoja = Nothing
        Set Obj_Libro = Nothing
        Set Obj_Excel = Nothing
    End If
End Sub

Public Sub LlenarHoja(pHoja, pEmpleado As blcemi.Empleado, pGuardias As blcemi.guardiaManager)
    Dim g As blcemi.guardia
    pHoja.Name = pEmpleado.NombreCompleto
    
    pHoja.range("A1:E1").merge
    Bordes pHoja.range("A1:E1")
    pHoja.range("A1:E1").HorizontalAlignment = -4108 'xlCenter
    pHoja.cells(1, 1) = pEmpleado.NombreCompleto
        
    pHoja.cells(3, 2) = "Monto"
    pHoja.cells(3, 3) = "Adelanto"
    pHoja.cells(3, 4) = "Plus"
    pHoja.cells(3, 5) = "Total"
    
    pHoja.cells(4, 2) = CCur(Replace(txtMonto, ".", ","))
    pHoja.cells(4, 3) = CCur(Replace(txtAdelanto, ".", ","))
    pHoja.cells(4, 4) = CCur(Replace(txtPlus, ".", ","))
    pHoja.cells(4, 5) = CCur(Replace(lblTotal, ".", ",")) 'reemplazar por la funcion
    
    pHoja.range("A6:E6").merge
    pHoja.range("A6:E6").HorizontalAlignment = -4108 'xlCenter
    Bordes pHoja.range("A6:E6")
    pHoja.cells(6, 1) = "Detalle de Guardias"

    pHoja.cells(7, 1) = "Dia"
    pHoja.cells(7, 2) = "Valor de Guardia"
    pHoja.cells(7, 3) = "Coseguro"
    pHoja.cells(7, 4) = "Plus"
    pHoja.cells(7, 5) = "Subtotal"
    
    Dim I As Integer
    I = 8
    
    For Each g In pGuardias
        pHoja.cells(I, 1) = g.fecha
        pHoja.cells(I, 2) = g.Monto
        pHoja.cells(I, 3) = g.Adelanto
        pHoja.cells(I, 4) = g.Plus
        pHoja.cells(I, 5) = g.GetSaldo
        I = I + 1
    Next
    Bordes pHoja.range("A8:E" + Trim(Str(I - 1)))

    pHoja.cells(I, 2).Formula = "=sum(B8:B" + Trim(Str(I - 1)) + ")"
    pHoja.cells(I, 3).Formula = "=sum(C8:C" + Trim(Str(I - 1)) + ")"
    pHoja.cells(I, 4).Formula = "=sum(D8:D" + Trim(Str(I - 1)) + ")"
    pHoja.cells(I, 5).Formula = "=sum(E8:E" + Trim(Str(I - 1)) + ")"
    
    'colocamos en negrita los enbezados en la hoja
    pHoja.Rows(1).Font.Bold = True
    pHoja.Rows(I).Font.Bold = True
    pHoja.Rows(3).Font.Bold = True
    pHoja.Rows(6).Font.Bold = True
    pHoja.Rows(7).Font.Bold = True
    'Autoajustamos
    pHoja.Columns("A:Z").autofit
            
End Sub
 

Public Sub ExportarLiquidacionToCalc()
    Dim g As blcemi.guardia
    On Error GoTo ErrSub
    
    Dim oSM                   'Root object for accessing OpenOffice from VB
    Dim oDesk As Object, oDoc As Object 'First objects from the API
    Dim pHoja As Object
    Dim arg()
    Dim I As Integer, j As Integer

    Set oSM = CreateObject("com.sun.star.ServiceManager")
    If Not oSM Is Nothing Then
        Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
        'Create a new doc
        Set oDoc = oDesk.loadComponentFromURL("private:factory/scalc", "_blank", 0, arg())
        Set pHoja = oDoc.getSheets().getByIndex(0)
        
        Dim enc As ColumnHeader
        'Hoja activa
        'poner los caption
    Set pEmpleado = mEmpleado
    Set pGuardias = mEmpleado.Guardias
        
        
    pHoja.getcellbyposition(0, 0).setFormula (pEmpleado.NombreCompleto)
    
        
    pHoja.getcellbyposition(1, 2).setFormula ("Monto")
    pHoja.getcellbyposition(2, 2).setFormula ("Adelanto")
    pHoja.getcellbyposition(3, 2).setFormula ("Plus")
    pHoja.getcellbyposition(4, 2).setFormula ("Total")
    
    pHoja.getcellbyposition(1, 3).setFormula (CCurrency(txtMonto))
    pHoja.getcellbyposition(2, 3).setFormula (CCurrency(txtAdelanto))
    pHoja.getcellbyposition(3, 3).setFormula (CCurrency(txtPlus))
    pHoja.getcellbyposition(4, 3).setFormula (CCurrency(lblTotal))
        
'    pHoja.range("A6:E6").merge
'    pHoja.range("A6:E6").HorizontalAlignment = -4108 'xlCenter
'    Bordes pHoja.range("A6:E6")
'    pHoja.cells(6, 1) = "Detalle de Guardias"
    
    

    pHoja.getcellbyposition(0, 6).setFormula ("Dia")
    pHoja.getcellbyposition(1, 6).setFormula ("Valor de Guardia")
    pHoja.getcellbyposition(2, 6).setFormula ("Coseguro")
    pHoja.getcellbyposition(3, 6).setFormula ("Plus")
    pHoja.getcellbyposition(4, 6).setFormula ("Subtotal")
    
    
    
    I = 7
    Dim aux As String
    
    For Each g In pGuardias
        aux = "'" + Trim(Str(g.fecha))
        pHoja.getcellbyposition(0, I).setFormula (aux)
     
        pHoja.getcellbyposition(1, I).setFormula (g.Monto)
        pHoja.getcellbyposition(2, I).setFormula (g.Adelanto)
        pHoja.getcellbyposition(3, I).setFormula (g.Plus)
        pHoja.getcellbyposition(4, I).setFormula (g.GetSaldo)
        I = I + 1
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

Public Function GetHelpContext() As String
    GetHelpContext = "empleados"
End Function
