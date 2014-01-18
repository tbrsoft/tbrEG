VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarLiqEmpleado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidaciones Empleados"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   8925
   Begin VB.CommandButton cmdResumen 
      Caption         =   "Exportar Resumen"
      Height          =   375
      Left            =   6600
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin ControlesPOO.ListViewConsulta lvwPeriodos 
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   8916
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      CampoKey        =   ""
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Periodo"
      MEncabezado0    =   "tolongstring"
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
   Begin ControlesPOO.ListViewConsulta lvwDetalle 
      Height          =   4575
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8070
      HideSelection   =   -1  'True
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Empleado"
      MEncabezado0    =   "nombreempleado"
      AEncabezado0    =   35
      NEncabezado1    =   "Monto"
      MEncabezado1    =   "monto"
      AEncabezado1    =   20
      NEncabezado2    =   "Adelanto"
      MEncabezado2    =   "adelanto"
      AEncabezado2    =   20
      NEncabezado3    =   "Plus"
      MEncabezado3    =   "plus"
      AEncabezado3    =   20
      NEncabezado4    =   "Observaciones"
      MEncabezado4    =   "observaciones"
      AEncabezado4    =   50
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
   Begin VB.Label lblPeriodo 
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
      Left            =   4080
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Detalle del periodo:"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmConsultarLiqEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdResumen_Click()

If Not lvwDetalle.Coleccion Is Nothing Then
    ExportarLiqEmpleados
Else
    MsgBox "Debe seleccionar un periodo para poder emitir el resumen.", vbOKOnly + vbInformation, "tbr Emergency Group"
End If

End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon

Dim periodos As New blcemi.MonthYearManager
periodos.LoadFromLiqEmpleados
Set lvwPeriodos.Coleccion = periodos

End Sub

Private Sub lvwPeriodos_ItemClick(Item As Object)
Dim liqs As New blcemi.LiqEmpleadoManager
liqs.LoadByPeriodo Item
Set lvwDetalle.Coleccion = liqs
lblPeriodo = Item.ToLongString
End Sub

Private Sub ExportarLiqEmpleados()
    Dim Obj_Excel As Object
    Dim Obj_Libro As Object
    Dim Obj_Hoja As Object
    Dim obj_HojaResumen As Object
    
    Set Obj_Excel = CreateObject("excel.application")
    If Not Obj_Excel Is Nothing Then
        Set Obj_Libro = Obj_Excel.Workbooks.Add()
        Set obj_HojaResumen = Obj_Excel.activesheet
        'lleno la info del resumen
        obj_HojaResumen.Name = "Resumen"
        obj_HojaResumen.cells(2, 2) = "Empleado"
        obj_HojaResumen.cells(2, 3) = "Monto"
        obj_HojaResumen.cells(2, 4) = "Adelanto"
        obj_HojaResumen.cells(2, 5) = "Plus"
        obj_HojaResumen.cells(2, 6) = "Total"
        
        obj_HojaResumen.Rows(2).Font.Bold = True

        Dim I As Integer
        I = 3
        Dim liqE As blcemi.LiquidacionEmpleado
        For Each liqE In lvwDetalle.Coleccion
            Set Obj_Hoja = Obj_Excel.sheets.Add
            
            LlenarHoja Obj_Hoja, liqE
            obj_HojaResumen.cells(I, 2) = liqE.GetEmpleado.NombreCompleto
            obj_HojaResumen.cells(I, 3) = liqE.Monto
            obj_HojaResumen.cells(I, 4) = liqE.Adelanto
            obj_HojaResumen.cells(I, 5) = liqE.Plus
            obj_HojaResumen.cells(I, 6) = liqE.GetSaldo
            I = I + 1
        Next
        
        
        
        obj_HojaResumen.cells(I, 3).Formula = "=Sum(C3:C" + Trim(Str(I - 1))
        obj_HojaResumen.cells(I, 4).Formula = "=Sum(D3:D" + Trim(Str(I - 1))
        obj_HojaResumen.cells(I, 5).Formula = "=Sum(E3:E" + Trim(Str(I - 1))
        obj_HojaResumen.cells(I, 6).Formula = "=Sum(F3:F" + Trim(Str(I - 1))
              
        obj_HojaResumen.Activate
         'Ponemos la aplicación excel visible
        Obj_Excel.Visible = True
        
        'Eliminamos las variables de objeto excel
        Set Obj_Hoja = Nothing
        Set Obj_Libro = Nothing
        Set Obj_Excel = Nothing
    End If
End Sub

Public Sub LlenarHoja(pHoja, pLiq As blcemi.LiquidacionEmpleado)
    Dim g As blcemi.guardia
    Dim mEmpleado As blcemi.Empleado
    Dim mGuardias As blcemi.guardiaManager
    Set mEmpleado = pLiq.GetEmpleado
    Set mGuardias = pLiq.Detalle
    
    pHoja.Name = mEmpleado.NombreCompleto

    
    pHoja.range("A1:E1").merge
    Bordes pHoja.range("A1:E1")
    pHoja.range("A1:E1").HorizontalAlignment = -4108 'xlCenter
    pHoja.cells(1, 1) = mEmpleado.NombreCompleto
        
    pHoja.cells(3, 2) = "Monto"
    pHoja.cells(3, 3) = "Adelanto"
    pHoja.cells(3, 4) = "Plus"
    pHoja.cells(3, 5) = "Total"
    
    pHoja.cells(4, 2) = CCur(Replace(pLiq.Monto, ".", ","))
    pHoja.cells(4, 3) = CCur(Replace(pLiq.Adelanto, ".", ","))
    pHoja.cells(4, 4) = CCur(Replace(pLiq.Plus, ".", ","))
    pHoja.cells(4, 5) = pLiq.GetSaldo 'reemplazar por la funcion
    
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
    
    If mGuardias.Count <> 0 Then
        For Each g In mGuardias
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
    Else
        pHoja.range("A8:E8").HorizontalAlignment = -4108 'xlCenter
        pHoja.range("A8:E8").merge
        pHoja.cells(8, 1) = "No Existen Guardias Registradas."
    End If
    'colocamos en negrita los enbezados en la hoja
    pHoja.Rows(1).Font.Bold = True
    pHoja.Rows(I).Font.Bold = True
    pHoja.Rows(3).Font.Bold = True
    pHoja.Rows(6).Font.Bold = True
    pHoja.Rows(7).Font.Bold = True
    'Autoajustamos
    pHoja.Columns("A:Z").autofit
            
End Sub

