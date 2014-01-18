VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConsultarLiqEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Liquidaciones"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   9000
   Begin VB.CommandButton cmdResumen 
      Caption         =   "Exportar Resumen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin ControlesPOO.ListViewConsulta lvwLiquidaciones 
      Height          =   3135
      Left            =   2640
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5530
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
      NEncabezado1    =   "Empresa"
      MEncabezado1    =   "Empresa"
      AEncabezado1    =   35
      NEncabezado2    =   "Tipo"
      MEncabezado2    =   "tipotostring"
      AEncabezado2    =   30
      NEncabezado3    =   "Saldo"
      MEncabezado3    =   "saldo"
      AEncabezado3    =   15
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
   Begin ControlesPOO.ListViewConsulta lvwPeriodos 
      Height          =   7455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   13150
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
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   3375
      Left            =   2640
      TabIndex        =   6
      Top             =   4080
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5953
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
      NEncabezado0    =   "Fecha"
      MEncabezado0    =   "fecha"
      AEncabezado0    =   20
      NEncabezado1    =   "Servicio"
      MEncabezado1    =   "servicio"
      AEncabezado1    =   20
      NEncabezado2    =   "Abonado"
      MEncabezado2    =   "montoabonado"
      AEncabezado2    =   20
      NEncabezado3    =   "IVA"
      MEncabezado3    =   "montoiva"
      AEncabezado3    =   20
      NEncabezado4    =   "Subtotal"
      MEncabezado4    =   "getsaldo"
      AEncabezado4    =   20
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
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Liquidaciones del periodo:"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Detalle:"
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   3840
      Width           =   1335
   End
End
Attribute VB_Name = "frmConsultarLiqEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim liquidaciones As New blcemi.LiqEmpresaManager
Dim mPeriodo As blcemi.MonthYear

Private Sub cmdResumen_Click()
Exportar
End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon

Dim periodos As New blcemi.MonthYearManager
periodos.LoadFromLiqEmpresa
Set lvwPeriodos.Coleccion = periodos
End Sub

Private Sub lvwLiquidaciones_ItemClick(Item As Object)
Dim liqE As blcemi.LiquidacionEmpresa
Set liqE = Item
Set lvw.Coleccion = liqE.Detalle
End Sub

Private Sub lvwPeriodos_ItemClick(Item As Object)
Dim liqs As New blcemi.LiqEmpresaManager
liqs.LoadByPeriodo Item
Set lvwLiquidaciones.Coleccion = liqs
Set mPeriodo = Item
lblPeriodo = mPeriodo.ToLongString
cmdResumen.Enabled = True
End Sub

Private Sub Exportar()
    Dim Obj_Excel As Object
    Dim Obj_Libro As Object
    Dim Obj_Hoja As Object
    Dim obj_HojaResumen As Object
    Dim mFila As Integer
    Set Obj_Excel = CreateObject("excel.application")
    If Not Obj_Excel Is Nothing Then
        Set Obj_Libro = Obj_Excel.Workbooks.Add()
        Set obj_HojaResumen = Obj_Excel.activesheet
        obj_HojaResumen.Name = "Resumen"
        obj_HojaResumen.cells(1, 1) = "Resumen de Liquidaciones de Servicios"
        obj_HojaResumen.Rows(1).Font.Bold = True
        UnirYCentrar obj_HojaResumen.range("A1:D1")
        
        obj_HojaResumen.cells(3, 1) = "Periodo:"
        UnirYCentrar obj_HojaResumen.range("B3")
        
        obj_HojaResumen.cells(3, 2) = mPeriodo.ToShortString
        
        obj_HojaResumen.cells(3, 3) = "Fecha:"
        obj_HojaResumen.cells(3, 3).HorizontalAlignment = -4152 'rigth
        UnirYCentrar obj_HojaResumen.range("D3")
        obj_HojaResumen.cells(3, 4) = Date

        mFila = 5
        obj_HojaResumen.cells(mFila, 1) = "Fecha"
        obj_HojaResumen.cells(mFila, 2) = "Empresa"
        obj_HojaResumen.cells(mFila, 3) = "Tipo"
        obj_HojaResumen.cells(mFila, 4) = "Saldo"
        Bordes obj_HojaResumen.range("A5:D5"), eXLThick
        
        Dim liq As blcemi.LiquidacionEmpresa
        For Each liq In lvwLiquidaciones.Coleccion
            mFila = mFila + 1
            obj_HojaResumen.cells(mFila, 1) = liq.fecha
            obj_HojaResumen.cells(mFila, 2) = liq.GetProperty("empresa")
            obj_HojaResumen.cells(mFila, 3) = liq.GetProperty("tipotostring")
            obj_HojaResumen.cells(mFila, 4) = liq.total
            LlenarHoja Obj_Excel.sheets.Add, liq.Detalle, liq.GetProperty("empresa")
        Next
            'totales
        Dim r As String
        r = "A" + Trim(Str(mFila + 1)) + ":D" + Trim(Str(mFila + 1))
        Bordes obj_HojaResumen.range(r), eXLThick
        'contenido
        r = "A6:D" + Trim(Str(mFila))
        Bordes obj_HojaResumen.range(r), eXLThick, True

        
        obj_HojaResumen.Columns("A:Z").autofit
        Obj_Excel.Visible = True
    End If
End Sub

'esta funcion exporta los datos de una sola liquidacion

Private Sub LlenarHoja(ObjHoja, pDetalle As blcemi.InfoContableManager, pEmpresa As String)
    'es una hoja del excel
    Dim ultimaFila As Integer
    ultimaFila = 7
    ObjHoja.Name = pEmpresa
    
    '-----exporta el detalle------------
    ObjHoja.cells(ultimaFila, 1) = "Fecha"
    ObjHoja.cells(ultimaFila, 2) = "Servicio"
    ObjHoja.cells(ultimaFila, 3) = "Abonado"
    ObjHoja.cells(ultimaFila, 4) = "IVA"
    ObjHoja.cells(ultimaFila, 5) = "Subtotal"
    Dim infoC As blcemi.InfoContable
    For Each infoC In pDetalle
        ultimaFila = ultimaFila + 1
        ObjHoja.cells(ultimaFila, 1) = infoC.GetProperty("fecha")
        ObjHoja.cells(ultimaFila, 2) = infoC.Servicio
        ObjHoja.cells(ultimaFila, 3) = infoC.MontoAbonado
        ObjHoja.cells(ultimaFila, 4) = infoC.MontoIVA
        ObjHoja.cells(ultimaFila, 5) = infoC.GetSaldo
    Next
    '------hasta aca exporta detalle------------
    
    
    ObjHoja.cells(1, 1) = "Liquidacion de Servicios"
    UnirYCentrar ObjHoja.range("A1:E1")
    
    ObjHoja.cells(3, 1) = "Periodo:"
    UnirYCentrar ObjHoja.range("B3")
   
    ObjHoja.cells(3, 2) = mPeriodo.ToShortString
    
    ObjHoja.cells(3, 4) = "Fecha:"
    UnirYCentrar ObjHoja.range("E3")
    ObjHoja.cells(3, 5) = Date
    
    ObjHoja.cells(5, 1) = "Empresa:"
    UnirYCentrar ObjHoja.range("B5:E5")
    ObjHoja.cells(5, 2) = pEmpresa
    
    Bordes ObjHoja.range("A7:E" + Trim(Str(ultimaFila))), eXLThick
    ObjHoja.cells(ultimaFila + 1, 1) = "Bruto:"
    ObjHoja.cells(ultimaFila + 1, 4) = "Total:"
    ObjHoja.cells(ultimaFila + 1, 2).Formula = "=Sum(B8:B" + Trim(Str(ultimaFila)) + ")"
    ObjHoja.cells(ultimaFila + 1, 5).Formula = "=Sum(E8:E" + Trim(Str(ultimaFila)) + ")"
    ObjHoja.Rows(ultimaFila + 1).Font.Bold = True
    'totales
    Bordes ObjHoja.range("A" + Trim(Str(ultimaFila + 1)) + ":E" + Trim(Str(ultimaFila + 1))), eXLThick
    'contenido
    Bordes ObjHoja.range("A8:E" + Trim(Str(ultimaFila))), eXLThick, True
    
    ObjHoja.Columns("A:Z").autofit
End Sub

