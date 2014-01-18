VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmLiquidacionEmpresa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidacion de Servicios a Empresa"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8010
   Begin VB.Frame Frame2 
      Caption         =   "Periodo"
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   7815
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Año:"
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Mes:"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Empresa"
      Height          =   615
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Doble  click aqui para ver detalles de la empresa"
      Top             =   120
      Width           =   7815
      Begin VB.Label lblIVA 
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
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblEmpresa 
         Caption         =   "Label1"
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   6015
      End
      Begin VB.Label Label4 
         Caption         =   "IVA:           %"
         Height          =   255
         Left            =   6480
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CheckBox chkInforme 
      Caption         =   "Emitir informe."
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   6960
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   7435
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
      MEncabezado3    =   "iva"
      AEncabezado3    =   20
      NEncabezado4    =   "Subtotal"
      MEncabezado4    =   "subtotal"
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
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6360
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   6840
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "300"
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
      Left            =   6960
      TabIndex        =   6
      Top             =   6240
      Width           =   855
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   7920
      Y1              =   6730
      Y2              =   6730
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      X1              =   120
      X2              =   7920
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Label Label3 
      Caption         =   "Total: $"
      Height          =   255
      Left            =   6120
      TabIndex        =   5
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Detalle servicios:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
End
Attribute VB_Name = "frmLiquidacionEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mOS As blcemi.ObraSocial
Dim mSE As blcemi.ServicioEmergencia
Dim infoCMan As blcemi.InfoContableManager
Dim mInfoCEmp As blcemi.InfoContableEmp
Dim mTotal As Currency

Private Sub cmdAceptar_Click()
    Dim mLiquidaciones As New blcemi.LiqEmpresaManager
    Dim mMes As Integer
    Dim mYear As Integer
    mMes = cmbMes.ListIndex + 1
    mYear = CInt(cmbYear.Text)
    
    If Not mOS Is Nothing Then
        mLiquidaciones.Nuevo infoCMan, False, Date, blcemi.eDCObraSocial, mOS.id, mMes, mYear
    Else
        mLiquidaciones.Nuevo infoCMan, False, Date, blcemi.eDCServicioEmergencia, mSE.id, mMes, mYear
    End If
    
    If chkInforme.Value = Checked Then
        cmdAceptar.Enabled = False
        cmdCancelar.Caption = "Cerrar"
        If Not mOS Is Nothing Then
            EmitirInforme mOS.Nombre
        Else
            EmitirInforme mSE.Nombre
        End If
    Else
        Unload Me
    End If
    
End Sub

Private Sub EmitirInforme(pEmpresa As String)
    'es una hoja del excel
    Dim ObjHoja As Object
    Dim ultimaFila As Integer
    ultimaFila = lvw.Coleccion.Count + 6 '4 porq incluye el encabezado
    Set ObjHoja = lvw.ExportToExcel("", 6)
            
    ObjHoja.cells(1, 1) = "Liquidacion de Servicios"
    UnirYCentrar ObjHoja.range("A1:E1")
    
    ObjHoja.cells(3, 1) = "Periodo:"
    UnirYCentrar ObjHoja.range("B3")
    Dim p As New blcemi.MonthYear
    p.Month = cmbMes.ListIndex + 1
    p.Year = CInt(cmbYear.Text)
    ObjHoja.cells(3, 2) = p.ToShortString
    
    ObjHoja.cells(3, 4) = "Fecha:"
    UnirYCentrar ObjHoja.range("E3")
    ObjHoja.cells(3, 5) = Date
    
    ObjHoja.cells(5, 1) = "Empresa:"
    UnirYCentrar ObjHoja.range("B5:E5")
    ObjHoja.cells(5, 2) = pEmpresa
    
    'recuadro empresa
    'Bordes ObjHoja.range("B5:E5"), eXLThick
    'titulo liquidacion
    'Bordes ObjHoja.range("A1:E1"), eXLThick
    'encabezados
    Bordes ObjHoja.range("A7:E" + Trim(Str(ultimaFila))), eXLThick
    ObjHoja.cells(ultimaFila + 2, 1) = "Bruto:"
    ObjHoja.cells(ultimaFila + 2, 4) = "Total:"
    ObjHoja.cells(ultimaFila + 2, 2).Formula = "=Sum(B8:B" + Trim(Str(ultimaFila + 1)) + ")"
    ObjHoja.cells(ultimaFila + 2, 5).Formula = "=Sum(E8:E" + Trim(Str(ultimaFila + 1)) + ")"
    ObjHoja.Rows(ultimaFila + 2).Font.Bold = True
    'totales
    Bordes ObjHoja.range("A" + Trim(Str(ultimaFila + 2)) + ":E" + Trim(Str(ultimaFila + 2))), eXLThick
    'contenido
    Bordes ObjHoja.range("A8:E" + Trim(Str(ultimaFila + 1))), eXLThick, True
    
    ObjHoja.Columns("A:Z").autofit
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-empresas"
End Function

Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
    mTotal = 0
    
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


Public Sub LiquidarServiciosOS(pos As blcemi.ObraSocial)
    Set mOS = pos
    lblEmpresa = pos.Nombre
    Set infoCMan = New blcemi.InfoContableManager
    infoCMan.LoadNoRendidosByEmpresa mOS.id, blcemi.eDCObraSocial
    Set mInfoCEmp = pos.InfoContable
    Set lvw.Coleccion = infoCMan
    lblTotal = mTotal
    lblIVA = mInfoCEmp.IVA
    Me.Show
End Sub

Public Sub LiquidarServiciosSE(pSE As blcemi.ServicioEmergencia)
    Set mSE = pSE
    
    lblEmpresa = pSE.Nombre
    Set infoCMan = New blcemi.InfoContableManager
    infoCMan.LoadNoRendidosByEmpresa mSE.id, blcemi.eDCServicioEmergencia
    Set mInfoCEmp = pSE.InfoContable
    Set lvw.Coleccion = infoCMan
    'mTotal se carga en el databound
    lblTotal = mTotal
    lblIVA = mInfoCEmp.IVA
    Me.Show
End Sub


Private Sub Frame1_DblClick()
    If Not mOS Is Nothing Then
        Dim frmABMOS As New frmABMObraSocial
        frmABMOS.VerDatos mOS
    Else
        Dim frmABMSE As New frmABMServicioEmergencia
        frmABMSE.VerDatos mSE
    End If
End Sub

Private Sub lblEmpresa_DblClick()
    Frame1_DblClick
End Sub

Private Sub lvw_ItemDataBound(Item As Object, listItem As MSComctlLib.listItem)
    Dim infoC As blcemi.InfoContable
    Set infoC = Item
    listItem.ListSubItems(3).Text = infoC.GetIva(mInfoCEmp)
    listItem.ListSubItems(4).Text = infoC.Servicio - infoC.MontoAbonado + infoC.GetIva(mInfoCEmp)
    mTotal = mTotal + infoC.GetSaldo
End Sub
