VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmFiltroAtencionesB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Siniestros"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Fechas"
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5295
      Begin VB.OptionButton optCualquierFecha 
         Caption         =   "Cualquier fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optEntreDia 
         Caption         =   "Entre el dia:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optMes 
         Caption         =   "Del mes:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbMes 
         Height          =   315
         ItemData        =   "frmFiltroAtencionesB.frx":0000
         Left            =   2040
         List            =   "frmFiltroAtencionesB.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1320
         Width           =   3135
      End
      Begin VB.OptionButton optYear 
         Caption         =   "Del año:"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1680
         Width           =   975
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1680
         Width           =   3135
      End
      Begin MSComCtl2.DTPicker dtpHasta 
         Height          =   315
         Left            =   2040
         TabIndex        =   15
         Top             =   960
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45678592
         CurrentDate     =   39352
      End
      Begin MSComCtl2.DTPicker dtpDesde 
         Height          =   315
         Left            =   2040
         TabIndex        =   16
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         Format          =   45678592
         CurrentDate     =   39352
      End
      Begin VB.Label Label1 
         Caption         =   "y el:"
         Height          =   195
         Left            =   1080
         TabIndex        =   19
         Top             =   960
         Width           =   285
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Despachador"
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   5295
      Begin VB.OptionButton optCualquierDespachador 
         Caption         =   "Cualquiera"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optDespachador 
         Caption         =   "Solo las de:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin ControlesPOO.Combo cmbDespachador 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         AtributoAMostrar=   "nombrecompleto"
         Enabled         =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo de Siniestro"
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   5295
      Begin VB.OptionButton optCualquierCodigo 
         Caption         =   "Cualquiera"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton optFiltroPorCodigo 
         Caption         =   "Solo codigo:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   1695
      End
      Begin ControlesPOO.Combo cmbCodigo 
         Height          =   315
         Left            =   2040
         TabIndex        =   3
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         AtributoAMostrar=   "nombrecompuesto"
         Enabled         =   -1  'True
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Ver listado"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4800
      Width           =   2535
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   4800
      Width           =   2535
   End
End
Attribute VB_Name = "frmFiltroAtencionesB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()

Dim vFechaDesde As String
Dim vFechaHasta As String
Dim vMes As Integer
Dim vYear As Integer
Dim vDespachador As blcemi.Empleado
Dim vCodigo As blcemi.CodigoEmergencia
Dim vDescripcion As String

vDescripcion = "Listado de Emergencias"

'filtro por fecha
If optCualquierFecha.Value = True Then

ElseIf optEntreDia.Value = True Then
    vFechaDesde = Str(dtpDesde.Value)
    vFechaHasta = Str(dtpHasta.Value)
    vDescripcion = vDescripcion + " entre el dia " + vFechaDesde + " hasta el dia " + vFechaHasta
ElseIf optMes.Value = True Then
    vMes = cmbMes.ListIndex + 1
    vDescripcion = vDescripcion + " de " + cmbMes.Text
ElseIf optYear.Value = True Then
    vYear = Val(cmbYear.Text)
    vDescripcion = vDescripcion + " de " + cmbYear.Text
End If

'filtro por despachador
If optDespachador.Value = True Then
    Set vDespachador = cmbDespachador.SelectedItem
    If vDespachador Is Nothing Then
        MsgBox "Seleccione un Despachador!", vbInformation
        Exit Sub
    End If
    vDescripcion = vDescripcion + ", despachado por " + cmbDespachador.SelectedItem.NombreCompleto
End If

'filtro por codigo de emergencia
If optFiltroPorCodigo.Value = True Then
    Set vCodigo = cmbCodigo.SelectedItem
    If vCodigo Is Nothing Then
        MsgBox "Seleccione un Codigo de Emergencia!", vbInformation
        Exit Sub
    End If
    vDescripcion = vDescripcion + ", del tipo " + cmbCodigo.SelectedItem.nombrecompuesto
End If

Dim frm As New frmListadoAtenciones
vDescripcion = vDescripcion + "."

frm.ConsultarB GBL.AtencionesBGBL.Filter(vFechaDesde, vFechaHasta, vMes, vYear, vDespachador, vCodigo), vDescripcion
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
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
    
    dtpHasta.Value = Date
    dtpDesde.Value = DateAdd("m", -1, Date)
    
    Set cmbDespachador.Coleccion = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eDespachador)
    Set cmbCodigo.Coleccion = GBL.CodigoEmergenciaGBL
    optCualquierFecha.Value = True 'porq modifico el txtyear y me marca el año
    Set Me.Icon = MDI.Icon

End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "filtrosiniestros"
End Function

Public Sub Refrescar()
    Set cmbDespachador.Coleccion = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eDespachador)
    cmbDespachador.Refresh
End Sub

Private Sub dtpHasta_Change()
    optEntreDia.Value = True
End Sub

Private Sub dtpDesde_Change()
    optEntreDia.Value = True
End Sub

Private Sub txtYear_Change()
optYear.Value = True
End Sub

Private Sub cmbCodigo_ItemSeleccionado(Item As Object)
    optFiltroPorCodigo.Value = True
End Sub

Private Sub cmbDespachador_ItemSeleccionado(Item As Object)
    optDespachador.Value = True
End Sub
Private Sub cmbMes_Click()
    optMes.Value = True
End Sub
