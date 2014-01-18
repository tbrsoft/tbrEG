VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmEmitirListadoCobros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emitir Listado de Cuotas a Cobrar"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3810
   Begin VB.ComboBox cmbYear 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Generar Listado"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1695
   End
   Begin ControlesPOO.ListViewConsulta lvwCobradores 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5741
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   -1  'True
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Cobrador"
      MEncabezado0    =   "nombrecompleto"
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
   Begin VB.ComboBox cmbMes 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Año:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Mes:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmEmitirListadoCobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAceptar_Click()
'validar año
If lvwCobradores.CheckedItems.Count = 0 Then
    MsgBox "Debe seleccionar al menos un Cobrador!", vbInformation
    Exit Sub
Else 'si esta todo bien...

    If MsgBox("¿Esta seguro que desea generar lAs blcemi.Cuotas?", vbQuestion + vbOKCancel) = vbOK Then
        Dim e As blcemi.Empleado
        Dim af As blcemi.Afiliado
        Dim ap As blcemi.AreaProtegida
        Dim aux As New blcemi.CuotaManager
        Dim nroRecibo As Long
        Dim mCuota As blcemi.Cuota
        Dim mMes As Integer
        Dim mYear As Integer
        mMes = cmbMes.ListIndex + 1
        mYear = CInt(cmbYear.Text)
        
        GBL.ResetearColecciones 'para que se vuelvan a cargar los afiliados, por cualquier cosa
        
        Dim frm As frmListadoCuotas
        For Each e In lvwCobradores.CheckedItems 'para cada empleado checkeado...
            Set frm = New frmListadoCuotas
            Set Afiliados = New blcemi.AfiliadoManager
        
            'separo los afiliados de este empleado...
            For Each af In GBL.AfiliadosGBL.GetAfiliadosTitulares
                If af.Cobrador.id = e.id Then
                    'me fijo si tiene generada la cuota...
                    If af.Cuotas.ItemByPeriodo(mMes, mYear) Is Nothing Then
                        nroRecibo = aux.GetUltimoNroRecibo
                        Set mCuota = af.Cuotas.Nuevo(af, Nothing, nroRecibo, mMes, mYear, af.Importe, UsuarioActual)
                        ImprimirReciboAfiliado2 mCuota, False
                        aux.AddItem mCuota
                    End If
                End If
            Next
            
            For Each ap In GBL.AreasProtegidasGBL
                If ap.Cobrador.id = e.id Then
                    'me fijo si tiene generada la cuota...
                    If ap.Cuotas.ItemByPeriodo(mMes, mYear) Is Nothing Then
                       nroRecibo = aux.GetUltimoNroRecibo
                       Set mCuota = ap.Cuotas.Nuevo(Nothing, ap, nroRecibo, mMes, mYear, ap.Importe, UsuarioActual)
                       ImprimirReciboAreaProtegida2 mCuota, False
                       aux.AddItem mCuota
                    End If
                End If
            Next
            'para q avise a la red
            
            Printer.EndDoc
            
            GBL.CuotasByEstadoGBL(blcemi.eImpaga).CuotasModificadasoAgregadas
            frm.MostrarListadoDeCuotas e, aux, mMes, mYear
            Unload Me
        Next
    End If
End If

End Sub

Private Sub cmdCancelar_Click()
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
    
    Set lvwCobradores.Coleccion = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eCobrador)
    lvwCobradores.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
    Set Me.Icon = MDI.Icon

End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-cobradores"
End Function

Public Sub Refrescar()
    Dim colChecked As New Collection
    Dim e As blcemi.Empleado
    For Each e In lvwCobradores.CheckedItems
        colChecked.Add e
    Next
    Set lvwCobradores.Coleccion = GBL.EmpleadosGBL.GetByCargoFijo(blcemi.eCobrador)
    Set lvwCobradores.CheckedItems = colChecked
    lvwCobradores.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesOtros
End Sub
