VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmRegistrarCobroXCobrador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   7260
   Begin VB.Frame fraTotales 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   7095
      Begin VB.Label lblNeto 
         Caption         =   "$0"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Ingreso Neto:"
         Height          =   195
         Left            =   5040
         TabIndex        =   11
         Top             =   240
         Width           =   960
      End
      Begin VB.Label lblComision 
         Caption         =   "$0"
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
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Comision Cobrador:"
         Height          =   195
         Left            =   2400
         TabIndex        =   9
         Top             =   240
         Width           =   1365
      End
      Begin VB.Label lblMontoTotal 
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
         Height          =   195
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "Total recaudado:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox picbotones 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3720
      ScaleHeight     =   495
      ScaleWidth      =   3495
      TabIndex        =   3
      Top             =   6120
      Width           =   3495
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Registrar Cobros"
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   1800
         TabIndex        =   4
         Top             =   0
         Width           =   1695
      End
   End
   Begin VB.TextBox txtComision 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "10"
      Top             =   120
      Width           =   495
   End
   Begin ControlesPOO.ListViewConsulta lvwCuotas 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8070
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   -1  'True
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Recibo"
      MEncabezado0    =   "NroRecibo"
      AEncabezado0    =   15
      NEncabezado1    =   "NroAfiliado"
      MEncabezado1    =   "nro"
      AEncabezado1    =   15
      NEncabezado2    =   "Apellido y Nombre"
      MEncabezado2    =   "nombre"
      AEncabezado2    =   40
      NEncabezado3    =   "Mes"
      MEncabezado3    =   "mes"
      AEncabezado3    =   10
      NEncabezado4    =   "Año"
      MEncabezado4    =   "ayear"
      AEncabezado4    =   10
      NEncabezado5    =   "Monto"
      MEncabezado5    =   "$monto"
      AEncabezado5    =   10
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
   Begin VB.Label Label1 
      Caption         =   "Porcentaje de Comision:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1785
   End
End
Attribute VB_Name = "frmRegistrarCobroXCobrador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mCobrador As blcemi.Empleado
Dim WithEvents mCuotas As blcemi.CuotaManager
Attribute mCuotas.VB_VarHelpID = -1

Public Sub MostrarListadoCuotasImpagas(pCobrador As blcemi.Empleado)
    Set mCobrador = pCobrador
    Me.Show
    Me.Caption = "Registrar cobros de " + pCobrador.NombreCompleto
    Set mCuotas = New blcemi.CuotaManager
    Dim c As blcemi.Cuota
    For Each c In GBL.CuotasByEstadoGBL(blcemi.eImpaga)
        If c.Cobrador.id = pCobrador.id Then mCuotas.AddItem c
    Next
    Set lvwCuotas.Coleccion = mCuotas
    cmdAceptar.Enabled = False
End Sub

Private Sub cmdAceptar_Click()
If DatosCorrectos Then
    If MsgBox("Esta seguro que desea registrar los pagos?", vbQuestion + vbOKCancel) = vbOK Then
        Dim c As blcemi.Cuota
        Dim cm As New blcemi.CuotaManager
        For Each c In lvwCuotas.CheckedItems
            c.RegistrarCobro UsuarioActual
            cm.AddItem c
        Next
                  
        Dim doc As Object
        Dim r As Object
        Dim t As Object
        
        Set lvwCuotas.Coleccion = cm
        
        Set doc = lvwCuotas.ExportToWord("Cobrador: " + mCobrador.NombreCompleto, , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont)
        Set r = doc.Tables(1).Rows.Add
        '           wdBorderTop=-1      wdLineStyleSingle =1
        r.cells.Borders(-1).LineStyle = 1
        Set t = doc.Tables.Add(r.range.Next(), 1, 3)
        
        t.Cell(r.Index + 1, 1).range.InsertAfter "Total Recaudado: " + lblMontoTotal
        t.Cell(r.Index + 1, 2).range.InsertAfter "Comision: " + lblComision
        t.Cell(r.Index + 1, 3).range.InsertAfter "Ingreso Neto: " + lblNeto
        
        'encabezado
        Set t = doc.Sections(1).Headers.Item(1).range.Tables.Add(doc.Sections(1).Headers.Item(1).range(), 1, 2)
        t.Cell(1, 1).range.InsertAfter ("Recibo de Cobranzas")
        t.Cell(1, 2).range.InsertAfter ("Fecha: " + Trim(Str(Date)))
        t.Cell(1, 2).range.Paragraphs(1).Alignment = 2 ' wdAlignParagraphRight

        Unload Me
    End If
End If
End Sub

Private Function DatosCorrectos() As Boolean
Dim aux As String
    If lvwCuotas.CheckedItems.Count = 0 Then
        'esto esta de mas porq desabilito el boton, pero por las dudas...
        aux = "Debe seleccionar alguna cuota!" + vbCrLf
    End If
    If TextBoxValidado(txtComision, eInteger) Then
        If CDbl(txtComision.Text) < 0 Or CDbl(txtComision.Text) > 100 Then
            aux = aux + "El porcentaje debe estar entre 0 y 100!" + vbCrLf
        End If
    Else
        aux = aux + "El porcentaje ingresado es incorrecto!" + vbCrLf
        DatosCorrectos = False
        Exit Function
    End If

    If aux <> "" Then
        MsgBox aux, vbInformation
        DatosCorrectos = False
    Else
        DatosCorrectos = True
    End If
End Function


Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Dim cantAt As Long
    Dim cantC As Long
    cantAt = GBL.GetCantidadRegistros("atencion")
    cantC = GBL.GetCantidadRegistros("cuota")
    If modo = eModoDemo Then
        If (cantAt > 800 Or cantC > 500) Then
            MsgBox "La version demo ha caducado. Consulte el manual de usuario para registrar el presente software.", vbExclamation
            Unload Me
        End If
    End If
    Set Me.Icon = MDI.Icon

End Sub

Private Sub mCuotas_HasChanged()
    'esto deberia hacerce dentro del control probablemente
    Dim colChecked As New Collection
    Dim c As blcemi.Cuota
    For Each c In lvwCuotas.CheckedItems
        colChecked.Add c
    Next
    lvwCuotas.Refresh
    Set lvwCuotas.CheckedItems = colChecked
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-cobradores"
End Function

Public Sub Refrescar()
    Dim colChecked As New Collection
    Dim c As blcemi.Cuota
    For Each c In lvwCuotas.CheckedItems
        colChecked.Add c
    Next
    mCuotas.Reload
    lvwCuotas.Refresh
    Set lvwCuotas.CheckedItems = colChecked
End Sub

Private Sub lvwCuotas_ItemCheck(Item As Object)
Dim c As blcemi.Cuota
Dim montoaux As Currency
Dim mComision As Currency

For Each c In lvwCuotas.CheckedItems
    montoaux = montoaux + c.Monto
Next
lblMontoTotal = "$" + Trim(Str(Round(montoaux, 2)))
mComision = (montoaux * Val(txtComision) / 100)
lblComision = "$" + Trim(Str(Round(mComision, 2)))
lblNeto = "$" + Trim(Str(Round(montoaux - mComision, 2)))
cmdAceptar.Enabled = (lvwCuotas.CheckedItems.Count <> 0)
End Sub

Private Sub txtComision_Change()
    On Error Resume Next
    lvwCuotas_ItemCheck Nothing 'para q refresque
End Sub

Private Sub txtComision_KeyPress(KeyAscii As Integer)
    SoloNumeros KeyAscii
End Sub
