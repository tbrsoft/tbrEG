VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmListadoCuotas 
   Caption         =   "Listado de Cuotas a Cobrar de "
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   7245
   Begin VB.Frame fraDetalles 
      Height          =   975
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      Begin VB.CommandButton cmdImprimir 
         Height          =   735
         Left            =   6360
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Imprime la planilla de cobranzas para el cobrador"
         Top             =   160
         Width           =   735
      End
      Begin VB.Label lblFechaLbl 
         Caption         =   "Fecha Emision:"
         Height          =   255
         Left            =   5040
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblFecha 
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
         Left            =   5040
         TabIndex        =   7
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label lblCobrador 
         AutoSize        =   -1  'True
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
         Left            =   2040
         TabIndex        =   6
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label Label2 
         Caption         =   "Cobrador:"
         Height          =   195
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lblMes 
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
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Listado de Cobranzas de: "
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4683
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Recibo"
      MEncabezado0    =   "NroRecibo"
      AEncabezado0    =   13
      NEncabezado1    =   "Apellido y Nombre"
      MEncabezado1    =   "nombre"
      AEncabezado1    =   28
      NEncabezado2    =   "Domicilio"
      MEncabezado2    =   "calle"
      AEncabezado2    =   14
      NEncabezado3    =   "Barrio"
      MEncabezado3    =   "barrio"
      AEncabezado3    =   15
      NEncabezado4    =   "Ciudad"
      MEncabezado4    =   "ciudad"
      AEncabezado4    =   15
      NEncabezado5    =   "Importe"
      MEncabezado5    =   "$monto"
      AEncabezado5    =   15
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
Attribute VB_Name = "frmListadoCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ANCHOMIN = 7400
Private Const ALTOMIN = 4200

Private mCobrador As blcemi.Empleado
Private mCuotas As blcemi.CuotaManager

Public Sub MostrarListadoDeCuotas(pCobrador As blcemi.Empleado, pCuotas As blcemi.CuotaManager, pMes As Integer, pYear As Integer)
Set mCuotas = pCuotas
Set lvw.Coleccion = mCuotas
Set mCobrador = pCobrador
Me.Show

Me.Caption = "Listado de Cobranzas"
lblCobrador = pCobrador.NombreCompleto
lblMes = MonthName(CLng(pMes)) + Str(pYear)
lblFecha = Date
End Sub

Private Sub cmdImprimir_Click()
    Dim doc As Object 'documento de word
    Dim t As Object
           
    Set doc = lvw.ExportToWord("Cobrador: " + mCobrador.NombreCompleto, , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont)
                               'wdHeaderFooterPrimary=1
    Set t = doc.Sections(1).Headers.Item(1).range.Tables.Add(doc.Sections(1).Headers.Item(1).range(), 1, 2)
    t.Cell(1, 1).range.InsertAfter ("Planilla de Cobranzas de: " + lblMes)
    t.Cell(1, 2).range.InsertAfter ("Fecha Emision: " + Trim(Str(Date)))
    t.Cell(1, 2).range.Paragraphs(1).Alignment = 2 ' wdAlignParagraphRight
    
    Dim tCont As Object 'Word.Table 'la tabla de contenidos
    Set tCont = doc.Tables(1)
    ancho = tCont.Cell(2, 1).Width - 40
    For I = 2 To tCont.Rows.Count
        tCont.Cell(I, 1).Width = 40
        tCont.Cell(I, 2).Width = tCont.Cell(I, 2).Width + ancho
    Next
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-cobradores"
End Function

Private Sub Form_Load()
    Set cmdImprimir.Picture = MDI.il32.ListImages("word").Picture
    Set Me.Icon = MDI.Icon

End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN
        
        lvw.Width = Me.Width
        lvw.Height = Me.ScaleHeight - lvw.Top
        fraDetalles.Width = Me.Width - 120
        cmdImprimir.Left = fraDetalles.Width - 800
        lblFecha.Left = cmdImprimir.Left - 1440
        lblFechaLbl.Left = lblFecha.Left
    End If
End Sub

