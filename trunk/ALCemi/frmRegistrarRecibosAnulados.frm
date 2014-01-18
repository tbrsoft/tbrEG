VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmRegistrarRecibosAnulados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Devolucion de Recibos Anulados"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6075
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   4560
      Width           =   2055
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Registrar Devolucion"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   4560
      Width           =   2055
   End
   Begin ControlesPOO.ListViewConsulta lvwCuotas 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7646
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   -1  'True
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "NroRecibo"
      MEncabezado0    =   "NroRecibo"
      AEncabezado0    =   20
      NEncabezado1    =   "Mes"
      MEncabezado1    =   "mes"
      AEncabezado1    =   10
      NEncabezado2    =   "Año"
      MEncabezado2    =   "ayear"
      AEncabezado2    =   10
      NEncabezado3    =   "Cobrador"
      MEncabezado3    =   "cobrador"
      AEncabezado3    =   40
      NEncabezado4    =   "Monto"
      MEncabezado4    =   "monto"
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
End
Attribute VB_Name = "frmRegistrarRecibosAnulados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mCuotas As blcemi.CuotaManager
Attribute mCuotas.VB_VarHelpID = -1

Public Sub MostrarListadoRecibosAnulados(pCobrador As blcemi.Empleado)
    'si le mando un cobrador me los filtra sino me muestra todos los recibos anulados
    If Not pCobrador Is Nothing Then
        Set mCuotas = GBL.CuotasByEstadoGBL(blcemi.ePedirRecibo).GetCuotasByCobrador(pCobrador)
    Else
        Set mCuotas = GBL.CuotasByEstadoGBL(blcemi.ePedirRecibo)
    End If
    
    Set lvwCuotas.Coleccion = mCuotas
    Me.Show
    cmdAceptar.Enabled = False
End Sub

Private Sub cmdAceptar_Click()

    If MsgBox("Esta seguro que desea registrar la devolucion de los recibos seleccionados?", vbQuestion + vbOKCancel) = vbOK Then
        Dim c As blcemi.Cuota
        For Each c In lvwCuotas.CheckedItems
            c.RegistrarDevolucion UsuarioActual
        Next
        Unload Me
    End If

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set Me.Icon = MDI.Icon

End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "cobro-cobradores"
End Function

Private Sub mCuotas_HasChanged()
    Dim colChecked As New Collection
    Dim c As blcemi.Cuota
    For Each c In lvwCuotas.CheckedItems
        colChecked.Add c
    Next
    Set lvwCuotas.Coleccion = mCuotas
    Set lvwCuotas.CheckedItems = colChecked
End Sub

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
    cmdAceptar.Enabled = (lvwCuotas.CheckedItems.Count <> 0)
End Sub


