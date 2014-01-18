VERSION 5.00
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.0#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConfigEncabezados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar Columnas"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5265
   Begin VB.CommandButton cmdTodosDer 
      Caption         =   ">>"
      Height          =   255
      Left            =   2460
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdDer 
      Caption         =   ">"
      Height          =   255
      Left            =   2460
      TabIndex        =   6
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton cmdIzq 
      Caption         =   "<"
      Height          =   255
      Left            =   2460
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton cmdTodosIzq 
      Caption         =   "<<"
      Height          =   255
      Left            =   2460
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin ControlesPOO.ListViewConsulta lvwSeleccionados 
      Height          =   3255
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5741
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Seleccionados"
      MEncabezado0    =   "nombre"
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
   Begin ControlesPOO.ListViewConsulta lvwDisponibles 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5741
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Disponibles"
      MEncabezado0    =   "nombre"
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
End
Attribute VB_Name = "frmConfigEncabezados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event ColumnasSeleccionadas(pSeleccionadas As ControlesPOO.LVCEncabezadoManager)

Dim mDisponibles As ControlesPOO.LVCEncabezadoManager
Dim mSeleccionadas As ControlesPOO.LVCEncabezadoManager

Public Sub ConfigurarColumnas(pDisponibles As ControlesPOO.LVCEncabezadoManager, pSeleccionadas As ControlesPOO.LVCEncabezadoManager)
    Dim enc As ControlesPOO.LVCEncabezado
    Set mDisponibles = New ControlesPOO.LVCEncabezadoManager
    Set mSeleccionadas = New ControlesPOO.LVCEncabezadoManager
    For Each enc In pDisponibles
        mDisponibles.Add enc.Nombre, enc.miembro, enc.ancho
    Next
    For Each enc In pSeleccionadas
        mSeleccionadas.Add enc.Nombre, enc.miembro, enc.ancho
        'elimino los q ya tengo seleccionados
        mDisponibles.Remove enc.miembro
    Next
    
    Set lvwDisponibles.Coleccion = mDisponibles
    Set lvwSeleccionados.Coleccion = mSeleccionadas
    Me.Show
End Sub

Private Sub cmdAceptar_Click()
    RaiseEvent ColumnasSeleccionadas(lvwSeleccionados.Coleccion)
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDer_Click()
    If Not lvwDisponibles.SelectedItem Is Nothing Then
        lvwDisponibles_ItemDblClick lvwDisponibles.SelectedItem
    End If
End Sub

Private Sub cmdIzq_Click()
    If Not lvwSeleccionados.SelectedItem Is Nothing Then
        lvwSeleccionados_ItemDblClick lvwSeleccionados.SelectedItem
    End If
End Sub

Private Sub cmdTodosDer_Click()
    Dim enc As ControlesPOO.LVCEncabezado
    For Each enc In lvwDisponibles.Coleccion
        lvwSeleccionados.Coleccion.Add enc.Nombre, enc.miembro, enc.ancho
    Next
    lvwDisponibles.Coleccion.Clear
    lvwDisponibles.Refresh
    lvwSeleccionados.Refresh
End Sub

Private Sub cmdTodosIzq_Click()
    Dim enc As ControlesPOO.LVCEncabezado
    For Each enc In lvwSeleccionados.Coleccion
        lvwDisponibles.Coleccion.Add enc.Nombre, enc.miembro, enc.ancho
    Next
    lvwSeleccionados.Coleccion.Clear
    lvwDisponibles.Refresh
    lvwSeleccionados.Refresh
End Sub

Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
End Sub

Private Sub lvwDisponibles_ItemDblClick(Item As Object)
    lvwSeleccionados.Coleccion.Add Item.Nombre, Item.miembro, Item.ancho
    lvwDisponibles.Coleccion.Remove Item.miembro
    lvwDisponibles.Refresh
    lvwSeleccionados.Refresh
End Sub

Private Sub lvwSeleccionados_ItemDblClick(Item As Object)
    lvwDisponibles.Coleccion.Add Item.Nombre, Item.miembro, Item.ancho
    lvwSeleccionados.Coleccion.Remove Item.miembro
    lvwDisponibles.Refresh
    lvwSeleccionados.Refresh
End Sub
