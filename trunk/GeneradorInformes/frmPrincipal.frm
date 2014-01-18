VERSION 5.00
Object = "{1417CD23-5617-4303-9AEF-2418F695BFFF}#1.0#0"; "ListViewConsultaCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Diseñador de Informes"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   9465
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDeletePar 
      Caption         =   "X"
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   3600
      Width           =   375
   End
   Begin VB.CommandButton cmdAddParameter 
      Caption         =   "+"
      Height          =   375
      Left            =   9000
      TabIndex        =   13
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "X"
      Height          =   375
      Left            =   4320
      TabIndex        =   12
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton cmdCampos 
      Caption         =   "Obtener campos desde SQL"
      Height          =   375
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   3255
   End
   Begin ControlesPOO.ListViewConsulta lvwParametros 
      Height          =   3495
      Left            =   4800
      TabIndex        =   9
      Top             =   2880
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6165
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      CampoKey        =   ""
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Nombre"
      MEncabezado0    =   "Nombre"
      AEncabezado0    =   33
      NEncabezado1    =   "Tipo"
      MEncabezado1    =   "tipo"
      AEncabezado1    =   33
      NEncabezado2    =   "Descripcion"
      MEncabezado2    =   "Descripcion"
      AEncabezado2    =   34
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
   Begin MSComDlg.CommonDialog cd 
      Left            =   8280
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAbrir 
      Caption         =   "Abrir"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "Ejecutar"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5880
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   120
      Width           =   1575
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6165
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   -1  'True
      CampoKey        =   ""
      AllowModify     =   -1  'True
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Nombre"
      MEncabezado0    =   "nombre"
      AEncabezado0    =   50
      NEncabezado1    =   "Miembro"
      MEncabezado1    =   "miembro"
      AEncabezado1    =   50
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
   Begin VB.TextBox txtSQL 
      Height          =   765
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "frmPrincipal.frx":0000
      Top             =   1320
      Width           =   7815
   End
   Begin VB.TextBox txtTitulo 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label C 
      Caption         =   "Campos (Columnas)"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "SQL:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label T 
      Caption         =   "Titulo:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents frmP As frmParametro
Attribute frmP.VB_VarHelpID = -1

Private Sub cmdAddParameter_Click()
Set frmP = New frmParametro
frmP.Show
End Sub

Private Sub frmP_NuevoParametro(par As LParameter)
If lvwParametros.Coleccion Is Nothing Then
    Set lvwParametros.Coleccion = New LParameterManager
End If
lvwParametros.Coleccion.AddItem par
lvwParametros.Refresh
txtSQL = txtSQL + par.Nombre + "=@" + par.Nombre
End Sub

Private Sub cmdCampos_Click()
Dim rs As Recordset
Dim f As Field
Dim encs As New ControlesPOO.LVCEncabezadoManager
Set rs = ExecuteSQL(txtSQL)
For Each f In rs.Fields
    encs.Add f.Name, f.Name, 5
Next
Set lvw.Coleccion = encs
lvw.Refresh
End Sub

Private Sub cmdDelete_Click()
    If Not lvw.SelectedItem Is Nothing Then
        lvw.Coleccion.Remove lvw.SelectedItem.Nombre
        lvw.Refresh
    End If
End Sub

Private Sub cmdDeletePar_Click()
    If Not lvwParametros.SelectedItem Is Nothing Then
        lvwParametros.Coleccion.Remove lvwParametros.SelectedItem.Nombre
        lvwParametros.Refresh
    End If
End Sub

Private Sub cmdEjecutar_Click()
Dim l As New Listado
Set l.Encabezados = lvw.Coleccion
Set l.Parametros = lvwParametros.Coleccion
l.SQL = txtSQL
l.Titulo = txtTitulo
frmListadoGenerico.MostrarListado l
End Sub


Private Sub cmdGuardar_Click()
cd.ShowSave
Dim l As New Listado
Set l.Encabezados = lvw.Coleccion
l.SQL = txtSQL
l.Titulo = txtTitulo
l.Save cd.FileName
End Sub

Private Sub Command2_Click()
Unload Me
End Sub


