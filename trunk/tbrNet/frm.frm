VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.1#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Estado de la red"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3750
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3000
      Top             =   600
   End
   Begin MSWinsockLib.Winsock wskServer 
      Index           =   0
      Left            =   2520
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1064
   End
   Begin MSWinsockLib.Winsock wskCliente 
      Left            =   3000
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "192.168.0.131"
      RemotePort      =   1064
   End
   Begin ControlesPOO.ListViewConsulta lvwConexiones 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2566
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   -1  'True
      FullRowSelection=   -1  'True
      AutoDistribuirColumnas=   -1  'True
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
      NEncabezado0    =   "Id"
      MEncabezado0    =   "id"
      AEncabezado0    =   10
      NEncabezado1    =   "Nombre"
      MEncabezado1    =   "nombre"
      AEncabezado1    =   45
      NEncabezado2    =   "Usuario"
      MEncabezado2    =   "usuario"
      AEncabezado2    =   45
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
   Begin VB.Label Label3 
      Caption         =   "Conexiones entrantes:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblEstado 
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   420
      Width           =   1935
   End
   Begin VB.Label lblModo 
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Estado:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label1 
      Caption         =   "Modo:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mServer As Servidor
Attribute mServer.VB_VarHelpID = -1
Private mConexion As Conexion

Friend Property Set Server(pServer As Servidor)
    Set mServer = pServer
End Property

Friend Property Get Server() As Servidor
    Set Server = mServer
End Property

Friend Property Set Cliente(pCliente As Conexion)
    Set mConexion = pCliente
End Property

Friend Property Get Cliente() As Conexion
    Set Cliente = mConexion
End Property

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Me.Hide
    Case vbKeyF5
        Refrescar
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then
        Cancel = 1
        Me.Hide
    End If
End Sub

'no puedo interceptar los eventos de una matriz de controles
Private Sub mServer_RefrescarConexiones()
'lstConexiones.Clear
'Dim c As Conexion
'For Each c In mServer.Conexiones
'    lstConexiones.AddItem Str(c.Id) + " - " + c.Nombre
'Next
Set lvwConexiones.Coleccion = mServer.Conexiones
End Sub

Private Sub Timer_Timer()
'aca llega solo desde el cliente
If wskCliente.State <> sckConnected Then
    lblEstado = "Intentando conectar..."
    mConexion.CambiarEstado "Intentando conectar..."
    wskCliente.Close
    wskCliente.Connect
Else
    'no estoy seguro q esto haga falta, si lo de detener el timer
    lblEstado = "Conectado"
    mConexion.CambiarEstado "Conectado"
    Timer.Enabled = False
End If
End Sub

Private Sub wskCliente_Close()
    mConexion.CambiarEstado "Conexion Cerrada"
    Timer.Enabled = True
End Sub

Private Sub wskCliente_Connect()
    mConexion.CambiarEstado "Conectado"
End Sub

Private Sub wskCliente_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
    Dim aux As String
    wskCliente.GetData aux
    mConexion.wSock_DataArrival aux
End Sub

Private Sub wskCliente_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskCliente.Close
Timer.Enabled = True
End Sub

Private Sub wskServer_ConnectionRequest(Index As Integer, ByVal requestID As Long)

wskServer(Index).Close
wskServer(Index).Accept requestID
mServer.AddNewConexion Index
Load wskServer(Index + 1)
wskServer(Index + 1).LocalPort = 1064
wskServer(Index + 1).Listen

End Sub

Private Sub wskServer_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
wskServer(Index).Close
End Sub

Friend Sub wskServer_Close(Index As Integer)
'arreglar si quiere descargar el index 0
On Error Resume Next
wskServer(Index).Close
Unload wskServer(Index)
mServer.Conexiones.Remove Index
End Sub

Private Sub wskServer_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim aux As String
wskServer(Index).GetData aux
mServer.Conexiones.Item(Index).wSock_DataArrival aux
End Sub

Public Sub Refrescar()
    If Not mServer Is Nothing Then
        mServer.PedirNombresHost
        mServer_RefrescarConexiones
    ElseIf Not mConexion Is Nothing Then
    'ver si se puede actualizar el estado etc.
    Select Case wskCliente.State
        Case 6
            lblEstado = "Intentando conectar..."
            mConexion.CambiarEstado "Intentando conectar..."
        Case 7
            mConexion.CambiarEstado "Conectado"
            lblEstado = "Conectado"
    End Select
    End If
End Sub
