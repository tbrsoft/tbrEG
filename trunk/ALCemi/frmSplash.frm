VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3765
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":000C
   ScaleHeight     =   3765
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   3600
      Top             =   360
   End
   Begin VB.Label lblMensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   3480
      Width           =   5055
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos& Lib "user32" _
(ByVal hwnd&, ByVal hWndInsertAfter&, _
ByVal X&, ByVal Y&, ByVal Wid&, _
ByVal Hgt&, ByVal flags&)
Const SWP_NOSIZE = 1
Const SWP_NOMOVE = 2
Const HWND_TOPMOST = -1
Const WM_SYSCOMMAND As Long = &H112&
Const MOUSE_MOVE As Long = &HF012&
Const SWP_SHOWWINDOW = &H40

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'lo saque para ver los mensajes de error!!!
    'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Timer1_Timer()
    'Me.Hide
    'MDI.Show
    Unload frmSplash
End Sub

Public Property Get Mensaje() As String
    Mensaje = lblMensaje.Caption
End Property

Public Property Let Mensaje(ByVal vNewValue As String)
    lblMensaje.Caption = vNewValue
End Property
