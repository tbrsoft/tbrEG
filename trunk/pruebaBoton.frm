VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin Proyecto1.GraphicButton UserControl11 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   495
      _extentx        =   873
      _extenty        =   873
   End
   Begin MSComctlLib.ImageList il32 
      Left            =   240
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   19
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":0000
            Key             =   "excel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":6862
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":74B4
            Key             =   "registrarcobro"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":84C6
            Key             =   "word"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":ED28
            Key             =   "restaurar"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":F97A
            Key             =   "agregar"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":161DC
            Key             =   "imprimir"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":16E2E
            Key             =   "aceptar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":259D0
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":26622
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":29A14
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":30F16
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":37778
            Key             =   "qth"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":383CA
            Key             =   "vl"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":3901C
            Key             =   "recibosanulados"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":39C6E
            Key             =   "papeleravacia"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":3A8C0
            Key             =   "papelerallena"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":3B512
            Key             =   "detalles"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "pruebaBoton.frx":41D74
            Key             =   "registraratencion"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Set UserControl11.Picture = il32.ListImages.Item("agregar").Picture
End Sub

Private Sub UserControl11_Click()
MsgBox "asdasd"
End Sub
