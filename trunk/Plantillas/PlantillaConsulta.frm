VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form PlantillaConsulta 
   Caption         =   "Form2"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9180
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   9180
   Begin VB.Frame fraBotones 
      Height          =   1215
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   9200
      Begin MSForms.CommandButton cmdAceptar 
         Height          =   840
         Left            =   5505
         TabIndex        =   7
         Top             =   240
         Width           =   1700
         Caption         =   "Aceptar"
         PicturePosition =   327683
         Size            =   "2999;1482"
         Picture         =   "PlantillaConsulta.frx":0000
         FontName        =   "Comic Sans MS"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdCancelar 
         Height          =   840
         Left            =   7300
         TabIndex        =   6
         Top             =   240
         Width           =   1700
         Caption         =   " Cancelar"
         PicturePosition =   327683
         Size            =   "2999;1482"
         Picture         =   "PlantillaConsulta.frx":EBA2
         FontName        =   "Comic Sans MS"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdEliminar 
         Height          =   840
         Left            =   3710
         TabIndex        =   5
         Top             =   240
         Width           =   1700
         Caption         =   "Eliminar"
         PicturePosition =   327683
         Size            =   "2999;1482"
         Picture         =   "PlantillaConsulta.frx":11F94
         FontName        =   "Comic Sans MS"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdModificar 
         Height          =   840
         Left            =   1920
         TabIndex        =   4
         Top             =   240
         Width           =   1740
         VariousPropertyBits=   268435483
         Caption         =   "Modificar"
         PicturePosition =   327683
         Size            =   "3069;1482"
         Picture         =   "PlantillaConsulta.frx":2902E
         FontName        =   "Comic Sans MS"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdNuevo 
         Height          =   840
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1700
         Caption         =   "Nuevo"
         PicturePosition =   327683
         Size            =   "2999;1482"
         Picture         =   "PlantillaConsulta.frx":30530
         FontName        =   "Comic Sans MS"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2566
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   1535
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "il"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "asd"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList il 
      Left            =   2760
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PlantillaConsulta.frx":36D92
            Key             =   "eliminar"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PlantillaConsulta.frx":4DE2C
            Key             =   "aceptar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PlantillaConsulta.frx":5C9CE
            Key             =   "cancelar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PlantillaConsulta.frx":5FDC0
            Key             =   "modificar"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PlantillaConsulta.frx":672C2
            Key             =   "nuevo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PlantillaConsulta.frx":6DB24
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PlantillaConsulta.frx":73DBE
            Key             =   "papeleravacia"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PlantillaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'cambiar estas medidas segun corresponda
Private Const ANCHOMIN = 9300
Private Const ALTOMIN = 5000

Private tipo As eTipoFormulario

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If tipo = 0 Then Err.Raise 2010, , "No se puede mostrar el formulario con el metodo Show, utilice la funcion Consultar."
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN
        
        lvw.Top = IIf(tBar.Visible And tBar.Align = vbAlignTop, tBar.Height, 100)
        If fraBotones.Visible Then
            fraBotones.Top = Me.ScaleHeight - fraBotones.Height
            fraBotones.Width = Me.Width - 100
            lvw.Height = fraBotones.Top - 100
        Else
            lvw.Height = IIf(tBar.Visible And tBar.Align = vbAlignBottom, tBar.Top - 100, Me.ScaleHeight - tBar.Top)
        End If
        lvw.Width = Me.Width - 100
        
    End If
End Sub

Private Sub distribuirBotones()
'fijarse si hace falta
End Sub
