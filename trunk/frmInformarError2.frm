VERSION 5.00
Begin VB.Form frmInformarError2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informe de errores"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "generar Informe"
      Height          =   435
      Left            =   1770
      TabIndex        =   3
      Top             =   2730
      Width           =   1665
   End
   Begin VB.TextBox Text1 
      Height          =   1155
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1470
      Width           =   4515
   End
   Begin VB.Label Label2 
      Caption         =   "Si lo desea agregue sus comentarios"
      Height          =   225
      Left            =   90
      TabIndex        =   2
      Top             =   1230
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Se ha generado un error en el sistema"
      Height          =   855
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4005
   End
End
Attribute VB_Name = "frmInformarError2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    'juntar todos los *.log de aqui
    
    
    TERR.AppendSinHist "****************" + vbCrLf + _
                        "****************" + vbCrLf + _
                        "DIJO:" + vbCrLf + Text1.Text + vbCrLf + _
                        "****************" + vbCrLf + _
                        "****************"


    Dim JS As New tbrJUSE2.clsJUSE
    Dim Fch As String
    Fch = CStr(Year(Date)) + "." + CStr(Month(Date)) + "." + CStr(Day(Date)) + "." + CStr(Hour(Time)) + "." + CStr(Minute(Time)) + ".log"
    
    JS.Archivo = APh + "Info_tbrEG_" + Fch
    JS.AddFiles APh, "log"
    JS.AddFile "C:\widnows\system32\reg_tlv.log" 'TODO corregir y mejorar
    JS.Unir False
    
    MsgBox "Se ha grabado el registro de errores en:" + vbCrLf + _
        JS.Archivo + vbCrLf + vbCrLf + _
        "Envialo por email a info@tbrsoft.com"
    
    Unload Me
End Sub

Private Sub Form_Load()
    Label1.Caption = "Ha habido un error menor en el sistema" + vbCrLf + _
        "Por favor genere y envíenos este informe para revisar a info@tbrsoft.com"
        
End Sub

Public Sub ForzarEnvio(estabaEn As String)
    Me.Text1.Text = estabaEn
    Me.Show 1
End Sub

