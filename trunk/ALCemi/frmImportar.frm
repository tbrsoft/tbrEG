VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmImportar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar datos externos"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   5850
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "Siguiente ->"
      Height          =   375
      Left            =   1800
      TabIndex        =   16
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<- Anterior"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Importando datos"
      Height          =   3855
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   5655
      Begin VB.ListBox lstProgreso 
         Height          =   2010
         Left            =   480
         TabIndex        =   25
         Top             =   1680
         Width           =   4695
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   720
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblProgreso 
         Caption         =   "Label5"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   1440
         Width           =   4695
      End
      Begin VB.Label lblmensaje 
         Caption         =   "Aguarde mientras se inserta la informacion en la base de datos..."
         Height          =   195
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Width           =   4545
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paso 1/3"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton optDestino 
         Caption         =   "Parentezcos"
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   6
         Top             =   2280
         Width           =   1455
      End
      Begin VB.OptionButton optDestino 
         Caption         =   "Ocupaciones"
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton optDestino 
         Caption         =   "Medicamentos"
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   4
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton optDestino 
         Caption         =   "Enfermedades"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   3
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton optDestino 
         Caption         =   "Alergias"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "¿Que tipo de informacion desea importar?"
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Paso 2/3"
      Height          =   3855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5655
      Begin VB.OptionButton optClipboard 
         Caption         =   "Portapapeles"
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   840
         Value           =   -1  'True
         Width           =   4095
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "..."
         Height          =   255
         Left            =   4680
         TabIndex        =   12
         Top             =   2760
         Width           =   375
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   720
         TabIndex        =   11
         Top             =   2760
         Width           =   3975
      End
      Begin VB.OptionButton optPath 
         Caption         =   "Archivo de texto (Solo .txt, no se admiten archivos de Word)"
         Height          =   255
         Left            =   480
         TabIndex        =   10
         Top             =   2400
         Width           =   4815
      End
      Begin VB.Label Label3 
         Caption         =   $"frmImportar.frx":0000
         Height          =   735
         Left            =   840
         TabIndex        =   14
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Seleccione el origen de los datos:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Paso 3/3"
      Height          =   3855
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5655
      Begin MSComctlLib.ListView ListView1 
         Height          =   2535
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4471
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.CommandButton cmdNone 
         Caption         =   "Ninguno"
         Height          =   255
         Left            =   1680
         TabIndex        =   20
         Top             =   3480
         Width           =   1335
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "Todos"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Corrobore que los datos sean correctos. Puede modificar los valores que desee. Seleccione los valores que deben ser importados."
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   4935
      End
   End
End
Attribute VB_Name = "frmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indice As Integer

Private Sub cmdAll_Click()
Dim lvi As listItem
For Each lvi In ListView1.ListItems
   lvi.Checked = True
Next
End Sub

Private Sub cmdAnterior_Click()
Select Case indice
    Case 1
        
    Case 2
        Frame1.ZOrder 0
        cmdSiguiente.Caption = "Siguiente ->"
        indice = 1
        cmdAnterior.Enabled = False
    Case 3
        Frame2.ZOrder 0
        cmdSiguiente.Caption = "Siguiente ->"
        indice = 2
        
End Select

End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

'Private Sub cmdEliminar_Click()
'If Not ListView1.SelectedItem Is Nothing Then
'    ListView1.ListItems.Remove ListView1.SelectedItem.Index
'End If
'
'End Sub

Private Sub cmdNone_Click()
Dim lvi As listItem
For Each lvi In ListView1.ListItems
   lvi.Checked = False
Next
End Sub

Private Sub cmdSiguiente_Click()
Select Case indice
    Case 1
        Frame2.ZOrder 0
        indice = 2
        cmdAnterior.Enabled = True
    Case 2
        Frame3.ZOrder 0
        Importar
        cmdSiguiente.Caption = "Finalizar"
        indice = 3
        cmdAnterior.Enabled = True
    Case 3
        AgregarABD
End Select
End Sub

Private Sub AgregarABD()
'cada 10 elementos llamo a doevents
Dim b As Integer

Frame4.ZOrder 0
pb.Max = ListView1.ListItems.Count + 1
Dim lvi As listItem
cmdSiguiente.Enabled = False
cmdAnterior.Enabled = False
'ver si se puede cancelar
cmdCancelar.Enabled = False

If optDestino(0).Value And optDestino(0).Enabled Then
    
    For Each lvi In ListView1.ListItems
        lblProgreso.Caption = "Importando " + lvi.Text
        If Not GBL.AlergiasGBL.Nuevo(lvi.Text) Is Nothing Then
            lstProgreso.AddItem "Agregado " + lvi.Text
        Else
            lstProgreso.AddItem "No agregado " + lvi.Text + ", ya existe."
        End If
        pb.Value = pb.Value + 1
        b = b + 1
        If b = 10 Then
            DoEvents
            b = 0
        End If
    Next
    
End If

If optDestino(1).Value And optDestino(1).Enabled Then
    
    For Each lvi In ListView1.ListItems
        lblProgreso.Caption = "Importando " + lvi.Text
        If Not GBL.EnfermedadesGBL.Nuevo(lvi.Text) Is Nothing Then
            lstProgreso.AddItem "Agregado " + lvi.Text
        Else
            lstProgreso.AddItem "No agregado " + lvi.Text + ", ya existe."
        End If
        pb.Value = pb.Value + 1
        b = b + 1
        If b = 10 Then
            DoEvents
            b = 0
        End If
    Next
End If

If optDestino(2).Value And optDestino(2).Enabled Then
   
    For Each lvi In ListView1.ListItems
        lblProgreso.Caption = "Importando " + lvi.Text
        If Not GBL.MedicamentosGBL.Nuevo(lvi.Text) Is Nothing Then
            lstProgreso.AddItem "Agregado " + lvi.Text
        Else
            lstProgreso.AddItem "No agregado " + lvi.Text + ", ya existe."
        End If
        pb.Value = pb.Value + 1
        b = b + 1
        If b = 10 Then
            DoEvents
            b = 0
        End If
    Next
End If

If optDestino(3).Value And optDestino(3).Enabled Then
    
    For Each lvi In ListView1.ListItems
        lblProgreso.Caption = "Importando " + lvi.Text
        If Not GBL.OcupacionesGBL.Nuevo(lvi.Text) Is Nothing Then
            lstProgreso.AddItem "Agregado " + lvi.Text
        Else
            lstProgreso.AddItem "No agregado " + lvi.Text + ", ya existe."
        End If
        pb.Value = pb.Value + 1
        b = b + 1
        If b = 10 Then
            DoEvents
            b = 0
        End If
    Next
End If

If optDestino(4).Value And optDestino(4).Enabled Then
    
    For Each lvi In ListView1.ListItems
        lblProgreso.Caption = "Importando " + lvi.Text
        If Not GBL.ParentezcosGBL.Nuevo(lvi.Text) Is Nothing Then
            lstProgreso.AddItem "Agregado " + lvi.Text
        Else
            lstProgreso.AddItem "No agregado " + lvi.Text + ", ya existe."
        End If
        pb.Value = pb.Value + 1
        b = b + 1
        If b = 10 Then
            DoEvents
            b = 0
        End If
    Next
End If

cmdCancelar.Enabled = True
cmdCancelar.Caption = "Cerrar"
pb.Visible = False
lblMensaje = "Se finalizó la importación de los datos."
lblProgreso = "Resultado de la operación:"
End Sub

Private Sub Importar()
    If optClipboard.Value Then
        If Not Clipboard.GetFormat(1) Then
            MsgBox "Para poder importar desde el portapapeles debe seleccionar los datos y utilizar la opcion ""Copiar"". Luego seleccione ""Siguiente""."
        Else
            ImportFromClipboard
        End If
    Else
        If txtPath = "" Then
            MsgBox "Seleccione un archivo!", vbExclamation, "tbrEmergencyGroup"
            indice = 2
            Frame2.ZOrder 0
        Else
            ImportFromFile
        End If
        
    End If
    'selecciono todos
    cmdAll_Click
End Sub

Private Sub ImportFromClipboard()

    If Clipboard.GetFormat(1) Then
        Dim aux() As String
        aux = Split(Clipboard.GetText(1), vbCrLf)
        For I = 0 To UBound(aux)
            If Trim(aux(I)) <> "" Then
                ListView1.ListItems.Add , , Trim(aux(I))
            End If
        Next
    End If

End Sub

Private Sub ImportFromFile()
Dim cadena As String
cadena = LeerArchivo(txtPath)
If cadena <> "" Then
        Dim aux() As String
        aux = Split(cadena, vbCrLf)
        For I = 0 To UBound(aux)
            If Trim(aux(I)) <> "" Then
                ListView1.ListItems.Add , , Trim(aux(I))
            End If
        Next
    End If
End Sub

Private Function LeerArchivo(Path As String) As String
    Dim fso 'As FileSystemObject
    Dim f
        
    Dim s As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Path = "" Then
        LeerArchivo = ""
    Else
        Set f = fso.GetFile(Path)
    End If
    Set ts = f.OpenAsTextStream(1)
    
    s = ts.ReadAll
    
    ts.Close
    LeerArchivo = s
End Function


Private Sub Form_Load()
    Set Me.Icon = MDI.Icon
    Frame1.ZOrder 0
    indice = 1
    'permisos
    optDestino(0).Enabled = UsuarioActual.Permisos.Can(blcemi.AltaAlergia)
    optDestino(1).Enabled = UsuarioActual.Permisos.Can(blcemi.AltaEnfermedad)
    optDestino(2).Enabled = UsuarioActual.Permisos.Can(blcemi.AltaMedicamento)
    optDestino(3).Enabled = UsuarioActual.Permisos.Can(blcemi.AltaOcupacion)
    optDestino(4).Enabled = UsuarioActual.Permisos.Can(blcemi.AltaParentezco)
    
End Sub

Public Function GetHelpContext() As String
    GetHelpContext = "herramientas"
End Function

