VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConsola 
   Caption         =   "Consola"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock wsock 
      Left            =   7410
      Top             =   1290
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   7440
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtComando 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   4560
      Width           =   7335
   End
   Begin VB.TextBox txt 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   4335
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmConsola.frx":0000
      Top             =   0
      Width           =   7335
   End
   Begin MSScriptControlCtl.ScriptControl script 
      Left            =   7320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "frmConsola"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vbInt As InterfazVB

Private Sub Form_Load()
Inicializar
analizarLineaDeComando
End Sub
'/execute abrirbasedatos /execute setearpathenconfig
Private Sub Inicializar()
    txt = "Consola - Version 1.0" + vbCrLf + ">"
    Dim a As New Archivo
    modulo = a.LeerArchivo(App.path + "\script.bas")
    script.AddCode modulo
'    Set word = CreateObject("tbrconfig.clsconfiguracion")
'    script.AddObject "word", word
    Set vbInt = New InterfazVB
    script.AddObject "txt", txt
    script.AddObject "cd", cd
    script.AddObject "vbInt", vbInt
    script.AddObject "wSock", wsock
    'script.AddObject "frmGrilla", frmGrilla
'    script.AddCode "Public sub Print(cadena)" + vbCrLf + "txt.text=txt.text+cadena+vbcrlf+""> """ + vbCrLf + "End sub"
     
End Sub

Private Sub analizarLineaDeComando()
On Error GoTo errman

Dim cmds() As String
'Dim pars() As String
cmds = Split(Command$, "/")

For i = 1 To UBound(cmds) 'el 0 es siempre vacio
    'pars = Split(cmds(i), " ")
    EjecutarComando (cmds(i))
Next
Exit Sub
errman:
txt = txt + "Error: Argumento de linea de comandos incorrecto"
End Sub

Private Sub Form_Resize()
txt.Width = Me.Width
txt.Height = Me.ScaleHeight - txtComando.Height
txtComando.Width = Me.Width
txtComando.Top = txt.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set conf = Nothing

End Sub

Private Sub txtComando_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = 6 Then
        Dim p As Procedure
        For Each p In script.Procedures
             If InStr(1, p.Name, txtComando, vbTextCompare) = 1 Then
                txtComando = p.Name
                txtComando.SelStart = Len(p.Name)
            End If
        Next
    End If
End Sub

Private Sub EjecutarComando(comando As String)
Select Case comando
            Case "comandos", "?", "help"
                txt = txt + "Comandos disponibles:" + vbCrLf
                txt = txt + "Cls - Limpia la pantalla." + vbCrLf
                txt = txt + "Exit, Quit - Cierra la Consola." + vbCrLf
                txt = txt + "Reset - Reinicia la consola." + vbCrLf + vbCrLf
                txt = txt + "Comandos Externos:" + vbCrLf
                Dim p As Procedure
                For Each p In script.Procedures
                    txt = txt + p.Name + vbCrLf
                Next
                txt = txt + ">"
            Case "cls"
                txt = ">"
            Case "reset"
                script.Reset
                Inicializar
            Case "exit", "quit"
                Unload Me
                End
            Case Else
                txt = txt + comando + vbCrLf
                script.ExecuteStatement comando
                txt = txt + ">"
        End Select
End Sub

Private Sub txtComando_KeyPress(KeyAscii As Integer)
On Error GoTo errman
If KeyAscii = vbKeyReturn Then
        EjecutarComando txtComando
        txtComando = ""
        txtComando.SetFocus
        KeyAscii = 0
End If
Exit Sub
errman:
txt = txt + vbCrLf + "Error: " + Err.Description + " - Linea:" + Str(script.Error.Line) + vbCrLf + ">"
End Sub

