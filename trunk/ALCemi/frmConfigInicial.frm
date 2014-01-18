VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmConfigInicial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "tbrEmergencyGroup - Configuración Inicial"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14115
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   14115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   47
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame fra6 
      Caption         =   "Finalizado"
      Height          =   4815
      Left            =   8160
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   6255
      Begin MSComctlLib.ProgressBar pb 
         Height          =   375
         Left            =   360
         TabIndex        =   37
         Top             =   1680
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblOperacion 
         Caption         =   "lblOperacion"
         Height          =   255
         Left            =   360
         TabIndex        =   36
         Top             =   1320
         Width           =   5415
      End
      Begin VB.Label Label14 
         Caption         =   "Se estan realizando las siguientes operaciones:"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   5175
      End
   End
   Begin VB.Frame fra5 
      Caption         =   "Paso 5"
      Height          =   4815
      Left            =   3630
      TabIndex        =   16
      Top             =   3810
      Width           =   6255
      Begin VB.Frame Frame1 
         Caption         =   "Direccion por defecto"
         Height          =   1935
         Left            =   720
         TabIndex        =   19
         Top             =   1920
         Width           =   4935
         Begin ControlesPOO.Combo cmbPais 
            Height          =   315
            Left            =   960
            TabIndex        =   46
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbBarrio 
            Height          =   315
            Left            =   960
            TabIndex        =   20
            Top             =   1440
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            NuevoEnabled    =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbCiudad 
            Height          =   315
            Left            =   960
            TabIndex        =   21
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            NuevoEnabled    =   -1  'True
            Enabled         =   -1  'True
         End
         Begin ControlesPOO.Combo cmbProvincia 
            Height          =   315
            Left            =   960
            TabIndex        =   22
            Top             =   720
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            Enabled         =   -1  'True
         End
         Begin VB.Label Label16 
            Caption         =   "País:"
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblBarrio 
            Caption         =   "Barrio:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1440
            Width           =   615
         End
         Begin VB.Label lblCiudad 
            Caption         =   "Ciudad:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblProvincia 
            Caption         =   "Provincia:"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   720
            Width           =   855
         End
      End
      Begin VB.Label Label9 
         Caption         =   $"frmConfigInicial.frx":0000
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   1080
         Width           =   5655
      End
      Begin VB.Label Label8 
         Caption         =   "Seleccione su ubicacion."
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame fra4 
      Caption         =   "Paso 4"
      Height          =   4815
      Left            =   6600
      TabIndex        =   15
      Top             =   120
      Width           =   6255
      Begin ControlesPOO.TreeViewConsulta tvw 
         Height          =   2775
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4895
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HideSelection   =   -1  'True
         Indentation     =   566,929
         LabelEdit       =   1
         LineStyle       =   1
         Nodo.BackColor0 =   " 16777215"
         Nodo.Bold0      =   "False"
         Nodo.ChildCollectionField0=   "Sintomas"
         Nodo.Expanded0  =   "False"
         Nodo.ForeColor0 =   " 0"
         Nodo.IdField0   =   "id"
         Nodo.TextField0 =   "NombreCompuesto"
         Nodo.BackColor1 =   " 16777215"
         Nodo.Bold1      =   "False"
         Nodo.ChildCollectionField1=   ""
         Nodo.Expanded1  =   "False"
         Nodo.ForeColor1 =   " 0"
         Nodo.IdField1   =   "id"
         Nodo.TextField1 =   "NombreCompuesto"
         Nodo.BackColor0 =   "0"
         Nodo.Bold0      =   "False"
         Nodo.ChildCollectionField0=   ""
         Nodo.Expanded0  =   "False"
         Nodo.ForeColor0 =   "0"
         Nodo.IdField0   =   ""
         Nodo.TextField0 =   ""
         Nodo.BackColor1 =   "0"
         Nodo.Bold1      =   "False"
         Nodo.ChildCollectionField1=   ""
         Nodo.Expanded1  =   "False"
         Nodo.ForeColor1 =   "0"
         Nodo.IdField1   =   ""
         Nodo.TextField1 =   ""
         Nodo.BackColor2 =   "0"
         Nodo.Bold2      =   "False"
         Nodo.ChildCollectionField2=   ""
         Nodo.Expanded2  =   "False"
         Nodo.ForeColor2 =   "0"
         Nodo.IdField2   =   ""
         Nodo.TextField2 =   ""
         Nodo.BackColor3 =   "0"
         Nodo.Bold3      =   "False"
         Nodo.ChildCollectionField3=   ""
         Nodo.Expanded3  =   "False"
         Nodo.ForeColor3 =   "0"
         Nodo.IdField3   =   ""
         Nodo.TextField3 =   ""
         Nodo.BackColor4 =   "0"
         Nodo.Bold4      =   "False"
         Nodo.ChildCollectionField4=   ""
         Nodo.Expanded4  =   "False"
         Nodo.ForeColor4 =   "0"
         Nodo.IdField4   =   ""
         Nodo.TextField4 =   ""
         Nodo.BackColor5 =   "0"
         Nodo.Bold5      =   "False"
         Nodo.ChildCollectionField5=   ""
         Nodo.Expanded5  =   "False"
         Nodo.ForeColor5 =   "0"
         Nodo.IdField5   =   ""
         Nodo.TextField5 =   ""
         Nodo.BackColor6 =   "0"
         Nodo.Bold6      =   "False"
         Nodo.ChildCollectionField6=   ""
         Nodo.Expanded6  =   "False"
         Nodo.ForeColor6 =   "0"
         Nodo.IdField6   =   ""
         Nodo.TextField6 =   ""
         Nodo.BackColor7 =   "0"
         Nodo.Bold7      =   "False"
         Nodo.ChildCollectionField7=   ""
         Nodo.Expanded7  =   "False"
         Nodo.ForeColor7 =   "0"
         Nodo.IdField7   =   ""
         Nodo.TextField7 =   ""
         Nodo.BackColor8 =   "0"
         Nodo.Bold8      =   "False"
         Nodo.ChildCollectionField8=   ""
         Nodo.Expanded8  =   "False"
         Nodo.ForeColor8 =   "0"
         Nodo.IdField8   =   ""
         Nodo.TextField8 =   ""
         Nodo.BackColor9 =   "0"
         Nodo.Bold9      =   "False"
         Nodo.ChildCollectionField9=   ""
         Nodo.Expanded9  =   "False"
         Nodo.ForeColor9 =   "0"
         Nodo.IdField9   =   ""
         Nodo.TextField9 =   ""
         Nodo.BackColor10=   "0"
         Nodo.Bold10     =   "False"
         Nodo.ChildCollectionField10=   ""
         Nodo.Expanded10 =   "False"
         Nodo.ForeColor10=   "0"
         Nodo.IdField10  =   ""
         Nodo.TextField10=   ""
         Nodo.BackColor11=   "0"
         Nodo.Bold11     =   "False"
         Nodo.ChildCollectionField11=   ""
         Nodo.Expanded11 =   "False"
         Nodo.ForeColor11=   "0"
         Nodo.IdField11  =   ""
         Nodo.TextField11=   ""
         Nodo.BackColor12=   "0"
         Nodo.Bold12     =   "False"
         Nodo.ChildCollectionField12=   ""
         Nodo.Expanded12 =   "False"
         Nodo.ForeColor12=   "0"
         Nodo.IdField12  =   ""
         Nodo.TextField12=   ""
         Nodo.BackColor13=   "0"
         Nodo.Bold13     =   "False"
         Nodo.ChildCollectionField13=   ""
         Nodo.Expanded13 =   "False"
         Nodo.ForeColor13=   "0"
         Nodo.IdField13  =   ""
         Nodo.TextField13=   ""
         Nodo.BackColor14=   "0"
         Nodo.Bold14     =   "False"
         Nodo.ChildCollectionField14=   ""
         Nodo.Expanded14 =   "False"
         Nodo.ForeColor14=   "0"
         Nodo.IdField14  =   ""
         Nodo.TextField14=   ""
         Nodo.BackColor15=   "0"
         Nodo.Bold15     =   "False"
         Nodo.ChildCollectionField15=   ""
         Nodo.Expanded15 =   "False"
         Nodo.ForeColor15=   "0"
         Nodo.IdField15  =   ""
         Nodo.TextField15=   ""
         Nodo.BackColor16=   "0"
         Nodo.Bold16     =   "False"
         Nodo.ChildCollectionField16=   ""
         Nodo.Expanded16 =   "False"
         Nodo.ForeColor16=   "0"
         Nodo.IdField16  =   ""
         Nodo.TextField16=   ""
         Nodo.BackColor17=   "0"
         Nodo.Bold17     =   "False"
         Nodo.ChildCollectionField17=   ""
         Nodo.Expanded17 =   "False"
         Nodo.ForeColor17=   "0"
         Nodo.IdField17  =   ""
         Nodo.TextField17=   ""
         Nodo.BackColor18=   "0"
         Nodo.Bold18     =   "False"
         Nodo.ChildCollectionField18=   ""
         Nodo.Expanded18 =   "False"
         Nodo.ForeColor18=   "0"
         Nodo.IdField18  =   ""
         Nodo.TextField18=   ""
         Nodo.BackColor19=   "0"
         Nodo.Bold19     =   "False"
         Nodo.ChildCollectionField19=   ""
         Nodo.Expanded19 =   "False"
         Nodo.ForeColor19=   "0"
         Nodo.IdField19  =   ""
         Nodo.TextField19=   ""
      End
      Begin VB.OptionButton optSeleccionado 
         Caption         =   "Utilizar el seleccionado en la lista."
         Height          =   195
         Left            =   135
         TabIndex        =   32
         Top             =   840
         Width           =   3135
      End
      Begin VB.OptionButton optPredeterminado 
         Caption         =   "Utilizar el sistema predeterminado."
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Value           =   -1  'True
         Width           =   2775
      End
      Begin VB.ListBox List1 
         Height          =   840
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   6015
      End
      Begin VB.Label Label13 
         Caption         =   "Seleccione la codificacion a utilizar."
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame fra3 
      Caption         =   "Paso 3"
      Height          =   4815
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.OptionButton optB 
         Caption         =   "Cuerpo de Bomberos."
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1935
      End
      Begin VB.OptionButton optSE 
         Caption         =   "Servicio de Emergencias."
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label lbl 
         Caption         =   "Seleccione el tipo de Empresa/Organismo"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   5175
      End
   End
   Begin VB.Frame fra2 
      Caption         =   "Paso 2"
      Height          =   4815
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtPass2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox txtPass 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Repita la clave:"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   2400
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Clave:"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Puede no ingresar una clave en este momento, en cuyo caso la clave por defecto para el administrador es 1."
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   1320
         Width           =   5175
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrese la clave del administrador del sistema. Puede cambiar la clave ingresada en cualquier momento a traves del menu principal."
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   5055
      End
   End
   Begin VB.Frame fra1 
      Caption         =   "Paso 1"
      Height          =   4815
      Left            =   8010
      TabIndex        =   38
      Top             =   3720
      Width           =   6255
      Begin VB.Frame fraModoRed 
         Caption         =   "Modo de funcionamiento en red"
         Height          =   1095
         Left            =   600
         TabIndex        =   42
         Top             =   2040
         Width           =   5055
         Begin VB.OptionButton optCliente 
            Caption         =   "Este equipo se utilizara como cliente."
            Height          =   255
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   3855
         End
         Begin VB.OptionButton optServidor 
            Caption         =   "Este equipo se utilizará como servidor."
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   4695
         End
      End
      Begin VB.OptionButton optRed 
         Caption         =   "El sistema se utilizará en varios equipos dentro de mi empresa."
         Height          =   255
         Left            =   480
         TabIndex        =   41
         Top             =   1560
         Width           =   5295
      End
      Begin VB.OptionButton optMonoUsuario 
         Caption         =   "El sistema se utilizará en este equipo solamente, no se necesita red."
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   1200
         Width           =   5415
      End
      Begin VB.Label Label15 
         Caption         =   "Seleccione la configuración de red que utilizará."
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   480
         Width           =   5775
      End
   End
   Begin VB.Frame fra0 
      Height          =   4575
      Left            =   2520
      TabIndex        =   11
      Top             =   0
      Width           =   6255
      Begin VB.Label Label11 
         Caption         =   "aqui."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   4080
         TabIndex        =   49
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Para obtener ayuda presione la tecla F1. o haga click"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   2160
         Width           =   3855
      End
      Begin VB.Label Label7 
         Caption         =   "Para comenzar seleccione ""Siguiente""."
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2880
         Width           =   5775
      End
      Begin VB.Label Label6 
         Caption         =   $"frmConfigInicial.frx":008D
         Height          =   735
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   5775
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "tbrEmergencyGroup - Software para Gestión de Emergencias"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   6015
      End
   End
   Begin VB.CommandButton cmdFinalizar 
      Caption         =   "Finalizar"
      Height          =   375
      Left            =   7440
      TabIndex        =   28
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdSiguiente 
      Caption         =   "Siguiente >>"
      Height          =   375
      Left            =   6000
      TabIndex        =   27
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdAnterior 
      Caption         =   "<< Anterior"
      Height          =   375
      Left            =   3840
      TabIndex        =   26
      Top             =   4680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image img 
      Height          =   4950
      Left            =   120
      Picture         =   "frmConfigInicial.frx":0141
      Top             =   120
      Width           =   2250
   End
End
Attribute VB_Name = "frmConfigInicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Const MAX_CAMPOS = 2

Dim codigos As blcemi.CodigoEmergenciaManager
Dim cantCodySint As Integer
Dim pasoActual As Integer
Dim bGuardarConfig As Boolean

Private Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As HH_COMMAND, ByVal dwData As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Sub MostrarCodificacionesDisponibles(Tipo As String)
    Dim fso As New FileSystemObject
    Dim f As Folder
    Dim fil As File
    Set f = fso.GetFolder(APh + "Codificaciones\" + Tipo)
    If Not f Is Nothing Then
        List1.Clear
        For Each fil In f.Files
            List1.AddItem Mid(fil.Name, 1, Len(fil.Name) - 4)
        Next
    End If
End Sub

Private Sub cmdCancelar_Click()
    If MsgBox("Esta seguro que desea cancelar la configuracion?", vbQuestion + vbYesNo) = vbYes Then
        End 're chongo
    End If
End Sub

'Private Sub cmdAnterior_Click()
'    pasoActual = pasoActual - 1
'    If pasoActual <= 0 Then pasoActual = 0
'    MostrarPasoActual
'End Sub

Private Sub cmdFinalizar_Click()
    If cmdFinalizar.Caption = "Finalizar" Then
    
        If optSE.Value Then
            CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento = 1
            modoSoftware = eMFEmergencia
        Else
            CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento = 2
            modoSoftware = eMFBomberos
        End If
            
        'guardo direccion default
        'TODO no deberia quedar vacio
        If Not cmbBarrio.SelectedItem Is Nothing Then
            CCFFGG.Configuracion.Defaults.Barrio = cmbBarrio.SelectedItem.Nombre
        End If
        
        If Not cmbCiudad.SelectedItem Is Nothing Then
            CCFFGG.Configuracion.Defaults.Ciudad = cmbCiudad.SelectedItem.Nombre
        End If
        
        If Not cmbProvincia.SelectedItem Is Nothing Then
            CCFFGG.Configuracion.Defaults.Provincia = cmbProvincia.SelectedItem.Nombre
        End If
        
        If Not cmbPais.SelectedItem Is Nothing Then
            CCFFGG.Configuracion.Defaults.Pais = cmbPais.SelectedItem.Nombre
        End If
        'trato de adivinar la resolucion
        SetearFondoParaResolucion
        
        CCFFGG.Configuracion.Save
        
        fra4.Visible = False
        fra5.Visible = True
        'si selecciona la opcion utilizar seleccionado entonces guardo el nuevo sistema
        If optSeleccionado.Value Then GuardarNuevoSistemaCodificacion
        'optPredeterminado no esta, ver por que ...
        
        If bGuardarConfig Then
            Dim cd As New CommonDialog
            Dim sPath As String
            cd.DialogPrompt = "Seleccione la carpeta donde se guardara el archivo de ccffgg.configuracion."
            cd.DialogTitle = "Seleccione la carpeta"
            cd.InitDir = APh
            cd.ShowFolder
            If cd.SelectedDir = "" Then
                sPath = APh + "config.ini"
                MsgBox "El archivo de configuracion se guardo en: " + APh + "config.ini", vbInformation
            Else
                sPath = cd.SelectedDir + "config.ini"
            End If
            'aca guardo la config
            Dim conf As New ConfiguracionInicial
            conf.Pais = CCFFGG.Configuracion.Defaults.Pais
            conf.Barrio = CCFFGG.Configuracion.Defaults.Barrio
            conf.Ciudad = CCFFGG.Configuracion.Defaults.Ciudad
            conf.Provincia = CCFFGG.Configuracion.Defaults.Provincia
            'conf.IPServidor 'se carga sola en save
            conf.ModoFuncionamiento = CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento
            res = conf.Save(sPath)
            If res <> 0 And res <> 2118 Then '2118 es q ya existe el recurso compartido
                MsgBox "No se pudo compartir la carpeta de la base de datos. Debera compartirla y configurar cada estacion de trabajo manualmente. Puede continuar con la configuracion normalmente.", vbInformation
            End If
        End If
        
        cmdFinalizar.Caption = "Cerrar"
        cmdAnterior.Visible = False
        cmdSiguiente.Visible = False
    Else
        Unload Me
    End If

End Sub

Private Sub SetearFondoParaResolucion()
    Dim res As String
    Dim p As String
    Dim f As String
    res = GetSystemMetrics(SM_CXSCREEN) & "x" & GetSystemMetrics(SM_CYSCREEN)
     
    f = "\fondo01-" + res + ".jpg"
    p = APh + "..\Fondos\" + IIf(CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento = 1, "Emergencias", "Bomberos") + f
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (fso.FileExists(p)) Then
        CCFFGG.Configuracion.Apariencia.PathFondo = p
    End If
End Sub

Private Sub GuardarNuevoSistemaCodificacion()
    pb.Max = cantCodySint + 3
    
    lblOperacion.Caption = "Eliminando codigos anteriores..."
    DoEvents
    pb.Value = 1
    GBL.ExecuteSQL "Delete from CodigoEmergencia where id<100" 'elimino los codigos utilizados
    
    lblOperacion.Caption = "Eliminando sintomas anteriores..."
    DoEvents
    pb.Value = 2
    GBL.ExecuteSQL "Delete from Sintoma where idCodigoEmergencia<100" 'elimino los sintomas tmb
     
    lblOperacion.Caption = "Eliminando codigos por empresa..."
    DoEvents
    pb.Value = 3
    'gbl.ExecuteSQL "update Atencion set sintoma=0" 'para que
    GBL.ExecuteSQL "Delete from CodigoXEmpresa" 'borro todos los codigos por empresa
        
    Dim cod As blcemi.CodigoEmergencia
    Dim sint As blcemi.Sintoma
   
    For Each cod In codigos
    
        lblOperacion.Caption = "Agregando codigo " + cod.Nombre + "..."
        DoEvents
        pb.Value = pb.Value + 1
        GBL.ExecuteSQL _
            "Insert into CodigoEmergencia " + _
            "(id, nombre, bold, colorfuente, vencimiento) values(" + _
            Str(cod.id) + ", '" + cod.Nombre + "','False',0,0);"
            
        For Each sint In cod.Sintomas
            lblOperacion.Caption = "Agregando sintoma " + sint.Nombre + "..."
            pb.Value = pb.Value + 1
            DoEvents
            GBL.ExecuteSQL "Insert into Sintoma (id, nombre, idCodigoEmergencia) values(" + Str(sint.id) + ", '" + sint.Nombre + "'," + Str(cod.id) + ");"
        Next
    Next
          
    pb.Visible = False
    lblOperacion.Caption = "Finalizadas todas las operaciones, ya puede comenzar a utilizar el software."
End Sub

Private Sub cmdSiguiente_Click()
SetearFondoParaResolucion
    If ValidarPaso(pasoActual) Then
        OcultarFrames
        Select Case pasoActual
            Case 0:
                fra1.Visible = True 'configuracion de red
                pasoActual = 1
            Case 1:
                EstablecerConfigRed
            Case 2:
                fra3.Visible = True
                pasoActual = 3
            Case 3:
                fra4.Visible = True
                MostrarCodificacionesDisponibles IIf(optSE.Value, "Emergencias", "Bomberos")
                pasoActual = 4
            Case 4:
                fra5.Visible = True
                pasoActual = 5
            Case 5:
                fra6.Visible = True
                pasoActual = 6
        End Select
        HabilitarBotones
    End If
End Sub

Private Function ValidarPaso(pPaso As Integer) As Boolean
    Select Case pPaso
        Case 0: ValidarPaso = True
        Case 1:
            If optRed.Value Then
                ValidarPaso = (optServidor.Value Or optCliente.Value)
            Else
                ValidarPaso = True
            End If
        Case 2: 'validar passwords
            If txtPass.Text <> "" Then
                If txtPass.Text <> txtPass2.Text Then
                   MsgBox "Las claves no coinciden o falta ingresar la confirmación.", vbExclamation, "Aviso"
                    ValidarPaso = False
                Else
                    GBL.EmpleadosGBL.Item(1).Pass = txtPass
                    GBL.EmpleadosGBL.Item(1).GuardarModificaciones
                    MsgBox "Se actualizó la clave del administrador.", vbInformation, "Aviso"
                    ValidarPaso = True
                End If
            Else
                ValidarPaso = True
            End If
        Case 3: ValidarPaso = (optSE.Value Or optB.Value)
        Case 4: ValidarPaso = True
        Case 5: 'no valido la direccion, si ingresa o no depende del usuario
            ValidarPaso = True
                
    End Select
End Function

Private Sub EstablecerConfigRed()
If optMonoUsuario Then
    fra2.Visible = True
    pasoActual = 2
Else
    If optRed Then
        If optServidor Then
            fra2.Visible = True
            pasoActual = 2
            bGuardarConfig = True
        Else 'si eligio modo cliente
            fra6.Visible = True
            pasoActual = 5
            CargarConfiguracion
        End If
    End If
End If
End Sub

Private Sub CargarConfiguracion()
    pb.Max = cantCodySint + 3
    
    lblOperacion.Caption = "Buscando archivo de ccffgg.configuracion..."
    DoEvents
    pb.Value = 1
    
    Dim fso As New FileSystemObject
    Dim sPath As String
    If Not fso.FileExists(APh + "config.ini") Then
        Dim cd As New CommonDialog
abrirArchivo:
        cd.DialogPrompt = "Seleccione el archivo de configuracion"
        cd.DialogTitle = "Abrir archivo"
        cd.InitDir = APh
        cd.ShowOpen
        If cd.FileName = "" Then
            If MsgBox("Esta seguro que desea finalizar el asistente?.", vbQuestion + vbOKCancel) = vbOK Then
                Unload Me
            Else
                GoTo abrirArchivo 'WTF!
            End If
        Else
            sPath = cd.FileName
        End If
    Else
        sPath = APh + "config.ini"
        lblOperacion.Caption = "Cargando archivo de ccffgg.configuracion..."
        DoEvents
        pb.Value = 1
        
    End If
    
    Dim conf As New ConfiguracionInicial
    conf.Cargar sPath
    CCFFGG.Configuracion.Defaults.Barrio = conf.Barrio
    CCFFGG.Configuracion.Defaults.Ciudad = conf.Ciudad
    CCFFGG.Configuracion.Defaults.Provincia = conf.Provincia
    CCFFGG.Configuracion.Defaults.Pais = conf.Pais
    CCFFGG.Configuracion.Red.DirIPRemota = conf.IPServidor
    CCFFGG.Configuracion.Comportamiento.ModoFuncionamiento = conf.ModoFuncionamiento
    CCFFGG.Configuracion.DBLayer.PathDB = conf.PathDB
    CCFFGG.Configuracion.Save
          
    pb.Visible = False
    lblOperacion.Caption = "Finalizadas todas las operaciones, ya puede comenzar a utilizar el software."
End Sub

Private Sub HabilitarBotones()
    Select Case pasoActual
        Case 0:
            cmdFinalizar.Enabled = False
            'cmdAnterior.Enabled = False
            cmdSiguiente.Enabled = True
        Case 1:
            cmdFinalizar.Enabled = False
            'cmdAnterior.Enabled = True
            cmdSiguiente.Enabled = False
        Case 2:
            cmdFinalizar.Enabled = False
            cmdSiguiente.Enabled = True
            'cmdAnterior.Enabled = True
        Case 3:
            cmdFinalizar.Enabled = False
            'cmdAnterior.Enabled = True
            cmdSiguiente.Enabled = True
        Case 4:
            cmdFinalizar.Enabled = False
            'cmdAnterior.Enabled = True
            cmdSiguiente.Enabled = True
        Case 5:
            cmdFinalizar.Enabled = True
            'cmdAnterior.Enabled = True
            cmdSiguiente.Enabled = False
        
    End Select
End Sub

Private Sub OcultarFrames()
    fra0.Visible = False
    fra1.Visible = False
    fra2.Visible = False
    fra3.Visible = False
    fra4.Visible = False
    fra5.Visible = False
    fra6.Visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        MostrarAyuda
    End If
End Sub

Private Sub Label11_Click()
    MostrarAyuda
End Sub

Private Sub MostrarAyuda()
On Error GoTo errman
    Dim h As Long
    h = HtmlHelp(hWndAyudaHTML, APh + "Ayuda.chm" + "::/Configuracioninicial.htm", HH_DISPLAY_TOPIC, 0&)
    Exit Sub
errman:
End Sub

Private Sub Form_Load()
    'Set Me.Icon = MDI.Icon
    'Set cmbProvincia.Coleccion = gbl.ProvinciasGBL
    
    Set cmbPais.Coleccion = GBL.PaisesGBL
    'ASEGURARSE QUE PAISESGBL este cargado !
    If cmbPais.Coleccion.Count = 0 Then
        GBL.PrintToErrorLog "cfgInit", "SinPaises", "ErrorPGBL"
    End If
    
    fra1.Top = fra0.Top
    fra2.Top = fra0.Top
    fra3.Top = fra0.Top
    fra4.Top = fra0.Top
    fra5.Top = fra0.Top
    fra6.Top = fra0.Top
    fra1.Left = fra0.Left
    fra2.Left = fra0.Left
    fra3.Left = fra0.Left
    fra4.Left = fra0.Left
    fra5.Left = fra0.Left
    fra6.Left = fra0.Left
    fra1.Height = fra0.Height
    fra2.Height = fra0.Height
    fra3.Height = fra0.Height
    fra4.Height = fra0.Height
    fra5.Height = fra0.Height
    fra6.Height = fra0.Height
        
    Me.Height = cmdAnterior.Top + cmdAnterior.Height + 500
    Me.Width = fra0.Width + img.Width + img.Left * 4
    
    OcultarFrames
    HabilitarBotones
    fra0.Visible = True
    
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
End Sub


'-------------auxiliares de direcciones----------------------------------

Private Sub cmbBarrio_NuevoSeleccionado()

If Not cmbCiudad.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Barrio")
    If Nombre <> "" Then
        If cmbBarrio.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim b As blcemi.Barrio
            Set b = cmbBarrio.Coleccion.Nuevobarrio(Nombre, cmbCiudad.SelectedItem)
            cmbBarrio.Refresh
            Set cmbBarrio.SelectedItem = b
        End If
    End If
Else
    'debe seleccionar una ciudad primero
End If

End Sub

Private Sub cmbCiudad_ItemSeleccionado(Item As Object)
    Dim c As blcemi.Ciudad
    Set c = Item
    Set cmbBarrio.Coleccion = c.Barrios
End Sub

Private Sub cmbCiudad_NuevoSeleccionado()
If Not cmbProvincia.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Ciudad")
    If Nombre <> "" Then
        If cmbCiudad.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim c As blcemi.Ciudad
            Set c = cmbCiudad.Coleccion.Nuevaciudad(Nombre, cmbProvincia.SelectedItem)
            cmbCiudad.Refresh
            Set cmbCiudad.SelectedItem = c
        End If
    End If
Else
    'debe seleccionar una provincia antes
End If

End Sub

Private Sub cmbProvincia_ItemSeleccionado(Item As Object)
    Dim p As blcemi.Provincia
    Set p = Item
    Set cmbCiudad.Coleccion = p.Ciudades
    Set cmbBarrio.Coleccion = Nothing
End Sub

Private Sub cmbProvincia_NuevoSeleccionado()
If Not cmbPais.SelectedItem Is Nothing Then
    Dim Nombre As String
    Nombre = frmABMGenerico.Nuevo("Agregar Provincia")
    If Nombre <> "" Then
        If cmbProvincia.Coleccion.ItemByName(Nombre) Is Nothing Then
            Dim c As blcemi.Provincia
            Set c = cmbProvincia.Coleccion.NuevaProvincia(Nombre, cmbPais.SelectedItem)
            cmbProvincia.Refresh
            Set cmbProvincia.SelectedItem = c
        End If
    End If
Else
    'debe seleccionar un pais antes
End If

End Sub

Private Sub cmbPais_ItemSeleccionado(Item As Object)
    Dim p As blcemi.Pais
    Set p = Item
    If p.PrimerOrden <> "" Then lblProvincia.Caption = p.PrimerOrden + ":"
    If p.SegundoOrden <> "" Then lblCiudad.Caption = p.SegundoOrden + ":"
    If p.TercerOrden <> "" Then lblBarrio.Caption = p.TercerOrden + ":"

    Set cmbProvincia.Coleccion = p.Provincias
    Set cmbCiudad.Coleccion = Nothing
    Set cmbBarrio.Coleccion = Nothing
End Sub

'-----------------------auxiliares de codigos y sintomas--------------------------
Private Sub MostrarCodificacionSeleccionada(pPath As String)
    Dim XMLDoc As MSXML.DOMDocument
    Set XMLDoc = New DOMDocument
    XMLDoc.async = False
    
    XMLDoc.Load pPath
    
    If XMLDoc.parseError.errorCode = 0 Then
        If XMLDoc.readyState = 4 Then
            'TreeView1.Nodes.Clear
            AddNode XMLDoc.documentElement
        End If
    Else
        MsgBox "Hubo un error intentando mostrar el sistema de codificación de emergencias, por favor intente con otro sistema.", vbExclamation, "tbrEmergencyGroup"
    End If
End Sub

Private Sub AddNode(ByRef XML_Node As IXMLDOMNode)
    On Error Resume Next
    Dim xNode As Node
    Dim xNodeList As IXMLDOMNodeList
    Dim I As Long

    Set codigos = New blcemi.CodigoEmergenciaManager
    cantCodySint = 0
    Set xNodeList = XML_Node.childNodes
    For I = 0 To xNodeList.length - 1
        AddCodigo xNodeList.Item(I)
    Next
    Set tvw.Coleccion = codigos
End Sub
Private Sub AddCodigo(ByRef XML_Node As IXMLDOMNode)
    On Error Resume Next
    Dim xCod As blcemi.CodigoEmergencia
    Dim xNodeList As IXMLDOMNodeList
    Dim I As Long

    Set xCod = New blcemi.CodigoEmergencia
    xCod.id = Val(XML_Node.Attributes.getNamedItem("id").nodeValue)
    xCod.Nombre = XML_Node.Attributes.getNamedItem("nombre").nodeValue
    Set xCod.Sintomas = New blcemi.SintomaManager
    
    codigos.AddItem xCod
    cantCodySint = cantCodySint + 1
    Set xNodeList = XML_Node.childNodes
    For I = 0 To xNodeList.length - 1
        AddSintoma xNodeList.Item(I), xCod
    Next
End Sub

Private Sub AddSintoma(ByRef XML_Node As IXMLDOMNode, ByRef pCOD As blcemi.CodigoEmergencia)
    On Error Resume Next
    Dim xSin As blcemi.Sintoma
    Dim xNodeList As IXMLDOMNodeList
    Dim I As Long

    Set xSin = New blcemi.Sintoma
    cantCodySint = cantCodySint + 1
    
    xSin.id = Val(XML_Node.Attributes.getNamedItem("id").nodeValue)
    xSin.Nombre = XML_Node.Attributes.getNamedItem("nombre").nodeValue
    pCOD.Sintomas.AddItem xSin
End Sub


Private Sub List1_Click()
    MostrarCodificacionSeleccionada APh + "codificaciones\" + IIf(optSE.Value, "Emergencias", "Bomberos") + "\" + List1.Text + ".xml"
End Sub

'---------------activo el boton siguiente------------
Private Sub optCliente_Click()
    cmdSiguiente.Enabled = True
End Sub

Private Sub optMonoUsuario_Click()
    cmdSiguiente.Enabled = True
    fraModoRed.Visible = False
End Sub

Private Sub optRed_Click()
    cmdSiguiente.Enabled = True
    fraModoRed.Visible = True
End Sub

Private Sub optServidor_Click()
cmdSiguiente.Enabled = True
End Sub
