VERSION 5.00
Begin VB.Form PlantillaABM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   6360
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Frame fraDatos 
      Height          =   3615
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "PlantillaABM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hacer:
    'Declarar los eventos Nuevo, Modificar, Eliminar
    'Implementar las funciones Nuevo,Modificar, Consulta, Eliminar

Private tipo As eTipoAMB 'enumeracion definida en Modulo

'declarar variable privada del tipo de objeto del que se quiera hacer abmc

Private Sub cmdAceptar_Click()
    
    Select Case tipo
        Case etALTA
            'implementar
        Case etBAJA
            'implementar
        Case etMODIFICACION
            'implementar
        Case etCONSULTA
            'implementar
    End Select
    
End Sub

Private Sub cmdCancelar_Click()
'implementar
End Sub

Public Sub Nuevo() 'mandar la coleccion como parametro
'implementar
tipo = etALTA
Me.Show
End Sub

Public Sub Modificar() 'mandar como parametro el elemento a modificar
'implementar
tipo = etMODIFICACION
Me.Show
End Sub

Public Sub Eliminar() 'mandar como parametro el elemento a eliminar
'implementar
tipo = etBAJA
Me.Show
End Sub

Public Sub VerDatos() 'mandar como parametro el elemento del q queremos ver los datos
'implementar
tipo = etCONSULTA
Me.Show
End Sub

Private Sub Form_Load()
'levanta un error si quiere usar el metodo show
If tipo = 0 Then Err.Raise 2009, , "No se puede mostrar el formulario con el metodo Show, utilice las funciones Nuevo, Modificar, Eliminar o VerDatos."
End Sub
