VERSION 5.00
Begin VB.Form frmprueba 
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmprueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mFormul As tbrcamposimpresion.Formulario

Public Sub Imprimir(formul As tbrcamposimpresion.Formulario)
Me.Show
Set mFormul = formul
End Sub

Private Sub Form_Paint()
mFormul.Imprimir Me

End Sub
