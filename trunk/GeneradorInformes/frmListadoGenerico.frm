VERSION 5.00
Object = "{1417CD23-5617-4303-9AEF-2418F695BFFF}#1.0#0"; "ListViewConsultaCtl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoGenerico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   8730
   Begin VB.Frame fraParametros 
      Caption         =   "Parametros"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8535
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar Listado"
         Height          =   375
         Left            =   6360
         TabIndex        =   4
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txt 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtp 
         Height          =   375
         Index           =   0
         Left            =   2880
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         Format          =   20971521
         CurrentDate     =   39870
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Height          =   375
         Index           =   0
         Left            =   200
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   45
      End
   End
   Begin ControlesPOO.ListViewConsulta lvw 
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8705
      HideSelection   =   0   'False
      HideEncabezados =   0   'False
      GridLines       =   0   'False
      FullRowSelection=   0   'False
      AutoDistribuirColumnas=   0   'False
      CampoKey        =   ""
      AllowModify     =   0   'False
      ShowCheckBoxes  =   0   'False
      MultiSelect     =   0   'False
      CampoImage      =   ""
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
End
Attribute VB_Name = "frmListadoGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lpars As New LParameterManager
Dim mListado As Listado

Private Sub cmdMostrar_Click()
'On Error GoTo errman
Dim C As Control
For Each C In dtp
    If C.Index <> 0 Then mListado.Parametros.Item(C.Tag).Valor = C.Value
Next

For Each C In txt
    If C.Index <> 0 Then mListado.Parametros.Item(C.Tag).Valor = C.Text
Next

Mostrar
End Sub

Private Sub Mostrar()
    
    Set lvw.Encabezados = mListado.Encabezados
    Set lvw.Coleccion = mListado.GetColeccion

End Sub

Public Sub MostrarListado(pListado As Listado)
Dim lblCount As Integer
Dim dtpCount As Integer
Dim txtCount As Integer
Dim maxWidth As Single

Set mListado = pListado
Me.Caption = "Listado - " + mListado.Titulo

If Not mListado.Parametros Is Nothing Then
    Dim lp As LParameter
    For Each lp In mListado.Parametros
        
        lblCount = lblCount + 1
        Load lbl(lblCount)
        With lbl(lblCount)
            .AutoSize = True
            .Caption = lp.Descripcion
            .AutoSize = False
            .Height = 375
            .Visible = True
            .Top = .Height * (lblCount - 1) + 250
            If .Width > maxWidth Then maxWidth = .Width
        End With
        Select Case LCase(lp.Tipo)
            Case "date"
                dtpCount = dtpCount + 1
                Load dtp(dtpCount)
                With dtp(dtpCount)
                    .Visible = True
                    .Top = .Height * (lblCount - 1) + 250
                    .Value = Date
                    .Tag = lp.Nombre
                End With
            Case "string"
                txtCount = txtCount + 1
                Load txt(txtCount)
                With txt(txtCount)
                    .Visible = True
                    .Top = .Height * (lblCount - 1) + 250
                    .Text = ""
                    .Tag = lp.Nombre
                End With
        End Select
        
    Next
    Dim C As Control
    For Each C In dtp
        C.Left = maxWidth + 300
    Next
    
    For Each C In txt
        C.Left = maxWidth + 300
    Next
    cmdMostrar.Left = fraParametros.Width - cmdMostrar.Width - 100
    fraParametros.Top = 100
    fraParametros.Height = lbl(lblCount).Top + lbl(lblCount).Height + 150
    lvw.Top = fraParametros.Height + 200 + fraParametros.Top
    lvw.Height = Me.ScaleHeight - lvw.Top
    'terminar de arreglar los tamaños
Else
    fraParametros.Visible = False
    lvw.Top = 0
    lvw.Height = Me.ScaleHeight - lvw.Top
    Mostrar
End If

Me.Show
End Sub

