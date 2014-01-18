VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{03F93260-B914-4BA7-8E50-6C4C3BB2BAD9}#1.2#0"; "ListViewConsultaCtl2.ocx"
Begin VB.Form frmListadoGenerico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
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
         Format          =   45088769
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
   Begin MSComctlLib.Toolbar tBar 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "word"
            Object.ToolTipText     =   "Exporta el listado a MS Word"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "excel"
            Object.ToolTipText     =   "Exporta el listado a MS Excel"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "writer"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Writer"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "calc"
            Object.ToolTipText     =   "Exporta el listado a OpenOffice Calc"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ph"
            Style           =   4
            Object.Width           =   5000
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cierra el formulario"
         EndProperty
      EndProperty
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
Dim mListado As listado

Private Sub cmdMostrar_Click()
'On Error GoTo errman
Dim c As Control
For Each c In dtp
    If c.Index <> 0 Then mListado.Parametros.Item(c.Tag).Valor = c.Value
Next

For Each c In txt
    If c.Index <> 0 Then mListado.Parametros.Item(c.Tag).Valor = c.Text
Next

Mostrar
End Sub

Private Sub Mostrar()

    Dim b As Button
    For Each b In tBar.Buttons
        If b.Style = tbrDefault Then b.Enabled = True
    Next
    
    Set lvw.Encabezados = mListado.Encabezados
    Set lvw.Coleccion = mListado.GetColeccion

End Sub

Public Sub MostrarListado(pListado As listado)
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
            Case "string", "integer", "long", "double", "single"
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
    Dim c As Control
    For Each c In dtp
        c.Left = maxWidth + 300
    Next
    
    For Each c In txt
        c.Left = maxWidth + 300
    Next
    cmdMostrar.Left = fraParametros.Width - cmdMostrar.Width - 100
    fraParametros.Top = tBar.Height + 100
    fraParametros.Height = lbl(lblCount).Top + lbl(lblCount).Height + 150
    lvw.Top = fraParametros.Height + 200 + fraParametros.Top
    lvw.Height = Me.ScaleHeight - lvw.Top
    'terminar de arreglar los tamaños
Else
    fraParametros.Visible = False
    lvw.Top = tBar.Height
    lvw.Height = Me.ScaleHeight - lvw.Top
    Mostrar
End If

Me.Show
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set tBar.ImageList = MDI.il32 'ver si esta o otra il
    Dim b As Button
    For Each b In tBar.Buttons
        If b.Style = tbrDefault Then
            b.Image = b.Key
            If b.Key <> "cancelar" Then b.Enabled = False
        End If
    Next

    Set Me.Icon = MDI.Icon
End Sub

Private Sub AplicarConfiguracion()
    lvw.GridLines = CCFFGG.Configuracion.Apariencia.GridLinesConsultas
    tBar.Buttons("calc").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToCalc
    tBar.Buttons("writer").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWrite
    tBar.Buttons("word").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToWord
    tBar.Buttons("excel").Visible = CCFFGG.Configuracion.Comportamiento.AllowExportToExcel
    
End Sub

'Private Sub Form_Resize()
'    If Me.WindowState = vbNormal Or Me.WindowState = vbMaximized Then
'        If Me.Width < ANCHOMIN Then Me.Width = ANCHOMIN
'        If Me.Height < ALTOMIN Then Me.Height = ALTOMIN
'
'        txtFiltro.Top = tBar.Height
'        lvw.Top = txtFiltro.Top + txtFiltro.Height
'        lvw.Height = Me.ScaleHeight - lvw.Top
'        txtFiltro.Width = Me.Width - 100
'        lvw.Width = Me.Width - 100
'
'        DistribuirBotones tBar
'
'    End If
'End Sub
Private Sub tBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "word"
            lvw.ExportToWord mListado.Titulo, , CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
        Case "excel"
            lvw.ExportToExcel mListado.Titulo
        Case "calc"
            lvw.ExportToOOCalc mListado.Titulo
        Case "write"
            lvw.ExportToOOWriter mListado.Titulo, CCFFGG.Configuracion.Apariencia.ContentsFont, CCFFGG.Configuracion.Apariencia.TitleFont
    End Select
End Sub
