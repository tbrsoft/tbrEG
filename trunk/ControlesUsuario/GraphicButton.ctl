VERSION 5.00
Begin VB.UserControl GraphicButton 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   ScaleHeight     =   2385
   ScaleWidth      =   3600
   ToolboxBitmap   =   "GraphicButton.ctx":0000
   Begin VB.Image pic 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   50
      Stretch         =   -1  'True
      Top             =   50
      Width           =   1095
   End
   Begin VB.Line lShadowBottom 
      BorderColor     =   &H80000003&
      Visible         =   0   'False
      X1              =   30
      X2              =   3360
      Y1              =   2160
      Y2              =   1200
   End
   Begin VB.Line lShadowRight 
      BorderColor     =   &H80000010&
      Visible         =   0   'False
      X1              =   2640
      X2              =   3360
      Y1              =   360
      Y2              =   2160
   End
   Begin VB.Line lShadowTop 
      BorderColor     =   &H80000003&
      Visible         =   0   'False
      X1              =   50
      X2              =   3340
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line lShadowLeft 
      BorderColor     =   &H80000003&
      Visible         =   0   'False
      X1              =   30
      X2              =   30
      Y1              =   30
      Y2              =   2160
   End
   Begin VB.Line lRight 
      X1              =   3480
      X2              =   3480
      Y1              =   120
      Y2              =   2400
   End
   Begin VB.Line lBottom 
      X1              =   0
      X2              =   3480
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line lTop 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   3480
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line lLeft 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   2280
   End
End
Attribute VB_Name = "GraphicButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Click()

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.Enabled Then
        lShadowBottom.Visible = True
        lShadowLeft.Visible = True
        lShadowRight.Visible = True
        lShadowTop.Visible = True
    End If
End Sub

Private Sub pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lShadowBottom.Visible = False
    lShadowLeft.Visible = False
    lShadowRight.Visible = False
    lShadowTop.Visible = False
    If X > 0 And X < UserControl.Width And Y > 0 And Y < UserControl.Height And UserControl.Enabled Then RaiseEvent Click
End Sub

Private Sub UserControl_Resize()
    lTop.Y1 = 0
    lTop.Y2 = 0
    lTop.X2 = UserControl.Width
    lRight.X1 = UserControl.Width - 10
    lRight.X2 = UserControl.Width - 10
    lRight.Y1 = 30
    lRight.Y2 = UserControl.Height
    lBottom.X2 = UserControl.Width
    lBottom.Y1 = UserControl.Height - 10
    lBottom.Y2 = UserControl.Height - 10
    
    lShadowTop.Y1 = 30
    lShadowTop.Y2 = 30
    lShadowTop.X2 = UserControl.Width
    
    lShadowRight.X1 = UserControl.Width - 30
    lShadowRight.X2 = UserControl.Width - 30
    lShadowRight.Y1 = 30
    lShadowRight.Y2 = UserControl.Height
    lShadowBottom.X2 = UserControl.Width
    lShadowBottom.Y1 = UserControl.Height - 30
    lShadowBottom.Y2 = UserControl.Height - 30
    lShadowLeft.Y2 = UserControl.Height - 20
    pic.Height = UserControl.Height - 100
    pic.Width = UserControl.Width - 100
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal blnEnabled As Boolean)
    UserControl.Enabled = blnEnabled
    'PropertyChanged "Enabled"
    Refresh
End Property

Public Property Get Picture() As IPictureDisp

Set Picture = pic.Picture
End Property

Public Property Set Picture(pict As IPictureDisp)
    Set pic.Picture = pict
End Property


Private Sub UserControl_Show()
pic.ToolTipText = Extender.ToolTipText

End Sub
