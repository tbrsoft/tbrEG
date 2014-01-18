VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsApariencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private vGridLinesConsultas As Boolean
Private vGridLinesOtros As Boolean
Private mTitleFont As StdFont
Private mContentsFont As StdFont
Dim vPathFondo As String

Private vToolBarStyleFlat As Boolean

Public Property Get PathFondo() As String
    PathFondo = vPathFondo
End Property

Public Property Let PathFondo(pValue As String)
    vPathFondo = pValue
End Property

Public Property Get GridLinesConsultas() As Boolean
    GridLinesConsultas = vGridLinesConsultas
End Property

Public Property Let GridLinesConsultas(pValor As Boolean)
     vGridLinesConsultas = pValor
End Property

Public Property Get GridLinesOtros() As Boolean
    GridLinesOtros = vGridLinesOtros
End Property

Public Property Let GridLinesOtros(pValor As Boolean)
     vGridLinesOtros = pValor
End Property

Public Property Get ToolBarStyleFlat() As Boolean
    ToolBarStyleFlat = vToolBarStyleFlat
End Property

Public Property Let ToolBarStyleFlat(pValor As Boolean)
     vToolBarStyleFlat = pValor
End Property

Public Property Get TitleFont() As StdFont
    Set TitleFont = mTitleFont
End Property

Public Property Set TitleFont(pValor As StdFont)
     Set mTitleFont = pValor
End Property

Public Property Get ContentsFont() As StdFont
    Set ContentsFont = mContentsFont
End Property

Public Property Set ContentsFont(pValor As StdFont)
     Set mContentsFont = pValor
End Property

Private Sub Class_Initialize()
    Load
End Sub

Friend Sub Load()
    vGridLinesConsultas = IIf(GetSetting("TbrEmergencyGroup", "Apariencia", "GridLinesConsulta", "True") = "True", True, False)
    vGridLinesOtros = IIf(GetSetting("TbrEmergencyGroup", "Apariencia", "GridLinesOtros", "True") = "True", True, False)
    
    'cargo las fuentes
    
    Dim pb As New PropertyBag
    Dim byteArr() As Byte
    aux = GetSetting("TbrEmergencyGroup", "Apariencia", "FontTitle")
    If aux <> "" Then
        byteArr = aux
        pb.Contents = byteArr
        Set mTitleFont = pb.ReadProperty("font")
    Else
        Set mTitleFont = GetFont("Times New Roman", 14, True)
    End If
    
    aux = GetSetting("TbrEmergencyGroup", "Apariencia", "FontContents")
    If aux <> "" Then
        byteArr = aux
        pb.Contents = byteArr
        Set mContentsFont = pb.ReadProperty("font")
    Else
        Set mContentsFont = GetFont("Times New Roman", 12)
    End If
    'MR01
    'el app.path de una dll es generalmente c:/windows/system32
    'en el instalador cambie el destino de config.dll de {sys} a {app}
    'además saque el "../" del renglon de abajo
    vPathFondo = GetSetting("TbrEmergencyGroup", "Apariencia", "PathFondo", App.Path + "\Fondos\fondo_01.jpg")
End Sub

Friend Sub Save()
    Dim pb As New PropertyBag
    pb.WriteProperty "font", mTitleFont
    SaveSetting "TbrEmergencyGroup", "Apariencia", "FontTitle", pb.Contents
    
    Set pb = New PropertyBag
    pb.WriteProperty "font", mContentsFont
    SaveSetting "TbrEmergencyGroup", "Apariencia", "FontContents", pb.Contents
    
    SaveSetting "TbrEmergencyGroup", "Apariencia", "GridLinesConsulta", IIf(vGridLinesConsultas, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Apariencia", "GridLinesOtros", IIf(vGridLinesOtros, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Apariencia", "PathFondo", vPathFondo
End Sub

Private Sub Class_Terminate()
    Save
End Sub

Private Function GetFont(pFontName As String, Optional pFontSize As Integer, Optional pBold As Boolean = False, Optional pItalic As Boolean = False, Optional pUnderline As Boolean = False) As StdFont
    Dim mFont As New StdFont
    Dim founded As Boolean
    
    For i = 0 To Screen.FontCount
        If Screen.Fonts(i) = pFontName Then
            mFont.Name = Screen.Fonts(i)
            founded = True
            Exit For
        End If
    Next
    
    If founded Then
        mFont.Bold = pBold
        mFont.Italic = pItalic
        mFont.Underline = pUnderline
        mFont.Size = pFontSize
        
        Set GetFont = mFont
    Else
        Set GetFont = Nothing
    End If
    
End Function
