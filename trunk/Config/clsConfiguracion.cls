VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Event ConfigChanged()
Private vApariencia As clsApariencia
Private vComportamiento As clsComportamiento
Private vRed As clsRed
Private vDBLayer As clsBDLayer
Private vCodigo As clsCodigo
Private vDefault As clsDefaults
Private vRequeridos As clsCamposRequeridos

Public Property Get Requeridos() As clsCamposRequeridos
    If vRequeridos Is Nothing Then Set vRequeridos = New clsCamposRequeridos
    Set Requeridos = vRequeridos
End Property

Public Property Set Requeridos(pValor As clsCamposRequeridos)
    Set vRequeridos = pValor
End Property

Public Property Get Defaults() As clsDefaults
    If vDefault Is Nothing Then Set vDefault = New clsDefaults
    Set Defaults = vDefault
End Property

Public Property Set Defaults(pValor As clsDefaults)
    Set vDefault = pValor
End Property

Public Property Get Apariencia() As clsApariencia
    If vApariencia Is Nothing Then Set vApariencia = New clsApariencia
    Set Apariencia = vApariencia
End Property

Public Property Set Apariencia(pValor As clsApariencia)
    Set vApariencia = pValor
End Property

Public Property Get Comportamiento() As clsComportamiento
    If vComportamiento Is Nothing Then
        Set vComportamiento = New clsComportamiento
    End If
    Set Comportamiento = vComportamiento
End Property

Public Property Set Comportamiento(pValor As clsComportamiento)
    Set vComportamiento = pValor
End Property

Public Property Get Red() As clsRed
    If vRed Is Nothing Then Set vRed = New clsRed
    Set Red = vRed
End Property

Public Property Set Red(pValor As clsRed)
    Set vRed = pValor
End Property

Public Property Get DBLayer() As clsBDLayer
    If vDBLayer Is Nothing Then Set vDBLayer = New clsBDLayer
    Set DBLayer = vDBLayer
End Property

Public Property Set DBLayer(pValor As clsBDLayer)
    Set vDBLayer = pValor
End Property

Public Property Get Codigo() As clsCodigo
    If vCodigo Is Nothing Then Set vCodigo = New clsCodigo
    Set Codigo = vCodigo
End Property

Public Property Set Codigo(pValor As clsCodigo)
    Set vCodigo = pValor
End Property

Public Sub ConfiguracionModificada()
    'guardar aca, creo...
    RaiseEvent ConfigChanged
End Sub

Public Sub Save()
    Apariencia.Save
    Comportamiento.Save
    Red.Save
    DBLayer.Save
    Codigo.Save
    Defaults.Save
    Requeridos.Save
End Sub

Public Sub Load()
    Apariencia.Load
    Comportamiento.Load
    Red.Load
    Codigo.Load
    DBLayer.Load
    Defaults.Load
    Requeridos.Load
End Sub

