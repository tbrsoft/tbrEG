VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComportamiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private vSepararAtenciones As Boolean
Private vMostrarSugerenciasDatosFaltantes As Boolean
Private vMostrarBarraMenu As Boolean
Private vAllowExportToWord As Boolean
Private vAllowExportToExcel As Boolean
Private vAllowExportToWrite As Boolean
Private vAllowExportToCalc As Boolean
Private vMostrarAvisoAtencionesPendientes As Boolean

Private vModoFuncionamiento As Integer
Private vEnviarErrores As Integer

Public Property Get ModoFuncionamiento() As Integer
    ModoFuncionamiento = vModoFuncionamiento
End Property

Public Property Let ModoFuncionamiento(pValor As Integer)
     vModoFuncionamiento = pValor
End Property

Public Property Get EnviarErrores() As Integer
    EnviarErrores = vEnviarErrores
End Property

Public Property Let EnviarErrores(pValor As Integer)
     vEnviarErrores = pValor
End Property

Public Property Get SepararAtenciones() As Boolean
    SepararAtenciones = vSepararAtenciones
End Property

Public Property Let SepararAtenciones(pValor As Boolean)
     vSepararAtenciones = pValor
End Property

Public Property Get MostrarSugerenciasDatosFaltantes() As Boolean
    MostrarSugerenciasDatosFaltantes = vMostrarSugerenciasDatosFaltantes
End Property

Public Property Let MostrarSugerenciasDatosFaltantes(pValor As Boolean)
     vMostrarSugerenciasDatosFaltantes = pValor
End Property

Public Property Get MostrarAvisoAtencionesPendientes() As Boolean
    MostrarAvisoAtencionesPendientes = vMostrarAvisoAtencionesPendientes
End Property

Public Property Let MostrarAvisoAtencionesPendientes(pValor As Boolean)
     vMostrarAvisoAtencionesPendientes = pValor
End Property

Public Property Get MostrarBarraMenu() As Boolean
    MostrarBarraMenu = vMostrarBarraMenu
End Property

Public Property Let MostrarBarraMenu(pValor As Boolean)
     vMostrarBarraMenu = pValor
End Property

Public Property Get AllowExportToWord() As Boolean
    AllowExportToWord = vAllowExportToWord
End Property

Public Property Let AllowExportToWord(pValor As Boolean)
     vAllowExportToWord = pValor
End Property

Public Property Get AllowExportToWrite() As Boolean
    AllowExportToWrite = vAllowExportToWrite
End Property

Public Property Let AllowExportToWrite(pValor As Boolean)
     vAllowExportToWrite = pValor
End Property

Public Property Get AllowExportToCalc() As Boolean
    AllowExportToCalc = vAllowExportToCalc
End Property

Public Property Let AllowExportToCalc(pValor As Boolean)
     vAllowExportToCalc = pValor
End Property

Public Property Get AllowExportToExcel() As Boolean
    AllowExportToExcel = vAllowExportToExcel
End Property

Public Property Let AllowExportToExcel(pValor As Boolean)
     vAllowExportToExcel = pValor
End Property

Private Sub Class_Initialize()
    Load
End Sub

Friend Sub Load()
    vSepararAtenciones = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "SepararAtenciones", "True") = "True", True, False)
    vAllowExportToExcel = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "AllowExportToExcel", "True") = "True", True, False)
    vAllowExportToWord = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "AllowExportToWord", "True") = "True", True, False)
    vAllowExportToWrite = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "AllowExportToWrite", "True") = "True", True, False)
    vAllowExportToCalc = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "AllowExportToCalc", "True") = "True", True, False)
    vMostrarSugerenciasDatosFaltantes = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "MostrarSugerencias", "True") = "True", True, False)
    vMostrarBarraMenu = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "MostrarBarraMenu", "False") = "True", True, False)
    vMostrarAvisoAtencionesPendientes = IIf(GetSetting("TbrEmergencyGroup", "Comportamiento", "MostrarAvisoAtencionesPendientes", "True") = "True", True, False)
    vModoFuncionamiento = CInt(GetSetting("TbrEmergencyGroup", "Comportamiento", "ModoFuncionamiento", "0"))
    vEnviarErrores = CInt(GetSetting("TbrEmergencyGroup", "Comportamiento", "EnviarErrores", "0"))

End Sub

Private Sub Class_Terminate()
    Save
End Sub

Friend Sub Save()
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "SepararAtenciones", IIf(vSepararAtenciones, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "AllowExportToExcel", IIf(vAllowExportToExcel, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "AllowExportToWord", IIf(vAllowExportToWord, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "AllowExportToWrite", IIf(vAllowExportToWrite, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "AllowExportToCalc", IIf(vAllowExportToCalc, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "MostrarSugerencias", IIf(vMostrarSugerenciasDatosFaltantes, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "MostrarBarraMenu", IIf(vMostrarBarraMenu, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "MostrarAvisoAtencionesPendientes", IIf(vMostrarAvisoAtencionesPendientes, "True", "False")
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "ModoFuncionamiento", Trim(Str(vModoFuncionamiento))
    SaveSetting "TbrEmergencyGroup", "Comportamiento", "EnviarErrores", Trim(Str(vEnviarErrores))
End Sub
