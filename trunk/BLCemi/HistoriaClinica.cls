VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistoriaClinica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Public Enum eHCTipoAfiliado
'    eHCAfiliadoPropio = 1
'    eHCAfiliadoExterno = 2
'End Enum

Private mMedicamentos As MedicamentoManager
Private mAlergias As AlergiaManager
Private mEnfermedades As EnfermedadManager

Private mAfiliado As Object

Private bCargada As Boolean

Public Property Get Medicamentos() As MedicamentoManager
    On Error GoTo MedicamentoErr

    Set Medicamentos = mMedicamentos
    
    Exit Property
MedicamentoErr:
    Call ErrorLog("HistoriaClinica", "Medicamentos Property Get", Err.Description)
End Property

Public Property Get Enfermedades() As EnfermedadManager
    On Error GoTo EnfermedadesErr

    Set Enfermedades = mEnfermedades
    
    Exit Property
EnfermedadesErr:
    Call ErrorLog("HistoriaClinica", "Enfermedades Property Get", Err.Description)
End Property

Public Property Get Alergias() As AlergiaManager
    On Error GoTo AlergiasErr

    Set Alergias = mAlergias

    Exit Property
AlergiasErr:
    Call ErrorLog("HistoriaClinica", "Alergias Property Get", Err.Description)
End Property

Friend Property Get Cargada() As Boolean
    Cargada = bCargada
End Property

Friend Sub Cargar(pAfiliado As Object)
    Set mAfiliado = pAfiliado
    bCargada = True
    Set mAlergias = New AlergiaManager
    Set mEnfermedades = New EnfermedadManager
    Set mMedicamentos = New MedicamentoManager
    Alergias.CargarXAfiliado mAfiliado
    Enfermedades.CargarXAfiliado mAfiliado
    Medicamentos.CargarXAfiliado mAfiliado
End Sub

Public Sub Inicializar(pAlergias As AlergiaManager, pEnfermedades As EnfermedadManager, pMedicamentos As MedicamentoManager)
    Set mAlergias = pAlergias
    Set mEnfermedades = pEnfermedades
    Set mMedicamentos = pMedicamentos
    bCargada = True
End Sub

Friend Sub Save(pTrans As Transaccion, pAfiliado As Object)
    Set mAfiliado = pAfiliado
    If TypeOf mAfiliado Is Afiliado Then
        Alergias.Guardar pTrans, pAfiliado.Id, tAlergiaXAfiliado
        Enfermedades.Guardar pTrans, pAfiliado.Id, tEnfermedadXAfiliado
        Medicamentos.Guardar pTrans, pAfiliado.Id, tMedicamentoXAfiliado
    ElseIf TypeOf mAfiliado Is AfiliadoExterno Then
        Alergias.Guardar pTrans, pAfiliado.Id, tAlergiaXAfiliadoExterno
        Enfermedades.Guardar pTrans, pAfiliado.Id, tEnfermedadXAfiliadoExterno
        Medicamentos.Guardar pTrans, pAfiliado.Id, tMedicamentoXAfiliadoExterno
    End If
End Sub

Public Sub BeginEdit()
    Alergias.BeginEdit
    Enfermedades.BeginEdit
    Medicamentos.BeginEdit
End Sub

Public Sub CancelChanges()
    Alergias.CancelChanges
    Enfermedades.CancelChanges
    Medicamentos.CancelChanges
End Sub

Friend Sub SaveChanges(pTrans As Transaccion)
    If TypeOf mAfiliado Is Afiliado Then
        Alergias.SaveChanges pTrans, mAfiliado.Id, tAlergiaXAfiliado
        Enfermedades.SaveChanges pTrans, mAfiliado.Id, tEnfermedadXAfiliado
        Medicamentos.SaveChanges pTrans, mAfiliado.Id, tMedicamentoXAfiliado
    ElseIf TypeOf mAfiliado Is AfiliadoExterno Then
        Alergias.SaveChanges pTrans, mAfiliado.Id, tAlergiaXAfiliadoExterno
        Enfermedades.SaveChanges pTrans, mAfiliado.Id, tEnfermedadXAfiliadoExterno
        Medicamentos.SaveChanges pTrans, mAfiliado.Id, tMedicamentoXAfiliadoExterno
    End If
End Sub


'Friend Property Set Enfermedades(ByVal vNewValue As EnfermedadManager)
'    On Error GoTo EnfermedadesErr
'
'    Set mEnfermedades = vNewValue
'
'    Exit Property
'EnfermedadesErr:
'    Call ErrorLog("HistoriaClinica", "Enfermedades Property Set")
'End Property

'Friend Property Set Alergias(ByVal vNewValue As AlergiaManager)
'    On Error GoTo AlergiasErr
'
'    Set mAlergias = vNewValue
'
'    Exit Property
'AlergiasErr:
'    Call ErrorLog("HistoriaClinica", "Alergias Property Set")
'End Property


'Friend Property Set Medicamentos(ByVal vNewValue As EnfermedadManager)
'    On Error GoTo MedicamentoErr
'
'    Set mMedicamentos = vNewValue
'
'    Exit Property
'MedicamentoErr:
'    Call ErrorLog("HistoriaClinica", "Medicamentos Property Set")
'End Property
