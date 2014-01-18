VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private mEsPK As Boolean
Private mNombre As String
Private mValue As Variant

Public Property Get Nombre() As String
    Nombre = mNombre
End Property

Public Property Get Valor() As Variant
    Valor = mValue
End Property

Public Property Let Nombre(pNombre As String)
    mNombre = pNombre
End Property

Public Property Let Valor(pValue As Variant)
    mValue = pValue
End Property

Public Property Get EsPK() As Boolean
    EsPK = mEsPK
End Property

Public Property Let EsPK(pValue As Boolean)
    mEsPK = pValue
End Property
