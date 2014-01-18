VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InterfazVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub SaveSettingA(appName As String, section As String, key As String, setting As Variant)
    SaveSetting appName, section, key, CStr(setting)
End Sub

Public Function ShellA(PathName, Optional WindowStyle As VbAppWinStyle = vbMinimizedFocus) As Double
    ShellA = Shell(PathName, WindowStyle)
End Function

