Attribute VB_Name = "modErrorHandling"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"46A921750245"
Attribute VB_Ext_KEY = "RVB_ModelStereotype" ,"Debug.ErrorHandling"
Option Explicit

Public Const MyUnhandledError = 9999

Private mLog As New Log
Private PathDef As String

'Public Sub RaiseError(ErrorNumber As Long, Message As String)
'      mLog.WriteToLog App.path + "\errorLog.txt", Trim(Str(Now)) + " - " + Message
'End Sub

Public Sub ErrorLog(pClass As String, pPropertyOrFunctionName As String, pError As String, Optional sPath As String = "")
    On Error GoTo errman:
    
    If sPath = "" Then
        If PathDef = "" Then
            PathDef = App.Path
            If Right(PathDef, 1) <> "\" Then PathDef = PathDef + "\"
        Else
            sPath = PathDef
        End If
    Else
        If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
        PathDef = sPath
    End If
    
    Dim informe As String
    informe = "Fecha - Hora: " + Trim(Str(Now)) + vbCrLf + "Software: tbrEmergencyGroup" + vbCrLf + "Version: " + Str(App.Major) + "." + Str(App.Minor) + "." + Str(App.Revision) + vbCrLf
     
    'informe = informe + "Version BD:" + BD.GetDBVersion + vbCrLf
    informe = informe + "Clase:" + pClass + vbCrLf + "Propiedad o Funcion: " + pPropertyOrFunctionName + vbCrLf + "Error: " + pError
    ErrHandlerLocal.InformarError informe
    mLog.WriteToLog PathDef + "regTbrEG.log", informe + vbCrLf + vbCrLf
    
    TERR.AppendLog "ErrLOG: ****" + vbCrLf + informe + vbCrLf
    
    Exit Sub
    
errman:
    mLog.WriteToLog PathDef + "errorLog.txt", "error en registrar error!" + vbCrLf + _
                                                Err.Description + vbCrLf + _
                                                CStr(Err.Number)
End Sub
