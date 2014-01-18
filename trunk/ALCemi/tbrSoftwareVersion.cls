VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrSoftwareVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event Done(pLastVersion As String)
Public Event Error(pMessage As String)

Dim WithEvents stateHandler As MyReadyStateHandler
Attribute stateHandler.VB_VarHelpID = -1
Dim oHttReq As XMLHTTPRequest

Private Const cSOAP = _
    "<?xml version=""1.0"" encoding=""utf-8""?>" & _
    "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
        "<soap:Body>" & _
            "<GetLastVersion xmlns=""http://www.tbrsoft.com.ar/"" > " & _
      "<pSoftwareId>string</pSoftwareId>" & _
     " </GetLastVersion>" & _
       "</soap:Body>" & _
    "</soap:Envelope>"

Public Sub GetLastVersion(pSoftwareID As String)
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    ' cargar el código SOAP para CaF
    parser.loadXML cSOAP
    '
    ' Indicar el parámetro a enviar
    parser.selectSingleNode("/soap:Envelope/soap:Body/GetLastVersion/pSoftwareId").Text = pSoftwareID
    '
    ' Mostrar el código XML enviado al servicio Web
    'Text1.Text = parser.xml
    '
    '
        
  ' Enviar el comando al servicio Web
    '
    ' usar XMLHTTPRequest para enviar la información al servicio Web
    

    Set oHttReq = New XMLHTTPRequest
    
    Set stateHandler = New MyReadyStateHandler
    oHttReq.OnReadyStateChange = stateHandler
    
    ' Enviar el comando de forma asíncrona (no se espera a que se reciba la respuesta)
    oHttReq.open "POST", "http://www.tbrsoft.com.ar/ws/softwareversion/tbrSoftwareVersionWS.asmx", True
    ' las cabeceras a enviar al servicio Web
    ' (no incluir los dos puntos en el nombre de la cabecera)
    oHttReq.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
    oHttReq.setRequestHeader "SOAPAction", "http://www.tbrsoft.com.ar/GetLastVersion"
    ' enviar el comando
    oHttReq.send parser.xml

    
End Sub

Private Function procesarRespuesta(ByVal s As String) As String
    ' procesar la respuesta recibida del servicio Web
      
    ' Poner los datos en el analizador de XML
    Dim parser As DOMDocument
    Set parser = New DOMDocument
    parser.loadXML s
    '
    On Error GoTo errman
    '
    Dim respuesta As String
    respuesta = parser.selectSingleNode("/soap:Envelope/soap:Body/GetLastVersionResponse/GetLastVersionResult").Text
    
    procesarRespuesta = respuesta
    
    Exit Function
errman:
     procesarRespuesta = "Error."
End Function

Private Sub stateHandler_ReadyStateChanged()
    If oHttReq.readyState = 4 Then
        Select Case oHttReq.Status
        Case 200:
            RaiseEvent Done(procesarRespuesta(oHttReq.responseText))
        Case Else:
            RaiseEvent Error("Error en la solicitud.")
        End Select
    End If
End Sub
