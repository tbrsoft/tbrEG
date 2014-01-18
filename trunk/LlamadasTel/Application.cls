VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim mClientToNotify As IClient

'esta funcion es llamada por el cliente
'para setear el objeto al cual envio las notificaciones de llamadas
Public Sub NotifyMe(ByVal pClient As IClient)
    frmLlamada.Show
    frmLlamada.setApp Me
    Set mClientToNotify = pClient
End Sub

'esta funcion es llamada por este mismo programa para avisar a los clientes q ocurrio una llamada
Friend Sub NotifyIncomingCall(pNumero As String)
    mClienteANotificar.LlamadaEntrante 1, pNumero
End Sub

'----------------miembros accesibles desde otros programas---------

Public Function TelephoneCalls() As TelephoneCallManager
    Set TelephoneCalls = LocalTelCalls
End Function

Public Function Users() As UserManager
    Set Users = LocalUsers
End Function




