VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcNcPreguntaRespuesta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private PREGUNTA As String
Private RESPUESTA As String
Private TIPO_PREGUNTA_RESPUESTA As enumTipoPreguntaRespuesta
Private ID_PREGUNTA_RESPUESTA As Long
Private REQUERIDA As Boolean
Private ID_AUX As Integer
Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property

Public Property Let setID_AUX(ByVal intID_AUX As Integer)
    ID_AUX = intID_AUX
End Property


Public Property Get getREQUERIDA() As Boolean

    getREQUERIDA = REQUERIDA

End Property

Public Property Let setREQUERIDA(ByVal valor As Boolean)

    REQUERIDA = valor

End Property

Public Property Get getPREGUNTA() As String

    getPREGUNTA = PREGUNTA

End Property

Public Property Let setPREGUNTA(ByVal valor As String)

    PREGUNTA = valor

End Property

Public Property Get getRESPUESTA() As String

    getRESPUESTA = RESPUESTA

End Property

Public Property Let setRESPUESTA(ByVal valor As String)

    RESPUESTA = valor

End Property

Public Property Get getTIPO_PREGUNTA_RESPUESTA() As Integer

    getTIPO_PREGUNTA_RESPUESTA = TIPO_PREGUNTA_RESPUESTA

End Property

Public Property Let setTIPO_PREGUNTA_RESPUESTA(ByVal valor As Integer)

    TIPO_PREGUNTA_RESPUESTA = valor

End Property

Public Property Get getID_PREGUNTA_RESPUESTA() As Long

    getID_PREGUNTA_RESPUESTA = ID_PREGUNTA_RESPUESTA

End Property

Public Property Let setID_PREGUNTA_RESPUESTA(ByVal valor As Long)

    ID_PREGUNTA_RESPUESTA = valor

End Property

Private Sub Class_Initialize()
' elementos Por Defecto
REQUERIDA = False
TIPO_PREGUNTA_RESPUESTA = RESP_ALFANUMERICO
End Sub

