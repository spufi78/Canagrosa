VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlantilla_determinaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***********************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE plantilla_determinaciones      *
'***********************************************************************
Private PLANTILLA_ID As Integer
Private TIPO_DETERMINACION_ID As Integer
Private ORDEN_MUESTRA As String
Private ORDEN As Integer
Private ESTADO As Integer
Private ES_ORIGINAL As Integer


Public Property Let setPLANTILLA_ID(ByVal dato As Integer)
    PLANTILLA_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Integer)
    TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setORDEN_MUESTRA(ByVal dato As String)
    ORDEN_MUESTRA = Trim(dato)
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = dato
End Property
Public Property Let setESTADO(ByVal dato As Integer)
    ESTADO = dato
End Property
Public Property Let setES_ORIGINAL(ByVal dato As Integer)
    ES_ORIGINAL = dato
End Property
Public Property Get getPLANTILLA_ID() As Integer
    getPLANTILLA_ID = PLANTILLA_ID
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Integer
    getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getORDEN_MUESTRA() As String
    getORDEN_MUESTRA = ORDEN_MUESTRA
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Property Get getESTADO() As Integer
    getESTADO = ESTADO
End Property
Public Property Get getES_ORIGINAL() As Integer
    getES_ORIGINAL = ES_ORIGINAL
End Property
