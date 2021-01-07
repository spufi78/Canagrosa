VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcNcAdjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsProcNcAdjuntos
'***************************************************************
Private PROCNC_ID As Long
Private ORDEN As Long
Private ruta As String
Private RUTA_TEMPORAL As String
Private OBSERVACIONES As String
Private fecha As String
Private EMPLEADO_ID As Integer
Private EMPLEADO_NOMBRE_APELLIDOS As String
Private ID_AUX As Integer
Private ID_ADJUNTO As Integer


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsProcNcAdjuntos
'***************************************************************
Public Property Let setPROCNC_ID(ByVal dato As Long)
        PROCNC_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        ruta = dato
End Property
Public Property Let setRUTA_TEMPORAL(ByVal dato As String)
        RUTA_TEMPORAL = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
        EMPLEADO_ID = dato
End Property
Public Property Let setEMPLEADO_NOMBRE_APELLIDOS(ByVal dato As String)
        EMPLEADO_NOMBRE_APELLIDOS = dato
End Property
Public Property Let setID_ADJUNTO(ByVal dato As Integer)
        ID_ADJUNTO = dato
End Property
Public Property Let setID_AUX(ByVal dato As Integer)
        ID_AUX = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsProcNcAdjuntos
'***************************************************************
Public Property Get getPROCNC_ID() As Long
        getPROCNC_ID = PROCNC_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getRUTA() As String
        getRUTA = ruta
End Property
Public Property Get getNOMBRE_ARCHIVO() As String
    Dim strArr() As String
        
    getNOMBRE_ARCHIVO = ""
    
    If ruta <> "" Then
        strArr = Split(ruta, "\")
        getNOMBRE_ARCHIVO = strArr(UBound(strArr))
    End If
        
End Property

Public Property Get getRUTA_TEMPORAL() As String
        getRUTA_TEMPORAL = RUTA_TEMPORAL
End Property
Public Property Get getNOMBRE_ARCHIVO_TEMP() As String
    Dim strArr() As String
        
    getNOMBRE_ARCHIVO_TEMP = ""
        
    If RUTA_TEMPORAL <> "" Then
        ' lA RUTA QUE VALE ES la temporal
        strArr = Split(RUTA_TEMPORAL, "\")
        getNOMBRE_ARCHIVO_TEMP = strArr(UBound(strArr))
    ElseIf ruta <> "" Then
        strArr = Split(ruta, "\")
        getNOMBRE_ARCHIVO_TEMP = strArr(UBound(strArr))
    End If
    
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getEMPLEADO_ID() As Integer
        getEMPLEADO_ID = EMPLEADO_ID
End Property

Public Property Get getEMPLEADO_NOMBRE_APELLIDOS() As String
        getEMPLEADO_NOMBRE_APELLIDOS = EMPLEADO_NOMBRE_APELLIDOS
End Property
Public Property Get getID_ADJUNTO() As Integer
        getID_ADJUNTO = ID_ADJUNTO
End Property
Public Property Get getID_AUX() As Integer
        getID_AUX = ID_AUX
End Property
