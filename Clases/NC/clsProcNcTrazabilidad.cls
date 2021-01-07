VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcNcTrazabilidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID_TRAZABILIDAD As Long
Private ID_PROCNC As Long
Private ID_ACCION As Long
Private ID_USUARIO As Long
Private DESCRIPCION As String
Private ACCION As String
Private USUARIO As String
Private FECHAHORA As Date
Private ID_AUX As Integer

Public Property Let setID_AUX(ByVal valor As Integer)

    ID_AUX = valor

End Property

Public Property Get getID_AUX() As Integer

    getID_AUX = ID_AUX

End Property

Public Property Let setID_TRAZABILIDAD(ByVal valor As Long)

    ID_TRAZABILIDAD = valor

End Property


Public Property Get getID_TRAZABILIDAD() As Long

    getID_TRAZABILIDAD = ID_TRAZABILIDAD

End Property

Public Property Get getID_PROCNC() As Long

    getD_PROCNC = ID_PROCNC

End Property

Public Property Let setID_PROCNC(ByVal valor As Long)

    ID_PROCNC = valor

End Property

Public Property Get getID_ACCION() As Long

    getID_ACCION = ID_ACCION

End Property

Public Property Let setID_ACCION(ByVal valor As Long)

    ID_ACCION = valor

End Property

Public Property Get getID_USUARIO() As Long

    getID_USUARIO = ID_USUARIO

End Property

Public Property Let setID_USUARIO(ByVal valor As Long)

    ID_USUARIO = valor

End Property

Public Property Get getACCION() As String

    getACCION = ACCION

End Property

Public Property Let setACCION(ByVal valor As String)

    ACCION = valor

End Property

Public Property Get getUSUARIO() As String

    getUSUARIO = USUARIO

End Property

Public Property Let setUSUARIO(ByVal valor As String)

    USUARIO = valor

End Property

Public Property Get getDESCRIPCION() As String

    getDESCRIPCION = DESCRIPCION

End Property

Public Property Let setDESCRIPCION(ByVal valor As String)

    DESCRIPCION = valor

End Property

Public Property Get getFECHAHORA() As Date

    getFECHAHORA = FECHAHORA

End Property

Public Property Let setFECHAHORA(ByVal valor As Date)

    FECHAHORA = valor

End Property



Public Function Insertar() As Boolean
Dim strSQL As String

On Error GoTo Insertar_Error

    strSQL = "INSERT INTO procnc_trazabilidad "
    strSQL = strSQL & "(ID_TRAZABILIDAD, ID_PROCNC, ID_ACCION, ID_USUARIO, DESCRIPCION, FECHAHORA)"
    strSQL = strSQL & " SELECT coalesce(MAX(ID_TRAZABILIDAD),0)+1 AS ID_TRAZABILIDAD"
    strSQL = strSQL & ", " & ID_PROCNC & " AS ID_PROCNC"
    strSQL = strSQL & ", " & ID_ACCION & " AS ID_ACCION"
    strSQL = strSQL & ", " & ID_USUARIO & " AS ID_USUARIO"
    strSQL = strSQL & ", '" & DESCRIPCION & "' AS DESCRIPCION"
    strSQL = strSQL & ", '" & Format(FECHAHORA, "yyyy-mm-dd Hh:Nn:Ss") & "' AS FECHAHORA"
    strSQL = strSQL & " FROM procnc_trazabilidad"

    execute_bd strSQL

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcTrazabilidad.Insertar"
    Exit Function
Insertar_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNcTrazabilidad.Insertar"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Insertar of Módulo de clase clsProcNcTrazabilidad" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function
