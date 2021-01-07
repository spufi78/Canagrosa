VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Calibracion_Historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_CALIBRACION_HISTORICO
'***********************************************************************
Private EQUIPO_ID As Long
Private fecha As String
Private MODALIDAD As String
Private PROCEDIMIENTO As String
Private OPERADOR_ID As Long
Private OPERADOR As String
Private LIMITACIONES_USO As String
Private RANGO_MIN As String
Private RANGO_MAX As String
Private UNIDADES As String
Private CONFORME As Long
Private RUTA_DOCUMENTO As String

'*************************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_CALIBRACION_HISTORICO
'*************************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setMODALIDAD(ByVal dato As String)
        MODALIDAD = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
        PROCEDIMIENTO = dato
End Property
Public Property Let setOPERADOR_ID(ByVal dato As Long)
        OPERADOR_ID = dato
End Property
Public Property Let setOPERADOR(ByVal dato As String)
        OPERADOR = dato
End Property
Public Property Let setLIMITACIONES_USO(ByVal dato As String)
        LIMITACIONES_USO = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As String)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As String)
        RANGO_MAX = dato
End Property
Public Property Let setUNIDADES(ByVal dato As String)
        UNIDADES = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
End Property
Public Property Let setRUTA_DOCUMENTO(ByVal dato As String)
        RUTA_DOCUMENTO = dato
End Property
'***********************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_CALIBRACION_HISTORICO
'***********************************************************************
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getMODALIDAD() As String
        getMODALIDAD = MODALIDAD
End Property
Public Property Get getPROCEDIMIENTO() As String
        getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Get getOPERADOR_ID() As Long
        getOPERADOR_ID = OPERADOR_ID
End Property
Public Property Get getOPERADOR() As String
        getOPERADOR = OPERADOR
End Property
Public Property Get getLIMITACIONES_USO() As String
        getLIMITACIONES_USO = LIMITACIONES_USO
End Property
Public Property Get getRANGO_MIN() As String
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As String
        getRANGO_MAX = RANGO_MAX
End Property
Public Property Get getUNIDADES() As String
        getUNIDADES = UNIDADES
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
End Property
Public Property Get getRUTA_DOCUMENTO() As String
        getRUTA_DOCUMENTO = RUTA_DOCUMENTO
End Property
'***************************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_CALIBRACION_HISTORICO
'***************************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_CALIBRACION_HISTORICO WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            EQUIPO_ID = rs("EQUIPO_ID")
            fecha = rs("fecha")
            MODALIDAD = rs("MODALIDAD")
            PROCEDIMIENTO = rs("PROCEDIMIENTO")
            OPERADOR_ID = rs("OPERADOR_ID")
            OPERADOR = rs("OPERADOR")
            LIMITACIONES_USO = rs("LIMITACIONES_USO")
            RANGO_MIN = rs("RANGO_MIN")
            RANGO_MAX = rs("RANGO_MAX")
            UNIDADES = rs("UNIDADES")
            CONFORME = rs("CONFORME")
            RUTA_DOCUMENTO = rs("RUTA_DOCUMENTO")
            
            Carga = True
        End If
        
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_calibracion_historico)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "INSERT INTO EQUIPOS_CALIBRACION_HISTORICO " & _
                   "VALUES (" & _
                        EQUIPO_ID & "," & _
                        "'" & Format(fecha, "yyyy-mm-dd hh:nn:ss") & "'" & "," & "'" & MODALIDAD & "'" & "," & _
                        "'" & PROCEDIMIENTO & "'" & ", " & OPERADOR_ID & ", " & "'" & OPERADOR & "'" & "," & _
                        "'" & LIMITACIONES_USO & "'" & ", " & _
                        "'" & RANGO_MIN & "'" & ", " & "'" & RANGO_MAX & "'" & ", " & _
                        "'" & UNIDADES & "'" & ", " & CONFORME & ", " & _
                        "'" & RUTA_DOCUMENTO & "'" & ", " & _
                        "'" & usuario.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_calibracion_historico)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long, strFecha As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_CALIBRACION_HISTORICO " & _
                   "   SET MODALIDAD = '" & MODALIDAD & "'" & _
                   "     , PROCEDIMIENTO = '" & PROCEDIMIENTO & "'" & _
                   "     , OPERADOR_ID = " & OPERADOR_ID & _
                   "     , OPERADOR = '" & OPERADOR & "'" & _
                   "     , LIMITACIONES_USO = '" & LIMITACIONES_USO & "'" & _
                   "     , RANGO_MIN = '" & RANGO_MIN & "'" & _
                   "     , RANGO_MAX = '" & RANGO_MAX & "'" & _
                   "     , UNIDADES = '" & UNIDADES & "'" & _
                   "     , CONFORME = " & CONFORME & _
                   "     , RUTA_DOCUMENTO = '" & RUTA_DOCUMENTO & "'" & _
                   "     , CUSERID = '" & usuario.getUSUARIO & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE EQUIPO_ID = " & ID & _
                   "   AND FECHA = '" & strFecha & "'"
        
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_calibracion_historico)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long, strFecha As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_CALIBRACION_HISTORICO " & _
                   " WHERE EQUIPO_ID = " & ID & _
                   "   AND FECHA = '" & strFecha & "'"
                   
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_calibracion_historico)", vbCritical, Err.Description
End Function

' listado de historico de calibraciones dela equipo pasado por parámetro
Public Function Listado(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ECH.EQUIPO_ID, ECH.FECHA AS FECHA_COMPLETA, ECH.RUTA_DOCUMENTO, ECH.FECHA, ECH.MODALIDAD, ECH.OPERADOR   " & _
                   "     , ECH.LIMITACIONES_USO, ECH.RANGO_MIN, ECH.RANGO_MAX, ECH.UNIDADES " & _
                   "     , IF(ECH.CONFORME='1','SI', 'NO') AS CONFORME " & _
                   "     , ECH.OPERADOR_ID " & _
                   "  FROM EQUIPOS E, EQUIPOS_CALIBRACION EC, EQUIPOS_CALIBRACION_HISTORICO ECH " & _
                   " WHERE E.ID_EQUIPO = EC.EQUIPO_ID " & _
                   "   AND EC.EQUIPO_ID = ECH.EQUIPO_ID " & _
                   "   AND E.ID_EQUIPO = " & ID & _
                   " ORDER BY FECHA DESC"
        Set Listado = datos_bd(consulta)
End Function
