VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_calibracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_CALIBRACION
'*************************************************************
Private EQUIPO_ID As Long
Private ID_CALIBRACION As Long
Private MODALIDAD_ID As Long
Private PERIODICIDAD_ID As Long
Private PROCEDIMIENTO_ID As Long
Private CALIBRADOR_EXTERNO_ID As Long
Private CALIBRADOR_INTERNO_ID As Long
Private fecha_actual As String
Private fecha_proxima As String
Private RANGO_MIN As String
Private RANGO_MAX As String
Private UNIDADES_ID As Long
Private RUTA_PLANTILLA As String
Private EFECTIVA As Long
Private ID_AUX As Integer
Public Property Let setID_AUX(ByVal dato As Integer)
    ID_AUX = dato
End Property
Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property
Public Property Let setID_CALIBRACION(ByVal dato As Integer)
    ID_CALIBRACION = dato
End Property
Public Property Get getID_CALIBRACION() As Integer
    getID_CALIBRACION = ID_CALIBRACION
End Property

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_CALIBRACION
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setMODALIDAD_ID(ByVal dato As Long)
        MODALIDAD_ID = dato
End Property
Public Property Let setPERIODICIDAD_ID(ByVal dato As Long)
        PERIODICIDAD_ID = dato
End Property
Public Property Let setPROCEDIMIENTO_ID(ByVal dato As Long)
        PROCEDIMIENTO_ID = dato
End Property
Public Property Let setCALIBRADOR_EXTERNO_ID(ByVal dato As Long)
        CALIBRADOR_EXTERNO_ID = dato
End Property
Public Property Let setCALIBRADOR_INTERNO_ID(ByVal dato As Long)
        CALIBRADOR_INTERNO_ID = dato
End Property
Public Property Let setFECHA_ACTUAL(ByVal dato As String)
        fecha_actual = dato
End Property
Public Property Let setFECHA_PROXIMA(ByVal dato As String)
        fecha_proxima = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As String)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As String)
        RANGO_MAX = dato
End Property
Public Property Let setUNIDADES_ID(ByVal dato As Long)
        UNIDADES_ID = dato
End Property
Public Property Let setRUTA_PLANTILLA(ByVal dato As String)
        RUTA_PLANTILLA = dato
End Property
Public Property Let setEFECTIVA(ByVal dato As Long)
        EFECTIVA = dato
End Property
'*************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_CALIBRACION
'*************************************************************
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getMODALIDAD_ID() As Long
        getMODALIDAD_ID = MODALIDAD_ID
End Property
Public Property Get getPERIODICIDAD_ID() As Long
        getPERIODICIDAD_ID = PERIODICIDAD_ID
End Property
Public Property Get getPROCEDIMIENTO_ID() As Long
        getPROCEDIMIENTO_ID = PROCEDIMIENTO_ID
End Property
Public Property Get getCALIBRADOR_EXTERNO_ID() As Long
        getCALIBRADOR_EXTERNO_ID = CALIBRADOR_EXTERNO_ID
End Property
Public Property Get getCALIBRADOR_INTERNO_ID() As Long
        getCALIBRADOR_INTERNO_ID = CALIBRADOR_INTERNO_ID
End Property
Public Property Get getFECHA_ACTUAL() As String
        getFECHA_ACTUAL = fecha_actual
End Property
Public Property Get getFECHA_PROXIMA() As String
        getFECHA_PROXIMA = fecha_proxima
End Property
Public Property Get getRANGO_MIN() As String
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As String
        getRANGO_MAX = RANGO_MAX
End Property
Public Property Get getUNIDADES_ID() As Long
        getUNIDADES_ID = UNIDADES_ID
End Property
Public Property Get getRUTA_PLANTILLA() As String
        getRUTA_PLANTILLA = RUTA_PLANTILLA
End Property
Public Property Get getEFECTIVA() As Long
        getEFECTIVA = EFECTIVA
End Property
'*****************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_CALIBRACION
'*****************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        Dim CONT As Long
        
        consulta = "SELECT * FROM EQUIPOS_CALIBRACION WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        CONT = 1
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            ID_CALIBRACION = CONT
            CONT = CONT + 1
            EQUIPO_ID = rs("EQUIPO_ID")
            MODALIDAD_ID = rs("MODALIDAD_ID")
            PERIODICIDAD_ID = rs("PERIODICIDAD_ID")
            PROCEDIMIENTO_ID = rs("PROCEDIMIENTO_ID")
            CALIBRADOR_EXTERNO_ID = rs("CALIBRADOR_EXTERNO_ID")
            CALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
            fecha_actual = Format(rs("FECHA_ACTUAL"), "yyyy-mm-dd")
            fecha_proxima = Format(rs("FECHA_PROXIMA"), "yyyy-mm-dd")
            RANGO_MIN = rs("RANGO_MIN")
            RANGO_MAX = rs("RANGO_MAX")
            UNIDADES_ID = rs("UNIDADES_ID")
            RUTA_PLANTILLA = rs("RUTA_PLANTILLA")
            EFECTIVA = rs("EFECTIVA")
            
            Carga = True
        End If
        
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_calibracion)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Me.Eliminar EQUIPO_ID
        
        consulta = "INSERT INTO EQUIPOS_CALIBRACION " & _
                   "VALUES (" & _
                        EQUIPO_ID & "," & _
                        MODALIDAD_ID & "," & PERIODICIDAD_ID & "," & _
                        PROCEDIMIENTO_ID & "," & CALIBRADOR_EXTERNO_ID & ", " & _
                        CALIBRADOR_INTERNO_ID & ", " & "'" & Format(fecha_actual, "yyyy-mm-dd") & "'" & ", " & _
                        "'" & Format(fecha_proxima, "yyyy-mm-dd") & "'" & ", " & "'" & RANGO_MIN & "'" & ", " & _
                        "'" & RANGO_MAX & "'" & ", " & UNIDADES_ID & ", " & _
                        "'" & RUTA_PLANTILLA & "'" & ", " & EFECTIVA & ", " & _
                        "'" & usuario.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_calibracion)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_CALIBRACION " & _
                   "   SET MODALIDAD_ID = " & MODALIDAD_ID & _
                   "     , PERIODICIDAD_ID = " & PERIODICIDAD_ID & _
                   "     , PROCEDIMIENTO_ID = " & PROCEDIMIENTO_ID & _
                   "     , CALIBRADOR_EXTERNO_ID = " & CALIBRADOR_EXTERNO_ID & _
                   "     , CALIBRADOR_INTERNO_ID = " & CALIBRADOR_INTERNO_ID & _
                   "     , FECHA_ACTUAL = '" & Format(fecha_actual, "yyyy-mm-dd") & "'" & _
                   "     , FECHA_PROXIMA = '" & Format(fecha_proxima, "yyyy-mm-dd") & "'" & _
                   "     , RANGO_MIN = '" & RANGO_MIN & "'" & _
                   "     , RANGO_MAX = '" & RANGO_MAX & "'" & _
                   "     , UNIDADES_ID = " & UNIDADES_ID & _
                   "     , RUTA_PLANTILLA = '" & RUTA_PLANTILLA & "'" & _
                   "     , EFECTIVA = " & EFECTIVA & _
                   "     , CUSERID = '" & usuario.getUSUARIO & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE EQUIPO_ID = " & ID
        
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_calibracion)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_CALIBRACION " & _
                   " WHERE EQUIPO_ID = " & ID
                   
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_calibracion)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * " & _
                   "  FROM EQUIPOS_CALIBRACION " & _
                   " WHERE EFECTIVA = 1 " & _
                   " ORDER BY EQUIPO_ID"
        
        Set Listado = datos_bd(consulta)
End Function

Public Function total_calibraciones(lngEquipo As Long) As Long
        Dim rs As New ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT COUNT(*) AS TOTAL " & _
                   "  FROM EQUIPOS_CALIBRACION " & _
                   " WHERE EQUIPO_ID = " & lngEquipo
        Set rs = datos_bd(consulta)
        total_calibraciones = rs("TOTAL")
        Set rs = Nothing
End Function

' función que marca como efectiva la calibración del equipo pasado por parámetro
Public Function marcar_como_efectiva(lngEquipo As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE EQUIPOS_CALIBRACION " & _
                   "   SET EFECTIVA = 1 " & _
                   "     , CUSERID = '" & usuario.getUSUARIO & "'" & _
                   "     , FTIMESTP = CURRENT_TIMESTAMP " & _
                   " WHERE EQUIPO_ID = " & lngEquipo
        
        execute_bd consulta
        marcar_como_efectiva = True
        Exit Function
fallo:
        marcar_como_efectiva = False
        MsgBox "Error al marcar como efectiva (clsEquipos_calibracion)", vbCritical, Err.Description
End Function

Public Function Eliminar_no_efectiva(lngEquipo As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_CALIBRACION " & _
                   " WHERE EFECTIVA = 0 " & _
                   "   AND EQUIPO_ID = " & lngEquipo
                   
        execute_bd consulta
        Eliminar_no_efectiva = True
        Exit Function
fallo:
        Eliminar_no_efectiva = False
        MsgBox "Error al Eliminar no efectiva (clsEquipos_calibracion)", vbCritical, Err.Description
End Function

