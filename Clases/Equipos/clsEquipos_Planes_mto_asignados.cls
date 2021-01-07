VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Planes_mto_asignados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_PLANES_MTO_ASIGNADOS
'***************************************************************
Private PLAN_MTO_ID As Long
Private EQUIPO_ID As Long
Private OPERADOR_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_PLANES_MTO_ASIGNADOS
'***************************************************************
Public Property Let setPLAN_MTO_ID(ByVal dato As Long)
        PLAN_MTO_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setOPERADOR_ID(ByVal dato As Long)
        OPERADOR_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_PLANES_MTO_ASIGNADOS
'***************************************************************
Public Property Get getPLAN_MTO_ID() As Long
        getPLAN_MTO_ID = PLAN_MTO_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getOPERADOR_ID() As Long
        getOPERADOR_ID = OPERADOR_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_PLANES_MTO_ASIGNADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_PLANES_MTO_ASIGNADOS WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            PLAN_MTO_ID = rs("PLAN_MTO_ID")
            EQUIPO_ID = rs("EQUIPO_ID")
            OPERADOR_ID = rs("OPERADOR_ID")
            Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_Planes_mto_asignados)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        If Not duplicado Then
            consulta = "INSERT INTO EQUIPOS_PLANES_MTO_ASIGNADOS VALUES (" & _
                                    EQUIPO_ID & ", " & PLAN_MTO_ID & ", " & OPERADOR_ID & ", " & _
                                    "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                       ")"
            execute_bd consulta
            Insertar = PLAN_MTO_ID
        End If
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_Planes_mto_asignados)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID_P As Long, ID_E As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_PLANES_MTO_ASIGNADOS " & _
                   " WHERE PLAN_MTO_ID = " & ID_P & _
                   "   AND EQUIPO_ID = " & ID_E
        execute_bd consulta
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_Planes_mto_asignados)", vbCritical, Err.Description
End Function

Public Function Eliminar_planes(ID_E As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_PLANES_MTO_ASIGNADOS " & _
                   " WHERE EQUIPO_ID = " & ID_E
        execute_bd consulta
        Eliminar_planes = True
        
        Exit Function
fallo:
        Eliminar_planes = False
        MsgBox "Error al Eliminar_planes (clsEquipos_Planes_mto_asignados)", vbCritical, Err.Description
End Function

Public Function lista_por_equipos(ID As Long) As ADODB.RecordSet
        Dim consulta As String

'E0306-I
        'consulta = "SELECT PLANES.ID_PLAN_MTO, PLANES.NOMBRE, PLANES.DESCRIPCION, FREQ.DESCRIPCION, CONCAT(USUARIOS.NOMBRE,' ', USUARIOS.APELLIDOS) EMPLEADO,USUARIOS.ID_EMPLEADO, PLANES.ORDEN " & _
                   "  FROM EQUIPOS_PLANES_MTO_ASIGNADOS ASIG, eq_planes_mantenimiento PLANES, DECODIFICADORA FREQ, USUARIOS " & _
                   " WHERE ASIG.PLAN_MTO_ID = PLANES.ID_PLAN_MTO " & _
                   "   AND FREQ.CODIGO = " & decodificadora.EQ_periodicidad & _
                   "   AND PLANES.FRECUENCIA_ID = FREQ.VALOR " & _
                   "   AND ASIG.OPERADOR_ID = USUARIOS.ID_EMPLEADO " & _
                   "   AND ASIG.EQUIPO_ID = " & ID & _
                   " ORDER BY ORDEN "
                   
        consulta = "SELECT PLANES.ID_PLAN_MTO, PLANES.NOMBRE, FREQ.DESCRIPCION, CONCAT(USUARIOS.NOMBRE,' ', USUARIOS.APELLIDOS) EMPLEADO,USUARIOS.ID_EMPLEADO, PLANES.ORDEN " & _
                   "  FROM EQUIPOS_PLANES_MTO_ASIGNADOS ASIG, eq_planes_mantenimiento PLANES, DECODIFICADORA FREQ, USUARIOS " & _
                   " WHERE ASIG.PLAN_MTO_ID = PLANES.ID_PLAN_MTO " & _
                   "   AND FREQ.CODIGO = " & decodificadora.EQ_periodicidad & _
                   "   AND PLANES.FRECUENCIA_ID = FREQ.VALOR " & _
                   "   AND ASIG.OPERADOR_ID = USUARIOS.ID_EMPLEADO " & _
                   "   AND ASIG.EQUIPO_ID = " & ID & _
                   " ORDER BY ORDEN "
'E0306-F
        Set lista_por_equipos = datos_bd(consulta)
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT COUNT(*) AS TOTAL " & _
               "  FROM EQUIPOS_PLANES_MTO_ASIGNADOS " & _
               " WHERE PLAN_MTO_ID = " & PLAN_MTO_ID & _
               "   AND EQUIPO_ID = " & EQUIPO_ID
    Set rs = datos_bd(consulta)
    If rs.Fields("TOTAL") = 0 Then
        duplicado = False
    Else
        MsgBox "Este plan de mantenimiento ya está asociada al equipo.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function

Public Function Modificar(EQUIPO_ID As Long, PLAN_ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE EQUIPOS_PLANES_MTO_ASIGNADOS " & _
               "   SET OPERADOR_ID = " & OPERADOR_ID & ", " & _
               "       CUSERID = '" & USUARIO.getUSUARIO & "', " & _
               "       FTIMESTP = CURRENT_TIMESTAMP " & _
               " WHERE EQUIPO_ID = " & EQUIPO_ID & _
               "   AND PLAN_MTO_ID = " & PLAN_ID
    execute_bd consulta
    
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar (clsEquipos_Planes_mto_asignados)", vbCritical, Err.Description
End Function

' Función que devuelve un recordset con los planes de mto asignados
' al equipo pasado por parámetro
Public Function planes_mto_asignados(ID As Long) As ADODB.RecordSet
'        Dim consulta As String
'
'        consulta = "SELECT ASIG.ID_PLAN_MTO " & _
'                   "  FROM EQUIPOS_PLANES_MTO_ASIGNADOS ASIG " & _
'                   " WHERE "
'
'        Set planes_mto_asignados = datos_bd(consulta)
End Function

