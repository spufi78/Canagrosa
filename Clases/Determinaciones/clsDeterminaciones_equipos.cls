VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeterminaciones_equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDETERMINACIONES_EQUIPOS
'***************************************************************
Private DETERMINACION_ID As Long
Private EQUIPO_ID As Long
Private ORDEN As Long
'M1385-I
Private EN_INFORME As Integer
'M1385-F
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDETERMINACIONES_EQUIPOS
'***************************************************************
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
        DETERMINACION_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'M1385-I
Public Property Let setEN_INFORME(ByVal dato As Integer)
        EN_INFORME = dato
End Property
'M1385-f
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDETERMINACIONES_EQUIPOS
'***************************************************************
Public Property Get getDETERMINACION_ID() As Long
        getDETERMINACION_ID = DETERMINACION_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'M1385-I
Public Property Get getEN_INFORME() As Long
        getEN_INFORME = EN_INFORME
End Property
'M1385-F
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDETERMINACIONES_EQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DETERMINACIONES_EQUIPOS WHERE DETERMINACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DETERMINACION_ID = rs("DETERMINACION_ID")
                EQUIPO_ID = rs("EQUIPO_ID")
                ORDEN = rs("ORDEN")
                'M1385-I
                EN_INFORME = rs("EN_INFORME")
                'M1385-F
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function
Public Function Insertar(lDETERMINACION As Long, lTIPO_DETERMINACION As Long) As Long
        On Error GoTo fallo
        Dim consulta As String
        'M1137-I : INSERTAR SOLO LOS EQUIPOS ACTIVOS
        'consulta = " INSERT INTO DETERMINACIONES_EQUIPOS " & _
        '           " SELECT " & lDETERMINACION & ",EQUIPO_ID,ORDEN " & _
        '           "   FROM TIPOS_DETERMINACION_EQUIPOS " & _
        '           "  WHERE TIPO_DETERMINACION_ID = " & lTIPO_DETERMINACION
        consulta = " INSERT INTO DETERMINACIONES_EQUIPOS " & _
                   " SELECT " & lDETERMINACION & ",EQUIPO_ID,ORDEN,0 " & _
                   "   FROM TIPOS_DETERMINACION_EQUIPOS A, EQUIPOS B" & _
                   "  WHERE TIPO_DETERMINACION_ID = " & lTIPO_DETERMINACION & _
                   "    AND A.EQUIPO_ID = B.ID_EQUIPO " & _
                   "    AND B.ESTADO_ID IN ('" & ENUM_EQ_ESTADOS_ACTIVO & "','" & ENUM_EQ_ESTADOS_NRC & "')"
        'M1137-F
        execute_bd consulta
        Insertar = lDETERMINACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function
Public Function Insertar_Determinacion() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = " INSERT INTO DETERMINACIONES_EQUIPOS VALUES (" & _
                     DETERMINACION_ID & "," & EQUIPO_ID & "," & ORDEN & "," & EN_INFORME & ")"
        execute_bd consulta
        Insertar_Determinacion = DETERMINACION_ID
        Exit Function
fallo:
        Insertar_Determinacion = 0
        MsgBox "Error al Insertar_Determinacion (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DETERMINACIONES_EQUIPOS SET " & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " ORDEN = " & ORDEN & "," & _
                        " EN_INFORME = " & EN_INFORME & "" & _
                " WHERE DETERMINACION_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DETERMINACIONES_EQUIPOS " & _
                "    WHERE DETERMINACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_EQUIPO,B.NOMBRE,B.SERIE, A.EN_INFORME " & _
                   " FROM DETERMINACIONES_EQUIPOS A,EQUIPOS B " & _
                   " WHERE A.EQUIPO_ID = B.ID_EQUIPO " & _
                   "   AND A.DETERMINACION_ID = " & ID & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function

Public Function duplicar(ID_DETERMINACION_ORIGEN As Long, ID_DETERMINACION_DESTINO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO DETERMINACIONES_EQUIPOS " & _
                   "SELECT " & ID_DETERMINACION_DESTINO & ",EQUIPO_ID,ORDEN,EN_INFORME FROM DETERMINACIONES_EQUIPOS WHERE DETERMINACION_ID = " & ID_DETERMINACION_ORIGEN
        execute_bd consulta
        duplicar = True
        Exit Function
fallo:
        duplicar = False
        MsgBox "Error al duplicar (clsDeterminaciones_equipos)", vbCritical, Err.Description
End Function

