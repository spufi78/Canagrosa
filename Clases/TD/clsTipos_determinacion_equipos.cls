VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_determinacion_equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_DETERMINACION_EQUIPOS
'***************************************************************
Private TIPO_DETERMINACION_ID As Long
Private EQUIPO_ID As Long
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_DETERMINACION_EQUIPOS
'***************************************************************
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_DETERMINACION_EQUIPOS
'***************************************************************
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_DETERMINACION_EQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DETERMINACION_EQUIPOS WHERE TIPO_DETERMINACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                EQUIPO_ID = rs("EQUIPO_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_determinacion_equipos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TIPOS_DETERMINACION_EQUIPOS " & _
                   "  VALUES (" & _
                        TIPO_DETERMINACION_ID & "," & EQUIPO_ID & "," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = TIPO_DETERMINACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_determinacion_equipos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_DETERMINACION_EQUIPOS SET " & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE TIPO_DETERMINACION_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTipos_determinacion_equipos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_DETERMINACION_EQUIPOS " & _
                "    WHERE TIPO_DETERMINACION_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_determinacion_equipos)", vbCritical, Err.Description
End Function
Public Function Listado(TD As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_EQUIPO,B.NOMBRE,B.SERIE " & _
                   " FROM TIPOS_DETERMINACION_EQUIPOS A,EQUIPOS B " & _
                   " WHERE A.EQUIPO_ID = B.ID_EQUIPO " & _
                   "   AND A.TIPO_DETERMINACION_ID = " & TD & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
'M1137-I
Public Function ListadoActivos(TD As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_EQUIPO,B.NOMBRE,B.SERIE " & _
                   " FROM TIPOS_DETERMINACION_EQUIPOS A,EQUIPOS B " & _
                   " WHERE A.EQUIPO_ID = B.ID_EQUIPO " & _
                   "   AND A.TIPO_DETERMINACION_ID = " & TD & _
                   "   AND B.ESTADO_ID IN ('" & ENUM_EQ_ESTADOS_ACTIVO & "','" & ENUM_EQ_ESTADOS_NRC & "')" & _
                   " ORDER BY A.ORDEN"
        Set ListadoActivos = datos_bd(consulta)
End Function
'M1137-F


