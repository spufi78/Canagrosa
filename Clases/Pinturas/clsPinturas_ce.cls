VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPinturas_ce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPINTURAS_CE
'***************************************************************
Private PINTURA_ID As Long
Private TIPO_ENSAYO_ID As Long
Private ORDEN As Long

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSPINTURAS_CE
'***************************************************************
Public Enum LISTA_CLSPINTURAS_CE_COLS
        ENUM_PINTURA_ID = 1
        ENUM_TIPO_ENSAYO_ID = 2
        ENUM_ORDEN = 3
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPINTURAS_CE
'***************************************************************
Public Property Let setPINTURA_ID(ByVal dato As Long)
        PINTURA_ID = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPINTURAS_CE
'***************************************************************
Public Property Get getPINTURA_ID() As Long
        getPINTURA_ID = PINTURA_ID
End Property
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPINTURAS_CE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM PINTURAS_CE WHERE PINTURA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PINTURA_ID = rs("PINTURA_ID")
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPinturas_ce)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PINTURAS_CE " & _
                   "  VALUES (" & _
                        PINTURA_ID & "," & TIPO_ENSAYO_ID & "," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = PINTURA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPinturas_ce)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PINTURAS_CE SET " & _
                        " TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE PINTURA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPinturas_ce)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PINTURAS_CE " & _
                "    WHERE PINTURA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPinturas_ce)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.ORDEN,B.NOMBRE,C.NOMBRE,B.ID_TIPO_ENSAYO,B.TIPO_ANALISIS_ID,B.INCLUYE_ESPESOR,B.LOTE_PROBETAS,B.NORMA " & _
                   "  FROM PINTURAS_CE A, CE_TIPOS_ENSAYOS B " & _
                   " LEFT JOIN PROCESOS_BASE C ON B.PROCESO_BASE_ID = C.ID_PROCESO_BASE " & _
                   " WHERE A.PINTURA_ID = " & ID & _
                   "   AND A.TIPO_ENSAYO_ID = B.ID_TIPO_ENSAYO " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPINTURAS_CE
'***************************************************************
Public Function ListadoHerencia() As ADODB.RecordSet
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

