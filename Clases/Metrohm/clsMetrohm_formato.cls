VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMetrohm_formato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMETROHM_FORMATO
'***************************************************************
Private FORMATO As String
Private COL_MUESTRA As Long
Private COL_ENSAYO As Long
Private COL_FECHA As Long
Private COL_DETERMINACIONES As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSMETROHM_FORMATO
'***************************************************************
Public Enum LISTA_CLSMETROHM_FORMATO_COLS
        ENUM_FORMATO = 1
        ENUM_COL_MUESTRA = 2
        ENUM_COL_ENSAYO = 3
        ENUM_COL_FECHA = 4
        ENUM_COL_DETERMINACIONES = 5
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMETROHM_FORMATO
'***************************************************************
Public Property Let setFORMATO(ByVal dato As String)
        FORMATO = dato
End Property
Public Property Let setCOL_MUESTRA(ByVal dato As Long)
        COL_MUESTRA = dato
End Property
Public Property Let setCOL_ENSAYO(ByVal dato As Long)
        COL_ENSAYO = dato
End Property
Public Property Let setCOL_FECHA(ByVal dato As Long)
        COL_FECHA = dato
End Property
Public Property Let setCOL_DETERMINACIONES(ByVal dato As String)
        COL_DETERMINACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMETROHM_FORMATO
'***************************************************************
Public Property Get getFORMATO() As String
        getFORMATO = FORMATO
End Property
Public Property Get getCOL_MUESTRA() As Long
        getCOL_MUESTRA = COL_MUESTRA
End Property
Public Property Get getCOL_ENSAYO() As Long
        getCOL_ENSAYO = COL_ENSAYO
End Property
Public Property Get getCOL_FECHA() As Long
        getCOL_FECHA = COL_FECHA
End Property
Public Property Get getCOL_DETERMINACIONES() As String
        getCOL_DETERMINACIONES = COL_DETERMINACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMETROHM_FORMATO
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM METROHM_FORMATO WHERE FORMATO = '" & ID & "'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                FORMATO = rs("FORMATO")
                COL_MUESTRA = rs("COL_MUESTRA")
                COL_ENSAYO = rs("COL_ENSAYO")
                COL_FECHA = rs("COL_FECHA")
                COL_DETERMINACIONES = rs("COL_DETERMINACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMetrohm_formato)", vbCritical, Err.Description
End Function
Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO METROHM_FORMATO " & _
                   "  VALUES (" & _
                        "'" & FORMATO & "'" & "," & COL_MUESTRA & "," & COL_ENSAYO & "," & _
                        COL_FECHA & "," & "'" & COL_DETERMINACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = FORMATO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMetrohm_formato)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE METROHM_FORMATO SET " & _
                        " COL_MUESTRA = " & COL_MUESTRA & "," & _
                        " COL_ENSAYO = " & COL_ENSAYO & "," & _
                        " COL_FECHA = " & COL_FECHA & "," & _
                        " COL_DETERMINACIONES = '" & COL_DETERMINACIONES & "'" & _
                " WHERE FORMATO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMetrohm_formato)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM METROHM_FORMATO " & _
                "    WHERE FORMATO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMetrohm_formato)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM METROHM_FORMATO ORDER BY FORMATO"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMETROHM_FORMATO
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

