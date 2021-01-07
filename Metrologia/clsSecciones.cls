VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSecciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSECCIONES
'***************************************************************
Private ID_SECCION As Long
Private PLANTA_ID As String
Private NOMBRE As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSSECCIONES
'***************************************************************
Public Enum LISTA_CLSSECCIONES_COLS
        ENUM_ID_SECCION = 1
        ENUM_PLANTA_ID = 2
        ENUM_NOMBRE = 3
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSECCIONES
'***************************************************************
Public Property Let setID_SECCION(ByVal dato As Long)
        ID_SECCION = dato
End Property
Public Property Let setPLANTA_ID(ByVal dato As String)
        PLANTA_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSECCIONES
'***************************************************************
Public Property Get getID_SECCION() As Long
        getID_SECCION = ID_SECCION
End Property
Public Property Get getPLANTA_ID() As String
        getPLANTA_ID = PLANTA_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSECCIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM secciones WHERE ID_SECCION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SECCION = rs("ID_SECCION")
                PLANTA_ID = rs("PLANTA_ID")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSecciones)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_SECCION) FROM secciones"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SECCION = 1
        Else
                ID_SECCION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO secciones " & _
                   "  VALUES (" & _
                        ID_SECCION & "," & "'" & PLANTA_ID & "'" & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_SECCION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSecciones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE secciones SET " & _
                        " PLANTA_ID = '" & PLANTA_ID & "'," & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_SECCION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSecciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM secciones " & _
                "    WHERE ID_SECCION = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSecciones)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM secciones ORDER BY ID_SECCION"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSSECCIONES
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_SECCION,PLANTA_ID FROM secciones ORDER BY PLANTA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
