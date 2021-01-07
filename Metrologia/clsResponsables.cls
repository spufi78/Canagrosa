VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsResponsables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSRESPONSABLES
'***************************************************************
Private ID_RESPONSABLE As Long
Private PLANTA_ID As String
Private NOMBRE As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSRESPONSABLES
'***************************************************************
Public Enum LISTA_CLSRESPONSABLES_COLS
        ENUM_ID_RESPONSABLE = 1
        ENUM_PLANTA_ID = 2
        ENUM_NOMBRE = 3
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSRESPONSABLES
'***************************************************************
Public Property Let setID_RESPONSABLE(ByVal dato As Long)
        ID_RESPONSABLE = dato
End Property
Public Property Let setPLANTA_ID(ByVal dato As String)
        PLANTA_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSRESPONSABLES
'***************************************************************
Public Property Get getID_RESPONSABLE() As Long
        getID_RESPONSABLE = ID_RESPONSABLE
End Property
Public Property Get getPLANTA_ID() As String
        getPLANTA_ID = PLANTA_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSRESPONSABLES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM responsables WHERE ID_RESPONSABLE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_RESPONSABLE = rs("ID_RESPONSABLE")
                PLANTA_ID = rs("PLANTA_ID")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsResponsables)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_RESPONSABLE) FROM responsables"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_RESPONSABLE = 1
        Else
                ID_RESPONSABLE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO responsables " & _
                   "  VALUES (" & _
                        ID_RESPONSABLE & "," & "'" & PLANTA_ID & "'" & "," & "'" & NOMBRE & "',''" & _
                ")"
        execute_bd consulta
        Insertar = ID_RESPONSABLE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsResponsables)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE responsables SET " & _
                        " PLANTA_ID = '" & PLANTA_ID & "'," & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_RESPONSABLE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsResponsables)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM responsables " & _
                "    WHERE ID_RESPONSABLE = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsResponsables)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM responsables ORDER BY ID_RESPONSABLE"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSRESPONSABLES
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_RESPONSABLE,PLANTA_ID FROM responsables ORDER BY PLANTA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
