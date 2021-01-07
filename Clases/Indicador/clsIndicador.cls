VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSINDICADOR
'***************************************************************
Private ID_INDICADOR As Long
Private DEPARTAMENTO_ID As Long
Private APARTADO_ID As Long
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSINDICADOR
'***************************************************************
Public Property Let setID_INDICADOR(ByVal dato As Long)
        ID_INDICADOR = dato
End Property
Public Property Let setDEPARTAMENTO_ID(ByVal dato As Long)
        DEPARTAMENTO_ID = dato
End Property
Public Property Let setAPARTADO_ID(ByVal dato As Long)
        APARTADO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSINDICADOR
'***************************************************************
Public Property Get getID_INDICADOR() As Long
        getID_INDICADOR = ID_INDICADOR
End Property
Public Property Get getDEPARTAMENTO_ID() As Long
        getDEPARTAMENTO_ID = DEPARTAMENTO_ID
End Property
Public Property Get getAPARTADO_ID() As Long
        getAPARTADO_ID = APARTADO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSINDICADOR
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM INDICADOR WHERE ID_INDICADOR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_INDICADOR = rs("ID_INDICADOR")
                DEPARTAMENTO_ID = rs("DEPARTAMENTO_ID")
                APARTADO_ID = rs("APARTADO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsIndicador)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_INDICADOR) FROM INDICADOR"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_INDICADOR = 1
        Else
                ID_INDICADOR = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO INDICADOR " & _
                   "  VALUES (" & _
                        ID_INDICADOR & "," & DEPARTAMENTO_ID & "," & APARTADO_ID & "," & _
                        "'" & DESCRIPCION & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_INDICADOR
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsIndicador)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE INDICADOR SET " & _
                        " DEPARTAMENTO_ID = " & DEPARTAMENTO_ID & "," & _
                        " APARTADO_ID = " & APARTADO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID_INDICADOR = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsIndicador)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM INDICADOR " & _
                "    WHERE ID_INDICADOR = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsIndicador)", vbCritical, Err.Description
End Function
Public Function Listado(departamento As Long, apartado As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If departamento <> 0 Then
            filtro = filtro & " AND I.DEPARTAMENTO_ID = " & departamento
        End If
        If apartado <> 0 Then
            filtro = filtro & " AND I.APARTADO_ID = " & apartado
        End If
        consulta = "SELECT DISTINCT I.ID_INDICADOR,DEP.DESCRIPCION,APA.DESCRIPCION,I.DESCRIPCION " & _
                   "  FROM INDICADOR I " & _
                   "  LEFT JOIN decodificadora DEP ON I.DEPARTAMENTO_ID = DEP.VALOR AND DEP.CODIGO = " & DECODIFICADORA.INDICADOR_DEPARTAMENTOS & _
                   "  LEFT JOIN decodificadora APA ON I.APARTADO_ID = APA.VALOR AND APA.CODIGO = " & DECODIFICADORA.INDICADOR_APARTADOS & _
                   " WHERE 1 = 1 " & _
                   filtro & _
                   " ORDER BY I.DEPARTAMENTO_ID ASC, I.APARTADO_ID ASC, I.DESCRIPCION ASC"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSINDICADOR
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_INDICADOR,DEPARTAMENTO_ID FROM INDICADOR ORDER BY DEPARTAMENTO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
