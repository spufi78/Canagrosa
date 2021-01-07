VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEq_trazmet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQ_TRAZMET
'***************************************************************
Private ID_TRAZMET As Long
Private CODIGO As String
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQ_TRAZMET
'***************************************************************
Public Property Let setID_TRAZMET(ByVal dato As Long)
        ID_TRAZMET = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQ_TRAZMET
'***************************************************************
Public Property Get getID_TRAZMET() As Long
        getID_TRAZMET = ID_TRAZMET
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_TRAZMET
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_TRAZMET WHERE ID_TRAZMET = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TRAZMET = rs("ID_TRAZMET")
                CODIGO = rs("CODIGO")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEq_trazmet)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_TRAZMET) FROM EQ_TRAZMET"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_TRAZMET = 1
        Else
                ID_TRAZMET = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EQ_TRAZMET " & _
                   "  VALUES (" & _
                        ID_TRAZMET & "," & "'" & CODIGO & "'" & "," & "'" & DESCRIPCION & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_TRAZMET
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEq_trazmet)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_TRAZMET SET " & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID_TRAZMET = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEq_trazmet)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EQ_TRAZMET " & _
                "    WHERE ID_TRAZMET = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEq_trazmet)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_TRAZMET ORDER BY ID_TRAZMET"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSEQ_TRAZMET
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_TRAZMET,CODIGO FROM EQ_TRAZMET ORDER BY CODIGO"
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "eq_trazmet"
        .setDESCRIPCION = "Trazmet"
        .setPK = "ID_TRAZMET"
        .setFILTRO = filtro
        .setCAMPO = "CODIGO"
        .setMUESTRA_DETALLE = False
        .setFILTRO = ""
        Set .FORMULARIO = FORMULARIO
    End With
End Function
