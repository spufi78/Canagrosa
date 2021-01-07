VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPinturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPINTURAS
'***************************************************************
Private ID_PINTURA As Long
Private CODIGO As String
Private DESCRIPCION As String
Private PRECIO As String
Private OT_ID As Long
Private FICHA_ID As Long
Private ACTIVO As Long

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSPINTURAS
'***************************************************************
Public Enum LISTA_CLSPINTURAS_COLS
        ENUM_ID_PINTURA = 1
        ENUM_CODIGO = 2
        ENUM_DESCRIPCION = 3
        ENUM_PRECIO = 4
        ENUM_OT_ID = 5
        ENUM_FICHA_ID = 6
        ENUM_ACTIVO = 7
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPINTURAS
'***************************************************************
Public Property Let setID_PINTURA(ByVal dato As Long)
        ID_PINTURA = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setOT_ID(ByVal dato As Long)
        OT_ID = dato
End Property
Public Property Let setFICHA_ID(ByVal dato As Long)
        FICHA_ID = dato
End Property
Public Property Let setACTIVO(ByVal dato As Long)
        ACTIVO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPINTURAS
'***************************************************************
Public Property Get getID_PINTURA() As Long
        getID_PINTURA = ID_PINTURA
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getOT_ID() As Long
        getOT_ID = OT_ID
End Property
Public Property Get getFICHA_ID() As Long
        getFICHA_ID = FICHA_ID
End Property
Public Property Get getACTIVO() As Long
        getACTIVO = ACTIVO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPINTURAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM PINTURAS WHERE ID_PINTURA = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PINTURA = rs("ID_PINTURA")
                CODIGO = rs("CODIGO")
                DESCRIPCION = rs("DESCRIPCION")
                PRECIO = rs("PRECIO")
                OT_ID = rs("OT_ID")
                FICHA_ID = rs("FICHA_ID")
                ACTIVO = rs("ACTIVO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPinturas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_PINTURA) FROM PINTURAS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PINTURA = 10000
        Else
                ID_PINTURA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO PINTURAS " & _
                   "  VALUES (" & _
                        ID_PINTURA & "," & "'" & CODIGO & "'" & "," & "'" & DESCRIPCION & "'" & ",'" & _
                        PRECIO & "'," & OT_ID & "," & FICHA_ID & "," & ACTIVO & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_PINTURA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPinturas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE PINTURAS SET " & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " PRECIO = '" & PRECIO & "'," & _
                        " OT_ID = " & OT_ID & "," & _
                        " FICHA_ID = " & FICHA_ID & "," & _
                        " ACTIVO = " & ACTIVO & "" & _
                " WHERE ID_PINTURA = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPinturas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        execute_bd "DELETE FROM PINTURAS_TA WHERE PINTURA_ID = " & ID
        execute_bd "DELETE FROM PINTURAS_CE WHERE PINTURA_ID = " & ID
        execute_bd "DELETE FROM PINTURAS WHERE ID_PINTURA = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPinturas)", vbCritical, Err.Description
End Function
Public Function Listado(CODIGO As String, TA As Long, CE As Long, ACTIVO As Integer) As ADODB.Recordset
        Dim CONSULTA As String
        Dim filtro As String
        filtro = " WHERE 1 = 1 "
        If CODIGO <> "" Then
            filtro = filtro & " AND A.CODIGO LIKE '%" & CODIGO & "%'"
        End If
        If ACTIVO = 1 Then
            filtro = filtro & " AND A.ACTIVO = 1 "
        End If
        If TA <> 0 Then
            filtro = filtro & " AND C.TIPO_ANALISIS_ID = " & TA
        End If
        If CE <> 0 Then
            filtro = filtro & " AND B.TIPO_ENSAYO_ID = " & CE
        End If
        CONSULTA = "SELECT DISTINCT A.ID_PINTURA,A.CODIGO,A.DESCRIPCION,A.ACTIVO " & _
                   "  FROM PINTURAS A " & _
                   "  LEFT JOIN PINTURAS_CE B ON A.ID_PINTURA = B.PINTURA_ID " & _
                   "  LEFT JOIN PINTURAS_TA C ON A.ID_PINTURA = C.PINTURA_ID " & _
                   filtro & _
                   " ORDER BY A.CODIGO"
        Set Listado = datos_bd(CONSULTA)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPINTURAS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim CONSULTA As String
                CONSULTA = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(CONSULTA)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_PINTURA,CODIGO FROM PINTURAS ORDER BY CODIGO"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "PINTURAS"
        .setDESCRIPCION = "Pinturas"
        .setPK = "ID_PINTURA"
        .setCAMPO = "CONCAT(CODIGO,' : ',DESCRIPCION)"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function

