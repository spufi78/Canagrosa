VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_ensayos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_ENSAYOS
'***************************************************************
Private ID_ENSAYO As Long
Private nombre As String
Private REQUIREMENT As String
Private TARIFA_CODIGO_ID As Long
Private TIPO_ID As Long
Private UNIDAD_ID As Long


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_ENSAYOS
'***************************************************************
Public Property Let setID_ENSAYO(ByVal dato As Long)
        ID_ENSAYO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setREQUIREMENT(ByVal dato As String)
        REQUIREMENT = dato
End Property
Public Property Let setTARIFA_CODIGO_ID(ByVal dato As Long)
        TARIFA_CODIGO_ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_ENSAYOS
'***************************************************************
Public Property Get getID_ENSAYO() As Long
        getID_ENSAYO = ID_ENSAYO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getREQUIREMENT() As String
        getREQUIREMENT = REQUIREMENT
End Property
Public Property Get getTARIFA_CODIGO_ID() As Long
        getTARIFA_CODIGO_ID = TARIFA_CODIGO_ID
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_ENSAYOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_ENSAYOS WHERE ID_ENSAYO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ENSAYO = rs("ID_ENSAYO")
                nombre = rs("NOMBRE")
                REQUIREMENT = rs("REQUIREMENT")
                TARIFA_CODIGO_ID = rs("TARIFA_CODIGO_ID")
                TIPO_ID = rs("TIPO_ID")
                UNIDAD_ID = rs("UNIDAD_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_ensayos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_ENSAYO) FROM PLASMA_ENSAYOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ENSAYO = 1
        Else
                ID_ENSAYO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PLASMA_ENSAYOS " & _
                   "  VALUES (" & _
                        ID_ENSAYO & "," & "'" & nombre & "'" & "," & "'" & REQUIREMENT & "'" & "," & _
                        TARIFA_CODIGO_ID & "," & TIPO_ID & "," & UNIDAD_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_ENSAYO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_ensayos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_ENSAYOS SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " REQUIREMENT = '" & REQUIREMENT & "'," & _
                        " TARIFA_CODIGO_ID = " & TARIFA_CODIGO_ID & "," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "," & _
                        " TIPO_ID = " & TIPO_ID & "" & _
                " WHERE ID_ENSAYO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_ensayos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Me.Equipos_Eliminar ID
        consulta = "DELETE FROM PLASMA_ENSAYOS " & _
                "    WHERE ID_ENSAYO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_ensayos)", vbCritical, Err.Description
End Function
Public Function Listado(sNOMBRE As String, sTIPO As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE 1 = 1"
        If sNOMBRE <> "" Then
            filtro = filtro & " AND A.NOMBRE LIKE '%" & sNOMBRE & "%'"
        End If
        If sTIPO <> "0" Then
            filtro = filtro & " AND A.TIPO_ID = " & sTIPO
        End If
        consulta = "SELECT A.ID_ENSAYO,B.DESCRIPCION,A.NOMBRE,A.REQUIREMENT,C.NOMBRE " & _
                   "  FROM PLASMA_ENSAYOS A " & _
                   "  LEFT JOIN decodificadora B ON A.TIPO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PLASMA_TIPOS & _
                   "  LEFT JOIN UNIDADES C ON A.UNIDAD_ID = C.ID_UNIDAD " & _
                   filtro & _
                   " ORDER BY A.NOMBRE "
        Set Listado = datos_bd(consulta)
End Function
Public Function ListadoId(ID As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = " WHERE A.ID_ENSAYO = " & ID
        consulta = "SELECT A.ID_ENSAYO,B.DESCRIPCION,A.NOMBRE,A.REQUIREMENT,C.NOMBRE " & _
                   "  FROM PLASMA_ENSAYOS A " & _
                   "  LEFT JOIN decodificadora B ON A.TIPO_ID = B.VALOR AND B.CODIGO = " & DECODIFICADORA.DECODIFICADORA_PLASMA_TIPOS & _
                   "  LEFT JOIN UNIDADES C ON A.UNIDAD_ID = C.ID_UNIDAD " & _
                   filtro & _
                   " ORDER BY A.NOMBRE "
        Set ListadoId = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_ENSAYOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_ENSAYO,NOMBRE FROM PLASMA_ENSAYOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setFILTRO = filtro
        .setTABLA = "PLASMA_ENSAYOS"
        .setDESCRIPCION = "Ensayos de Plasma"
        .setPK = "ID_ENSAYO"
        .setCAMPO = "CONCAT(NOMBRE,' (',REQUIREMENT,')')"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Equipos_Eliminar(ENSAYO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd ("DELETE FROM plasma_ENSAYOS_EQUIPOS WHERE ENSAYO_ID = " & ENSAYO_ID)
        Equipos_Eliminar = True
        Exit Function
fallo:
        Equipos_Eliminar = False
        MsgBox "Error al Equipos_Eliminar (clsCe_tipos_ensayos)", vbCritical, Err.Description
End Function
Public Function Equipos_Insertar(ENSAYO_ID As Long, EQUIPO As Long, ORDEN As Integer, informe As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "INSERT INTO plasma_ENSAYOS_EQUIPOS VALUES (" & ENSAYO_ID & "," & EQUIPO & "," & ORDEN & "," & Abs(informe) & ")"
        Equipos_Insertar = True
        Exit Function
fallo:
        Equipos_Insertar = False
        MsgBox "Error al Equipos_Insertar (clsCe_tipos_ensayos)", vbCritical, Err.Description
End Function

Public Function Equipos_Listado(ENSAYO_ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.EQUIPO_ID,B.NOMBRE,B.SERIE,A.EN_INFORME,A.ORDEN " & _
                   "  FROM plasma_ENSAYOS_EQUIPOS A, EQUIPOS B" & _
                   " WHERE A.ENSAYO_ID = " & ENSAYO_ID & _
                   "   AND A.EQUIPO_ID = B.ID_EQUIPO " & _
                   " ORDER BY A.ORDEN"
        Set Equipos_Listado = datos_bd(consulta)
End Function
Public Function Equipos_Duplicar(ENSAYO_ID_ORIGEN As Long, ENSAYO_ID_DESTINO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "INSERT INTO plasma_ensayos_equipos " & _
                   "SELECT " & ENSAYO_ID_DESTINO & ",A.EQUIPO_ID,A.ORDEN,A.EN_INFORME " & _
                   "  FROM plasma_ensayos_equipos A " & _
                   " WHERE A.ENSAYO_ID = " & ENSAYO_ID_ORIGEN
        Set Equipos_Duplicar = datos_bd(consulta)
End Function


