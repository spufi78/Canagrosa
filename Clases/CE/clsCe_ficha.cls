VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_ficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_FICHA
'***************************************************************
Private ID_FICHA As Long
Private PROCESO As String
Private ACTIVO As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_FICHA
'***************************************************************
Public Property Let setID_FICHA(ByVal dato As Long)
        ID_FICHA = dato
End Property
Public Property Let setPROCESO(ByVal dato As String)
        PROCESO = dato
End Property
Public Property Let setACTIVO(ByVal dato As Integer)
        ACTIVO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_FICHA
'***************************************************************
Public Property Get getID_FICHA() As Long
        getID_FICHA = ID_FICHA
End Property
Public Property Get getPROCESO() As String
        getPROCESO = PROCESO
End Property
Public Property Get getACTIVO() As Integer
        getACTIVO = ACTIVO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_FICHA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM CE_FICHA WHERE ID_FICHA = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FICHA = rs("ID_FICHA")
                PROCESO = rs("PROCESO")
                ACTIVO = rs("ACTIVO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_ficha)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_FICHA) FROM CE_FICHA"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FICHA = 1
        Else
                ID_FICHA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO CE_FICHA " & _
                   "  VALUES (" & _
                        ID_FICHA & "," & "'" & PROCESO & "'," & ACTIVO & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_FICHA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_ficha)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE CE_FICHA SET " & _
                        " PROCESO = '" & PROCESO & "', " & _
                        " ACTIVO = " & ACTIVO & _
                   " WHERE ID_FICHA = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_ficha)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM CE_ENSAYOS WHERE FICHA_ID = " & ID
        execute_bd CONSULTA
        CONSULTA = "DELETE FROM CE_FICHA WHERE ID_FICHA = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_ficha)", vbCritical, Err.Description
End Function
Public Function Listado(filtro As String, BANO As String, ACTIVA As Boolean) As ADODB.Recordset
        Dim CONSULTA As String
        Dim s1 As String
        Dim s2 As String
        If BANO <> "" Then
            s1 = " AND B.ID_BANO = " & CLng(BANO)
        End If
        If ACTIVA = True Then
            s2 = " AND A.ACTIVO = 1 "
        End If
        CONSULTA = "SELECT DISTINCT A.PROCESO,A.ID_FICHA,A.ACTIVO " & _
                   "  FROM CE_FICHA A LEFT JOIN BANOS B " & _
                   "    ON A.ID_FICHA = B.FICHA_ID " & _
                   " WHERE A.PROCESO LIKE '%" & filtro & "%'" & _
                   s1 & s2 & _
                   " ORDER BY ID_FICHA"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_por_ID(ID As Long) As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT A.PROCESO,A.ID_FICHA,A.ACTIVO " & _
                   "  FROM CE_FICHA A " & _
                   " WHERE A.ID_FICHA = " & ID
        Set Listado_por_ID = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_FICHA,PROCESO FROM CE_FICHA ORDER BY PROCESO"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CE_FICHA"
        .setDESCRIPCION = "Fichas de Controles de Eficacia"
        .setFILTRO = filtro
        .setPK = "ID_FICHA"
        .setCAMPO = "PROCESO"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
