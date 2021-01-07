VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluidos_normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFLUIDOS_NORMAS
'***************************************************************
Private ID_NORMA As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFLUIDOS_NORMAS
'***************************************************************
Public Property Let setID_NORMA(ByVal dato As Long)
        ID_NORMA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFLUIDOS_NORMAS
'***************************************************************
Public Property Get getID_NORMA() As Long
        getID_NORMA = ID_NORMA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFLUIDOS_NORMAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM FLUIDOS_NORMAS WHERE ID_NORMA = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_NORMA = rs("ID_NORMA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFLUIDOS_NORMAS)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_NORMA) FROM FLUIDOS_NORMAS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_NORMA = 1
        Else
                ID_NORMA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO FLUIDOS_NORMAS " & _
                   "  VALUES (" & _
                        ID_NORMA & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_NORMA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFLUIDOS_NORMAS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE FLUIDOS_NORMAS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_NORMA = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFLUIDOS_NORMAS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM FLUIDOS_NORMAS " & _
                "    WHERE ID_NORMA = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFLUIDOS_NORMAS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM FLUIDOS_NORMAS ORDER BY ID_NORMA"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_NORMA,NOMBRE FROM FLUIDOS_NORMAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "FLUIDOS_NORMAS"
        .setDESCRIPCION = "Normas de Fluidos"
        .setPK = "ID_NORMA"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
