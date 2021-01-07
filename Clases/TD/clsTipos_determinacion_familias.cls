VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_determinacion_familias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_DETERMINACION_FAMILIAS
'***************************************************************
Private ID_FAMILIA As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_DETERMINACION_FAMILIAS
'***************************************************************
Public Property Let setID_FAMILIA(ByVal dato As Long)
        ID_FAMILIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_DETERMINACION_FAMILIAS
'***************************************************************
Public Property Get getID_FAMILIA() As Long
        getID_FAMILIA = ID_FAMILIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_DETERMINACION_FAMILIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM TIPOS_DETERMINACION_FAMILIAS WHERE ID_FAMILIA = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FAMILIA = rs("ID_FAMILIA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_determinacion_familias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_FAMILIA) FROM TIPOS_DETERMINACION_FAMILIAS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FAMILIA = 1
        Else
                ID_FAMILIA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO TIPOS_DETERMINACION_FAMILIAS " & _
                   "  VALUES (" & _
                        ID_FAMILIA & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_FAMILIA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_determinacion_familias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE TIPOS_DETERMINACION_FAMILIAS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_FAMILIA = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTipos_determinacion_familias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        Dim rs As ADODB.Recordset
        CONSULTA = "SELECT * FROM TIPOS_DETERMINACION WHERE FAMILIA_ID = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar. Existen determinaciones con esta familia.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        CONSULTA = "DELETE FROM TIPOS_DETERMINACION_FAMILIAS " & _
                "    WHERE ID_FAMILIA = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_determinacion_familias)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM TIPOS_DETERMINACION_FAMILIAS " & _
                   " WHERE ID_FAMILIA <> 0 " & _
                   " ORDER BY ID_FAMILIA"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_FAMILIA,NOMBRE FROM TIPOS_DETERMINACION_FAMILIAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TIPOS_DETERMINACION_FAMILIAS"
        .setDESCRIPCION = "Familias de determinaciones"
        .setPK = "ID_FAMILIA"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
