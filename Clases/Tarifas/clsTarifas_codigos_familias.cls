VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTarifas_codigos_familias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTARIFAS_CODIGOS_FAMILIAS
'***************************************************************
Private ID_FAMILIA_CODIGOS As Long
Private DESCRIPCION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTARIFAS_CODIGOS_FAMILIAS
'***************************************************************
Public Property Let setID_FAMILIA_CODIGOS(ByVal dato As Long)
        ID_FAMILIA_CODIGOS = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTARIFAS_CODIGOS_FAMILIAS
'***************************************************************
Public Property Get getID_FAMILIA_CODIGOS() As Long
        getID_FAMILIA_CODIGOS = ID_FAMILIA_CODIGOS
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTARIFAS_CODIGOS_FAMILIAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM TARIFAS_CODIGOS_FAMILIAS WHERE ID_FAMILIA_CODIGOS = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FAMILIA_CODIGOS = rs("ID_FAMILIA_CODIGOS")
                DESCRIPCION = rs("DESCRIPCION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTarifas_codigos_familias)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT MAX(ID_FAMILIA_CODIGOS) FROM TARIFAS_CODIGOS_FAMILIAS"
        Set rs = datos_bd(CONSULTA)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FAMILIA_CODIGOS = 1
        Else
                ID_FAMILIA_CODIGOS = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Me.CrearID
        CONSULTA = "INSERT INTO TARIFAS_CODIGOS_FAMILIAS " & _
                   "  VALUES (" & _
                        ID_FAMILIA_CODIGOS & "," & "'" & DESCRIPCION & "'" & _
                ")"
        execute_bd CONSULTA
        Insertar = ID_FAMILIA_CODIGOS
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTarifas_codigos_familias)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE TARIFAS_CODIGOS_FAMILIAS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE ID_FAMILIA_CODIGOS = " & ID
        execute_bd CONSULTA
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTarifas_codigos_familias)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        Dim rs As ADODB.Recordset
        CONSULTA = "SELECT * FROM TARIFAS_CODIGOS WHERE FAMILIA_CODIGO_ID = " & ID
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount > 0 Then
            MsgBox "No se puede eliminar. Existen códigos con esta familia.", vbCritical, App.Title
            Eliminar = False
            Exit Function
        End If
        CONSULTA = "DELETE FROM TARIFAS_CODIGOS_FAMILIAS " & _
                "    WHERE ID_FAMILIA_CODIGOS = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTarifas_codigos_familias)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT * FROM TARIFAS_CODIGOS_FAMILIAS ORDER BY ID_FAMILIA_CODIGOS"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT ID_FAMILIA_CODIGOS,DESCRIPCION FROM TARIFAS_CODIGOS_FAMILIAS ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TARIFAS_CODIGOS_FAMILIAS"
        .setDESCRIPCION = "Familias"
        .setPK = "ID_FAMILIA_CODIGOS"
        .setCAMPO = "DESCRIPCION"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
