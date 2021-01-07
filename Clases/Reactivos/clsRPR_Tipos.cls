VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPR_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_REACTIVO_PR As Integer
Private NOMBRE As String
Private CODIGO As String
Private ALMACENAMIENTO As String
Private Equipos As String
Private cantidad As String
Private PROTOCOLO As String
Private PROC_REFERENCIA As String
Private TIPO_CADUCIDAD_ID As Integer
Private CENTRO_ID As Integer
'E0151-I
Private tipo As Long
'E0151-F
Private anulado As Integer

Public Property Let setID_TIPO_REACTIVO_PR(ByVal dato As Integer)
    ID_TIPO_REACTIVO_PR = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setCODIGO(ByVal dato As String)
    CODIGO = Trim(dato)
End Property
Public Property Let setALMACENAMIENTO(ByVal dato As String)
    ALMACENAMIENTO = Trim(dato)
End Property
Public Property Let setEQUIPOS(ByVal dato As String)
    Equipos = Trim(dato)
End Property
Public Property Let setCANTIDAD(ByVal dato As String)
    cantidad = Trim(dato)
End Property
Public Property Let setPROTOCOLO(ByVal dato As String)
    PROTOCOLO = Trim(dato)
End Property
Public Property Let setPROC_REFERENCIA(ByVal dato As String)
    PROC_REFERENCIA = Trim(dato)
End Property
Public Property Let setTIPO_CADUCIDAD_ID(ByVal dato As Integer)
    TIPO_CADUCIDAD_ID = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Integer)
    CENTRO_ID = dato
End Property
'E0152-I
Public Property Let setTIPO(ByVal dato As Integer)
    tipo = dato
End Property
'E0152-F
Public Property Let setANULADO(ByVal dato As Integer)
    anulado = dato
End Property
Public Property Get getID_TIPO_REACTIVO_PR() As Integer
    getID_TIPO_REACTIVO_PR = ID_TIPO_REACTIVO_PR
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getCODIGO() As String
    getCODIGO = CODIGO
End Property
Public Property Get getALMACENAMIENTO() As String
    getALMACENAMIENTO = ALMACENAMIENTO
End Property
Public Property Get getEQUIPOS() As String
    getEQUIPOS = Equipos
End Property
Public Property Get getCANTIDAD() As String
    getCANTIDAD = cantidad
End Property
Public Property Get getPROTOCOLO() As String
    getPROTOCOLO = PROTOCOLO
End Property
Public Property Get getPROC_REFERENCIA() As String
    getPROC_REFERENCIA = PROC_REFERENCIA
End Property
Public Property Get getTIPO_CADUCIDAD_ID() As Integer
    getTIPO_CADUCIDAD_ID = TIPO_CADUCIDAD_ID
End Property
Public Property Get getCENTRO_ID() As Integer
    getCENTRO_ID = CENTRO_ID
End Property
'E0153-I
Public Property Get getTIPO() As Integer
    getTIPO = tipo
End Property
Public Property Get getANULADO() As Integer
    getANULADO = anulado
End Property

Public Function Listado(NOMBRE As String, PROCEDIMIENTO As String, strTipo As Integer, anulado As Boolean, CENTRO_ID As Integer) As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT A.*,B.NOMBRE AS CENTRO " & _
               "  FROM RPR_TIPOS A,CENTROS B " & _
               " WHERE A.CENTRO_ID = B.ID_CENTRO AND A.NOMBRE LIKE '%" & NOMBRE & "%'" & _
               " AND A.PROC_REFERENCIA LIKE '%" & PROCEDIMIENTO & "%'" & _
               IIf(strTipo <> 0, " AND A.TIPO = " & strTipo, "") & _
               IIf(anulado = False, " AND A.ANULADO = 0 ", "") & _
               IIf(CENTRO_ID = 0, "", " AND A.CENTRO_ID = " & CENTRO_ID) & _
               " ORDER BY A.NOMBRE"
    Set Listado = datos_bd(CONSULTA)
End Function
'E0153-F
Public Function CARGAR(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM RPR_TIPOS WHERE ID_TIPO_REACTIVO_PR=" & ID
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount > 0 Then
        ID_TIPO_REACTIVO_PR = rs("ID_TIPO_REACTIVO_PR")
        NOMBRE = rs("NOMBRE")
        CODIGO = rs("CODIGO")
        ALMACENAMIENTO = rs("ALMACENAMIENTO")
        Equipos = rs("EQUIPOS")
        cantidad = rs("CANTIDAD")
        PROTOCOLO = rs("PROTOCOLO")
        PROC_REFERENCIA = rs("PROC_REFERENCIA")
        TIPO_CADUCIDAD_ID = rs("TIPO_CADUCIDAD_ID")
        CENTRO_ID = rs("CENTRO_ID")
        'E0154-I
        tipo = rs("TIPO")
        anulado = rs("ANULADO")
        'E0154-F
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
    MsgBox "Error al cargar el reactivo.", vbCritical, Err.Description
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_TIPO_REACTIVO_PR) FROM RPR_TIPOS"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_REACTIVO_PR = 1
    Else
        ID_TIPO_REACTIVO_PR = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Private Function validarCodigoDuplicado(ID As Long, CODIGO As String) As Boolean
    Dim CONSULTA As String
    Dim rs As ADODB.Recordset
    CONSULTA = "SELECT * FROM RPR_TIPOS WHERE ID_TIPO_REACTIVO_PR <> " & ID & " AND CODIGO = '" & CODIGO & "'"
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount > 0 Then
        validarCodigoDuplicado = True
    Else
        validarCodigoDuplicado = False
    End If
End Function

Public Function Insertar() As Long
    On Error GoTo fallo
    Dim CONSULTA As String
    If validarCodigoDuplicado(CLng(ID_TIPO_REACTIVO_PR), CODIGO) Then
        MsgBox "El CODIGO ya existe en otro reactivo propio.", vbCritical, App.Title
        Insertar = 0
        Exit Function
    End If
    Me.CrearID
    'E0155-I
    CONSULTA = "Insert into RPR_TIPOS " & _
               " values(" & _
               ID_TIPO_REACTIVO_PR & ",'" & NOMBRE & "','" & _
               CODIGO & "','" & _
               ALMACENAMIENTO & "','" & _
               Equipos & "','" & _
               cantidad & "','" & _
               PROTOCOLO & "','" & _
               PROC_REFERENCIA & "'," & _
               TIPO_CADUCIDAD_ID & ", " & _
               CENTRO_ID & "," & _
               tipo & ",0)"
    'E0155-I
    execute_bd CONSULTA
    Insertar = ID_TIPO_REACTIVO_PR
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Reactivo (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
'    Dim RS As New ADODB.Recordset
    Dim CONSULTA As String
'    Dim oComponentes As New clsRPR_Componentes
'    If oComponentes.Eliminar(id) = True Then
        CONSULTA = "UPDATE RPR_TIPOS SET ANULADO = " & usuario.getID_EMPLEADO & " WHERE ID_TIPO_REACTIVO_PR=" & ID
        execute_bd CONSULTA
        Eliminar = True
'    End If
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el Reactivo (Eliminar)", vbCritical, Err.Description
End Function
Public Function Reactivar(ID As Long) As Boolean
    On Error GoTo fallo
'    Dim RS As New ADODB.Recordset
    Dim CONSULTA As String
'    Dim oComponentes As New clsRPR_Componentes
'    If oComponentes.Eliminar(id) = True Then
        CONSULTA = "UPDATE RPR_TIPOS SET ANULADO = 0 WHERE ID_TIPO_REACTIVO_PR=" & ID
        execute_bd CONSULTA
        Reactivar = True
'    End If
    Exit Function
fallo:
    Reactivar = False
    MsgBox "Error al Reactivar el Reactivo (Eliminar)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim CONSULTA As String
    If validarCodigoDuplicado(ID, CODIGO) Then
        MsgBox "El CODIGO ya existe en otro reactivo propio.", vbCritical, App.Title
        Modificar = False
        Exit Function
    End If
    'E0156-I
    CONSULTA = "UPDATE RPR_TIPOS SET " & _
               " NOMBRE = '" & NOMBRE & "'," & _
               " CODIGO = '" & CODIGO & "'," & _
               " ALMACENAMIENTO = '" & ALMACENAMIENTO & "'," & _
               " EQUIPOS = '" & Equipos & "'," & _
               " CANTIDAD = '" & cantidad & "'," & _
               " PROTOCOLO = '" & PROTOCOLO & "'," & _
               " PROC_REFERENCIA = '" & PROC_REFERENCIA & "'," & _
               " TIPO_CADUCIDAD_ID = " & TIPO_CADUCIDAD_ID & ", " & _
               " CENTRO_ID = " & CENTRO_ID & ", " & _
               " TIPO = " & tipo & _
               " WHERE ID_TIPO_REACTIVO_PR = " & ID
    'E0156-I
    execute_bd CONSULTA
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar el Reactivo (Modificar)", vbCritical, Err.Description
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT id_tipo_reactivo_pr,nombre FROM RPR_TIPOS order by nombre"
    Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "RPR_TIPOS"
        .setDESCRIPCION = "Reactivos / Suministros"
        .setPK = "ID_TIPO_REACTIVO_PR"
        .setCAMPO = "NOMBRE"
        .setFILTRO = filtro
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function
