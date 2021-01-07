VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_Frecuencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_FRECUENCIA As Integer
Private NOMBRE As String
Public Property Let setID_TIPO_FRECUENCIA(ByVal dato As Integer)
    ID_TIPO_FRECUENCIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getID_TIPO_FRECUENCIA() As Integer
    getID_TIPO_FRECUENCIA = ID_TIPO_FRECUENCIA
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Function Listado() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM TIPOS_FRECUENCIA order by nombre"
    Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_TIPO_FRECUENCIA,NOMBRE FROM tipos_frecuencia order by nombre"
    Set Listado_Combo = datos_bd(CONSULTA)
End Function

Public Function CARGAR(TIPO_FRECUENCIA As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM TIPOS_FRECUENCIA WHERE ID_TIPO_FRECUENCIA=" & TIPO_FRECUENCIA
    Set rs = datos_bd(CONSULTA)
    ID_TIPO_FRECUENCIA = rs("ID_TIPO_FRECUENCIA")
    NOMBRE = rs("NOMBRE")
    CARGAR = True
    Exit Function
fallo:
    CARGAR = False
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_TIPO_FRECUENCIA) FROM TIPOS_FRECUENCIA"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_FRECUENCIA = 1
    Else
        ID_TIPO_FRECUENCIA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim CONSULTA As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        CONSULTA = "Insert into TIPOS_FRECUENCIA " & _
                   " values(" & _
                   ID_TIPO_FRECUENCIA & ",'" & NOMBRE & "')"
        execute_bd CONSULTA
        Insertar = ID_TIPO_FRECUENCIA
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar las TIPOS_FRECUENCIA (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "DELETE FROM TIPOS_FRECUENCIA WHERE ID_TIPO_FRECUENCIA=" & ID
    execute_bd CONSULTA
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Modificar(TIPO_FRECUENCIA As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "UPDATE TIPOS_FRECUENCIA " & _
               " SET NOMBRE = '" & NOMBRE & "'" & _
               " WHERE ID_TIPO_FRECUENCIA=" & TIPO_FRECUENCIA
    execute_bd CONSULTA
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_TIPO_FRECUENCIA FROM TIPOS_FRECUENCIA WHERE NOMBRE = '" & NOMBRE & "'"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe una frecuencia con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "TIPOS_FRECUENCIA"
        .setDESCRIPCION = "Periodicidades"
        .setPK = "ID_TIPO_FRECUENCIA"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function

