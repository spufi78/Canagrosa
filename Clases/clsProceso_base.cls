VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProceso_base"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_PROCESO_BASE As Integer
Private nombre As String
Private NOMBRE_INGLES As String
Public Property Let setID_PROCESO_BASE(ByVal dato As Integer)
    ID_PROCESO_BASE = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setNOMBRE_INGLES(ByVal dato As String)
    NOMBRE_INGLES = Trim(dato)
End Property
Public Property Get getID_PROCESO_BASE() As Integer
    getID_PROCESO_BASE = ID_PROCESO_BASE
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getNOMBRE_INGLES() As String
    getNOMBRE_INGLES = NOMBRE_INGLES
End Property
Public Function Listado(nombre As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM PROCESOS_BASE " & _
               " WHERE NOMBRE LIKE '%" & nombre & "%'" & _
               " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PROCESO_BASE,NOMBRE FROM PROCESOS_BASE order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function CARGAR(PROCESO_BASE As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM PROCESOS_BASE WHERE ID_PROCESO_BASE=" & PROCESO_BASE
    Set rs = datos_bd(consulta)
    ID_PROCESO_BASE = rs("ID_PROCESO_BASE")
    nombre = rs("NOMBRE")
    NOMBRE_INGLES = rs("NOMBRE_INGLES")
    CARGAR = True
    Exit Function
fallo:
    CARGAR = False
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PROCESO_BASE) FROM PROCESOS_BASE"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PROCESO_BASE = 1
    Else
        ID_PROCESO_BASE = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into PROCESOS_BASE " & _
                   " values(" & _
                   ID_PROCESO_BASE & ",'" & nombre & "','" & NOMBRE_INGLES & "')"
        execute_bd consulta
        Insertar = ID_PROCESO_BASE
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Proceso Base (InsertarProceso_Base)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM PROCESOS_BASE WHERE ID_PROCESO_BASE=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function

Public Function Modificar(PROCESO_BASE As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE PROCESOS_BASE " & _
               " SET NOMBRE = '" & nombre & "'," & _
               "     NOMBRE_INGLES = '" & NOMBRE_INGLES & "'" & _
               " WHERE ID_PROCESO_BASE=" & PROCESO_BASE
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PROCESO_BASE FROM PROCESOS_BASE WHERE NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe un proceso con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "PROCESOS_BASE"
        .setDESCRIPCION = "Procesos"
        .setPK = "ID_PROCESO_BASE"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = frmProcesosBaseDetalle
    End With
End Function
Public Function imprimir() As Boolean
        With frmReport
            .iniciar
            .informe = "\Banos\rptprocesos"
            .criterio = ""
            .imprimir = False
            .generar
            .Show 1
        End With
        Unload frmReport
End Function

