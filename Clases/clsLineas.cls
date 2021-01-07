VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_LINEA As Integer
Private NOMBRE As String
Public Property Let setID_LINEA(ByVal dato As Integer)
    ID_LINEA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getID_LINEA() As Integer
    getID_LINEA = ID_LINEA
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Function Listado() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM LINEAS order by nombre"
    Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Historico() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT distinct a.id_linea,a.nombre " & _
               "  FROM LINEAS a, banos b,banos_control c " & _
               " where a.id_linea = b.linea_id " & _
               "   and b.id_bano = c.bano_id " & _
               " order by nombre"
    Set Listado_Historico = datos_bd(CONSULTA)
End Function
Public Function CARGAR(linea As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM LINEAS WHERE ID_LINEA=" & linea
    Set rs = datos_bd(CONSULTA)
    ID_LINEA = rs("ID_LINEA")
    NOMBRE = rs("NOMBRE")
    CARGAR = True
    Exit Function
fallo:
    CARGAR = False
End Function

Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_LINEA) FROM LINEAS"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_LINEA = 1
    Else
        ID_LINEA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function InsertarLinea() As Long
    On Error GoTo fallo
    Dim CONSULTA As String
    InsertarLinea = 0
    If duplicado = False Then
        Me.CrearID
        CONSULTA = "Insert into LINEAS " & _
                   " values(" & _
                   ID_LINEA & ",'" & NOMBRE & "')"
        execute_bd CONSULTA
        InsertarLinea = ID_LINEA
    End If
    Exit Function
fallo:
    InsertarLinea = 0
    MsgBox "Error al insertar la Linea (InsertarLinea)", vbCritical, Err.Description
End Function
Public Function Eliminar(linea As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM BANOS WHERE LINEA_ID=" & linea
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        CONSULTA = "DELETE FROM lineas WHERE ID_LINEA=" & linea
        execute_bd CONSULTA
    Else
        MsgBox "Existen baños con esta linea. No se puede eliminar.", vbExclamation, App.Title
    End If
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function

Public Function Modificar(linea As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "UPDATE lineas " & _
               " SET NOMBRE = '" & NOMBRE & "'" & _
               " WHERE ID_LINEA=" & linea
    execute_bd CONSULTA
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_LINEA FROM LINEAS WHERE NOMBRE = '" & NOMBRE & "'"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe una linea con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function Max_Linea(documento As Integer) As Integer
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(LINEA_ID) FROM LINEAS_BANOS WHERE DOCUMENTO=" & documento
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Max_Linea = 0
    Else
        Max_Linea = rs(0)
    End If
    Exit Function
fallo:
    Max_Linea = 0
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_LINEA, NOMBRE FROM LINEAS " & _
               " ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "LINEAS"
        .setDESCRIPCION = "Lineas de Proceso"
        .setPK = "ID_LINEA"
        .setCAMPO = "NOMBRE"
        .setFILTRO = filtro
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
