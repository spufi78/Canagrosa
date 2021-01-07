VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_Control_Tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_CONTROL As Integer
Private NOMBRE As String
Public Property Let setID_CONTROL(ByVal dato As Integer)
    ID_CONTROL = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getID_CONTROL() As Integer
    getID_CONTROL = ID_CONTROL
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_CONTROL_TIPOS order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function cargar(estado As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_CONTROL_TIPOS WHERE ID_CONTROL=" & estado
    Set rs = datos_bd(consulta)
    ID_CONTROL = rs("ID_CONTROL")
    NOMBRE = rs("NOMBRE")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_CONTROL) FROM EMPLEADOS_CONTROL_TIPOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_CONTROL = 1
    Else
        ID_CONTROL = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into EMPLEADOS_CONTROL_TIPOS " & _
                   " values(" & _
                   ID_CONTROL & ",'" & NOMBRE & "')"
        execute_bd consulta
        Insertar = ID_CONTROL
    Else
        MsgBox "Ya existe ese estado.", vbInformation, App.Title
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar es estado (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(estado As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ESTADO_ID FROM OPERARIOS WHERE ESTADO_ID = " & estado
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen empleados con ese estado. No se puede eliminar.", vbCritical, App.Title
        Eliminar = False
        Exit Function
    End If
    consulta = "DELETE FROM EMPLEADOS_CONTROL_TIPOS WHERE ID_CONTROL=" & estado
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_CONTROL FROM EMPLEADOS_CONTROL_TIPOS WHERE NOMBRE = '" & NOMBRE & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        duplicado = True
    End If
    Set rs = Nothing
End Function

Public Function Modificar(estado As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE EMPLEADOS_CONTROL_TIPOS " & _
               " SET NOMBRE = '" & NOMBRE & "'" & _
               " WHERE ID_CONTROL=" & estado
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function
Public Function Listado_por_defecto() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_CONTROL_TIPOS order by ID_CONTROL"
    Set Listado_por_defecto = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_CONTROL, NOMBRE FROM EMPLEADOS_CONTROL_TIPOS " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function

