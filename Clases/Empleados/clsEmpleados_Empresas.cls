VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_Empresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_EMPRESA As Integer
Private DESCRIPCION As String
Public Property Let setID_EMPRESA(ByVal dato As Integer)
    ID_EMPRESA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Get getID_EMPRESA() As Integer
    getID_EMPRESA = ID_EMPRESA
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_EMPRESAS order by DESCRIPCION"
    Set Listado = datos_bd(consulta)
End Function
Public Function CARGAR(ESTADO As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_EMPRESAS WHERE ID_EMPRESA=" & ESTADO
    Set rs = datos_bd(consulta)
    ID_EMPRESA = rs("ID_EMPRESA")
    DESCRIPCION = rs("DESCRIPCION")
    CARGAR = True
    Exit Function
fallo:
    CARGAR = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_EMPRESA) FROM EMPLEADOS_EMPRESAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_EMPRESA = 1
    Else
        ID_EMPRESA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into EMPLEADOS_EMPRESAS " & _
                   " values(" & _
                   ID_EMPRESA & ",'" & DESCRIPCION & "')"
        execute_bd consulta
        Insertar = ID_EMPRESA
    Else
        MsgBox "La empresa ya existe.", vbInformation, App.Title
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la empresa (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(empresa As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_EMPLEADO FROM empleados WHERE ID_EMPRESA = " & empresa
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen empleados vinculados a esta empresa. No se puede eliminar.", vbCritical, App.Title
        Eliminar = False
        Exit Function
    End If
    consulta = "DELETE FROM EMPLEADOS_EMPRESAS WHERE ID_EMPRESA=" & empresa
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_EMPRESA FROM EMPLEADOS_EMPRESAS WHERE DESCRIPCION = '" & DESCRIPCION & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        duplicado = True
    End If
    Set rs = Nothing
End Function

Public Function Modificar(empresa As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE EMPLEADOS_EMPRESAS " & _
               " SET DESCRIPCION = '" & DESCRIPCION & "'" & _
               " WHERE ID_EMPRESA=" & empresa
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function
Public Function Listado_por_defecto() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_EMPRESAS order by ID_EMPRESA"
    Set Listado_por_defecto = datos_bd(consulta)
End Function


