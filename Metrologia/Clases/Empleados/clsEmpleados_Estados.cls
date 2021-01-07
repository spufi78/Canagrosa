VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_Estados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_ESTADO As Integer
Private nombre As String
Public Property Let setID_ESTADO(ByVal dato As Integer)
    ID_ESTADO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Get getID_ESTADO() As Integer
    getID_ESTADO = ID_ESTADO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_ESTADOS order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function cargar(estado As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_ESTADOS WHERE ID_ESTADO=" & estado
    Set rs = datos_bd(consulta)
    ID_ESTADO = rs("ID_ESTADO")
    nombre = rs("NOMBRE")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_ESTADO) FROM EMPLEADOS_ESTADOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_ESTADO = 1
    Else
        ID_ESTADO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into EMPLEADOS_ESTADOS " & _
                   " values(" & _
                   ID_ESTADO & ",'" & nombre & "')"
        execute_bd consulta
        Insertar = ID_ESTADO
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
    consulta = "DELETE FROM EMPLEADOS_ESTADOS WHERE ID_ESTADO=" & estado
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_ESTADO FROM EMPLEADOS_ESTADOS WHERE NOMBRE = '" & nombre & "'"
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
    consulta = "UPDATE EMPLEADOS_ESTADOS " & _
               " SET NOMBRE = '" & nombre & "'" & _
               " WHERE ID_ESTADO=" & estado
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function
Public Function Listado_por_defecto() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_ESTADOS order by ID_ESTADO"
    Set Listado_por_defecto = datos_bd(consulta)
End Function
