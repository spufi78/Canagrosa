VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_SECTOR As Long
Private NOMBRE As String
Private DEPARTAMENTO As String
Public Property Let setID_SECTOR(ByVal dato As Long)
    ID_SECTOR = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setDEPARTAMENTO(ByVal dato As String)
    DEPARTAMENTO = Trim(dato)
End Property
Public Property Get getID_SECTOR() As Long
    getID_SECTOR = ID_SECTOR
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getDEPARTAMENTO() As String
    getDEPARTAMENTO = DEPARTAMENTO
End Property

Public Function CARGAR(ByVal ID_SECTOR) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM SECTORES " & _
             " WHERE ID_SECTOR=" & ID_SECTOR
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        ID_SECTOR = rs("ID_SECTOR")
        NOMBRE = rs("nombre")
        DEPARTAMENTO = rs("DEPARTAMENTO")
    End If
    Set rs = Nothing
    Exit Function
fallo:
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM SECTORES WHERE ID_SECTOR <> 0 order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_SECTOR,NOMBRE FROM SECTORES order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function

Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into SECTORES " & _
                   " values(" & _
                   ID_SECTOR & "," & _
                   "'" & NOMBRE & "'," & _
                   "'" & DEPARTAMENTO & "')"
        execute_bd consulta
        Insertar = ID_SECTOR
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el SECTOR (Insertar)", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_SECTOR) FROM SECTORES"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_SECTOR = 1
    Else
        ID_SECTOR = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Eliminar(SECTOR As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FAMILIAS WHERE SECTOR_ID = " & SECTOR
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        consulta = "DELETE FROM SECTORES WHERE ID_SECTOR=" & SECTOR
        execute_bd consulta
        Eliminar = True
    Else
        MsgBox "Existen familias con este sector. Familia:" & rs("NOMBRE"), vbCritical, App.Title
        Eliminar = False
    End If
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Modificar(SECTOR As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE SECTORES " & _
               " SET NOMBRE = '" & NOMBRE & "'" & _
               "   , DEPARTAMENTO = '" & DEPARTAMENTO & "'" & _
               " WHERE ID_SECTOR=" & SECTOR
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_SECTOR FROM SECTORES WHERE NOMBRE = '" & NOMBRE & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe un sector con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function


