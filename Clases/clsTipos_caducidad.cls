VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_caducidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_CADUCIDAD As Integer
Private CADUCIDAD As String
Private dias As Integer
Public Property Let setID_TIPO_CADUCIDAD(ByVal dato As Integer)
    ID_TIPO_CADUCIDAD = dato
End Property
Public Property Let setCADUCIDAD(ByVal dato As String)
    CADUCIDAD = Trim(dato)
End Property
Public Property Let setDIAS(ByVal dato As Integer)
    dias = Trim(dato)
End Property
Public Property Get getID_TIPO_CADUCIDAD() As Integer
    getID_TIPO_CADUCIDAD = ID_TIPO_CADUCIDAD
End Property
Public Property Get getCADUCIDAD() As String
    getCADUCIDAD = CADUCIDAD
End Property
Public Property Get getDIAS() As Integer
    getDIAS = dias
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_CADUCIDAD order by CADUCIDAD"
    Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_CADUCIDAD) FROM TIPOS_CADUCIDAD"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_CADUCIDAD = 1
    Else
        ID_TIPO_CADUCIDAD = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into TIPOS_CADUCIDAD " & _
               " values(" & _
               ID_TIPO_CADUCIDAD & ",'" & CADUCIDAD & "'," & dias & ")"
    execute_bd consulta
    Insertar = ID_TIPO_CADUCIDAD
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la caducidad (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM TIPOS_CADUCIDAD WHERE ID_TIPO_CADUCIDAD=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function cargar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_CADUCIDAD WHERE ID_TIPO_CADUCIDAD=" & ID
    Set rs = datos_bd(consulta)
    ID_TIPO_CADUCIDAD = rs("ID_TIPO_CADUCIDAD")
    CADUCIDAD = rs("CADUCIDAD")
    dias = rs("DIAS")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_CADUCIDAD, CADUCIDAD FROM TIPOS_CADUCIDAD " & _
               " ORDER BY CADUCIDAD"
    Set Listado_Combo = datos_bd(consulta)
End Function


