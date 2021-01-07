VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_Permisividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE unidades           *
'***********************************************************
Private ID_TIPO_PERMISIVIDAD As Integer
Private DESCRIPCION As String
Public Property Let setID_TIPO_PERMISIVIDAD(ByVal dato As Integer)
    ID_TIPO_PERMISIVIDAD = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Get getID_TIPO_PERMISIVIDAD() As Integer
    getID_TIPO_PERMISIVIDAD = ID_TIPO_PERMISIVIDAD
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_PERMISIVIDAD order by DESCRIPCION"
    Set Listado = datos_bd(consulta)
End Function
Public Function Cargar(TIPO_PERMISIVIDAD As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_PERMISIVIDAD WHERE ID_TIPO_PERMISIVIDAD=" & TIPO_PERMISIVIDAD
    Set rs = datos_bd(consulta)
    ID_TIPO_PERMISIVIDAD = rs("ID_TIPO_PERMISIVIDAD")
    DESCRIPCION = rs("DESCRIPCION")
    Cargar = True
    Exit Function
fallo:
    Cargar = False
End Function
Public Sub CrearId()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_PERMISIVIDAD) FROM TIPOS_PERMISIVIDAD"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_PERMISIVIDAD = 1
    Else
        ID_TIPO_PERMISIVIDAD = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearId
    consulta = "Insert into TIPOS_PERMISIVIDAD " & _
               " values(" & _
               ID_TIPO_PERMISIVIDAD & ",'" & DESCRIPCION & "')"
    execute_bd consulta
    Insertar = ID_TIPO_PERMISIVIDAD
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la unidad (InsertarUnidad)", vbCritical, Err.Description
End Function
Public Function Eliminar(TIPO_PERMISIVIDAD As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM TIPOS_PERMISIVIDAD WHERE ID_TIPO_PERMISIVIDAD=" & TIPO_PERMISIVIDAD
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function


