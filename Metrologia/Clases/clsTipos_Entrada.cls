VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_Entrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsTIPOS_ENTRADA
'***************************************************************
Private ID_TIPO_ENTRADA As Integer
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsTIPOS_ENTRADA
'***************************************************************
Public Property Let setID_TIPO_ENTRADA(ByVal dato As Integer)
        ID_TIPO_ENTRADA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsTIPOS_ENTRADA
'***************************************************************
Public Property Get getID_TIPO_ENTRADA() As Integer
        getID_TIPO_ENTRADA = ID_TIPO_ENTRADA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsTIPOS_ENTRADA
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_ENTRADA WHERE ID_TIPO_ENTRADA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TIPO_ENTRADA = rs("ID_TIPO_ENTRADA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTIPOS_ENTRADA)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO TIPOS_ENTRADA " & _
                   "  VALUES (" & _
                        ID_TIPO_ENTRADA & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_TIPO_ENTRADA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTIPOS_ENTRADA)", vbCritical, Err.Description
End Function
Public Function modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_ENTRADA SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_TIPO_ENTRADA = " & ID
        execute_bd consulta
        modificar = True
        Exit Function
fallo:
        modificar = False
        MsgBox "Error al Modificar (clsTIPOS_ENTRADA)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_ENTRADA " & _
                "    WHERE ID_TIPO_ENTRADA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTIPOS_ENTRADA)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_ENTRADA ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_ENTRADA) FROM TIPOS_ENTRADA"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_ENTRADA = 1
    Else
        ID_TIPO_ENTRADA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_ENTRADA, NOMBRE FROM TIPOS_ENTRADA " & _
               " order by ID_TIPO_ENTRADA"
    Set Listado_Combo = datos_bd(consulta)
End Function



