VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_dato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_TIPO_DATO As Integer
Private nombre As String
Private OBLIGATORIO As Integer
Private UNIDAD_ID As Integer
Public Property Let setID_TIPO_DATO(ByVal dato As Integer)
    ID_TIPO_DATO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setOBLIGATORIO(ByVal dato As Integer)
    OBLIGATORIO = Trim(dato)
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Integer)
    UNIDAD_ID = Trim(dato)
End Property
Public Property Get getID_TIPO_DATO() As Integer
    getID_TIPO_DATO = ID_TIPO_DATO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getOBLIGATORIO() As Integer
    getOBLIGATORIO = OBLIGATORIO
End Property
Public Property Get getUNIDAD_ID() As Integer
    getUNIDAD_ID = UNIDAD_ID
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM tipos_dato order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_DATO) FROM TIPOS_DATO"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_DATO = 1
    Else
        ID_TIPO_DATO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into TIPOS_DATO " & _
                   " values(" & _
                   ID_TIPO_DATO & ",'" & nombre & "'," & OBLIGATORIO & "," & UNIDAD_ID & ")"
        execute_bd consulta
        Insertar = ID_TIPO_DATO
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar las ID_TIPO_DATO (ID_TIPO_DATO)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM TIPOS_DATO WHERE ID_TIPO_DATO=" & ID
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
    consulta = "SELECT * FROM TIPOS_DATO WHERE ID_TIPO_DATO=" & ID
    Set rs = datos_bd(consulta)
    ID_TIPO_DATO = rs("ID_TIPO_DATO")
    nombre = rs("NOMBRE")
    OBLIGATORIO = rs("OBLIGATORIO")
    UNIDAD_ID = rs("UNIDAD_ID")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Function Modificar(TIPO_DATO As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE TIPOS_DATO " & _
               " SET NOMBRE = '" & nombre & "'," & _
               "     OBLIGATORIO = " & OBLIGATORIO & "," & _
               "     UNIDAD_ID = " & UNIDAD_ID & _
               " WHERE ID_TIPO_DATO=" & TIPO_DATO
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_DATO FROM TIPOS_DATO WHERE NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe una frecuencia con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_DATO, NOMBRE FROM TIPOS_DATO " & _
               " ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function
