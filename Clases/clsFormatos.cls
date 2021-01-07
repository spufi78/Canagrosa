VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsformatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLIENTE      *
'*****************************************************
Private ID_FORMATO As Integer
Private DESCRIPCION As String
Private PRECINTADO As Integer
Public Property Let setID_FORMATO(ByVal dato As Integer)
    ID_FORMATO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Let setPRECINTADO(ByVal dato As Integer)
    PRECINTADO = dato
End Property
Public Property Get getID_FORMATO() As Integer
    getID_FORMATO = ID_FORMATO
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPRECINTADO() As Integer
    getPRECINTADO = PRECINTADO
End Property

Public Function EsPrecintado(ByVal ID_FORMATO) As Boolean
    Dim CONSULTA As String
    Dim i As Integer
    CONSULTA = "SELECT precintado FROM formatos" & _
             " WHERE id_formato=" & ID_FORMATO
    i = datos_bd(CONSULTA).Fields("precintado")
    EsPrecintado = False
    If i = 1 Then
        EsPrecintado = True
    End If
End Function
Public Function CargarFormato(ByVal formato As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "select * from formatos where id_formato = " & formato
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount = 0 Then
        CargarFormato = False
    Else
        CargarFormato = True
        ID_FORMATO = rs("ID_FORMATO")
        DESCRIPCION = rs("DESCRIPCION")
        PRECINTADO = rs("PRECINTADO")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CargarFormato = False
    MsgBox "Error al cargar el formato (clsFormato)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM formatos order by descripcion"
    Set Listado = datos_bd(CONSULTA)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_FORMATO,DESCRIPCION FROM formatos order by descripcion"
    Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_FORMATO) FROM FORMATOS"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FORMATO = 1
    Else
        ID_FORMATO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim CONSULTA As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        CONSULTA = "Insert into FORMATOS " & _
                   " values(" & _
                   ID_FORMATO & ",'" & DESCRIPCION & "'," & PRECINTADO & ")"
        execute_bd CONSULTA
        Insertar = ID_FORMATO
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar las ID_FORMATO (ID_FORMATO)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "DELETE FROM FORMATOS WHERE ID_FORMATO=" & ID
    execute_bd CONSULTA
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function

Public Function Modificar(formato As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "UPDATE FORMATOS " & _
               " SET DESCRIPCION = '" & DESCRIPCION & "'," & _
               "     PRECINTADO = " & PRECINTADO & _
               " WHERE ID_FORMATO=" & formato
    execute_bd CONSULTA
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_FORMATO FROM FORMATOS WHERE DESCRIPCION = '" & DESCRIPCION & "'"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe un envase con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "FORMATOS"
        .setDESCRIPCION = "Envases"
        .setPK = "ID_FORMATO"
        .setCAMPO = "DESCRIPCION"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function


