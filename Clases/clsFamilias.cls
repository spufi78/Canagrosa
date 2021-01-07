VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_FAMILIA As Long
Private SECTOR_ID As Long
Private nombre As String
Private CC As String
Private CODIGO_CONTAPLUS As String
Private pedido As Integer
Public Property Let setID_FAMILIA(ByVal dato As Long)
    ID_FAMILIA = dato
End Property
Public Property Let setSECTOR_ID(ByVal dato As Long)
    SECTOR_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setCC(ByVal dato As String)
    CC = dato
End Property
Public Property Let setCODIGO_CONTAPLUS(ByVal dato As String)
    CODIGO_CONTAPLUS = dato
End Property
Public Property Let setPEDIDO(ByVal dato As Integer)
    pedido = dato
End Property

Public Property Get getID_FAMILIA() As Long
    getID_FAMILIA = ID_FAMILIA
End Property
Public Property Get getSECTOR_ID() As Long
    getSECTOR_ID = SECTOR_ID
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getCC() As String
    getCC = CC
End Property
Public Property Get getCODIGO_CONTAPLUS() As String
    getCODIGO_CONTAPLUS = CODIGO_CONTAPLUS
End Property
Public Property Get getPEDIDO() As Integer
    getPEDIDO = pedido
End Property

Public Function cargar(ByVal ID_FAMILIA As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FAMILIAS " & _
             " WHERE ID_FAMILIA=" & ID_FAMILIA
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        ID_FAMILIA = rs("ID_FAMILIA")
        SECTOR_ID = rs("SECTOR_ID")
        nombre = rs("nombre")
        CC = rs("CC")
        CODIGO_CONTAPLUS = rs("CODIGO_CONTAPLUS")
        pedido = rs("PEDIDO")
    End If
    Set rs = Nothing
    Exit Function
fallo:
End Function
Public Function Listado(SECTOR As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FAMILIAS where SECTOR_ID = " & SECTOR & " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function ListadoPedido() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT nombre,id_familia FROM FAMILIAS where PEDIDO = 1 order by nombre "
    Set ListadoPedido = datos_bd(consulta)
End Function

Public Function Listado_completo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FAMILIAS WHERE ID_FAMILIA <> 0 order by nombre"
    Set Listado_completo = datos_bd(consulta)
End Function
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into FAMILIAS " & _
                   " values(" & _
                   ID_FAMILIA & "," & _
                   SECTOR_ID & ",'" & _
                   nombre & "','" & CC & "','" & CODIGO_CONTAPLUS & "'," & pedido & ")"
        execute_bd consulta
        Insertar = ID_FAMILIA
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la familia (Insertar)", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_FAMILIA) FROM FAMILIAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FAMILIA = 1
    Else
        ID_FAMILIA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Eliminar(familia As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_MUESTRA WHERE FAMILIA_ID = " & familia
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        consulta = "DELETE FROM FAMILIAS WHERE ID_FAMILIA=" & familia
        execute_bd consulta
        Eliminar = True
    Else
        MsgBox "No se puede eliminar la familia, existen tipos muestras con esta familia asignada. T.M.:" & rs("NOMBRE"), vbCritical, App.Title
        Eliminar = False
    End If
    Exit Function
fallo:
    Eliminar = False
End Function

Public Function Modificar(ID_FAMILIA As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE FAMILIAS " & _
               " SET NOMBRE = '" & nombre & "'," & _
               "     CC = '" & CC & "'," & _
               "     CODIGO_CONTAPLUS = '" & CODIGO_CONTAPLUS & "'," & _
               "     PEDIDO = " & pedido & "," & _
               "     SECTOR_ID = " & SECTOR_ID & _
               " WHERE ID_FAMILIA=" & ID_FAMILIA
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FAMILIA FROM FAMILIAS WHERE NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe una familia con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "FAMILIAS"
        .setDESCRIPCION = "Familias (Cuentas contables)"
        .setPK = "ID_FAMILIA"
        .setCAMPO = "NOMBRE"
        .setFILTRO = filtro
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FAMILIA,NOMBRE FROM FAMILIAS ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function
