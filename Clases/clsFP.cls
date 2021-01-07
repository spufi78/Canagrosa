VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_FP As Integer
Private nombre As String
Private dias As Integer
Private ccc As String
Public Property Let setID_FP(ByVal dato As Integer)
    ID_FP = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Let setDIAS(ByVal dato As Integer)
    dias = Trim(dato)
End Property
Public Property Let setCCC(ByVal dato As String)
    ccc = Trim(dato)
End Property
Public Property Get getID_FP() As Integer
    getID_FP = ID_FP
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getDIAS() As Integer
    getDIAS = dias
End Property
Public Property Get getCCC() As String
    getCCC = nombre
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FORMAS_PAGO order by NOMBRE"
    Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_FP) FROM FORMAS_PAGO"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FP = 1
    Else
        ID_FP = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into FORMAS_PAGO " & _
               " values(" & _
               ID_FP & ",'" & nombre & "'," & dias & ",'" & ccc & "')"
    execute_bd consulta
    Insertar = ID_FP
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la NOMBRE (Insertar)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE FORMAS_PAGO SET " & _
               " NOMBRE = '" & nombre & "'," & _
               " DIAS = " & dias & "," & _
               " CCC = '" & ccc & "'" & _
               " WHERE ID_FP = " & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar la NOMBRE (Modificar)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM CLIENTES WHERE FP_ID = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        MsgBox "Existen clientes con la forma de pago seleccionada. No se puede eliminar.", vbExclamation, App.Title
        Eliminar = False
        Exit Function
    End If
    consulta = "DELETE FROM FORMAS_PAGO WHERE ID_FP=" & ID
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
    consulta = "SELECT * FROM FORMAS_PAGO WHERE ID_FP=" & ID
    Set rs = datos_bd(consulta)
    ID_FP = rs("ID_FP")
    nombre = rs("NOMBRE")
    dias = rs("DIAS")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FP, NOMBRE FROM FORMAS_PAGO " & _
               " ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setFILTRO = filtro
        .setTABLA = "FORMAS_PAGO"
        .setDESCRIPCION = "Formas de Pago"
        .setPK = "ID_FP"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function


