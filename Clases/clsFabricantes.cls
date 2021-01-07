VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFabricantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE unidades           *
'***********************************************************
Private ID_FABRICANTE As Integer
Private nombre As String
Public Property Let setID_FABRICANTE(ByVal dato As Integer)
    ID_FABRICANTE = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Get getID_FABRICANTE() As Integer
    getID_FABRICANTE = ID_FABRICANTE
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FABRICANTES order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function cargar(unidad As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM FABRICANTES WHERE ID_FABRICANTE=" & unidad
    Set rs = datos_bd(consulta)
    ID_FABRICANTE = rs("ID_FABRICANTE")
    nombre = rs("NOMBRE")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_FABRICANTE) FROM FABRICANTES"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_FABRICANTE = 1
    Else
        ID_FABRICANTE = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into FABRICANTES " & _
                   " values(" & _
                   ID_FABRICANTE & ",'" & nombre & "')"
        execute_bd consulta
        Insertar = ID_FABRICANTE
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el fabricante (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT FABRICANTE_ID FROM TIPOS_BOTE_EX WHERE FABRICANTE_ID = " & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        MsgBox "Existen botes con dicho fabricante. No se puede eliminar.", vbInformation, App.Title
        Eliminar = True
        Exit Function
    End If
    consulta = "DELETE FROM FABRICANTES WHERE ID_FABRICANTE=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function duplicado() As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_FABRICANTE FROM FABRICANTES WHERE NOMBRE = '" & nombre & "'"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        duplicado = False
    Else
        ID_FABRICANTE = rs(0)
        duplicado = True
    End If
    Exit Function
fallo:
    duplicado = False
End Function
Public Function Modificar(FABRICANTE As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE FABRICANTES " & _
               " SET NOMBRE = '" & nombre & "'" & _
               " WHERE ID_FABRICANTE=" & FABRICANTE
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function


