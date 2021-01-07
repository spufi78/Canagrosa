VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private id_pais As Long
Private NOMBRE As String
Public Property Let setid_pais(ByVal dato As Long)
    id_pais = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getid_pais() As Long
    getid_pais = id_pais
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property

Public Function CargarPais(ByVal id_pais) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM paises " & _
             " WHERE id_pais=" & id_pais
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        id_pais = rs("id_pais")
        NOMBRE = rs("nombre")
    End If
    Set rs = Nothing
    Exit Function
fallo:
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT nombre,id_pais FROM paises order by nombre"
    Set Listado = datos_bd(consulta)
End Function

Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearId
    consulta = "Insert into paises " & _
               " values(" & _
               id_pais & ",'" & _
               UCase(NOMBRE) & "')"
    execute_bd consulta
    Insertar = id_pais
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el pais (Insertar)", vbCritical, Err.Description
End Function
Public Sub CrearId()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PAIS) FROM PAISES"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        id_pais = 1
    Else
        id_pais = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function buscar(pais As String) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_pais FROM paises where nombre = '" & UCase(pais) & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        buscar = 0
    Else
        buscar = rs.Fields(0)
    End If
    Set rs = Nothing
End Function



