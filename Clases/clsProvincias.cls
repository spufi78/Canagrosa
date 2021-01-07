VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProvincias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private id_PROVINCIA As Long
Private PAIS_ID As Long
Private nombre As String
Public Property Let setid_PROVINCIA(ByVal dato As Long)
    id_PROVINCIA = dato
End Property
Public Property Let setPAIS_ID(ByVal dato As Long)
    PAIS_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = Trim(dato)
End Property
Public Property Get getid_PROVINCIA() As Long
    getid_PROVINCIA = id_PROVINCIA
End Property
Public Property Get getPAIS_ID() As Long
    getPAIS_ID = PAIS_ID
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property

Public Function CargarProvincia(ByVal id_PROVINCIA) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM PROVINCIAS " & _
             " WHERE id_PROVINCIA=" & id_PROVINCIA
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        id_PROVINCIA = rs("id_PROVINCIA")
        PAIS_ID = rs("PAIS_ID")
        nombre = rs("nombre")
    End If
    Set rs = Nothing
    Exit Function
fallo:
End Function
Public Function Listado(pais As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT nombre,id_Provincia FROM Provincias where pais_id = " & pais & " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PROVINCIA,NOMBRE FROM PROVINCIAS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into PROVINCIAS " & _
               " values(" & _
               id_PROVINCIA & "," & _
               PAIS_ID & ",'" & _
               UCase(nombre) & "')"
    execute_bd consulta
    Insertar = id_PROVINCIA
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la provincia (Insertar)", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(id_PROVINCIA) FROM PROVINCIAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        id_PROVINCIA = 1
    Else
        id_PROVINCIA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function buscar(PROVINCIA As String) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_provincia FROM Provincias where nombre = '" & UCase(PROVINCIA) & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        buscar = 0
    Else
        buscar = rs.Fields(0)
    End If
    Set rs = Nothing
End Function
