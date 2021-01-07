VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMunicipios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private id_municipio As Long
Private PROVINCIA_ID As Long
Private NOMBRE As String
Public Property Let setID_MUNICIPIO(ByVal dato As Long)
    id_municipio = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getID_MUNICIPIO() As Long
    getID_MUNICIPIO = id_municipio
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property

Public Function CargarMunicipio(ByVal id_municipio) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM municipios " & _
             " WHERE id_municipio=" & id_municipio
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        id_municipio = rs("id_municipio")
        PROVINCIA_ID = rs("provincia_id")
        NOMBRE = rs("nombre")
    End If
    Set rs = Nothing
    Exit Function
fallo:
End Function
Public Function Listado(ByVal provincia As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT nombre,id_municipio FROM municipios where provincia_id = " & provincia & " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearId
    consulta = "Insert into municipios " & _
               " values(" & _
               id_municipio & "," & _
               PROVINCIA_ID & ",'" & _
               UCase(NOMBRE) & "')"
    execute_bd consulta
    Insertar = id_municipio
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el municipio (Insertar)", vbCritical, Err.Description
End Function
Public Sub CrearId()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_MUNICIPIO) FROM MUNICIPIOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        id_municipio = 1
    Else
        id_municipio = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function buscar(municipio As String) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_municipio FROM municipios where nombre = '" & UCase(municipio) & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        buscar = 0
    Else
        buscar = rs.Fields(0)
    End If
    Set rs = Nothing
End Function

