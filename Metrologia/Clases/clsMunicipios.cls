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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private ID_MUNICIPIO As Long
Private PROVINCIA_ID As Long
Private NOMBRE As String
Public Property Let setID_MUNICIPIO(ByVal dato As Long)
    ID_MUNICIPIO = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
    PROVINCIA_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getID_MUNICIPIO() As Long
    getID_MUNICIPIO = ID_MUNICIPIO
End Property
Public Property Get getPROVINCIA_ID() As Long
    getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property

Public Function Cargar(ByVal ID_MUNICIPIO) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM MUNICIPIOS " & _
               " WHERE ID_MUNICIPIO=" & ID_MUNICIPIO
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        ID_MUNICIPIO = rs("ID_MUNICIPIO")
        PROVINCIA_ID = rs("provincia_id")
        NOMBRE = rs("nombre")
        Cargar = True
    Else
        Cargar = False
    End If
    Set rs = Nothing
    Exit Function
fallo:
End Function
Public Function Listado(ByVal PROVINCIA As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT NOMBRE,ID_MUNICIPIO FROM municipios where provincia_id = " & PROVINCIA & " order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function Modificar(ByVal Municipio As Long) As Boolean
   On Error GoTo Modificar_Error

    execute_bd ("update municipios set nombre = '" & NOMBRE & "' where id_municipio = " & Municipio)
    Modificar = True
    Exit Function

   On Error GoTo 0
   Exit Function

Modificar_Error:
    Modificar = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificar of Módulo de clase clsMunicipios", vbCritical, App.Title
End Function
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into municipios " & _
               " values(" & _
               ID_MUNICIPIO & "," & _
               PROVINCIA_ID & ",'" & _
               UCase(NOMBRE) & "')"
    execute_bd consulta
    Insertar = ID_MUNICIPIO
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el municipio (Insertar)", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_MUNICIPIO) FROM MUNICIPIOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_MUNICIPIO = 1
    Else
        ID_MUNICIPIO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function buscar(Municipio As String) As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_MUNICIPIO FROM MUNICIPIOS where nombre = '" & UCase(Municipio) & "'"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        buscar = 0
    Else
        buscar = rs.Fields(0)
    End If
    Set rs = Nothing
End Function
Public Function Listado_Combo_FK(pk As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_MUNICIPIO, NOMBRE FROM MUNICIPIOS " & _
               " WHERE PROVINCIA_ID = " & pk & _
               " ORDER BY NOMBRE"
    Set Listado_Combo_FK = datos_bd(consulta)
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES WHERE MUNICIPIO_ID = " & ID
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No puede eliminarse el municipio. Esta asignado a " & rs.RecordCount & " cliente/s", vbExclamation, App.Title
        Else
            consulta = "DELETE FROM MUNICIPIOS WHERE ID_MUNICIPIO = " & ID
            execute_bd consulta
        End If
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (MUNICIPIOS)", vbCritical, Err.Description
End Function

