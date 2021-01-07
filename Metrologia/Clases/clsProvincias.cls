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
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsPROVINCIAS
'***************************************************************
Private ID_PROVINCIA As Integer
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsPROVINCIAS
'***************************************************************
Public Property Let setID_PROVINCIA(ByVal dato As Integer)
        ID_PROVINCIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsPROVINCIAS
'***************************************************************
Public Property Get getID_PROVINCIA() As Integer
        getID_PROVINCIA = ID_PROVINCIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsPROVINCIAS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVINCIAS WHERE ID_PROVINCIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PROVINCIA = rs("ID_PROVINCIA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPROVINCIAS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO PROVINCIAS " & _
                   "  VALUES (" & _
                        ID_PROVINCIA & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_PROVINCIA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPROVINCIAS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVINCIAS SET " & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_PROVINCIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPROVINCIAS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES WHERE PROVINCIA_ID = " & ID
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "No puede eliminarse la provincia. Esta asignado a " & rs.RecordCount & " cliente/s", vbExclamation, App.Title
        Else
            consulta = "DELETE FROM PROVINCIAS " & _
                    "    WHERE ID_PROVINCIA = " & ID
            execute_bd consulta
        End If
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPROVINCIAS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROVINCIAS ORDER BY NOMBRE"
        Set Listado = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PROVINCIA) FROM PROVINCIAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PROVINCIA = 1
    Else
        ID_PROVINCIA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PROVINCIA, NOMBRE FROM PROVINCIAS " & _
               " order by nombre"
    Set Listado_Combo = datos_bd(consulta)
End Function

