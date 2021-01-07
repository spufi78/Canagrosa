VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentos_tipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE DOCUMENTOS_TIPOS           *
'***********************************************************
Private ID_TIPO_DOCUMENTO As Integer
Private NOMBRE As String
Public Property Let setID_TIPO_DOCUMENTO(ByVal dato As Integer)
    ID_TIPO_DOCUMENTO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Get getID_TIPO_DOCUMENTO() As Integer
    getID_TIPO_DOCUMENTO = ID_TIPO_DOCUMENTO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM DOCUMENTOS_TIPOS order by ID_TIPO_DOCUMENTO"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_TIPO_DOCUMENTO, NOMBRE FROM DOCUMENTOS_TIPOS " & _
               " order by ID_TIPO_DOCUMENTO"
    Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Cargar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM DOCUMENTOS_TIPOS WHERE ID_TIPO_DOCUMENTO=" & ID
    Set rs = datos_bd(consulta)
    ID_TIPO_DOCUMENTO = rs("ID_TIPO_DOCUMENTO")
    NOMBRE = rs("NOMBRE")
    Cargar = True
    Exit Function
fallo:
    Cargar = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_DOCUMENTO) FROM DOCUMENTOS_TIPOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_DOCUMENTO = 1
    Else
        ID_TIPO_DOCUMENTO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into DOCUMENTOS_TIPOS " & _
               " values(" & _
               ID_TIPO_DOCUMENTO & ",'" & NOMBRE & "')"
    execute_bd consulta
    Insertar = ID_TIPO_DOCUMENTO
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el COLOR (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM DOCUMENTOS_TIPOS WHERE ID_TIPO_DOCUMENTO=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
