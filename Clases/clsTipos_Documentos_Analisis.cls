VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_Documentos_Analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***********************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE tipos_especial     *
'***********************************************************
Private ID_TIPO_DOCUMENTO_ANALISIS As Integer
Private NOMBRE As String
Private PLANTILLA As String
Private TIPO_ANALISIS_ID As Integer
Public Property Let setID_TIPO_DOCUMENTO_ANALISIS(ByVal dato As Integer)
    ID_TIPO_DOCUMENTO_ANALISIS = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setPLANTILLA(ByVal dato As String)
    PLANTILLA = Trim(dato)
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Integer)
    TIPO_ANALISIS_ID = Trim(dato)
End Property
Public Property Get getID_TIPO_DOCUMENTO_ANALISIS() As Integer
    getID_TIPO_DOCUMENTO_ANALISIS = ID_TIPO_DOCUMENTO_ANALISIS
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getPLANTILLA() As String
    getPLANTILLA = PLANTILLA
End Property
Public Property Get getTIPO_ANALISIS_ID() As Integer
    getTIPO_ANALISIS_ID = TIPO_ANALISIS_ID
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_DOCUMENTOS_ANALISIS order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Sub CrearId()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_TIPO_DOCUMENTO_ANALISIS) FROM TIPOS_DOCUMENTOS_ANALISIS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_TIPO_DOCUMENTO_ANALISIS = 1
    Else
        ID_TIPO_DOCUMENTO_ANALISIS = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Insertar = 0
    Me.CrearId
    consulta = "Insert into TIPOS_DOCUMENTOS_ANALISIS " & _
               " values(" & _
              ID_TIPO_DOCUMENTO_ANALISIS & ",'" & NOMBRE & "','" & PLANTILLA & "'," & TIPO_ANALISIS_ID & ")"
    execute_bd consulta
    Insertar = ID_TIPO_DOCUMENTO_ANALISIS
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar las ID_TIPO_DOCUMENTO_ANALISIS (ID_TIPO_DOCUMENTO_ANALISIS)", vbCritical, Err.Description
End Function
Public Function Eliminar(id As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM TIPOS_DOCUMENTOS_ANALISIS WHERE ID_TIPO_DOCUMENTO_ANALISIS=" & id
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Cargar(id As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM TIPOS_DOCUMENTOS_ANALISIS WHERE TIPO_ANALISIS_ID=" & id
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Cargar = False
    Else
        ID_TIPO_DOCUMENTO_ANALISIS = rs("ID_TIPO_DOCUMENTO_ANALISIS")
        NOMBRE = rs("NOMBRE")
        PLANTILLA = rs("PLANTILLA")
        TIPO_ANALISIS_ID = rs("TIPO_ANALISIS_ID")
        Cargar = True
    End If
    Exit Function
fallo:
    Cargar = False
End Function
Public Function Modificar(TIPO_DATO As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE TIPOS_DOCUMENTOS_ANALISIS " & _
               " SET NOMBRE = '" & NOMBRE & "'," & _
               "     PLANTILLA = '" & PLANTILLA & "'," & _
               "     TIPO_ANALISIS_ID = " & TIPO_ANALISIS_ID & _
               " WHERE ID_TIPO_DOCUMENTO_ANALISIS=" & TIPO_DATO
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function
