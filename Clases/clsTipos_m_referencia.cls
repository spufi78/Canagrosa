VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_m_referencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_M_REFERENCIA
'***************************************************************
Private ID_TIPO_M_REFERENCIA As Long
Private nombre As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_M_REFERENCIA
'***************************************************************
Public Property Let setID_TIPO_M_REFERENCIA(ByVal dato As Long)
        ID_TIPO_M_REFERENCIA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_M_REFERENCIA
'***************************************************************
Public Property Get getID_TIPO_M_REFERENCIA() As Long
        getID_TIPO_M_REFERENCIA = ID_TIPO_M_REFERENCIA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_M_REFERENCIA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_M_REFERENCIA WHERE ID_TIPO_M_REFERENCIA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_TIPO_M_REFERENCIA = rs("ID_TIPO_M_REFERENCIA")
                nombre = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_m_referencia)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_TIPO_M_REFERENCIA) FROM TIPOS_M_REFERENCIA"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_TIPO_M_REFERENCIA = 1
        Else
                ID_TIPO_M_REFERENCIA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO TIPOS_M_REFERENCIA " & _
                   "  VALUES (" & _
                        ID_TIPO_M_REFERENCIA & "," & "'" & nombre & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_TIPO_M_REFERENCIA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_m_referencia)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_M_REFERENCIA SET " & _
                        " NOMBRE = '" & nombre & "'" & _
                " WHERE ID_TIPO_M_REFERENCIA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTipos_m_referencia)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_M_REFERENCIA " & _
                "    WHERE ID_TIPO_M_REFERENCIA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_m_referencia)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_M_REFERENCIA ORDER BY ID_TIPO_M_REFERENCIA"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_TIPO_M_REFERENCIA,NOMBRE FROM TIPOS_M_REFERENCIA ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
