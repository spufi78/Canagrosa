VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDOCUMENTOS_SUBTIPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDOCUMENTOS_SUBTIPOS
'***************************************************************
Private ID_SUBTIPO_DOCUMENTO As Long
Private IVA As Long
Private NOMBRE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDOCUMENTOS_SUBTIPOS
'***************************************************************
Public Property Let setID_SUBTIPO_DOCUMENTO(ByVal dato As Long)
        ID_SUBTIPO_DOCUMENTO = dato
End Property
Public Property Let setIVA(ByVal dato As Long)
        IVA = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDOCUMENTOS_SUBTIPOS
'***************************************************************
Public Property Get getID_SUBTIPO_DOCUMENTO() As Long
        getID_SUBTIPO_DOCUMENTO = ID_SUBTIPO_DOCUMENTO
End Property
Public Property Get getIVA() As Long
        getIVA = IVA
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDOCUMENTOS_SUBTIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_SUBTIPOS WHERE ID_SUBTIPO_DOCUMENTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SUBTIPO_DOCUMENTO = rs("ID_SUBTIPO_DOCUMENTO")
                IVA = rs("IVA")
                NOMBRE = rs("NOMBRE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDOCUMENTOS_SUBTIPOS)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_SUBTIPO_DOCUMENTO) FROM DOCUMENTOS_SUBTIPOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SUBTIPO_DOCUMENTO = 1
        Else
                ID_SUBTIPO_DOCUMENTO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO DOCUMENTOS_SUBTIPOS " & _
                   "  VALUES (" & _
                        ID_SUBTIPO_DOCUMENTO & "," & IVA & "," & "'" & NOMBRE & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_SUBTIPO_DOCUMENTO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDOCUMENTOS_SUBTIPOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_SUBTIPOS SET " & _
                        " IVA = " & IVA & "," & _
                        " NOMBRE = '" & NOMBRE & "'" & _
                " WHERE ID_SUBTIPO_DOCUMENTO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDOCUMENTOS_SUBTIPOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS_SUBTIPOS " & _
                "    WHERE ID_SUBTIPO_DOCUMENTO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDOCUMENTOS_SUBTIPOS)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_SUBTIPOS ORDER BY ID_SUBTIPO_DOCUMENTO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_SUBTIPO_DOCUMENTO,NOMBRE FROM DOCUMENTOS_SUBTIPOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
