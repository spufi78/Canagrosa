VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDOCUMENTOS_OBSERVACIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDOCUMENTOS_OBSERVACIONES
'***************************************************************
Private DOCUMENTO_ID As Long
Private OBSERVACION As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDOCUMENTOS_OBSERVACIONES
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setOBSERVACION(ByVal dato As String)
        OBSERVACION = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDOCUMENTOS_OBSERVACIONES
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getOBSERVACION() As String
        getOBSERVACION = OBSERVACION
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDOCUMENTOS_OBSERVACIONES
'***************************************************************
Public Function Carga(DOCUMENTO_LO As Long, ORDEN_L As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_OBSERVACIONES " & _
                   " WHERE DOCUMENTO_ID = " & DOCUMENTO_LO & " AND ORDEN = " & ORDEN_L
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                OBSERVACION = rs("OBSERVACION")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDOCUMENTOS_OBSERVACIONES)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DOCUMENTOS_OBSERVACIONES " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & ORDEN & ",'" & OBSERVACION & "'" & _
                ")"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDOCUMENTOS_OBSERVACIONES)", vbCritical, Err.Description
End Function
Public Function Modificar(DOCUMENTO_L As Long, ORDEN_L As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_OBSERVACIONES SET " & _
                        " OBSERVACION = '" & OBSERVACION & "'" & _
                " WHERE DOCUMENTO_ID = " & DOCUMENTO_L & " AND ORDEN = " & ORDEN_L
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDOCUMENTOS_OBSERVACIONES)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS_OBSERVACIONES " & _
                "    WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDOCUMENTOS_OBSERVACIONES)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_OBSERVACIONES ORDER BY DOCUMENTO_ID"
        Set Listado = datos_bd(consulta)
End Function
