VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcesos_base_normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsProcesos_base_normas
'***************************************************************
Private PROCESO_BASE_ID As Long
Private NORMA_ID As Long
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsProcesos_base_normas
'***************************************************************
Public Property Let setPROCESO_BASE_ID(ByVal dato As Long)
        PROCESO_BASE_ID = dato
End Property
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsProcesos_base_normas
'***************************************************************
Public Property Get getPROCESO_BASE_ID() As Long
        getPROCESO_BASE_ID = PROCESO_BASE_ID
End Property
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsProcesos_base_normas
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PROCESOS_BASE_NORMAS WHERE PROCESO_BASE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PROCESO_BASE_ID = rs("PROCESO_BASE_ID")
                NORMA_ID = rs("NORMA_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProcesos_base_normas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PROCESOS_BASE_NORMAS " & _
                   "  VALUES (" & _
                        PROCESO_BASE_ID & "," & NORMA_ID & "," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = PROCESO_BASE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProcesos_base_normas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROCESOS_BASE_NORMAS " & _
                "    WHERE PROCESO_BASE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProcesos_base_normas)", vbCritical, Err.Description
End Function
Public Function Listado(TD As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.NORMA_ID,B.NOMBRE,B.CODIGO,B.EDICION " & _
                   " FROM PROCESOS_BASE_NORMAS A, CA_NORMAS B " & _
                   " WHERE A.PROCESO_BASE_ID = " & TD & _
                   "   AND A.NORMA_ID = B.ID_NORMA " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function


