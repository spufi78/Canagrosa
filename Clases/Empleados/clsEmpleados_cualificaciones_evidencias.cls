VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_cualificaciones_e"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_EVIDENCIAS
'***************************************************************
Private CUALIFICACION_ID As Long
Private ORDEN As Long
Private ANNO As Integer
Private DESCRIPCION As String
Private ruta As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_EVIDENCIAS
'***************************************************************
Public Property Let setCUALIFICACION_ID(ByVal dato As Long)
        CUALIFICACION_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setANNO(ByVal dato As Integer)
        ANNO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        ruta = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_EVIDENCIAS
'***************************************************************
Public Property Get getCUALIFICACION_ID() As Long
        getCUALIFICACION_ID = CUALIFICACION_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getANNO() As Integer
        getANNO = ANNO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getRUTA() As String
        getRUTA = ruta
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES_EVIDENCIAS
'***************************************************************
Public Function Carga(CUALIFICACION_ID As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM empleados_cualificaciones_evidencias WHERE CUALIFICACION_ID = " & CUALIFICACION_ID & " AND ORDEN = " & ORDEN
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CUALIFICACION_ID = rs("CUALIFICACION_ID")
                ORDEN = rs("ORDEN")
                ANNO = rs("ANNO")
                DESCRIPCION = rs("DESCRIPCION")
                ruta = rs("RUTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_cualificaciones_evidencias)", vbCritical, Err.Description
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        CrearORDEN
        consulta = "INSERT INTO empleados_cualificaciones_evidencias " & _
                   "  VALUES (" & _
                        CUALIFICACION_ID & "," & ORDEN & ",YEAR(CURRENT_DATE)," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & ruta & "'" & _
                ")"
        execute_bd consulta
        Insertar = ORDEN
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_cualificaciones_evidencias)", vbCritical, Err.Description
End Function
Private Sub CrearORDEN()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ORDEN) FROM empleados_cualificaciones_evidencias where CUALIFICACION_ID = " & CUALIFICACION_ID
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ORDEN = 1
    Else
        ORDEN = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Eliminar(CUALIFICACION_ID As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM empleados_cualificaciones_evidencias " & _
                   " WHERE CUALIFICACION_ID = " & CUALIFICACION_ID & _
                   "   AND ORDEN = " & ORDEN
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_cualificaciones_evidencias)", vbCritical, Err.Description
End Function
Public Function Listado(ID_CUALIFICACION As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM empleados_cualificaciones_evidencias " & _
                   " WHERE CUALIFICACION_ID = " & ID_CUALIFICACION & _
                   " ORDER BY CUALIFICACION_ID"
        Set Listado = datos_bd(consulta)
End Function
