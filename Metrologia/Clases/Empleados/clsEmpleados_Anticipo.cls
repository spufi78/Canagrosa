VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_Anticipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID As Long
Private EMPLEADO_ID As Integer
Private CANTIDAD As Single
Private COMENTARIO As String
Private fecha As String
Public Property Let setID(ByVal dato As Long)
    ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
    EMPLEADO_ID = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Single)
    CANTIDAD = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
    COMENTARIO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = Format(dato, "yyyy-mm-dd")
End Property
Public Property Get getID() As Long
    getID = ID
End Property
Public Property Get getEMPLEADO_ID() As Integer
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getCANTIDAD() As Single
    getCANTIDAD = CANTIDAD
End Property
Public Property Get getCOMENTARIO() As String
    getCOMENTARIO = COMENTARIO
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Function Listado(operario As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_ANTICIPOS WHERE EMPLEADO_ID = " & operario & _
               " ORDER BY ID DESC"
    Set Listado = datos_bd(consulta)
End Function
Public Function cargar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_ANTICIPOS WHERE ID=" & ID
    Set rs = datos_bd(consulta)
    ID = rs("ID")
    EMPLEADO_ID = rs("EMPLEADO_ID")
    CANTIDAD = rs("CANTIDAD")
    COMENTARIO = rs("COMENTARIO")
    fecha = rs("FECHA")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Function Cargar_Ultimo(operario As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_ANTICIPOS WHERE EMPLEADO_ID = " & operario & _
               " ORDER BY ID DESC"
    Set rs = datos_bd(consulta)
    ID = rs("ID")
    EMPLEADO_ID = rs("EMPLEADO_ID")
    CANTIDAD = rs("CANTIDAD")
    COMENTARIO = rs("COMENTARIO")
    fecha = rs("FECHA")
    Cargar_Ultimo = True
    Exit Function
fallo:
    Cargar_Ultimo = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID) FROM EMPLEADOS_ANTICIPOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID = 1
    Else
        ID = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into EMPLEADOS_ANTICIPOS " & _
                   " values(" & ID & "," & _
                   EMPLEADO_ID & ",'" & _
                   CANTIDAD & "','" & _
                   COMENTARIO & "','" & fecha & "')"
    execute_bd consulta
    Insertar = ID
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el anticipo (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM EMPLEADOS_ANTICIPOS WHERE ID=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Totales_mes(operario As Integer, mes As Integer, anno As Integer) As Single
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim fd As Date
    Dim fh As Date
    fd = Format("1/" & mes & "/" & anno, "mm/dd/yyyy")
    fh = Format("1/" & mes + 1 & "/" & anno, "mm/dd/yyyy")
    consulta = "SELECT SUM(CANTIDAD) FROM EMPLEADOS_ANTICIPOS " & _
               " WHERE EMPLEADO_ID = " & operario & _
               "   AND fecha >= (#" & fd & "#) and fecha < (#" & fh & "#)"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Totales_mes = 0
    Else
        Totales_mes = rs.Fields(0)
    End If
    Set rs = Nothing
End Function
Public Function total(operario As Long) As Single
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT SUM(CANTIDAD) FROM EMPLEADOS_ANTICIPOS " & _
               " WHERE EMPLEADO_ID = " & operario
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        total = 0
    Else
        total = rs.Fields(0)
    End If
    Set rs = Nothing
End Function

