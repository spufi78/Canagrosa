VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_Control"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID As Long
Private EMPLEADO_ID As Integer
Private tipo As String
Private VALOR As String
Private JUSTIFICADA As Integer
Private COMENTARIO As String
Private Fecha As String
Public Property Let setID(ByVal dato As Long)
    ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
    EMPLEADO_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Integer)
    tipo = Trim(dato)
End Property
Public Property Let setVALOR(ByVal dato As String)
    VALOR = dato
End Property
Public Property Let setJUSTIFICADA(ByVal dato As Integer)
    JUSTIFICADA = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
    COMENTARIO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
    Fecha = Format(dato, "yyyy-mm-dd")
End Property
Public Property Get getID() As Long
    getID = ID
End Property
Public Property Get getEMPLEADO_ID() As Integer
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getTIPO() As Integer
    getTIPO = tipo
End Property
Public Property Get getVALOR() As String
    getVALOR = VALOR
End Property
Public Property Get getJUSTIFICADA() As Integer
    getJUSTIFICADA = JUSTIFICADA
End Property
Public Property Get getCOMENTARIO() As String
    getCOMENTARIO = COMENTARIO
End Property
Public Property Get getFECHA() As String
    getFECHA = Fecha
End Property
Public Function Listado(operario As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.ID, A.FECHA,B.NOMBRE,A.VALOR,A.COMENTARIO " & _
               "  FROM EMPLEADOS_CONTROL A,EMPLEADOS_CONTROL_TIPOS B " & _
               " WHERE A.EMPLEADO_ID = " & operario & _
               "   AND A.TIPO = B.ID_CONTROL " & _
               " ORDER BY A.ID DESC"
    Set Listado = datos_bd(consulta)
End Function
Public Function cargar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_CONTROL WHERE ID=" & ID
    Set rs = datos_bd(consulta)
    ID = rs("ID")
    EMPLEADO_ID = rs("EMPLEADO_ID")
    tipo = rs("TIPO")
    VALOR = rs("VALOR")
    JUSTIFICADA = rs("JUSTIFICADA")
    COMENTARIO = rs("COMENTARIO")
    Fecha = rs("FECHA")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID) FROM EMPLEADOS_CONTROL"
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
'    Insertar = 0
'    If duplicado = False Then
        Me.CrearID
        consulta = "Insert into EMPLEADOS_CONTROL " & _
                   " values(" & ID & "," & _
                   EMPLEADO_ID & "," & tipo & ",'" & _
                   VALOR & "'," & JUSTIFICADA & ",'" & _
                   COMENTARIO & "','" & Fecha & "')"
        execute_bd consulta
        Insertar = ID
'    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el expediente (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM EMPLEADOS_CONTROL WHERE ID=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
'Public Function duplicado() As Boolean
'    Dim rs As New ADODB.Recordset
'    Dim consulta As String
'    consulta = "SELECT EMPLEADO_ID FROM EMPLEADOS_CONTROL WHERE TIPO = '" & TIPO & "'"
'    Set rs = datos_bd(consulta)
'    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
'        duplicado = False
'    Else
'        duplicado = True
'    End If
'    Set rs = Nothing
'End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE EMPLEADOS_CONTROL " & _
               " SET TIPO = " & tipo & "," & _
               "     VALOR = '" & VALOR & "'," & _
               "     COMENTARIO = '" & COMENTARIO & "'," & _
               "     FECHA = '" & Fecha & "'," & _
               "     JUSTIFICADA = " & JUSTIFICADA & _
               " WHERE ID=" & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox Err.Description
End Function
Public Function Totales(operario As Long, tipo As Integer) As Single
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT SUM(VALOR) FROM EMPLEADOS_CONTROL " & _
               " WHERE EMPLEADO_ID = " & operario & _
               "   AND TIPO = " & tipo
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Totales = 0
    Else
        Totales = rs.Fields(0)
    End If
    Set rs = Nothing
End Function
Public Function Totales_mes(operario As Long, tipo As Integer, mes As Integer, anno As Integer) As Single
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim fd As Date
    Dim fh As Date
    fd = Format("1/" & mes & "/" & anno, "mm/dd/yyyy")
    fh = Format("1/" & mes + 1 & "/" & anno, "mm/dd/yyyy")
    consulta = "SELECT SUM(VALOR) FROM EMPLEADOS_CONTROL " & _
               " WHERE EMPLEADO_ID = " & operario & _
               "   AND TIPO = " & tipo & _
               "   AND fecha >= (#" & fd & "#) and fecha < (#" & fh & "#)"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Totales_mes = 0
    Else
        Totales_mes = rs.Fields(0)
    End If
    Set rs = Nothing
End Function

Public Function Total_día(operario As Long, tipo As Integer, Fecha As Date) As Single
    Dim consulta As String
    Dim rs As ADODB.Recordset
    If tipo = 4 Then
        consulta = "SELECT VALOR FROM EMPLEADOS_CONTROL " & _
                   " WHERE EMPLEADO_ID = " & operario & _
                   "   AND TIPO = " & tipo & _
                   "   AND FECHA = #" & Format(Fecha, "mm/dd/yyyy") & "#"
    Else
        consulta = "SELECT SUM(VALOR) FROM EMPLEADOS_CONTROL " & _
                   " WHERE EMPLEADO_ID = " & operario & _
                   "   AND TIPO = " & tipo & _
                   "   AND FECHA = #" & Format(Fecha, "mm/dd/yyyy") & "#"
    End If
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Total_día = 0
    Else
        If tipo = 4 Then
            Total_día = 1
        Else
            Total_día = rs.Fields(0)
        End If
    End If
    Set rs = Nothing
End Function

