VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_Expediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID As Long
Private EMPLEADO_ID As Integer
Private TIPO_CONTRATO As String
Private CATEGORIA As String
Private PLUS As Single
Private OTROS_INGRESOS As String
Private PRECIO_HORA_EXTRA As String
Private COMENTARIO As String
Private Fecha As String
Public Property Let setID(ByVal dato As Long)
    ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
    EMPLEADO_ID = dato
End Property
Public Property Let setTIPO_CONTRATO(ByVal dato As String)
    TIPO_CONTRATO = Trim(dato)
End Property
Public Property Let setCATEGORIA(ByVal dato As String)
    CATEGORIA = dato
End Property
Public Property Let setPLUS(ByVal dato As Single)
    PLUS = dato
End Property
Public Property Let setOTROS_INGRESOS(ByVal dato As String)
    OTROS_INGRESOS = dato
End Property
Public Property Let setPRECIO_HORA_EXTRA(ByVal dato As String)
    PRECIO_HORA_EXTRA = dato
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
Public Property Get getTIPO_CONTRATO() As String
    getTIPO_CONTRATO = TIPO_CONTRATO
End Property
Public Property Get getCATEGORIA() As String
    getCATEGORIA = CATEGORIA
End Property
Public Property Get getPLUS() As Single
    getPLUS = PLUS
End Property
Public Property Get getOTROS_INGRESOS() As String
    getOTROS_INGRESOS = OTROS_INGRESOS
End Property
Public Property Get getPRECIO_HORA_EXTRA() As String
    getPRECIO_HORA_EXTRA = PRECIO_HORA_EXTRA
End Property
Public Property Get getCOMENTARIO() As String
    getCOMENTARIO = COMENTARIO
End Property
Public Property Get getFECHA() As String
    getFECHA = Fecha
End Property
Public Function Listado(operario As Long) As ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_EXPEDIENTE WHERE EMPLEADO_ID = " & operario & _
               " ORDER BY FECHA DESC"
    Set Listado = datos_bd(consulta)
End Function
Public Function CARGAR(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_EXPEDIENTE WHERE ID=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID = rs("ID")
        EMPLEADO_ID = rs("EMPLEADO_ID")
        TIPO_CONTRATO = rs("TIPO_CONTRATO")
        CATEGORIA = rs("CATEGORIA")
        PLUS = rs("PLUS")
        OTROS_INGRESOS = rs("OTROS_INGRESOS")
        PRECIO_HORA_EXTRA = rs("PRECIO_HORA_EXTRA")
        COMENTARIO = rs("COMENTARIO")
        Fecha = rs("FECHA")
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
End Function
Public Function Cargar_Ultimo(operario As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_EXPEDIENTE WHERE EMPLEADO_ID = " & operario & _
               " ORDER BY ID DESC"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        ID = rs("ID")
        EMPLEADO_ID = rs("EMPLEADO_ID")
        TIPO_CONTRATO = rs("TIPO_CONTRATO")
        CATEGORIA = rs("CATEGORIA")
        PLUS = rs("PLUS")
        OTROS_INGRESOS = rs("OTROS_INGRESOS")
        PRECIO_HORA_EXTRA = rs("PRECIO_HORA_EXTRA")
        COMENTARIO = rs("COMENTARIO")
        Fecha = rs("FECHA")
        Cargar_Ultimo = True
    Else
        Cargar_Ultimo = False
    End If
    Exit Function
fallo:
    Cargar_Ultimo = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "SELECT MAX(ID) FROM EMPLEADOS_EXPEDIENTE"
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
        consulta = "Insert into EMPLEADOS_EXPEDIENTE " & _
                   " values(" & ID & "," & _
                   EMPLEADO_ID & ",'" & TIPO_CONTRATO & "','" & _
                   CATEGORIA & "','" & PLUS & "','" & _
                   OTROS_INGRESOS & "','" & PRECIO_HORA_EXTRA & "','" & _
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
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "DELETE FROM EMPLEADOS_EXPEDIENTE WHERE ID=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
'Public Function duplicado() As Boolean
'    Dim rs As New ADODB.Recordset
'    Dim consulta As String
'    consulta = "SELECT EMPLEADO_ID FROM EMPLEADOS_EXPEDIENTE WHERE TIPO_CONTRATO = '" & TIPO_CONTRATO & "'"
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
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    consulta = "UPDATE EMPLEADOS_EXPEDIENTE " & _
               " SET TIPO_CONTRATO = '" & TIPO_CONTRATO & "'," & _
               "     CATEGORIA = '" & CATEGORIA & "'," & _
               "     OTROS_INGRESOS = '" & OTROS_INGRESOS & "'," & _
               "     PRECIO_HORA_EXTRA = '" & PRECIO_HORA_EXTRA & "'," & _
               "     COMENTARIO = '" & COMENTARIO & "'," & _
               "     FECHA = '" & Fecha & "'," & _
               "     PLUS = '" & PLUS & "'" & _
               " WHERE ID=" & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox Err.Description
End Function
