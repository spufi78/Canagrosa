VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_Nominas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID As Long
Private EMPLEADO_ID As Integer
Private mes As Integer
Private anno As Integer
Private TOTAL_DEVENGADO As Single
Private APORTACIONES As Single
Private RETENCION As Single
Private TOTAL_LIQUIDO As Single
Private SS As Single
Private COMENTARIO As String
Public Property Let setID(ByVal dato As Long)
    ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
    EMPLEADO_ID = dato
End Property
Public Property Let setMES(ByVal dato As Integer)
    mes = Trim(dato)
End Property
Public Property Let setANNO(ByVal dato As Integer)
    anno = dato
End Property
Public Property Let setTOTAL_DEVENGADO(ByVal dato As Single)
    TOTAL_DEVENGADO = dato
End Property
Public Property Let setAPORTACIONES(ByVal dato As Single)
    APORTACIONES = dato
End Property
Public Property Let setRETENCION(ByVal dato As Single)
    RETENCION = dato
End Property
Public Property Let setTOTAL_LIQUIDO(ByVal dato As Single)
    TOTAL_LIQUIDO = dato
End Property
Public Property Let setSS(ByVal dato As Single)
    SS = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
    COMENTARIO = dato
End Property
Public Property Get getID() As Long
    getID = ID
End Property
Public Property Get getEMPLEADO_ID() As Integer
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getMES() As Integer
    getMES = mes
End Property
Public Property Get getANNO() As Integer
    getANNO = anno
End Property
Public Property Get getTOTAL_DEVENGADO() As Single
    getTOTAL_DEVENGADO = TOTAL_DEVENGADO
End Property
Public Property Get getAPORTACIONES() As Single
    getAPORTACIONES = APORTACIONES
End Property
Public Property Get getRETENCION() As Single
    getRETENCION = RETENCION
End Property
Public Property Get getTOTAL_LIQUIDO() As Single
    getTOTAL_LIQUIDO = TOTAL_LIQUIDO
End Property
Public Property Get getSS() As Single
    getSS = SS
End Property
Public Property Get getCOMENTARIO() As String
    getCOMENTARIO = COMENTARIO
End Property
Public Function Listado(operario As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_NOMINAS WHERE EMPLEADO_ID = " & operario & _
               " ORDER BY ID DESC"
    Set Listado = datos_bd(consulta)
End Function
Public Function cargar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_NOMINAS WHERE ID=" & ID
    Set rs = datos_bd(consulta)
    ID = rs("ID")
    EMPLEADO_ID = rs("EMPLEADO_ID")
    mes = rs("MES")
    anno = rs("ANNO")
    TOTAL_DEVENGADO = rs("TOTAL_DEVENGADO")
    APORTACIONES = rs("APORTACIONES")
    RETENCION = rs("RETENCION")
    TOTAL_LIQUIDO = rs("TOTAL_LIQUIDO")
    SS = rs("SS")
    COMENTARIO = rs("COMENTARIO")
    cargar = True
    Exit Function
fallo:
    cargar = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID) FROM EMPLEADOS_NOMINAS"
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
    Insertar = duplicado
    If Insertar = 0 Then
        Me.CrearID
        consulta = "Insert into EMPLEADOS_NOMINAS " & _
                   " values(" & ID & "," & _
                   EMPLEADO_ID & "," & mes & "," & _
                   anno & ",'" & TOTAL_DEVENGADO & "','" & _
                   APORTACIONES & "','" & RETENCION & "','" & _
                   TOTAL_LIQUIDO & "','" & SS & "','" & COMENTARIO & "')"
        execute_bd consulta
        Insertar = ID
    Else
        Me.Modificar (Insertar)
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la nomina (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM EMPLEADOS_NOMINAS WHERE ID=" & ID
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function duplicado() As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID FROM EMPLEADOS_NOMINAS " & _
              " WHERE EMPLEADO_ID =" & Me.getEMPLEADO_ID & _
              "   AND MES = " & Me.getMES & _
              "   AND ANNO = " & Me.getANNO
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = 0
    Else
        duplicado = rs(0)
    End If
    Set rs = Nothing
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE EMPLEADOS_NOMINAS " & _
               " SET MES = " & mes & "," & _
               "     ANNO = " & anno & "," & _
               "     APORTACIONES = '" & APORTACIONES & "'," & _
               "     RETENCION = '" & RETENCION & "'," & _
               "     TOTAL_LIQUIDO = '" & TOTAL_LIQUIDO & "'," & _
               "     SS = '" & SS & "'," & _
               "     COMENTARIO = '" & COMENTARIO & "'," & _
               "     TOTAL_DEVENGADO = '" & TOTAL_DEVENGADO & "'" & _
               " WHERE ID=" & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox Err.Description
End Function
Public Function Datos_Nomina(operario As Integer, mes As Integer, anno As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM EMPLEADOS_NOMINAS WHERE EMPLEADO_ID = " & operario & _
               "   AND MES= " & mes & " AND ANNO=" & anno & _
               " ORDER BY ID DESC"
    Set Datos_Nomina = datos_bd(consulta)
End Function
Public Function Liquido(operario As Integer, mes As Integer, anno As Integer) As Currency
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT total_liquido FROM EMPLEADOS_NOMINAS WHERE EMPLEADO_ID = " & operario & _
               "   AND MES= " & mes & " AND ANNO=" & anno & _
               " ORDER BY ID DESC"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        Liquido = Format("0", "currency")
    Else
        Liquido = Format(rs.Fields(0), "currency")
    End If
End Function
