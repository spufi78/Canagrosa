VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID As Long
Private MUESTRA_ID As Long
Private EMPLEADO_ID As Long
Private tipo As Integer
Private PUESTO As String
Private ESTADO As Integer
Private fecha As String
Private hora As String
Private RUTA_ORIGEN As String
Private RUTA_DESTINO As String

Public Property Let setID(ByVal dato As Long)
    ID = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
    MUESTRA_ID = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
    EMPLEADO_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Integer)
    tipo = dato
End Property
Public Property Let setPUESTO(ByVal dato As String)
    PUESTO = Trim(dato)
End Property
Public Property Let setESTADO(ByVal dato As Integer)
    ESTADO = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = Trim(dato)
End Property
Public Property Let setHORA(ByVal dato As String)
    hora = Trim(dato)
End Property
Public Property Let setRUTA_DESTINO(ByVal dato As String)
    RUTA_DESTINO = Trim(dato)
End Property
Public Property Let setRUTA_ORIGEN(ByVal dato As String)
    RUTA_ORIGEN = Trim(dato)
End Property


Public Property Get getID() As Integer
    getID = ID
End Property
Public Property Get getMUESTRA_ID() As Integer
    getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getEMPLEADO_ID() As Integer
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getTIPO() As Integer
    getTIPO = tipo
End Property
Public Property Get getPUESTO() As String
    getPUESTO = PUESTO
End Property
Public Property Get getESTADO() As Integer
    getESTADO = ESTADO
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getHORA() As String
    getHORA = hora
End Property
Public Property Get getRUTA_DESTINO() As String
    getRUTA_DESTINO = RUTA_DESTINO
End Property
Public Property Get getRUTA_ORIGEN() As String
    getRUTA_ORIGEN = RUTA_ORIGEN
End Property
Public Function Insertar() As Integer
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    consulta = "Insert into impresion " & _
               " values(" & _
               ID & "," & MUESTRA_ID & "," & EMPLEADO_ID & "," & _
               tipo & ",'" & PUESTO & "',0,'" & Format(Date, "dd/mm/yyyy") & "','" & _
               Format(Time, "hh:mm:ss") & "', '" & RUTA_ORIGEN & "', '" & RUTA_DESTINO & "')"
    execute_bd consulta
    Insertar = ID
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la Impresión de la muestra (Insertar)", vbCritical, Err.Description
End Function
Public Function CARGAR(CODIGO As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT * FROM impresion WHERE ID=" & CODIGO
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        CARGAR = False
    Else
        ID = rs("ID")
        MUESTRA_ID = rs("MUESTRA_ID")
        EMPLEADO_ID = rs("EMPLEADO_ID")
        tipo = rs("TIPO")
        PUESTO = rs("PUESTO")
        ESTADO = rs("ESTADO")
        hora = rs("HORA")
        fecha = rs("FECHA")
        RUTA_ORIGEN = rs("RUTA_ORIGEN")
        RUTA_DESTINO = rs("RUTA_DESTINO")
        
        CARGAR = True
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CARGAR = False
'    MsgBox "Error al cargar la muestra para su impresión.", vbCritical, Err.Description
End Function
Public Function Eliminar(CODIGO As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM impresion WHERE ID=" & CODIGO
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Imprimiendo(CODIGO As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE IMPRESION SET ESTADO = 1 WHERE ID=" & CODIGO
    execute_bd consulta
    Imprimiendo = True
    Exit Function
fallo:
    Imprimiendo = False
End Function
Public Function ERROR(CODIGO As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE IMPRESION SET ESTADO = 3 WHERE ID=" & CODIGO
    execute_bd consulta
    ERROR = True
    Exit Function
fallo:
    ERROR = False
End Function
Public Function OFFICE2007(CODIGO As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "UPDATE IMPRESION SET ESTADO = 5 WHERE ID=" & CODIGO
    execute_bd consulta
    OFFICE2007 = True
    Exit Function
fallo:
    OFFICE2007 = False
End Function
Public Function Impreso(CODIGO As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
'    consulta = "UPDATE IMPRESION SET ESTADO = 2 WHERE ID=" & CODIGO
'    execute_bd consulta
'    consulta = "Insert into impresion_historico " & _
'               "select muestra_id,empleado_id,tipo,puesto,2,fecha,hora from impresion where id = " & CODIGO
'    execute_bd consulta
    consulta = "delete from impresion where id = " & CODIGO
    execute_bd consulta
    Impreso = True
    Exit Function
fallo:
    Impreso = False
End Function
Public Function Listado(INFORMES As Boolean) As ADODB.Recordset
    Dim consulta As String
    Dim tipo As String
    If INFORMES = True Then
        tipo = " AND A.TIPO IN (" & TIPOS_DOCUMENTOS_INFORMES & ")"
    Else
        tipo = " AND A.TIPO NOT IN (" & TIPOS_DOCUMENTOS_INFORMES & ")"
    End If
    
    consulta = "SELECT A.MUESTRA_ID,A.TIPO,E.USUARIO,A.PUESTO,A.ESTADO,A.FECHA,A.HORA,E.ID_EMPLEADO,A.ID,A.RUTA_ORIGEN,A.RUTA_DESTINO " & _
        "  FROM IMPRESION A, usuarios AS E " & _
        " WHERE A.EMPLEADO_ID = E.ID_EMPLEADO " & _
        tipo & _
        "  ORDER BY FECHA ASC, HORA ASC"
    Set Listado = datos_bd(consulta)
End Function
Public Function ListadoTodos() As ADODB.Recordset
    Dim consulta As String
    Dim tipo As String
    consulta = "SELECT A.MUESTRA_ID,A.TIPO,E.USUARIO,A.PUESTO,A.ESTADO,A.FECHA,A.HORA,E.ID_EMPLEADO,A.ID,A.RUTA_ORIGEN,A.RUTA_DESTINO " & _
        "  FROM IMPRESION A, usuarios AS E " & _
        " WHERE A.EMPLEADO_ID = E.ID_EMPLEADO " & _
        "  ORDER BY FECHA ASC, HORA ASC"
    Set ListadoTodos = datos_bd(consulta)
End Function
Public Function ListadoPNT() As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT A.MUESTRA_ID,A.TIPO,E.USUARIO,A.PUESTO,A.ESTADO,A.FECHA,A.HORA,E.ID_EMPLEADO,A.ID,A.RUTA_ORIGEN,A.RUTA_DESTINO " & _
        "  FROM IMPRESION A, usuarios AS E " & _
        " WHERE A.EMPLEADO_ID = E.ID_EMPLEADO " & _
        "   AND A.TIPO = 40 " & _
        "  ORDER BY FECHA ASC, HORA ASC"
    Set ListadoPNT = datos_bd(consulta)
End Function
Public Function ListadoID(ID As Long) As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT A.MUESTRA_ID,A.TIPO,E.USUARIO,A.PUESTO,A.ESTADO,A.FECHA,A.HORA,E.ID_EMPLEADO,A.ID,A.RUTA_ORIGEN,A.RUTA_DESTINO " & _
        "  FROM IMPRESION A, usuarios AS E " & _
        " WHERE A.EMPLEADO_ID = E.ID_EMPLEADO " & _
        "   AND A.ID = " & ID & _
        "  ORDER BY FECHA ASC, HORA ASC"
    Set ListadoID = datos_bd(consulta)
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID) FROM IMPRESION"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID = 1
    Else
        ID = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Cargar_alodine(CODIGO As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT * FROM impresion WHERE ESTADO <> 3 AND TIPO = 20 AND MUESTRA_ID=" & CODIGO
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Cargar_alodine = False
    Else
        Cargar_alodine = True
    End If
    Set rs = Nothing
    Exit Function
fallo:
    Cargar_alodine = False
'    MsgBox "Error al cargar la muestra para su impresión.", vbCritical, Err.Description
End Function
Public Function cargar_documento(CODIGO As Long, tipo As Integer) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT * FROM impresion WHERE TIPO = " & tipo & " AND MUESTRA_ID=" & CODIGO
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        cargar_documento = False
    Else
        cargar_documento = True
    End If
    Set rs = Nothing
    Exit Function
fallo:
    cargar_documento = False
'    MsgBox "Error al cargar la muestra para su impresión.", vbCritical, Err.Description
End Function

