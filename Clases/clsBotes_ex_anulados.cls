VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBotes_ex_anulados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_BOTE_EX_ANULADO As Integer
Private TIPO_BOTE_EX_ID As Integer
Private LOTE As String
Private FECHA_PEDIDO As String
Private FECHA_RECEPCION As String
Private FECHA_CADUCIDAD As String
Private OBSERVACIONES As String
Public Property Let setID_BOTE_EX_ANULADO(ByVal dato As Integer)
    ID_BOTE_EX_ANULADO = dato
End Property
Public Property Let setTIPO_BOTE_EX_ID(ByVal dato As Integer)
    TIPO_BOTE_EX_ID = dato
End Property
Public Property Let setLOTE(ByVal dato As String)
    LOTE = dato
End Property
Public Property Let setFECHA_PEDIDO(ByVal dato As String)
    FECHA_PEDIDO = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
    FECHA_RECEPCION = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
    FECHA_CADUCIDAD = Trim(dato)
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Get getID_BOTE_EX_ANULADO() As Integer
    getID_BOTE_EX_ANULADO = ID_BOTE_EX_ANULADO
End Property
Public Property Get getTIPO_BOTE_EX_ID() As Integer
    getTIPO_BOTE_EX_ID = TIPO_BOTE_EX_ID
End Property
Public Property Get getLOTE() As String
    getLOTE = LOTE
End Property
Public Property Get getFECHA_PEDIDO() As String
    getFECHA_PEDIDO = FECHA_PEDIDO
End Property
Public Property Get getFECHA_RECEPCION() As String
    getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getFECHA_CADUCIDAD() As String
    getFECHA_CADUCIDAD = FECHA_CADUCIDAD
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM BOTES_EX_ANULADOS order by ID_BOTE_EX_ANULADO ASC"
    Set Listado = datos_bd(consulta)
End Function
Public Function Cargar(id As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT * FROM BOTES_EX_ANULADOS WHERE ID_BOTE_EX_ANULADO=" & id
    Set rs = datos_bd(consulta)
    ID_BOTE_EX_ANULADO = rs("ID_BOTE_EX_ANULADO")
    TIPO_BOTE_EX_ID = rs("TIPO_BOTE_EX_ID")
    LOTE = rs("LOTE")
    FECHA_PEDIDO = rs("FECHA_PEDIDO")
    FECHA_RECEPCION = rs("FECHA_RECEPCION")
    FECHA_CADUCIDAD = rs("FECHA_CADUCIDAD")
    OBSERVACIONES = rs("OBSERVACIONES")
    Cargar = True
    Exit Function
fallo:
    Cargar = False
    MsgBox "Error al cargar el bote de reactivo.", vbCritical, Err.Description
End Function

Public Sub CrearId()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_BOTE_EX_ANULADO) FROM BOTES_EX_ANULADOS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_BOTE_EX_ANULADO = 1
    Else
        ID_BOTE_EX_ANULADO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearId
    consulta = "Insert into BOTES_EX_ANULADOS " & _
               " values(" & _
               ID_BOTE_EX_ANULADO & "," & _
               TIPO_BOTE_EX_ID & ",'" & _
               LOTE & "','" & _
               FECHA_PEDIDO & "','" & _
               FECHA_RECEPCION & "','" & _
               FECHA_CADUCIDAD & "','" & _
               OBSERVACIONES & "')"
    execute_bd consulta
    Insertar = ID_BOTE_EX_ANULADO
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar el Reactivo (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(id As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM BOTES_EX_ANULADOS WHERE ID_BOTE_EX_ANULADO=" & id
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al eliminar el Bote de Reactivo (Eliminar)", vbCritical, Err.Description
End Function
Public Function Modificar(id As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE BOTES_EX_ANULADOS SET " & _
               " TIPO_BOTE_EX_ID = " & TIPO_BOTE_EX_ID & "," & _
               " LOTE = '" & LOTE & "'," & _
               " FECHA_PEDIDO = '" & FECHA_PEDIDO & "'," & _
               " FECHA_RECEPCION = '" & FECHA_RECEPCION & "'," & _
               " FECHA_CADUCIDAD = '" & FECHA_CADUCIDAD & "'," & _
               " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
               " WHERE ID_BOTE_EX_ANULADO = " & id
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar el Reactivo (Modificar)", vbCritical, Err.Description
End Function
