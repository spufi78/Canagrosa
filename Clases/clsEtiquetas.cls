VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_ETIQUETA As Long
Private CENTRO_ID As Long
Private TIPO_ID As Integer
Private ID As Long
Private ESTADO As Integer
Private USUARIO_ID As Long
Private TS As String

Public Property Let setID_ETIQUETA(ByVal dato As Long)
    ID_ETIQUETA = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
    CENTRO_ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Integer)
    TIPO_ID = dato
End Property
Public Property Let setID(ByVal dato As Long)
    ID = dato
End Property
Public Property Let setESTADO(ByVal dato As Integer)
    ESTADO = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
    USUARIO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
    TS = Trim(dato)
End Property

Public Property Get getID_ETIQUETA() As Long
    getID_ETIQUETA = ID_ETIQUETA
End Property
Public Property Get getCENTRO_ID() As Integer
    getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getTIPO_ID() As Integer
    getTIPO_ID = TIPO_ID
End Property
Public Property Get getID() As Long
    getID = ID
End Property
Public Property Get getUSUARIO_ID() As Integer
    getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getESTADO() As Integer
    getESTADO = ESTADO
End Property
Public Property Get getTS() As String
    getTS = TS
End Property
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    CrearID
    consulta = "Insert into etiquetas " & _
               " values(" & _
               ID_ETIQUETA & "," & CENTRO_ID & "," & TIPO_ID & "," & ID & ",0," & _
               USUARIO_ID & ",CURRENT_TIMESTAMP)"
    execute_bd consulta
    Insertar = ID_ETIQUETA
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar la Impresión de la muestra (Insertar)", vbCritical, Err.Description
End Function
Public Function CARGAR(ID_ETIQUETA As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT * FROM etiquetas WHERE ID_ETIQUETA=" & ID_ETIQUETA
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        CARGAR = False
    Else
        ID_ETIQUETA = rs("ID_ETIQUETA")
        CENTRO_ID = rs("CENTRO_ID")
        TIPO_ID = rs("TIPO_ID")
        ID = rs("ID")
        ESTADO = rs("ESTADO")
        USUARIO_ID = rs("USUARIO_ID")
        If Not IsNull(rs("TS")) Then
            TS = rs("TS")
        End If
        
        CARGAR = True
    End If
    Set rs = Nothing
    Exit Function
fallo:
    CARGAR = False
'    MsgBox "Error al cargar la muestra para su impresión.", vbCritical, Err.Description
End Function
Public Function Eliminar(ID_ETIQUETA As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "DELETE FROM etiquetas WHERE ID_ETIQUETA=" & ID_ETIQUETA
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function EliminarCentro(CENTRO_ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "DELETE FROM etiquetas WHERE CENTRO_ID=" & CENTRO_ID
    execute_bd consulta
    EliminarCentro = True
    Exit Function
fallo:
    EliminarCentro = False
End Function
Public Function Imprimiendo(ID_ETIQUETA As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE etiquetas SET ESTADO = 1 WHERE ID_ETIQUETA=" & ID_ETIQUETA
    execute_bd consulta
    Imprimiendo = True
    Exit Function
fallo:
    Imprimiendo = False
End Function
Public Function ERROR(ID_ETIQUETA As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE etiquetas SET ESTADO = 3 WHERE ID_ETIQUETA=" & ID_ETIQUETA
    execute_bd consulta
    ERROR = True
    Exit Function
fallo:
    ERROR = False
End Function
Public Function Impreso(ID_ETIQUETA As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "delete from etiquetas where ID_ETIQUETA = " & ID_ETIQUETA
    execute_bd consulta
    Impreso = True
    Exit Function
fallo:
    Impreso = False
End Function
Public Function Listado(Laboratorio As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.ID_ETIQUETA,A.TIPO_ID,A.ID,E.USUARIO,A.ESTADO,A.TS " & _
               "  FROM etiquetas A, usuarios AS E " & _
               " WHERE A.USUARIO_ID = E.ID_EMPLEADO " & _
               "   and A.CENTRO_ID IN (" & Laboratorio & ")" & _
               " ORDER BY A.TS "
    Set Listado = datos_bd(consulta)
End Function
Private Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_ETIQUETA) FROM etiquetas"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_ETIQUETA = 1
    Else
        ID_ETIQUETA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub


