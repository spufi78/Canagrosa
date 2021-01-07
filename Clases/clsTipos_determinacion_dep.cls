VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 1  'vbSimpleBound
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_determinacion_dep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private TIPO_DETERMINACION_ID As Integer
Private CAMPO_ID As Integer
Private TIPO_DETERMINACION_ID_DEP As Integer
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Integer)
    TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setCAMPO_ID(ByVal dato As Integer)
    CAMPO_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID_DEP(ByVal dato As Integer)
    TIPO_DETERMINACION_ID_DEP = dato
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Integer
    getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getCAMPO_ID() As Integer
    getCAMPO_ID = CAMPO_ID
End Property
Public Property Get getTIPO_DETERMINACION_ID_DEP() As Integer
    getTIPO_DETERMINACION_ID_DEP = TIPO_DETERMINACION_ID_DEP
End Property
Public Function cargar(ByVal deter As Integer, ByVal campo As Integer) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "select * from tipos_determinacion_dep where tipo_determinacion_id = " & deter & " and campo_id=" & campo
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        cargar = False
    Else
        cargar = True
        TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
        CAMPO_ID = rs("CAMPO_ID")
        TIPO_DETERMINACION_ID_DEP = rs("TIPO_DETERMINACION_ID_DEP")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    cargar = False
    MsgBox "Error al cargar el tipo determinacion dependiente (CargarTipoDeterminacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT b.nombre,d.nombre, c.nombre,a.tipo_determinacion_id,a.campo_id,a.tipo_determinacion_id_dep " & _
               " FROM tipos_determinacion_dep a, tipos_determinacion b, tipos_determinacion c, formulas_campos d " & _
               " Where a.TIPO_DETERMINACION_ID = b.ID_TIPO_DETERMINACION " & _
               "   and a.tipo_determinacion_id_dep = c.id_tipo_determinacion " & _
               "   and a.campo_id = d.id_campo " & _
               " order by a.tipo_determinacion_id "
'    consulta = "SELECT * FROM tipos_determinacion_dep order by tipo_determinacion_id"
    Set Listado = datos_bd(consulta)
End Function
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Insert into tipos_determinacion_dep " & _
               " values(" & _
               TIPO_DETERMINACION_ID & "," & CAMPO_ID & "," & _
               TIPO_DETERMINACION_ID_DEP & ")"
    execute_bd consulta
    Insertar = TIPO_DETERMINACION_ID
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar los tipos_determinacion_dep", vbCritical, Err.Description
End Function

Public Function Eliminar(deter As Long, campo As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "DELETE FROM tipos_determinacion_dep WHERE TIPO_DETERMINACION_ID=" & deter & " AND CAMPO_ID=" & campo
    execute_bd consulta
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Listado_Dependencias(deter As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = "select * from tipos_determinacion_dep where tipo_determinacion_id = " & deter
    Set Listado_Dependencias = datos_bd(consulta)
End Function



