VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_SOLUCION As Integer
Private NOMBRE As String
Private FACTURA_DETERMINACIONES As Integer
Private ANULADA As Integer

Public Property Let setID_SOLUCION(ByVal dato As Integer)
    ID_SOLUCION = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setFACTURA_DETERMINACIONES(ByVal dato As Integer)
    FACTURA_DETERMINACIONES = dato
End Property
Public Property Let setANULADA(ByVal dato As Integer)
    ANULADA = dato
End Property
Public Property Get getID_SOLUCION() As Integer
    getID_SOLUCION = ID_SOLUCION
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getFACTURA_DETERMINACIONES() As Integer
    getFACTURA_DETERMINACIONES = FACTURA_DETERMINACIONES
End Property
Public Property Get getANULADA() As Integer
    getANULADA = ANULADA
End Property
Public Function Listado(filtro As String) As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM SOLUCIONES " & _
               " WHERE NOMBRE LIKE '%" & filtro & "%'" & _
               "   AND ANULADA = 0 " & _
               " order by nombre"
    Set Listado = datos_bd(CONSULTA)
End Function
Public Function CARGAR(solucion As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT * FROM SOLUCIONES WHERE ID_SOLUCION=" & solucion
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount > 0 Then
        ID_SOLUCION = rs("ID_SOLUCION")
        NOMBRE = rs("NOMBRE")
        FACTURA_DETERMINACIONES = rs("FACTURA_DETERMINACIONES")
        ANULADA = rs("ANULADA")
        CARGAR = True
    Else
        CARGAR = False
    End If
    Exit Function
fallo:
    CARGAR = False
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT MAX(ID_SOLUCION) FROM SOLUCIONES"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_SOLUCION = 1
    Else
        ID_SOLUCION = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim CONSULTA As String
    Insertar = 0
    FACTURA_DETERMINACIONES = 0
    If duplicado = False Then
        Dim oTA As New clsTipos_analisis
        Dim ANALISIS As Integer
        oTA.setNOMBRE = NOMBRE
        oTA.setFACTURA_DETERMINACIONES = FACTURA_DETERMINACIONES
        ANALISIS = oTA.Insertar
        If ANALISIS > 0 Then
               CONSULTA = "Insert into SOLUCIONES " & _
                           " values(" & _
                           ANALISIS & ",'" & NOMBRE & "'," & FACTURA_DETERMINACIONES & ",0)"
               execute_bd CONSULTA
        End If
        Insertar = ANALISIS
    End If
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar las SOLUCIONES (Insertar)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    ' Miramos si hay baños asociados
    CONSULTA = "select id_bano from banos where solucion_id =" & ID
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount > 0 Then
'    If (Not IsNull(rs.Fields(0))) Or (Not (rs.EOF And rs.BOF)) Then  'si es nulo No se recupero ninguno
        MsgBox "No se puede eliminar la solución, existen baños con dicha solución asociada.", vbCritical, App.Title
        Exit Function
    End If
    ' PODER ANULAR LA SOLUCION AUNQUE EXISTAN MUESTRAS
'    consulta = "select max(id_muestra) from muestras where tipo_analisis_id =" & ID
'    Set rs = datos_bd(consulta)
'    If Not IsNull(rs.Fields(0)) Or Not (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
'        MsgBox "No se puede eliminar la solución, existen muestras con dicha solución asociada.", vbCritical, App.Title
'        Exit Function
'    End If
    ' Eliminamos el tipo de análisis asociado a la solución
'    consulta = "DELETE FROM TIPOS_ANALISIS WHERE ID_TIPO_ANALISIS=" & ID
    CONSULTA = "UPDATE TIPOS_ANALISIS SET ANULADO = 1 WHERE ID_TIPO_ANALISIS=" & ID
    execute_bd CONSULTA
    ' Eliminamos la solución
'    consulta = "DELETE FROM SOLUCIONES WHERE ID_SOLUCION=" & ID
    CONSULTA = "UPDATE SOLUCIONES SET ANULADA = 1 WHERE ID_SOLUCION=" & ID
    execute_bd CONSULTA
    MsgBox "Solución anulada correctamente.", vbInformation, App.Title
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Modificar(solucion As Integer) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
'    consulta = "UPDATE TIPOS_ANALISIS " & _
'               "   SET NOMBRE = '" & NOMBRE & "'," & _
'               "     FACTURA_DETERMINACIONES = " & FACTURA_DETERMINACIONES & _
'               " WHERE ID_TIPO_ANALISIS=" & solucion
'    execute_bd consulta
    CONSULTA = "UPDATE SOLUCIONES " & _
               " SET NOMBRE = '" & NOMBRE & "'," & _
               "     FACTURA_DETERMINACIONES = " & FACTURA_DETERMINACIONES & _
               " WHERE ID_SOLUCION=" & solucion
    execute_bd CONSULTA
    Modificar = True
    Exit Function
fallo:
    Modificar = False
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_SOLUCION FROM SOLUCIONES WHERE NOMBRE = '" & NOMBRE & "'"
    Set rs = datos_bd(CONSULTA)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        duplicado = False
    Else
        MsgBox "Existe una solución con la misma descripción.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function

Public Function Listado_Combo() As ADODB.Recordset
    Dim CONSULTA As String
    CONSULTA = "SELECT ID_SOLUCION,NOMBRE FROM SOLUCIONES order by nombre"
    Set Listado_Combo = datos_bd(CONSULTA)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "SOLUCIONES"
        .setDESCRIPCION = "Soluciones de Proceso"
        .setPK = "ID_SOLUCION"
        .setCAMPO = "NOMBRE"
        .setFILTRO = " ANULADA = 0 "
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
