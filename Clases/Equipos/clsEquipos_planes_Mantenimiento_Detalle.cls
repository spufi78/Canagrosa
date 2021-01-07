VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_PM_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsEquipos_PM_Detalle
'***************************************************************
Private plan_mto_id As Long
Private ACCION_ID As Long
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsEquipos_PM_Detalle
'***************************************************************
Public Property Let setPLAN_MTO_ID(ByVal dato As Long)
        plan_mto_id = dato
End Property
Public Property Let setACCION_ID(ByVal dato As Long)
        ACCION_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsEquipos_PM_Detalle
'***************************************************************
Public Property Get getPLAN_MTO_ID() As Long
        getPLAN_MTO_ID = plan_mto_id
End Property
Public Property Get getACCION_ID() As Long
        getACCION_ID = ACCION_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsEquipos_PM_Detalle
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * FROM eq_planes_mantenimiento_DETALLE WHERE PLAN_MTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                plan_mto_id = rs("PLAN_MTO_ID")
                ACCION_ID = rs("ACCION_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_PM_Detalle)", vbCritical, Err.Description
End Function
Public Function obtener_max_orden() As Long
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT MAX(ORDEN) " & _
               "  FROM eq_planes_mantenimiento_DETALLE " & _
               " WHERE PLAN_MTO_ID = " & plan_mto_id
    
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        obtener_max_orden = 1
    Else
        obtener_max_orden = CLng(rs.Fields(0)) + 1
    End If
    Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO eq_planes_mantenimiento_DETALLE " & _
                   "VALUES (" & _
                           plan_mto_id & "," & ACCION_ID & "," & ORDEN & ", " & _
                           "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        execute_bd consulta
        Insertar = plan_mto_id
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_PM_Detalle)", vbCritical, Err.Description
End Function

'Public Function Modificar(ID As Long) As Boolean
'        On Error GoTo fallo
'        Dim consulta As String
'        consulta = "UPDATE EQUIPOS_PLAN_MANTENIMIENTO_DETALLE SET " & _
'                   "       NOMBRE = '" & NOMBRE & "', " & _
'                   "       CUSERID = '" & USUARIO.getUSUARIO & "', " & _
'                   "       FTIMESTP = CURRENT_TIMESTAMP " & _
'                   " WHERE PLAN_MTO_ID = " & ID
'        execute_bd consulta
'        Modificar = True
'        Exit Function
'fallo:
'        Modificar = False
'        MsgBox "Error al Modificar (clsEquipos_PM_Detalle)", vbCritical, Err.Description
'End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM eq_planes_mantenimiento_DETALLE " & _
                   " WHERE PLAN_MTO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_PM_Detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar_Accion(lngPlanID As Long, lngAccionID As Long, lngOrden As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM eq_planes_mantenimiento_DETALLE " & _
                   " WHERE PLAN_MTO_ID = " & lngPlanID & _
                   "   AND ACCION_ID = " & lngAccionID & _
                   "   AND ORDEN = " & lngOrden
        execute_bd consulta
        Eliminar_Accion = True
        Exit Function
fallo:
        Eliminar_Accion = False
        MsgBox "Error al Eliminar (clsEquipos_PM_Detalle)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ACC.ID_ACCION, ACC.NOMBRE, ACC.T_PREVISTO, DET.ORDEN, D.DESCRIPCION FAMILIA " & _
                   "  FROM eq_planes_mantenimiento_DETALLE DET, eq_planes_acciones ACC, " & _
                   "       decodificadora D " & _
                   " WHERE ACC.FAMILIA_ACC_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.EQ_FAMILIAS_ACCIONES_PLANES_MTO & _
                   "   AND DET.ACCION_ID = ACC.ID_ACCION " & _
                   "   AND PLAN_MTO_ID = " & ID & _
                   " ORDER BY ORDEN "
        Set Listado = datos_bd(consulta)
End Function
