VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlanMantenimientoAcciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPlanMantenimientoAcciones
'***************************************************************
Private ID_ACCION As Long
Private FAMILIA_ACC_ID As Long
Private NOMBRE As String
Private T_PREVISTO As Long
Private ID_PLAN_MTO As Long
Private ID_AUX As Integer

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPlanMantenimientoAcciones
'***************************************************************
Public Property Let setID_ACCION(ByVal dato As Long)
        ID_ACCION = dato
End Property
Public Property Let setFAMILIA_ACC_ID(ByVal dato As Long)
        FAMILIA_ACC_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setT_PREVISTO(ByVal dato As String)
        T_PREVISTO = dato
End Property
Public Property Let setID_PLAN_MTO(ByVal dato As Long)
        ID_PLAN_MTO = dato
End Property
Public Property Let setID_AUX(ByVal dato As Integer)
        ID_AUX = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPlanMantenimientoAcciones
'***************************************************************
Public Property Get getID_ACCION() As Long
        getID_ACCION = ID_ACCION
End Property
Public Property Get getFAMILIA_ACC_ID() As Long
        getFAMILIA_ACC_ID = FAMILIA_ACC_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getT_PREVISTO() As Long
        getT_PREVISTO = T_PREVISTO
End Property
Public Property Get getID_PLAN_MTO() As Long
        getID_PLAN_MTO = ID_PLAN_MTO
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPlanMantenimientoAcciones
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM PlanMantenimientoAcciones WHERE ID_ACCION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ACCION = rs("ID_ACCION")
                FAMILIA_ACC_ID = rs("FAMILIA_ACC_ID")
                NOMBRE = rs("NOMBRE")
                T_PREVISTO = rs("T_PREVISTO")
                ID_PLAN_MTO = rs("ID_PLAN_MTO")
                Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        log ("Error al cargar los datos(clsPlanMantenimientoAcciones)")
End Function

Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_ACCION) FROM PlanMantenimientoAcciones"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ACCION = 1
        Else
                ID_ACCION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        
End Function

Public Function Insertar() As String
    On Error GoTo fallo
    Dim consulta As String
        
    Me.CrearID
    If Not duplicado(ID_ACCION) Then
        consulta = "INSERT INTO PlanMantenimientoAcciones " & _
                   "  VALUES (" & _
                        ID_ACCION & ", " & FAMILIA_ACC_ID & ", " & "'" & NOMBRE & "'" & ", " & T_PREVISTO & ", " & _
                        "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        execute_bd consulta
        Insertar = ID_ACCION
    End If
    Exit Function
    
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlanMantenimientoAcciones)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        Modificar = False
        If Not duplicado(ID) Then
            consulta = "UPDATE PlanMantenimientoAcciones SET " & _
                       "       FAMILIA_ACC_ID = " & FAMILIA_ACC_ID & ", " & _
                       "       NOMBRE = '" & NOMBRE & "'" & ", " & _
                       "       T_PREVISTO = " & T_PREVISTO & ", " & _
                       "       CUSERID = '" & USUARIO.getUSUARIO & "', " & _
                       "       FTIMESTP = CURRENT_TIMESTAMP " & _
                       " WHERE ID_ACCION = " & ID
            
            execute_bd consulta
            Modificar = True
        End If
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlanMantenimientoAcciones)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As New ADODB.RecordSet

        ' Se comprueba si hay planes con esa acción
        consulta = "SELECT COUNT(*) AS TOTAL " & _
                   "  FROM EQUIPOS_PLANES_MANTENIMIENTO_DETALLE " & _
                   " WHERE ACCION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs("TOTAL") <> 0 Then
            MsgBox "No se puede eliminar la acción. Forma parte de algún plan de mantenimiento.", vbExclamation, App.Title
            Eliminar = False
            Exit Function
        End If
        
        execute_bd ("DELETE FROM PlanMantenimientoAcciones WHERE ID_ACCION = " & ID)
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlanMantenimientoAcciones)", vbCritical, Err.Description
End Function

Public Function Listado(strNOMBRE As String, strTPrev As String, strFamAcc As String) As ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT EPA.*, A.DESCRIPCION FAMILIA " & _
               "  FROM PlanMantenimientoAcciones EPA, DECODIFICADORA A" & _
               " WHERE EPA.FAMILIA_ACC_ID = A.VALOR AND A.CODIGO = " & decodificadora.EQ_FAMILIAS_ACCIONES_PLANES_MTO & _
               "   AND NOMBRE LIKE '%" & strNOMBRE & "%' " & _
               "   AND T_PREVISTO LIKE '%" & strTPrev & "%' " & _
               "   AND A.DESCRIPCION LIKE '%" & strFamAcc & "%' " & _
               " ORDER BY NOMBRE "
               
    Set Listado = datos_bd(consulta)
End Function

Public Function Listado_completo() As ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT * " & _
               "  FROM PlanMantenimientoAcciones " & _
               " ORDER BY NOMBRE "
    Set Listado_completo = datos_bd(consulta)
End Function
Public Function Listado_por_Familia(ID_familia_acciones As Long) As ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT ID_ACCION, NOMBRE " & _
               "  FROM PlanMantenimientoAcciones " & _
               " WHERE FAMILIA_ACC_ID = " & ID_familia_acciones & _
               " ORDER BY NOMBRE "
    Set Listado_por_Familia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_ACCION, NOMBRE " & _
                   "  FROM PlanMantenimientoAcciones " & _
                   " ORDER BY NOMBRE "
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function duplicado(lngID_ACCION As Long) As Boolean
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT ID_ACCION " & _
               "  FROM PlanMantenimientoAcciones " & _
               " WHERE NOMBRE = '" & NOMBRE & "'" & _
               "   AND FAMILIA_ACC_ID = " & FAMILIA_ACC_ID & _
               "   AND ID_ACCION <> " & lngID_ACCION
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recuperó ninguno
        duplicado = False
    Else
        MsgBox "Existe una acción con el mismo nombre.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function

Public Function llenar_combo(combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "PlanMantenimientoAcciones"
        .setDESCRIPCION = "Acciones"
        .setPK = "ID_ACCION"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function
