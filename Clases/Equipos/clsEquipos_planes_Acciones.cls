VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_planes_Acciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_PLANES_ACCIONES
'***************************************************************
Private ID_ACCION As Long
Private FAMILIA_ACC_ID As Long
Private nombre As String
Private T_PREVISTO As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_PLANES_ACCIONES
'***************************************************************
Public Property Let setID_ACCION(ByVal dato As Long)
        ID_ACCION = dato
End Property
Public Property Let setFAMILIA_ACC_ID(ByVal dato As Long)
        FAMILIA_ACC_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setT_PREVISTO(ByVal dato As String)
        T_PREVISTO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_PLANES_ACCIONES
'***************************************************************
Public Property Get getID_ACCION() As Long
        getID_ACCION = ID_ACCION
End Property
Public Property Get getFAMILIA_ACC_ID() As Long
        getFAMILIA_ACC_ID = FAMILIA_ACC_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getT_PREVISTO() As Long
        getT_PREVISTO = T_PREVISTO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_PLANES_ACCIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * FROM eq_planes_acciones WHERE ID_ACCION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ACCION = rs("ID_ACCION")
                FAMILIA_ACC_ID = rs("FAMILIA_ACC_ID")
                nombre = rs("NOMBRE")
                T_PREVISTO = rs("T_PREVISTO")
                Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        log ("Error al cargar los datos(clsEQUIPOS_PLANES_ACCIONES)")
End Function

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_ACCION) FROM eq_planes_acciones"
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
        consulta = "INSERT INTO eq_planes_acciones " & _
                   "  VALUES (" & _
                        ID_ACCION & ", " & FAMILIA_ACC_ID & ", " & "'" & nombre & "'" & ", " & T_PREVISTO & ", " & _
                        "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                   ")"
        execute_bd consulta
        Insertar = ID_ACCION
    End If
    Exit Function
    
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEQUIPOS_PLANES_ACCIONES)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        Modificar = False
        If Not duplicado(ID) Then
            consulta = "UPDATE eq_planes_acciones SET " & _
                       "       FAMILIA_ACC_ID = " & FAMILIA_ACC_ID & ", " & _
                       "       NOMBRE = '" & nombre & "'" & ", " & _
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
        MsgBox "Error al Modificar (clsEQUIPOS_PLANES_ACCIONES)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As New ADODB.Recordset

        ' Se comprueba si hay planes con esa acción
        consulta = "SELECT COUNT(*) AS TOTAL " & _
                   "  FROM eq_planes_mantenimiento_DETALLE " & _
                   " WHERE ACCION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs("TOTAL") <> 0 Then
            MsgBox "No se puede eliminar la acción. Forma parte de algún plan de mantenimiento.", vbExclamation, App.Title
            Eliminar = False
            Exit Function
        End If
        
        execute_bd ("DELETE FROM eq_planes_acciones WHERE ID_ACCION = " & ID)
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEQUIPOS_PLANES_ACCIONES)", vbCritical, Err.Description
End Function

Public Function Listado(strNOMBRE As String, strTPrev As String, strFamAcc As String) As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT EPA.*, A.DESCRIPCION FAMILIA " & _
               "  FROM eq_planes_acciones EPA, decodificadora A" & _
               " WHERE EPA.FAMILIA_ACC_ID = A.VALOR AND A.CODIGO = " & DECODIFICADORA.EQ_FAMILIAS_ACCIONES_PLANES_MTO & _
               "   AND NOMBRE LIKE '%" & strNOMBRE & "%' " & _
               "   AND T_PREVISTO LIKE '%" & strTPrev & "%' " & _
               "   AND A.DESCRIPCION LIKE '%" & strFamAcc & "%' " & _
               " ORDER BY NOMBRE "
               
    Set Listado = datos_bd(consulta)
End Function

Public Function Listado_completo() As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT * " & _
               "  FROM eq_planes_acciones " & _
               " ORDER BY NOMBRE "
    Set Listado_completo = datos_bd(consulta)
End Function
Public Function Listado_por_Familia(ID_familia_acciones As Long) As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT ID_ACCION, NOMBRE " & _
               "  FROM eq_planes_acciones " & _
               " WHERE FAMILIA_ACC_ID = " & ID_familia_acciones & _
               " ORDER BY NOMBRE "
    Set Listado_por_Familia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT ID_ACCION, NOMBRE " & _
                   "  FROM eq_planes_acciones " & _
                   " ORDER BY NOMBRE "
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function duplicado(lngID_ACCION As Long) As Boolean
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT ID_ACCION " & _
               "  FROM eq_planes_acciones " & _
               " WHERE NOMBRE = '" & nombre & "'" & _
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

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "eq_planes_acciones"
        .setDESCRIPCION = "Acciones"
        .setPK = "ID_ACCION"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
End Function

Public Function Insertar_en_plan(ByVal ID_ACCION As Long, ID_PLAN_MTO As Long, ORDEN As Integer) As Boolean

    Dim consulta As String
    Dim rs As ADODB.Recordset
    
    Insertar_en_plan = True
    
    
    consulta = "SELECT * FROM eq_planes_mantenimiento_detalle WHERE ACCION_ID = " & CStr(ID_ACCION) & " AND PLAN_MTO_ID = " & CStr(ID_PLAN_MTO)
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount <> 0 Then
        Set rs = Nothing
        MsgBox "No se puede insertar la acción porque ya existe en el plan", vbInformation, "Insertar Acción en Plan de Mantenimiento"
        Insertar_en_plan = False
        Exit Function
    End If
    
    consulta = "insert into eq_planes_mantenimiento_detalle (plan_mto_id, accion_id, orden, cuserid,ftimestp) values ("
    consulta = consulta & ID_PLAN_MTO & ", " & ID_ACCION & ", " & ORDEN & " , " & USUARIO.getID_EMPLEADO & ",localtimestamp)"
    
    execute_bd consulta
    
    Set rs = Nothing
        
End Function

Public Function eliminar_de_plan(ByVal ID_ACCION As Long, ID_PLAN_MTO As Long, ORDEN As Integer) As Boolean

    Dim consulta As String
    
    consulta = "delete from eq_planes_mantenimiento_detalle where accion_id = " & ID_ACCION & " and plan_mto_id = " & ID_PLAN_MTO & " and orden = " & ORDEN
    
    execute_bd consulta
    
    consulta = "update eq_planes_mantenimiento_detalle set orden = orden -1 where plan_mto_id = " & ID_PLAN_MTO & " and orden > " & ORDEN
    
    execute_bd consulta
    

End Function

