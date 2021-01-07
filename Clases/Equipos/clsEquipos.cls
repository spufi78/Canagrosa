VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS
'***************************************************************
Private ID_EQUIPO As Long
Private particular As Long
Private NEQUIPO As Long
Private nombre As String
'M0459-I
Private UNIDADES_LOTE As Integer
'M0459-f
Private DESCRIPCION As String
Private CONDICIONES_AMBIENTALES As Long
Private PROVEEDOR_ID As Integer
Private FECHA_RECEPCION As String
Private FECHA_SERVICIO As String
Private ALTA_BAJA As Integer
'M1124-I
Private ESTADO_ID As String
'M1124-F
Private SERIE As String
Private NOTAS As String
Private FAMILIA_ID As Integer
Private familia As clsFamiliasEquipos
Private RESPONSABLE_ID As Long
Private MODELO As String
Private fabricante As String
Private ES_NADCAP As Integer
Private ES_MTL As Integer
Private ES_CP As Integer
Private ES_ENAC As Integer
Private RANGO_MEDIDA_MIN As String
Private RANGO_MEDIDA_MAX As String
Private UNIDAD_ID As Integer
Private TOLERANCIA_MAXIMA As String
Private INCERTIDUMBRE_MAXIMA As String
Private PRECISIONN As String
Private TEMPERATURA_MIN As String
Private TEMPERATURA_MAX As String
Private HUMEDAD_MIN As String
Private HUMEDAD_MAX As String
Private COND_AMBIENTALES_OTRAS As String
Private RANGO_TRABAJO_MIN As String
Private RANGO_TRABAJO_MAX As String
Private LIMITACIONES_USO As String
Private PERIODICIDAD_CALIBRACION_ID As Integer
Private FECHA_ULT_CALIBRACION As String
Private FECHA_PROX_CALIBRACION As String
Private PERIODICIDAD_VERIFICACION_ID As Integer
Private FECHA_ULT_VERIFICACION As String
Private FECHA_PROX_VERIFICACION As String
Private PERIODICIDAD_MANTENIMIENTO_ID As Integer
Private FECHA_ULT_MANTENIMIENTO As String
Private FECHA_PROX_MANTENIMIENTO As String
Private TIPO_CALIBRACION_ID As Integer
Private TIPO_VERIFICACION_ID As Integer
Private CALIBRADOR_ID As Integer
Private CALIBRADOR_INTERNO_ID As Integer
Private VERIFICADOR_ID As Integer
Private VERIFICADOR_INTERNO_ID As Integer
Private SITUACION_ID As Integer
Private CON_CALIBRACION As Integer
Private CON_VERIFICACION As Integer
Private CON_MANTENIMIENTO As Integer
Private REVISADO As Integer
Private ID_AUX As Integer
Private COORDX As Long
Private COORDY As Long
Private PROCEDENCIA_ID As Long
'Private RESPONSABLE_ID As Long
Private SITUACION_DESCRIPCION As String
Private LIMITACIONES_USO_COL As New clsGenericCollection
Private PLAN_MANTENIMIENTO_ID As Long
Private MANTENEDOR_ID As Long
Private ACCESORIOS_COL As New clsGenericCollection
Private MOSTRAR_EN_PLANO As Integer
Private PRIORITARIO As Integer
Private INSITU As Integer
Private TIPO_EQUIPO_ID As Integer
Private TIPO_EQUIPO As String
Private PROCEDIMIENTO_CALIBRACION_ID As Long
Private PROCEDIMIENTO_VERIFICACION_ID As Long
Private FUERA_SERVICIO As Integer
Private es_accesorio As Integer
Private NUMERO_USOS_MAXIMO As Integer
Private NUMERO_USOS_CONTADOR As Integer
'M1050-I
Private CLIENTE_ID As Long
Private TRAZMET_ID As Long
Private NUMERO_EQUIPO_CLIENTE As String
Private OBSERVACIONES As String
'M1050-F
Private CRITICO As Integer
Private CENTRO_ID As Integer

'Colecciones
' Colecciones de documentos Vinculados
'Private mvarobjAdjuntosDocumentacion As New clsGenericCollection
'J5 Private mvarobjAdjuntosInstruccionesTecnicas As New clsGenericCollection
'Private mvarobjCalibraciones As New clsGenericCollection
'Private mvarobjVerificaciones As New clsGenericCollection
'Private mvarobjMantenimientos As New clsGenericCollection
'Private mvarobjNormas As New clsGenericCollection
' Faltan los eventos, que van por su propia cuenta (ver clsEquipoEventos)
'
'M1050 Private view_VLIST_EQUIPOS_PREAVISO_CAL As String
'M1050 Private view_VLIST_EQUIPOS_PREAVISO_VER As String
'M1050 Private view_VLIST_EQUIPOS_PREAVISO_MTO As String
'M1050 Private view_VLIST_EQUIPOS_PERIODICIDAD_MANTENIMIENTO As String
'M1050 Private view_VLIST_DECODIFICADORA_FRECUENCIA_EQUIPOS As String

Public Function Actualizar_Fecha_Prox_Mantenimiento(ByVal prmIdEquipo As Long, Optional ByVal dtmFecha_Prox As Date = vbNull) As Boolean
    Dim sql As String
    Dim rs As New ADODB.Recordset
    
    If Not IsNull(dtmFecha_Prox) Then
    If IsDate(dtmFecha_Prox) Then
    If dtmFecha_Prox <> CDate("31/12/1899") Then
        sql = "update equipos set FECHA_PROX_MANTENIMIENTO = '" & Format(dtmFecha_Prox, "dd/mm/yyyy") & "' WHERE ID_EQUIPO = " & CStr(prmIdEquipo)
    
        execute_bd sql
        
        Actualizar_Fecha_Prox_Mantenimiento = True
        
        Exit Function
    End If
    End If
    End If
    
    
    ' cuando no se pasa la fecha por parámerto
    
    If prmIdEquipo > 0 Then
        
'        sql = "SELECT equipo_id, fecha_actual FROM eq_mantenimiento_equipos where equipo_id = " & CStr(prmIdEquipo) & " and estado = 0"
'        Set rs = datos_bd(sql)
'        If rs.RecordCount <> 0 Then
            ' Seleccionamos la fecha del próximo mantenimiento no realizado de todo el equipo
            ' Actualizamos fecha_prox_mantenimiento
            sql = "SELECT equipo_id, fecha_actual FROM eq_mantenimiento_equipos where equipo_id = " & CStr(prmIdEquipo) & " and estado = 0 order by fecha_actual asc limit 1"
            Set rs = datos_bd(sql)
            ' comprueba que existan datos
            If rs.RecordCount <> 0 Then
                On Error Resume Next
                ' comprueba que no sean nulos
                rs.MoveFirst
                If Err.Number = 0 Then
                    If Not IsNull(rs(0)) Then
                        sql = "UPDATE EQUIPOS SET FECHA_PROX_MANTENIMIENTO = '" & Format(rs("fecha_actual"), "YYYY-MM-DD") & "' WHERE EQUIPOS.ID_EQUIPO = " & CStr(prmIdEquipo)
                        execute_bd sql
                    End If
                End If
            Else
                sql = "UPDATE EQUIPOS SET FECHA_PROX_MANTENIMIENTO = null WHERE EQUIPOS.ID_EQUIPO = " & CStr(prmIdEquipo)
                execute_bd sql
            End If
            ' Actualizamos fecha_ult_mantenimiento
            sql = "SELECT equipo_id, fecha_actual FROM eq_mantenimiento_equipos where equipo_id = " & CStr(prmIdEquipo) & " and estado = 1 order by fecha_actual desc limit 1 "
            Set rs = datos_bd(sql)
            ' comprueba que existan datos
            If rs.RecordCount <> 0 Then
                On Error Resume Next
                ' comprueba que no sean nulos
                rs.MoveFirst
                If Err.Number = 0 Then
                    If Not IsNull(rs(0)) Then
                        sql = "UPDATE EQUIPOS SET FECHA_ULT_MANTENIMIENTO = '" & Format(rs("fecha_actual"), "YYYY-MM-DD") & "' WHERE EQUIPOS.ID_EQUIPO = " & CStr(prmIdEquipo)
                        execute_bd sql
                    End If
                End If
            Else
                sql = "UPDATE EQUIPOS SET FECHA_ULT_MANTENIMIENTO = null WHERE EQUIPOS.ID_EQUIPO = " & CStr(prmIdEquipo)
                execute_bd sql
            End If

'        End If

    End If


End Function
Public Function actualizar_plan_mantenimiento_por_defecto(ID_PLAN As Long) As Boolean

    Dim strCad As String
    
    
    strCad = "update eq_planes_mantenimiento_equipos set por_defecto = 0 where equipo_id = " & CStr(ID_EQUIPO)
    execute_bd strCad
        
    strCad = "update eq_planes_mantenimiento_equipos set por_defecto = 1 where equipo_id = " & CStr(ID_EQUIPO) & " and plan_mantenimiento_id = " & ID_PLAN
    execute_bd strCad
    
    ' lo pone tambien en EQUIPOS
    strCad = "UPDATE equipos SET PLAN_MANTENIMIENTO_ID = " & ID_PLAN & " where ID_equipo = " & CStr(ID_EQUIPO)
    execute_bd strCad
    
    
End Function

Public Function anadir_plan_mantenimiento_equipo(ID_PLAN As Long) As Boolean

    Dim strCad As String, rs As ADODB.Recordset
    
    strCad = "SELECT * from eq_planes_mantenimiento_equipos where equipo_id = " & CStr(ID_EQUIPO) & " and plan_mantenimiento_id = " & CStr(ID_PLAN)
    
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount <> 0 Then
        Set rs = Nothing
        Exit Function
    End If
    
    strCad = "insert into eq_planes_mantenimiento_equipos (equipo_id, plan_mantenimiento_id, por_defecto) values(" & ID_EQUIPO & ", " & ID_PLAN & ", 0)"
    execute_bd strCad
        
    'ahora evalua si alguno está puesto como POR DEFECTO
    strCad = "SELECT * from eq_planes_mantenimiento_equipos where equipo_id = " & CStr(ID_EQUIPO) & " and por_defecto = 1"
    
    Set rs = datos_bd(strCad)
    If rs.RecordCount = 0 Then
        actualizar_plan_mantenimiento_por_defecto ID_PLAN
    End If
    
End Function

Public Function AplicarAsignacionRapida(ByVal lista_seleccionados As String, ByVal familia As Long, ByVal tipo As Long, ByVal responsable As Long, ByVal Procedencia As Long, ByVal nadcap As Long, ByVal accesorio As Integer, ByVal calibracion As Integer, ByVal verificacion As Integer, _
    ByVal Cal_periodo As Long, ByVal cal_procedimiento As Long, ByVal Cal_Responsable As Long, ByVal Cal_Tipo As Long, _
    ByVal Ver_periodo As Long, ByVal ver_procedimiento As Long, ByVal Ver_Responsable As Long, ByVal ver_tipo As Long, _
    fabricante, localizacion, proveedor) As Boolean
    
On Error GoTo fallo
    
    Dim strSql As String
    Dim lista_equipos As String
    Dim rs As ADODB.Recordset
    Dim dias_periodo As Long
    dias_periodo = 0

    
    lista_equipos = "(" & Replace(Replace(lista_seleccionados, "::", ","), ":", "") & ")"
    
    
    strSql = ""
        
    If Trim(lista_seleccionados) = "" Then
        MsgBox "Debe indicar al menos 1 Equipo en la lista.", vbInformation, "Asignacion Rápida a Equipos"
        Exit Function
    End If
    
    
    If familia > -1 Then strSql = strSql & ", familia_id = " & familia
    If tipo > -1 Then strSql = strSql & ", tipo_equipo_id = " & tipo
    If responsable > -1 Then strSql = strSql & ", responsable_id = " & responsable
    If nadcap > -1 Then strSql = strSql & ", es_nadcap = " & nadcap
    If accesorio > -1 Then strSql = strSql & ", es_accesorio = " & accesorio
    If proveedor > -1 Then strSql = strSql & ", proveedor_id = " & proveedor
    If localizacion > -1 Then strSql = strSql & ", situacion_id = " & localizacion
    If fabricante <> "" Then strSql = strSql & ", fabricante = '" & fabricante & "'"
    
    If calibracion = 1 Then
        strSql = strSql & ", con_calibracion = 1"
        If Cal_periodo > -1 Then strSql = strSql & ", PERIODICIDAD_CALIBRACION_ID = " & Cal_periodo
        If cal_procedimiento > -1 Then strSql = strSql & ", PROCEDIMIENTO_CALIBRACION_ID = " & cal_procedimiento
        If Cal_Responsable > -1 Then strSql = strSql & ", CALIBRADOR_INTERNO_ID = " & Cal_Responsable
        If Cal_Tipo > -1 Then strSql = strSql & ", TIPO_CALIBRACION_ID = " & Cal_Tipo
    ElseIf calibracion = 0 Then
        strSql = strSql & ", con_calibracion = 0"
    End If
    
    If verificacion = 1 Then
        strSql = strSql & ", con_verificacion = 1"
        If Ver_periodo > -1 Then strSql = strSql & ", PERIODICIDAD_VERIFICACION_ID = " & Ver_periodo
        If ver_procedimiento > -1 Then strSql = strSql & ", PROCEDIMIENTO_VERIFICACION_ID = " & ver_procedimiento
        If Ver_Responsable > -1 Then strSql = strSql & ", VERIFICADOR_INTERNO_ID = " & Ver_Responsable
        If ver_tipo > -1 Then strSql = strSql & ", TIPO_VERIFICACION_ID = " & ver_tipo
    ElseIf verificacion = 0 Then
        strSql = strSql & ", con_verificacion = 0"
    End If
    
    
    If Trim(strSql) = "" Then
        MsgBox "No ha indicado que característica desea modificar para el grupo de Equipos.", vbInformation, "Asignacion Rápida a Equipos"
        Exit Function
    End If
        
    strSql = "UPDATE EQUIPOS SET " & Mid(strSql, 2) & " WHERE ID_EQUIPO IN " & lista_equipos
    execute_bd strSql
    
    ' Reasignar mantenimientos
    If responsable > -1 Then
        execute_bd "update eq_mantenimiento_equipos set mantenedor_id = " & responsable & " where equipo_id in " & lista_equipos & " and estado = 0"
        execute_bd "update eq_operaciones_pendientes set responsable_id = " & responsable & " where equipo_id in " & lista_equipos & " and tipo_cvm_id = 2"
    End If
    AplicarAsignacionRapida = True

Exit Function
fallo:

End Function

Public Function comprobar_es_accesorio_de_otros(ID As Long) As Boolean
Dim rs As ADODB.Recordset

    comprobar_es_accesorio_de_otros = False

    Set rs = datos_bd("SELECT * FROM eq_accesorios WHERE ID_ACCESORIO = " & CStr(ID))
    
    If rs.RecordCount <> 0 Then
        comprobar_es_accesorio_de_otros = True
    End If
        
    Set rs = Nothing
    
End Function
Public Function devolver_actual_mantenimiento(ByVal id_eq As Long) As Long
    Dim rs As ADODB.Recordset
    Set rs = datos_bd("select id_mantenimiento from eq_mantenimiento_equipos where estado <> 0 and equipo_id = " & id_eq & " order by fecha_actual desc")
    
    If rs.RecordCount <> 0 Then
        devolver_actual_mantenimiento = rs(0)
    Else
        devolver_actual_mantenimiento = 0
    End If
End Function

Public Function devolver_fecha_prox_mantenimiento() As ADODB.Recordset

    Dim strCad As String
    
    If CON_MANTENIMIENTO = 0 Then
        strCad = "SELECT 0 as ID_MANTENIMIENTO, 0 as PLANMTO_ID, 0 as MANTENEDOR_ID, '' AS FECHA_ACTUAL, '' as RESPONSABLE, '' as PLAN_MANTENIMIENTO, '' AS PROTOCOLO"
    Else
        strCad = " select ID_MANTENIMIENTO, PLANMTO_ID, MANTENEDOR_ID, min(FECHA_ACTUAL) AS FECHA_ACTUAL " & _
                " , coalesce(concat(usuarios.nombre, ' ', usuarios.APELLIDOS), 'N/A') as RESPONSABLE" & _
                " , coalesce (eq_planes_mantenimiento.nombre, 'N/A') as PLAN_MANTENIMIENTO" & _
                ", coalesce (ca_documentos.nombre, 'N/A') as PROTOCOLO" & _
                " from eq_mantenimiento_equipos " & _
                " left outer join usuarios on eq_mantenimiento_equipos.MANTENEDOR_ID = usuarios.ID_EMPLEADO" & _
                " left outer join eq_planes_mantenimiento on eq_mantenimiento_equipos.PLANMTO_ID = eq_planes_mantenimiento.ID_PLAN_MTO" & _
                " left outer join ca_documentos on eq_mantenimiento_equipos.PROCEDIMIENTO_ID = ca_documentos.id_documento" & _
                " WHERE eq_mantenimiento_equipos.equipo_id = " & CStr(ID_EQUIPO) & " AND estado = 0"
    End If

    Set devolver_fecha_prox_mantenimiento = datos_bd(strCad)

End Function
Public Function devolver_lista_calibraciones(ByVal ID_EQUIPO As Long, Optional ByRef fecha_ult As String = "", Optional ByRef responsable_ult As String = "", Optional ByRef responsable_ult_id As String = "", Optional ByVal prm_ver_previsto As Integer) As ADODB.Recordset
    Dim cad As String, rs As ADODB.Recordset, dt_ref As Date
    cad = "select a.fecha_actual as fecha, a.CALIBRADOR_interno_id as responsable_id, coalesce(concat(usuarios.NOMBRE,' ',usuarios.APELLIDOS), '') as responsable " & _
        " from eq_calibracion_equipos a " & _
        " left outer join usuarios on a.CALIBRADOR_interno_id = usuarios.ID_EMPLEADO " & _
        " where a.equipo_id = " & CStr(ID_EQUIPO) & " and a.estado > 0 order by fecha_actual desc"
    Set rs = datos_bd(cad)
    
    If rs.RecordCount = 0 Then
        fecha_ult = ""
    Else
        rs.MoveFirst
        fecha_ult = Format(rs("fecha"), "dd/mm/yyyy")
        responsable_ult = rs!responsable
        responsable_ult_id = rs!RESPONSABLE_ID
    End If
    
    cad = "select a.* " & _
            " , tipo_cal.descripcion as tipo_calibracion" & _
            " , coalesce(period.descripcion, '') as periodicidad" & _
            " , coalesce(calibrador_ext.NOMBRE,'') as calibrador_externo" & _
            " , coalesce(concat(usuarios.NOMBRE,' ',usuarios.APELLIDOS), '') as responsable" & _
            " , coalesce(proc.NOMBRE,'') as procedimiento" & _
            " , decoEstado.descripcion as ESTADO_TEXTO " & _
            " , decoResultado.descripcion as RESULTADO_TEXTO " & _
            " from eq_calibracion_equipos a" & _
            " inner join decodificadora tipo_cal on a.TIPO_ID = tipo_cal.VALOR and tipo_cal.CODIGO = 104" & _
            " left outer join eq_periodicidad period on a.PERIODICIDAD_ID = period.ID_PERIODICIDAD" & _
            " left outer join proveedores calibrador_ext on a.CALIBRADOR_EXTERNO_ID = calibrador_ext.ID_PROVEEDOR" & _
            " left outer join usuarios on a.CALIBRADOR_interno_id = usuarios.ID_EMPLEADO " & _
            " left outer join ca_documentos proc on a.PROCEDIMIENTO_ID = proc.ID_DOCUMENTO" & _
            " left join decodificadora decoEstado on a.ESTADO = decoEstado.VALOR and decoEstado.CODIGO = " & DECODIFICADORA_EQUIPOS_CALIBRACIONES_ESTADOS & _
            " left join decodificadora decoResultado on a.RESULTADO = decoResultado.VALOR and decoResultado.CODIGO = " & DECODIFICADORA_EQUIPOS_CALIBRACIONES_RESULTADOS & _
            " where a.equipo_id = " & CStr(ID_EQUIPO)
    If prm_ver_previsto = 1 Then
        cad = cad & " And a.estado = 0 "
        cad = cad & " order by id_calibracion asc"
    ElseIf prm_ver_previsto = 0 Then
        cad = cad & " And a.estado > 0"
        cad = cad & " order by id_calibracion desc"
    Else
        cad = cad & " order by id_calibracion desc"
    End If

    Set devolver_lista_calibraciones = datos_bd(cad)
    Set rs = Nothing


End Function
Public Function devolver_lista_verificaciones(ByVal ID_EQUIPO As Long, Optional ByRef fecha_ult As String = "", Optional ByRef responsable_ult As String = "", Optional ByRef responsable_ult_id As String = "", Optional ByVal prm_ver_previsto As Integer) As ADODB.Recordset
    Dim cad As String, rs As ADODB.Recordset, dt_ref As Date

     ' calcula de la última realizada
    cad = "select a.fecha_actual as fecha, a.verificADOR_interno_id as responsable_id, coalesce(concat(usuarios.NOMBRE,' ',usuarios.APELLIDOS), '') as responsable " & _
        " from eq_verificacion_equipos a " & _
        " left outer join usuarios on a.verificADOR_interno_id = usuarios.ID_EMPLEADO " & _
        " where a.equipo_id = " & CStr(ID_EQUIPO) & " and a.estado > 0 order by fecha_actual desc"
    Set rs = datos_bd(cad)
    
    If rs.RecordCount = 0 Then
        fecha_ult = ""
    Else
        rs.MoveFirst
        fecha_ult = Format(rs("fecha"), "dd/mm/yyyy")
        responsable_ult = rs!responsable
        responsable_ult_id = rs!RESPONSABLE_ID
    End If
    
    'recoge las últimas verificaciones
    
    cad = "select a.* " & _
            " , tipo_ver.descripcion as tipo_verificacion" & _
            " , coalesce(period.descripcion, '') as periodicidad" & _
            " , coalesce(verificador_ext.NOMBRE,'') as verificador_externo" & _
            " , coalesce(concat(usuarios.NOMBRE,' ',usuarios.APELLIDOS), '') as responsable" & _
            " , coalesce(proc.NOMBRE,'') as procedimiento" & _
            " , decoEstado.descripcion as ESTADO_TEXTO " & _
            " , decoResultado.descripcion as RESULTADO_TEXTO " & _
            " from eq_verificacion_equipos a" & _
            " inner join decodificadora tipo_ver on a.TIPO_ID = tipo_ver.VALOR and tipo_ver.CODIGO = 104" & _
            " left outer join eq_periodicidad period on a.PERIODICIDAD_ID = period.ID_PERIODICIDAD" & _
            " left outer join proveedores verificador_ext on a.VERIFICADOR_EXTERNO_ID = verificador_ext.ID_PROVEEDOR" & _
            " left outer join usuarios on a.verificador_interno_id = usuarios.ID_EMPLEADO " & _
            " left outer join ca_documentos proc on a.PROCEDIMIENTO_ID = proc.ID_DOCUMENTO" & _
            " left join decodificadora decoEstado on a.ESTADO = decoEstado.VALOR and decoEstado.CODIGO = " & DECODIFICADORA_EQUIPOS_VERIFICACIONES_ESTADOS & _
            " left join decodificadora decoResultado on a.RESULTADO = decoResultado.VALOR and decoResultado.CODIGO = " & DECODIFICADORA_EQUIPOS_VERIFICACIONES_RESULTADOS & _
            " where a.equipo_id = " & CStr(ID_EQUIPO)
    If prm_ver_previsto = 1 Then
        cad = cad & " And a.estado = 0 "
        cad = cad & " order by fecha_actual asc,id_verificacion asc"
    ElseIf prm_ver_previsto = 0 Then
        cad = cad & " And a.estado > 0"
        cad = cad & " order by fecha_actual desc,id_verificacion desc"
    Else
        cad = cad & " order by fecha_actual desc,id_verificacion desc"
    End If

    Set devolver_lista_verificaciones = datos_bd(cad)
    Set rs = Nothing


End Function
Public Function devolver_actual_calibracion(ByVal id_eq As Long) As Long
    
    Dim rs As ADODB.Recordset
'    Set rs = datos_bd("select id_calibracion from eq_calibracion_equipos where estado <> 0 and equipo_id = " & id_eq & " order by fecha_prevista desc")
    Set rs = datos_bd("select id_calibracion from eq_calibracion_equipos where estado <> 0 and equipo_id = " & id_eq & " order by fecha_actual desc")
    
    If rs.RecordCount <> 0 Then
        devolver_actual_calibracion = rs(0)
    Else
        devolver_actual_calibracion = 0
    End If
End Function
Public Function devolver_proxima_calibracion(ByVal id_eq As Long) As ADODB.Recordset
    
    Dim rs As ADODB.Recordset
    Dim cad As String
    
    Set rs = datos_bd("select id_calibracion from eq_calibracion_equipos where estado = 0 and equipo_id = " & id_eq & " order by fecha_prevista asc")
    
    If rs.RecordCount <> 0 Then
        
        rs.MoveFirst
        
        cad = "select a.* " & _
                " , tipo_cal.descripcion as tipo_calibracion" & _
                " , coalesce(period.descripcion, '') as periodicidad" & _
                " , coalesce(calibrador_ext.NOMBRE,'') as calibrador_externo" & _
                " , coalesce(concat(usuarios.NOMBRE,' ',usuarios.APELLIDOS), '') as responsable" & _
                " , coalesce(proc.NOMBRE,'') as procedimiento" & _
                " from eq_calibracion_equipos a" & _
                " inner join decodificadora tipo_cal on a.TIPO_ID = tipo_cal.VALOR and tipo_cal.CODIGO = 104" & _
                " left outer join eq_periodicidad period on a.PERIODICIDAD_ID = period.ID_PERIODICIDAD" & _
                " left outer join proveedores calibrador_ext on a.CALIBRADOR_EXTERNO_ID = calibrador_ext.ID_PROVEEDOR" & _
                " left outer join usuarios on a.CALIBRADOR_interno_id = usuarios.ID_EMPLEADO " & _
                " left outer join ca_documentos proc on a.PROCEDIMIENTO_ID = proc.ID_DOCUMENTO" & _
                " where a.id_calibracion = " & CStr(rs!ID_CALIBRACION)
        Set rs = datos_bd(cad)
    End If
    
    Set devolver_proxima_calibracion = rs

    


End Function
Public Function devolver_actual_verificacion(ByVal id_eq As Long) As Long
    
    Dim rs As ADODB.Recordset
'    Set rs = datos_bd("select id_verificacion from eq_verificacion_equipos where estado <> 0 and equipo_id = " & id_eq & " order by fecha_prevista desc")
    Set rs = datos_bd("select id_verificacion from eq_verificacion_equipos where estado <> 0 and equipo_id = " & id_eq & " order by fecha_actual desc")
    
    If rs.RecordCount <> 0 Then
        devolver_actual_verificacion = rs(0)
    Else
        devolver_actual_verificacion = 0
    End If
End Function

Public Function devolver_proxima_verificacion(ByVal id_eq As Long) As ADODB.Recordset
    
    Dim rs As ADODB.Recordset
    Dim cad As String
    
    Set rs = datos_bd("select id_verificacion from eq_verificacion_equipos where estado = 0 and equipo_id = " & id_eq & " order by fecha_prevista asc")
    
    If rs.RecordCount <> 0 Then
        
        rs.MoveFirst
        
        cad = "select a.* " & _
                " , tipo_ver.descripcion as tipo_verificacion" & _
                " , coalesce(period.descripcion, '') as periodicidad" & _
                " , coalesce(verificador_ext.NOMBRE,'') as verificador_externo" & _
                " , coalesce(concat(usuarios.NOMBRE,' ',usuarios.APELLIDOS), '') as responsable" & _
                " , coalesce(proc.NOMBRE,'') as procedimiento" & _
                " from eq_verificacion_equipos a" & _
                " inner join decodificadora tipo_ver on a.TIPO_ID = tipo_ver.VALOR and tipo_ver.CODIGO = 104" & _
                " left outer join eq_periodicidad period on a.PERIODICIDAD_ID = period.ID_PERIODICIDAD" & _
                " left outer join proveedores verificador_ext on a.VERIFICADOR_EXTERNO_ID = verificador_ext.ID_PROVEEDOR" & _
                " left outer join usuarios on a.verificADOR_interno_id = usuarios.ID_EMPLEADO " & _
                " left outer join ca_documentos proc on a.PROCEDIMIENTO_ID = proc.ID_DOCUMENTO" & _
                " where a.id_verificacion = " & CStr(rs!ID_VERIFICACION)
        Set rs = datos_bd(cad)
    End If
    
    Set devolver_proxima_verificacion = rs

    


End Function

Public Function duplicar_equipo(ByRef ID As String) As Boolean
On Error GoTo fallo

    carga CLng(ID)
    CrearID
    Me.setNUMERO_EQUIPO_CLIENTE = ""
    Insertar
    duplicar_limitaciones_uso ID, ID_EQUIPO
    duplicar_accesorios ID, ID_EQUIPO
'    duplicar_documentacion ID, ID_EQUIPO
    duplicar_normas ID, ID_EQUIPO
    duplicar_planes ID, ID_EQUIPO
    ID = ID_EQUIPO
    duplicar_equipo = True

Exit Function
fallo:


End Function

Private Sub duplicar_limitaciones_uso(ByVal id_eq As String, ByVal id_nuevo As String)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim max_id As Long
    
    consulta = "select cast(coalesce(max(ID_LIMITACION_USO), 0) as unsigned) as id from eq_limitaciones_uso"
    Set rs = datos_bd(consulta)
    
    rs.MoveFirst
    max_id = rs!ID
    
    consulta = "select * from eq_limitaciones_uso where equipo_id = " & id_eq
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount <> 0 Then
    
        While Not rs.EOF
            max_id = max_id + 1
            
            consulta = "insert into eq_limitaciones_uso (ID_LIMITACION_USO, EQUIPO_ID, DESCRIPCION)"
            consulta = consulta & " VALUES (" & max_id & ", " & id_nuevo & ", '" & rs!DESCRIPCION & "')"
            
            execute_bd consulta
            
            rs.MoveNext
        Wend
    End If
    
End Sub


Private Sub duplicar_normas(ByVal id_eq As String, ByVal id_nuevo As String)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim max_id As Long
       
    consulta = "select * from eq_normas_equipos where equipo_id = " & id_eq
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount <> 0 Then
    
        While Not rs.EOF
            
            
            consulta = "insert into eq_normas_equipos (documento_id, EQUIPO_ID)"
            consulta = consulta & " VALUES (" & rs!DOCUMENTO_ID & ", " & id_nuevo & ")"
            
            execute_bd consulta
            
            rs.MoveNext
        Wend
    End If
    
End Sub

Private Sub duplicar_planes(ByVal id_eq As String, ByVal id_nuevo As String)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim max_id As Long
       
    consulta = "select * from eq_planes_mantenimiento_equipos where equipo_id = " & id_eq
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount <> 0 Then
    
        While Not rs.EOF
            
            
            consulta = "insert into eq_planes_mantenimiento_equipos (PLAN_MANTENIMIENTO_ID, EQUIPO_ID, POR_DEFECTO)"
            consulta = consulta & " VALUES (" & rs!PLAN_MANTENIMIENTO_ID & ", " & id_nuevo & ", " & rs!POR_DEFECTO & ")"
            
            execute_bd consulta
            
            rs.MoveNext
        Wend
    End If
    
End Sub


'Private Sub duplicar_documentacion(ByVal id_eq As String, ByVal id_nuevo As String)
'End Sub
Private Sub duplicar_accesorios(ByVal id_eq As String, ByVal id_nuevo As String)
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim max_id As Long
       
    consulta = "select * from eq_accesorios where equipo_id = " & id_eq
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount <> 0 Then
    
        While Not rs.EOF
            
            
            consulta = "insert into eq_accesorios (ID_accesorio, EQUIPO_ID, orden, cuserid, ftimestp, fecha_alta, fecha_baja, en_uso)"
            consulta = consulta & " VALUES (" & rs!ID_ACCESORIO & ", " & id_nuevo & ", " & rs!ORDEN & ", '" & rs!CUSERID & "', localtimestamp, '" & Format(rs!FECHA_ALTA, "yyyy-mm-dd hh:nn:ss") & "', '" & Format(rs!FECHA_BAJA, "yyyy-mm-dd hh:nn:ss") & "', " & rs!EN_USO & ")"
            
            execute_bd consulta
            
            rs.MoveNext
        Wend
    End If
    
End Sub



Public Function eliminar_mantenimiento(ByVal prm_ID_MANTENIMIENTO As Long, ByVal prm_id_plan As Long) As Boolean
Dim strCad As String

    strCad = "delete from eq_mantenimiento_equipos where id_mantenimiento = " & prm_ID_MANTENIMIENTO
    execute_bd strCad
    
    strCad = "delete from eq_mantenimiento_equipos_acciones where mantenimiento_id = " & prm_ID_MANTENIMIENTO
    execute_bd strCad
    ' eq_operaciones_pendientes
'OOP    strCad = "delete from eq_operaciones_pendientes where tipo_cvm_id = 2 and id_cvm_origen = " & prm_ID_MANTENIMIENTO
'OOP    execute_bd strCad
    


' genera el evento
Dim oEV As New clsEquipoEventos

With oEV
    .setEQUIPO_ID = ID_EQUIPO
    .setCVM_ID = prm_ID_MANTENIMIENTO
    .setEVENTO_ID = EQUIPOS_EVENTOS.EVT_REG_PREVISTO_ELIMINADO
    .setRAZON_ID = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_MANTENIMIENTO_ELIMINADO
    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
    .Insertar
End With
Set oEV = Nothing

    
'Revisa las tareas pendientes, por si tiene que poner otra para el mismo equipo
'OOP Dim rs As ADODB.Recordset
 Dim oOP As New clsEquiposOperacionesPendientes
'OOP Dim id_origen As Long

    ' COMPRUEBA SI ESTA CALIBRACION ESTABA COMO OPERACION PENDIENTE
'OOP    If oOP.Carga(CLng(prm_ID_MANTENIMIENTO), EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO) Then
        ' estaba como operacion pendiente
    
'OOP        Set rs = datos_bd("select id_mantenimiento from eq_mantenimiento_equipos where estado = 0 and equipo_id = " & ID_EQUIPO & " and planmto_ID = " & prm_id_plan & " order by fecha_actual asc")
        
'OOP        If rs.RecordCount = 0 Then
            oOP.Eliminar CLng(prm_ID_MANTENIMIENTO), ID_EQUIPO, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
'OOP        Else
'OOP            rs.MoveFirst
            ' modifica la tarea pendiente
'OOP            id_origen = oOP.getID_CVM_ORIGEN
'OOP            oOP.Eliminar CLng(prm_ID_MANTENIMIENTO), ID_EQUIPO, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
'OOP            oOP.crear_mantenimiento_pendiente id_origen, rs!id_mantenimiento, ID_EQUIPO, New clsEquiposPeriodicidad
'OOP        End If
'OOP    End If
    
    Set oOP = Nothing

End Function

Public Function eliminar_plan_mantenimiento_equipo(ID_PLAN As Long) As Boolean

    Dim strCad As String, rs As ADODB.Recordset
    Dim lng_default As Long
    
    lng_default = -1
    
        
    strCad = "delete from eq_planes_mantenimiento_equipos where equipo_id = " & CStr(ID_EQUIPO) & " and plan_mantenimiento_id = " & CStr(ID_PLAN)
    execute_bd strCad
            
    ' cuando lo va a eliminar y es el plan por defecto, pone a otro en su lugar, el primero que pilla
    strCad = "select * from eq_planes_mantenimiento_equipos where equipo_id = " & CStr(ID_EQUIPO)
    Set rs = datos_bd(strCad)
        
    If rs.RecordCount <> 0 Then
        ' cuando elimina el plan que tiene por defecto
        ' evalua si existen más, y pone al primero que encuentra.
        
        ' si no encuentra ninguno, lo pone a -1

        rs.MoveFirst
        ' pone al primero de la lista
        lng_default = rs("plan_mantenimiento_id")
        While Not rs.EOF
            If CInt(rs("por_defecto")) = 1 Then
                lng_default = rs("plan_mantenimiento_id")
            End If
            rs.MoveNext
        Wend
    End If
    
    strCad = "update equipos set plan_mantenimiento_id = " & lng_default & " where id_equipo = " & CStr(ID_EQUIPO)
    execute_bd strCad
    
    strCad = "update eq_planes_mantenimiento_equipos set por_defecto = 1 where plan_mantenimiento_id = " & CStr(lng_default) & " and equipo_id = " & CStr(ID_EQUIPO)
    execute_bd strCad
End Function
Public Function Anadir_limitacionuso_equipo(prmLimitacion As String) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "select * from eq_limitaciones_uso where equipo_id = " & CStr(ID_EQUIPO) & " AND descripcion = '" & Trim(prmLimitacion) & "'"
    Set rs = datos_bd(sql)
    
    If rs.RecordCount <> 0 Then
        Set rs = Nothing
        Exit Function
    End If
    
    sql = ""
    sql = "INSERT INTO eq_limitaciones_uso ("
    sql = sql & "equipo_id, "
    sql = sql & "id_limitacion_uso, "
    sql = sql & "descripcion"
    sql = sql & ") "
    sql = sql & "SELECT " & CStr(ID_EQUIPO) & " as equipo_id,"
    sql = sql & "coalesce(max(id_limitacion_uso),0)+1 as id_limitacion_uso, "
    sql = sql & "'" & prmLimitacion & "' as descripcion "
    sql = sql & "FROM eq_limitaciones_uso"
    
    execute_bd sql
    
    Call Cargar_LimitacionesUso
    
End Function
Public Function Anadir_Norma_Equipo(prmlngidNorma As Long, tipo As Integer) As Boolean
    Dim sql As String
    Dim rs As ADODB.Recordset
    
    sql = "select * from eq_normas_equipos where equipo_id = " & CStr(ID_EQUIPO) & " AND documento_id = " & CStr(prmlngidNorma) & " and tipo = " & tipo
    Set rs = datos_bd(sql)
    
    If rs.RecordCount <> 0 Then
        Set rs = Nothing
        Exit Function
    End If
        
    sql = ""
    sql = "INSERT INTO eq_normas_equipos ("
    sql = sql & "equipo_id, "
    sql = sql & "documento_id, "
    sql = sql & "tipo "
    sql = sql & ") "
    sql = sql & "SELECT " & CStr(ID_EQUIPO) & " as equipo_id,"
    sql = sql & CStr(prmlngidNorma) & " as documento_id, "
    sql = sql & tipo
    
    execute_bd sql
    
    'Call Carga_Normas
    
End Function


Public Function devolver_id_ult_calibracion(prmIdEquipo As Long) As Long

    Dim strCad As String
    Dim rs As ADODB.Recordset
    devolver_id_ult_calibracion = 0
    
    strCad = "SELECT ID_CALIBRACION FROM EQ_CALIBRACION_EQUIPOS WHERE EQUIPO_ID = " & prmIdEquipo & " AND estado > 0 order by fecha_actual desc"
    Set rs = datos_bd(strCad)
    If Not rs Is Nothing Then
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            devolver_id_ult_calibracion = rs("ID_CALIBRACION")
        End If
    End If
    
    Set rs = Nothing

End Function

Public Function devolver_id_ult_verificacion(prmIdEquipo As Long) As Long

    Dim strCad As String
    Dim rs As ADODB.Recordset
    devolver_id_ult_verificacion = 0
    
    strCad = "SELECT ID_VERIFICACION FROM EQ_verificacion_EQUIPOS WHERE EQUIPO_ID = " & prmIdEquipo & " AND estado > 0 order by fecha_actual desc"
    Set rs = datos_bd(strCad)
    If Not rs Is Nothing Then
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            devolver_id_ult_verificacion = rs("ID_VERIFICACION")
        End If
    End If
    
    Set rs = Nothing

End Function

'M1050-I
'Public Function Eliminar_accesorio_equipo(ByVal prmlngidAccesorio As String) As Boolean
'Dim strSql As String
'
'    strSql = "DELETE FROM eq_accesorios "
'    strSql = strSql & " WHERE id_accesorio= " & CStr(prmlngidAccesorio)
'    strSql = strSql & " AND equipo_id = " & CStr(ID_EQUIPO)
'
'    execute_bd strSql
'
'    cargar_accesorios
'
'End Function
'M1050-F
Public Function Eliminar_Calibracion(ByVal prmId As String, ByVal prm_id_periodicidad As Long) As Boolean
    Dim sql As String
    sql = "delete from eq_calibracion_equipos where id_calibracion = " & prmId
    execute_bd sql
    sql = "delete from eq_calibracion_equipos_accesorios where id_calibracion = " & prmId
    execute_bd sql
    ' reactivos
    sql = "delete from eq_calibracion_reactivos where calibracion_id = " & prmId
    execute_bd sql
    ' PARÁMETROS RESULTADOS
    sql = "delete from eq_calibracion_parametros_resultados where calibracion_id = " & prmId
    execute_bd sql
    ' genera el evento
    Dim oEV As New clsEquipoEventos
    With oEV
        .setEQUIPO_ID = ID_EQUIPO
        .setCVM_ID = prmId
        .setEVENTO_ID = EQUIPOS_EVENTOS.EVT_REG_PREVISTO_ELIMINADO
        .setRAZON_ID = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_CALIBRACION_ELIMINADA
        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
        .Insertar
    End With
    Set oEV = Nothing
    ' revisa las tareas pendientes, por si tiene que poner otra para el mismo equipo
'OOP    Dim rs As ADODB.Recordset
    Dim oOP As New clsEquiposOperacionesPendientes
'OOP    Dim id_origen As Long

    ' COMPRUEBA SI ESTA CALIBRACION ESTABA COMO OPERACION PENDIENTE
'OOP    If oOP.Carga(CLng(prmId), EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION) Then
        ' estaba como operacion pendiente
    
'OOP        Set rs = datos_bd("select id_calibracion from eq_calibracion_equipos where estado = 0 and equipo_id = " & ID_EQUIPO & " and periodicidad_ID = " & prm_id_periodicidad & " order by fecha_actual asc")
        
'OOP        If rs.RecordCount = 0 Then
            oOP.Eliminar CLng(prmId), ID_EQUIPO, EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
'OOP        Else
'OOP            rs.MoveFirst
            ' modifica la tarea pendiente
'OOP            id_origen = oOP.getID_CVM_ORIGEN
'OOP            oOP.Eliminar CLng(prmId), ID_EQUIPO, EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
'OOP            oOP.crear_calibracion_pendiente id_origen, rs!ID_CALIBRACION, ID_EQUIPO, New clsEquiposPeriodicidad
'OOP        End If
'OOP    End If
    
    Set oOP = Nothing

End Function


Public Function Eliminar_Verificacion(ByVal prmId As String, ByVal prm_id_periodicidad As Long) As Boolean

Dim sql As String

sql = "delete from eq_verificacion_equipos where id_verificacion = " & prmId
execute_bd sql

'Parametros resultados
sql = "delete from eq_verificacion_parametros_resultados where verificacion_id = " & prmId
execute_bd sql

'OOP ' eq_operaciones_pendientes
'OOP sql = "delete from eq_operaciones_pendientes where tipo_cvm_id = 1 and id_cvm_origen = " & prmId
'OOP execute_bd sql


' genera el evento
Dim oEV As New clsEquipoEventos

With oEV
    .setEQUIPO_ID = ID_EQUIPO
    .setCVM_ID = prmId
    .setEVENTO_ID = EQUIPOS_EVENTOS.EVT_REG_PREVISTO_ELIMINADO
    .setRAZON_ID = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_VERIFICACION_ELIMINADA
    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
    .Insertar
End With
Set oEV = Nothing


'Revisa las tareas pendientes, por si tiene que poner otra para el mismo equipo
'OOPDim rs As ADODB.Recordset
Dim oOP As New clsEquiposOperacionesPendientes
'OOPDim id_origen As Long

    ' COMPRUEBA SI ESTA CALIBRACION ESTABA COMO OPERACION PENDIENTE
'OOP    If oOP.Carga(CLng(prmId), EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION) Then
        ' estaba como operacion pendiente
    
'OOP        Set rs = datos_bd("select id_verificacion from eq_verificacion_equipos where estado = 0 and equipo_id = " & ID_EQUIPO & " and periodicidad_ID = " & prm_id_periodicidad & " order by fecha_actual asc")
        
'OOP        If rs.RecordCount = 0 Then
            oOP.Eliminar CLng(prmId), ID_EQUIPO, EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
'OOP        Else
'OOP            rs.MoveFirst
            ' modifica la tarea pendiente
'OOP            id_origen = oOP.getID_CVM_ORIGEN
'OOP            oOP.Eliminar CLng(prmId), ID_EQUIPO, EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
'OOP            oOP.crear_verificacion_pendiente id_origen, rs!ID_VERIFICACION, ID_EQUIPO, New clsEquiposPeriodicidad
'OOP        End If
'OOP    End If
    
    Set oOP = Nothing



'Carga_Verificaciones

End Function

Public Function Eliminar_Norma_equipo(ByVal prmlngid As String, tipo As Integer) As Boolean
Dim strSql As String

    strSql = "DELETE FROM eq_normas_equipos "
    strSql = strSql & " WHERE documento_id = " & CStr(prmlngid)
    strSql = strSql & " AND equipo_id = " & CStr(ID_EQUIPO)
    strSql = strSql & " AND tipo = " & tipo
    
    execute_bd strSql
        
    'Carga_Normas
    
End Function

Public Function Eliminar_LimitacionUso_equipo(ByVal prmlngid As String) As Boolean
Dim strSql As String

    strSql = "DELETE FROM eq_limitaciones_uso "
    strSql = strSql & " WHERE id_limitacion_uso = " & CStr(prmlngid)
    strSql = strSql & " AND equipo_id = " & CStr(ID_EQUIPO)
    
    execute_bd strSql
        
    Cargar_LimitacionesUso
    'eq_limitaciones_uso
End Function


Public Function EliminarParametroResultado(ByVal id_param As String) As Boolean
Dim sql As String

sql = "delete from eq_parametros_resultados where id_parametro = " & id_param

execute_bd sql

End Function

Public Function generar_evento(ByVal prmIdEquipo As Long, ByVal prmEvento As EQUIPOS_EVENTOS, ByVal prmMotivoEvento As EQUIPOS_EVENTOS_MOTIVOS, Optional ByVal prmObservaciones As String = "", Optional ByVal prmIdCVM As Long = 0) As Boolean

    ' Guarda en la base de datos el evento
    ' Dependiendo del evento que sea, realiza acciones a tal efecto

    Dim strSql As String, objEvento As New clsEquipoEventos
    
    Select Case prmEvento
        Case EQUIPOS_EVENTOS.EVT_RECEPCION_EQUIPO
        Case EQUIPOS_EVENTOS.EVT_CALIBRACION_REALIZADA
            If prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_RESULTADO_NEG Then
                ' debe poner el equipo fuera de servicio
                PonerEquipoFueraServicio prmIdEquipo
                prmObservaciones = "Mensaje de Geslab: Las 2 últmas verificaciones fueron NO CONFORMES. Se pondrá el equipo Fuera de Servicio" & vbCrLf & prmObservaciones
                avisar_dpto_equipo_fuera_servicio True, EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
            ElseIf prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_RESULTADO_NEG_1A_VEZ Then
                avisar_dpto_equipo_fuera_servicio False, EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
                prmObservaciones = "Mensaje de Geslab: Es necesario realizar una nueva Calibración. " & vbCrLf & prmObservaciones
            ElseIf prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_RESULTADO_OK Then
                PonerEquipoEnServicio prmIdEquipo
            End If
        Case EQUIPOS_EVENTOS.EVT_VERIFICACION_REALIZADA
            If prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_RESULTADO_NEG Then
                ' debe poner el equipo fuera de servicio
                PonerEquipoFueraServicio prmIdEquipo
                prmObservaciones = "Mensaje de Geslab: Las 2 últmas verificaciones fueron NO CONFORMES. Se pondrá el equipo Fuera de Servicio" & vbCrLf & prmObservaciones
                avisar_dpto_equipo_fuera_servicio True, EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
            ElseIf prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_RESULTADO_NEG_1A_VEZ Then
                prmObservaciones = "Mensaje de Geslab: Es necesario realizar una nueva Verificacion. " & vbCrLf & prmObservaciones
                avisar_dpto_equipo_fuera_servicio False, EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
            ElseIf prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_RESULTADO_OK Then
                PonerEquipoEnServicio prmIdEquipo
            End If
            
        Case EQUIPOS_EVENTOS.EVT_MANTENIMIENTO_REALIZADO
            If prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_RESULTADO_NEG Then
                ' debe poner el equipo fuera de servicio
                PonerEquipoFueraServicio prmIdEquipo
                avisar_dpto_equipo_fuera_servicio True, -1
                prmObservaciones = "Mensaje de Geslab: Se pondrá el equipo Fuera de Servicio" & vbCrLf & prmObservaciones
            End If
        Case EQUIPOS_EVENTOS.EVT_PUESTA_EN_SERVICIO
            ' debe poner el equipo En servicio
            PonerEquipoEnServicio prmIdEquipo
            prmObservaciones = "Mensaje de Geslab: Se pondrá el equipo En Servicio" & vbCrLf & prmObservaciones
        Case EQUIPOS_EVENTOS.EVT_PUESTA_EN_FUERA_DE_SERVICIO
            ' debe poner el equipo fuera de servicio
            PonerEquipoFueraServicio prmIdEquipo
            avisar_dpto_equipo_fuera_servicio True, -1
            prmObservaciones = "Mensaje de Geslab: Se pondrá el equipo Fuera de Servicio" & vbCrLf & prmObservaciones
        Case EQUIPOS_EVENTOS.EVT_ENVIO_DE_EQUIPO
            ' debe poner el equipo fuera de servicio
            PonerEquipoFueraServicio prmIdEquipo
            avisar_dpto_equipo_fuera_servicio True, -1
            prmObservaciones = "Mensaje de Geslab: Se pondrá el equipo Fuera de Servicio" & vbCrLf & prmObservaciones
        Case EQUIPOS_EVENTOS.EVT_CAMBIO_EDICION_DE_EQUIPO
            'En principio no debe hacer nada
        'M1326-I
        Case EQUIPOS_EVENTOS.EVT_APERTURA_CIERRE_EQUIPO
            If prmMotivoEvento = EQUIPOS_EVENTOS_MOTIVOS.EVTR_APERTURA Then
               prmObservaciones = "Mensaje de Geslab: Equipo marcado como Abierto" & vbCrLf & prmObservaciones
            Else
               prmObservaciones = "Mensaje de Geslab: Equipo marcado como Cerrado" & vbCrLf & prmObservaciones
            End If
        'M1326-F
        Case Else
    End Select
    
    With objEvento
        .setCVM_ID = prmIdCVM
        .setEQUIPO_ID = prmIdEquipo
        .setEVENTO_ID = prmEvento
        .setOBSERVACIONES = prmObservaciones
        .setRAZON_ID = prmMotivoEvento
        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
        .Insertar
    End With
End Function

Public Function ImprimirEtiqueta(ID As Long) As Boolean
    Dim prnPrinter As Printer
    ' se mira si el equipo tiene impresora de etiquetas
    Dim oParametro As New clsParametros

   On Error GoTo ImprimirEtiqueta_Error

    Dim oEquipo As New clsEquipos
    oEquipo.carga ID
    log ("Impresion etiqueta equipo")
    Dim PC As String
'    If EQUIPO <> "" Then
'        PC = EQUIPO
'    Else
        PC = USUARIO.getUSO
'    End If
    log ("PC : " & PC)
    
    If PC = "IBERIA" Then
        Dim oEtiqueta As New clsEtiquetas
        With oEtiqueta
           .setTIPO_ID = C_ETIQUETAS_TIPOS.C_ETIQUETAS_TIPOS_EQ
           .setCENTRO_ID = CENTROS.CENTRO_MADRID
           .setID = ID
           .setUSUARIO_ID = oEquipo.getRESPONSABLE_ID
           .Insertar
        End With
        MsgBox "Se han enviado a imprimir las copias.", vbInformation, App.Title
        Exit Function
    End If
    
    If Not oParametro.carga(parametros.IMPRESORA_ETIQUETAS_PEQUENA, PC) Then
        MsgBox "Este equipo no tiene asignada impresora de etiquetas.", vbCritical, App.Title
        Exit Function
    End If
    log ("Parametro impresora : " & Replace(oParametro.getVALOR, "/", "\"))
    
    Dim impresora_encontrada As String
    Dim IMPRESORA As Printer
    impresora_encontrada = ""
    For Each prnPrinter In Printers
'        log "Impresora instalada : " & prnPrinter.DeviceName
        ' Quitamos del nombre de la impresora el #00X para las impresoras TSN del remoto
        Dim impresoraArray() As String
        Dim impresoraNombre As String
        impresoraArray = Split(UCase(prnPrinter.DeviceName), "#")
        impresoraNombre = Trim(impresoraArray(0))
        
'        If UCase(prnPrinter.DeviceName) = Replace(UCase(oParametro.getVALOR), "/", "\") Then
        If impresoraNombre = Replace(UCase(oParametro.getVALOR), "/", "\") Then
            Set Printer = prnPrinter
            Set IMPRESORA = prnPrinter
            impresora_encontrada = prnPrinter.DeviceName
            Exit For
        End If
    Next
    If impresora_encontrada <> "" Then
        Dim imp_ant As String
        imp_ant = Impresora_Predeterminada
        Establecer_Impresora impresora_encontrada
        
        Dim P1() As String
        Dim P2() As String
        ReDim P1(1) As String
        ReDim P2(1) As String
        P1(1) = "FIRMA"
        Dim oUsuario As New clsUsuarios
'        Dim oEquipo As New clsEquipos
'        oEquipo.Carga ID
        oUsuario.CARGAR oEquipo.getRESPONSABLE_ID
        If ReadINI(App.Path + "\config.ini", "Documentos", "firmas") = "c:\geslab\recursos\firmas" Then
            Dim firma As String
            Dim lista() As String
            If oUsuario.getFIRMA <> "" Then
                lista = Split(oUsuario.getFIRMA, "/")
                firma = ReadINI(App.Path + "\config.ini", "Documentos", "firmas") & "\" & lista(UBound(lista))
            End If
            P2(1) = firma
        Else
            P2(1) = Replace(oUsuario.getFIRMA, "/", "\")
        End If
        With frmReport
            .iniciar
            .informe = "\Equipos\rptEquipos_ETIQUETA_EquipoPEQ"
            .criterio = "{equipos.ID_EQUIPO} =" & ID
            .ParametrosNombre = P1
            .ParametrosValores = P2
            .imprimir = True
            .generar
        End With
        Set frmReport = Nothing
        Establecer_Impresora imp_ant
    Else
        MsgBox "No se localiza la impresora de etiquetas.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Function

ImprimirEtiqueta_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ImprimirEtiqueta of Módulo de clase clsEquipos"

End Function

Public Function ImprimirFichaEquipo(ID As Long) As Boolean
    
Dim sql As String
Dim rs As ADODB.Recordset

    If CON_MANTENIMIENTO = 1 Then
        sql = "select * from eq_planes_mantenimiento_equipos where equipo_id = " & ID & " and por_defecto=1"
        Set rs = datos_bd(sql)
        If rs.RecordCount = 0 Then
            MsgBox "El equipo está marcado CON MANTENIMIENTO, pero no tiene ninún Plan de Mantenimiento asignado por defecto." & vbCrLf & "Puede que si lo ha desmarcado ahora, deba guardar los datos del equipo.", vbInformation, "Ficha de Equipo"
            Exit Function
        End If
    End If
    
    
    Dim objRpt As New frmReport
    With objRpt
        .iniciar
        .informe = "Equipos\rptEquipo_FICHA"
'        .CRITERIO = "{equipos.ID_EQUIPO} = " & CStr(ID) & ".00 AND {eq_planes_mantenimiento_equipos.POR_DEFECTO} = 1"
        .criterio = "{equipos.ID_EQUIPO} = " & CStr(ID)
        .imprimir = False
        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objRpt
    Set objRpt = Nothing
End Function

Public Function Listado_asignacion_rapida(RESP As String, FAM As String, filtro As String, ByVal familia As Integer, ByVal tipo As Integer, ByVal Cal_alguno As Integer, ByVal Ver_alguno As Integer, ByVal EsNadcap As Integer, ByVal fabricante As Integer, ByVal proveedor As Integer, ByVal es_accesorio As Integer, ByVal no_es_accesorio As Integer, ByVal localizacion As Integer, ByVal Procedencia As Integer, ByVal Cal_periodo As Integer, ByVal Cal_Responsable As Integer, ByVal Cal_Tipo As Integer, ByVal Ver_periodo As Integer, ByVal Ver_Responsable As Integer, ByVal ver_tipo As Integer) As ADODB.Recordset
    Dim strCad As String
    
    If Trim(filtro) <> "" Then
        filtro = " and concat(id_equipo, lower(nombre), lower(descripcion), lower(serie), lower(modelo), lower(fabricante)) like '%" & LCase(filtro) & "%'"
    End If
    If Cal_alguno = 1 Or Cal_Tipo = 1 Or Cal_Responsable = 1 Or Cal_periodo = 1 Then
        filtro = filtro & " AND equipos.con_calibracion = 1"
    End If
    If Ver_alguno = 1 Or ver_tipo = 1 Or Ver_Responsable = 1 Or Ver_periodo = 1 Then
        filtro = filtro & " AND equipos.con_verificacion = 1"
    End If
    If familia = 1 Then
        filtro = filtro & " AND equipos.familia_id = 0"
    End If
    If tipo = 1 Then
        filtro = filtro & " AND equipos.tipo_equipo_id =0"
    End If
    If EsNadcap = 1 Then
    filtro = filtro & " AND equipos.es_nadcap = 0"
    End If
    If fabricante = 1 Then
        filtro = filtro & " AND trim(equipos.fabricante) = ''"
    End If
    If proveedor = 1 Then
        filtro = filtro & " AND equipos.proveedor_id = 0"
    End If
    If es_accesorio = 1 Then
        filtro = filtro & " AND equipos.es_accesorio = 1"
    End If
    If no_es_accesorio = 1 Then
        filtro = filtro & " AND equipos.es_accesorio = 0"
    End If
    If localizacion = 1 Then
        filtro = filtro & " AND equipos.situacion_id =0"
    End If
    If Procedencia = 1 Then
        filtro = filtro & " AND equipos.PROCEDENCIA_ID=0"
    End If
    If Cal_periodo = 1 Then
        filtro = filtro & " AND equipos.PERIODICIDAD_CALIBRACION_ID <= 0"
    End If
    If Cal_Responsable = 1 Then
        filtro = filtro & " AND equipos.CALIBRADOR_INTERNO_ID <=0"
    End If
    If Cal_Tipo = 1 Then
        filtro = filtro & " AND equipos.TIPO_CALIBRACION_ID <=0"
    End If
    If Ver_periodo = 1 Then
        filtro = filtro & " AND equipos.PERIODICIDAD_VERIFICACION_ID <=0"
    End If
    If Ver_Responsable = 1 Then
        filtro = filtro & " AND equipos.VERIFICADOR_INTERNO_ID <=0"
    End If
    If ver_tipo = 1 Then
        filtro = filtro & " AND equipos.TIPO_VERIFICACION_ID <=0"
    End If
    If Cal_alguno = 1 Then
        filtro = filtro & " AND (equipos.PERIODICIDAD_CALIBRACION_ID <=0 or equipos.CALIBRADOR_INTERNO_ID <=0 or equipos.TIPO_CALIBRACION_ID <=0)"
    End If
    If Ver_alguno = 1 Then
        filtro = filtro & " AND (equipos.PERIODICIDAD_VERIFICACION_ID <=0 or equipos.VERIFICADOR_INTERNO_ID <=0 or equipos.TIPO_VERIFICACION_ID <=0)"
    End If
    If RESP <> "" Then
        If IsNumeric(RESP) Then
            filtro = filtro & " and equipos.RESPONSABLE_ID = " & RESP
        End If
    End If
    If FAM <> "" Then
        If IsNumeric(FAM) Then
            filtro = filtro & " and equipos.FAMILIA_ID = " & FAM
        End If
    End If
    strCad = "select * from equipos where alta_baja = 0 " & filtro & " order by id_equipo"
    Set Listado_asignacion_rapida = datos_bd(strCad)
End Function

Public Sub devolver_filtros_listado_con_fechas_preavisos(ByRef inner_cal As String, ByRef inner_ver As String, ByRef inner_man As String, ByRef campos_cal As String, ByRef campos_ver As String, ByRef campos_man As String)
    inner_cal = " LEFT JOIN eq_periodicidad PER_CAL ON EQUIPOS.PERIODICIDAD_CALIBRACION_ID = PER_CAL.ID_PERIODICIDAD "
    inner_ver = " LEFT JOIN eq_periodicidad PER_VER ON EQUIPOS.PERIODICIDAD_VERIFICACION_ID = PER_VER.ID_PERIODICIDAD "
    inner_man = " LEFT JOIN eq_periodicidad PER_MAN ON EQUIPOS.PERIODICIDAD_MANTENIMIENTO_ID = PER_MAN.ID_PERIODICIDAD "
    
    Dim cal_fecha_preaviso As String
    Dim ver_fecha_preaviso As String
    Dim man_fecha_preaviso As String
    cal_fecha_preaviso = " DATE_SUB(EQUIPOS.FECHA_PROX_CALIBRACION, INTERVAL PER_CAL.DIAS_PREAVISO DAY) "
    ver_fecha_preaviso = " DATE_SUB(EQUIPOS.FECHA_PROX_VERIFICACION, INTERVAL PER_VER.DIAS_PREAVISO DAY) "
    man_fecha_preaviso = " DATE_SUB(EQUIPOS.FECHA_PROX_MANTENIMIENTO, INTERVAL PER_MAN.DIAS_PREAVISO DAY) "
    
    campos_cal = " CASE WHEN EQUIPOS.CON_CALIBRACION <= 0 OR ISNULL(EQUIPOS.FECHA_PROX_CALIBRACION) THEN 1 " & _
                 "      WHEN " & cal_fecha_preaviso & " > LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_CALIBRACION > LOCALTIMESTAMP THEN 2 " & _
                 "      WHEN " & cal_fecha_preaviso & " <= LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_CALIBRACION > LOCALTIMESTAMP THEN 4 " & _
                 "      WHEN EQUIPOS.FECHA_PROX_CALIBRACION <= LOCALTIMESTAMP THEN 3 ELSE 1  " & _
                 " END AS CAL_ESTADO, " & _
                 " EQUIPOS.FECHA_PROX_CALIBRACION AS CAL_FECHA_PREVISTA, " & _
                 " IFNULL(" & cal_fecha_preaviso & ",'') AS CAL_FECHA_PREAVISO, "
                 
    campos_ver = " CASE WHEN EQUIPOS.CON_VERIFICACION <= 0 OR ISNULL(EQUIPOS.FECHA_PROX_VERIFICACION) THEN 1 " & _
                 "      WHEN " & ver_fecha_preaviso & " > LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_VERIFICACION > LOCALTIMESTAMP THEN 2 " & _
                 "      WHEN " & ver_fecha_preaviso & " <= LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_VERIFICACION > LOCALTIMESTAMP THEN 4 " & _
                 "      WHEN EQUIPOS.FECHA_PROX_VERIFICACION <= LOCALTIMESTAMP THEN 3 ELSE 1  " & _
                 " END AS VER_ESTADO, " & _
                 " EQUIPOS.FECHA_PROX_VERIFICACION AS VER_FECHA_PREVISTA, " & _
                 " IFNULL(" & ver_fecha_preaviso & ",'') AS VER_FECHA_PREAVISO, "
                 
    campos_man = " CASE WHEN EQUIPOS.CON_MANTENIMIENTO <= 0 OR ISNULL(EQUIPOS.FECHA_PROX_MANTENIMIENTO) THEN 1 " & _
                 "      WHEN " & man_fecha_preaviso & " > LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_MANTENIMIENTO > LOCALTIMESTAMP THEN 2 " & _
                 "      WHEN " & man_fecha_preaviso & " <= LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_MANTENIMIENTO > LOCALTIMESTAMP THEN 4 " & _
                 "      WHEN EQUIPOS.FECHA_PROX_MANTENIMIENTO <= LOCALTIMESTAMP THEN 3 ELSE 2  " & _
                 " END AS MAN_ESTADO, " & _
                 " EQUIPOS.FECHA_PROX_MANTENIMIENTO AS MAN_FECHA_PREVISTA, " & _
                 " IFNULL(" & man_fecha_preaviso & ",'') AS MAN_FECHA_PREAVISO, "
                 
                 
'    campos_cal = " CASE WHEN EQUIPOS.CON_CALIBRACION = 0 OR EQUIPOS.FECHA_PROX_CALIBRACION = '' OR EQUIPOS.FECHA_PROX_CALIBRACION = '0000-00-00' THEN 1 " & _
'                 "      WHEN DATE_SUB(EQUIPOS.FECHA_PROX_CALIBRACION, INTERVAL PER_CAL.DIAS_PREAVISO DAY) > LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_CALIBRACION > LOCALTIMESTAMP THEN 2 " & _
'                 "      WHEN DATE_SUB(EQUIPOS.FECHA_PROX_CALIBRACION, INTERVAL PER_CAL.DIAS_PREAVISO DAY) <= LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_CALIBRACION > LOCALTIMESTAMP THEN 4 " & _
'                 "      WHEN EQUIPOS.FECHA_PROX_CALIBRACION <= LOCALTIMESTAMP THEN 3 ELSE 1  " & _
'                 " END AS CAL_ESTADO, " & _
'                 " IF(EQUIPOS.FECHA_PROX_CALIBRACION = '0000-00-00','',EQUIPOS.FECHA_PROX_CALIBRACION) AS CAL_FECHA_PREVISTA, " & _
'                 " IFNULL(DATE_SUB(EQUIPOS.FECHA_PROX_CALIBRACION, INTERVAL PER_CAL.DIAS_PREAVISO DAY),'') AS CAL_FECHA_PREAVISO, "
    
'    campos_ver = " CASE WHEN EQUIPOS.CON_VERIFICACION = 0 OR EQUIPOS.FECHA_PROX_VERIFICACION = '' OR EQUIPOS.FECHA_PROX_VERIFICACION = '0000-00-00' THEN 1 " & _
'                 "      WHEN DATE_SUB(EQUIPOS.FECHA_PROX_VERIFICACION, INTERVAL PER_VER.DIAS_PREAVISO DAY) > LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_VERIFICACION > LOCALTIMESTAMP THEN 2 " & _
'                 "      WHEN DATE_SUB(EQUIPOS.FECHA_PROX_VERIFICACION, INTERVAL PER_VER.DIAS_PREAVISO DAY) <= LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_VERIFICACION > LOCALTIMESTAMP THEN 4 " & _
'                 "      WHEN EQUIPOS.FECHA_PROX_VERIFICACION <= LOCALTIMESTAMP THEN 3 ELSE 1  " & _
'                 " END AS VER_ESTADO, " & _
'                 " IF(EQUIPOS.FECHA_PROX_VERIFICACION = '0000-00-00','',EQUIPOS.FECHA_PROX_VERIFICACION) AS VER_FECHA_PREVISTA, " & _
'                 " IFNULL(DATE_SUB(EQUIPOS.FECHA_PROX_VERIFICACION, INTERVAL PER_VER.DIAS_PREAVISO DAY),'') AS VER_FECHA_PREAVISO, "
'
'    campos_man = " CASE WHEN EQUIPOS.CON_MANTENIMIENTO = 0 OR EQUIPOS.FECHA_PROX_MANTENIMIENTO = '' OR EQUIPOS.FECHA_PROX_MANTENIMIENTO = '0000-00-00' THEN 1 " & _
'                 "      WHEN DATE_SUB(EQUIPOS.FECHA_PROX_MANTENIMIENTO, INTERVAL PER_MAN.DIAS_PREAVISO DAY) > LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_MANTENIMIENTO > LOCALTIMESTAMP THEN 2 " & _
'                 "      WHEN DATE_SUB(EQUIPOS.FECHA_PROX_MANTENIMIENTO, INTERVAL PER_MAN.DIAS_PREAVISO DAY) <= LOCALTIMESTAMP AND EQUIPOS.FECHA_PROX_MANTENIMIENTO > LOCALTIMESTAMP THEN 4 " & _
'                 "      WHEN EQUIPOS.FECHA_PROX_MANTENIMIENTO <= LOCALTIMESTAMP THEN 3 ELSE 1  " & _
'                 " END AS MAN_ESTADO, " & _
'                 " IF(EQUIPOS.FECHA_PROX_MANTENIMIENTO = '0000-00-00','',EQUIPOS.FECHA_PROX_MANTENIMIENTO) AS MAN_FECHA_PREVISTA, " & _
'                 " IFNULL(DATE_SUB(EQUIPOS.FECHA_PROX_MANTENIMIENTO, INTERVAL PER_MAN.DIAS_PREAVISO DAY),'') AS MAN_FECHA_PREAVISO, "

'    inner_cal = "  LEFT JOIN ( " & _
'                "   SELECT AA.EQUIPO_ID, MIN(AA.FECHA_PREVISTA) AS FECHA_PREVISTA, MIN(AA.FECHA_PREAVISO) AS FECHA_PREAVISO " & _
'                "     FROM EQ_OPERACIONES_PENDIENTES AA " & _
'                "    WHERE AA.TIPO_CVM_ID = 0 " & _
'                "    GROUP BY AA.EQUIPO_ID " & _
'                "  ) AS CAL ON EQUIPOS.ID_EQUIPO = CAL.EQUIPO_ID "
'    campos_cal = " CASE WHEN CAL.FECHA_PREAVISO  > LOCALTIMESTAMP AND CAL.FECHA_PREVISTA > LOCALTIMESTAMP THEN 2 " & _
'                 "      WHEN CAL.FECHA_PREAVISO <= LOCALTIMESTAMP AND CAL.FECHA_PREVISTA > LOCALTIMESTAMP THEN 4 " & _
'                 "      WHEN CAL.FECHA_PREVISTA <= LOCALTIMESTAMP THEN 3 ELSE 1 END AS CAL_ESTADO, " & _
'                 " CAL.FECHA_PREVISTA AS CAL_FECHA_PREVISTA,CAL.FECHA_PREAVISO AS CAL_FECHA_PREAVISO, "
'    inner_ver = "  LEFT JOIN ( " & _
'                "   SELECT AA.EQUIPO_ID, MIN(AA.FECHA_PREVISTA) AS FECHA_PREVISTA, MIN(AA.FECHA_PREAVISO) AS FECHA_PREAVISO " & _
'                "     FROM EQ_OPERACIONES_PENDIENTES AA " & _
'                "    WHERE AA.TIPO_CVM_ID = 1 " & _
'                "    GROUP BY AA.EQUIPO_ID " & _
'                "  ) AS VER ON EQUIPOS.ID_EQUIPO = VER.EQUIPO_ID "
'    campos_ver = " CASE WHEN VER.FECHA_PREAVISO  > LOCALTIMESTAMP AND VER.FECHA_PREVISTA > LOCALTIMESTAMP THEN 2 " & _
'                 "      WHEN VER.FECHA_PREAVISO <= LOCALTIMESTAMP AND VER.FECHA_PREVISTA > LOCALTIMESTAMP THEN 4 " & _
'                 "      WHEN VER.FECHA_PREVISTA <= LOCALTIMESTAMP THEN 3 ELSE 1 END AS VER_ESTADO, " & _
'                 " VER.FECHA_PREVISTA AS VER_FECHA_PREVISTA,VER.FECHA_PREAVISO AS VER_FECHA_PREAVISO, "
'
'    inner_man = "  LEFT JOIN ( " & _
'                "   SELECT AA.EQUIPO_ID, MIN(AA.FECHA_PREVISTA) AS FECHA_PREVISTA, MIN(AA.FECHA_PREAVISO) AS FECHA_PREAVISO " & _
'                "     FROM EQ_OPERACIONES_PENDIENTES AA " & _
'                "    WHERE AA.TIPO_CVM_ID = 2 " & _
'                "    GROUP BY AA.EQUIPO_ID " & _
'                "  ) AS MAN ON EQUIPOS.ID_EQUIPO = MAN.EQUIPO_ID "
'    campos_man = " CASE WHEN MAN.FECHA_PREAVISO  > LOCALTIMESTAMP AND MAN.FECHA_PREVISTA > LOCALTIMESTAMP THEN 2 " & _
'                 "      WHEN MAN.FECHA_PREAVISO <= LOCALTIMESTAMP AND MAN.FECHA_PREVISTA > LOCALTIMESTAMP THEN 4 " & _
'                 "      WHEN MAN.FECHA_PREVISTA <= LOCALTIMESTAMP THEN 3 ELSE 1 END AS MAN_ESTADO, " & _
'                 " MAN.FECHA_PREVISTA AS MAN_FECHA_PREVISTA,MAN.FECHA_PREAVISO AS MAN_FECHA_PREAVISO, "
End Sub
Public Function ListadoFueraFecha(ByVal prmIdUsuario As Long, SoloMisAvisos As Boolean, tipo As Integer, FUERA_SERVICIO As Boolean) As ADODB.Recordset
    ' TIPO. 0:Todos, 1:CAL, 2:VER, 3:MAN
    Dim strSql As String
    Dim blnUsuarioAutorizado As Boolean
    Dim oParam As New clsParametros
    
    ' Primero Mira en los parámetros del sistema si el usuario está autorizado a verlo todo
   On Error GoTo ListadoFueraFecha_Error

    oParam.carga parametros.USUARIOS_CUADERNO_AVISOS_EQUIPOS, ""
    blnUsuarioAutorizado = False
    
    If InStr(1, Replace(oParam.getVALOR, " ", ""), "," & CStr(prmIdUsuario)) > 0 Or _
       InStr(1, Replace(oParam.getVALOR, " ", ""), CStr(prmIdUsuario) & ",") > 0 Then
        blnUsuarioAutorizado = True
    End If
    
    If Not SoloMisAvisos Then
        blnUsuarioAutorizado = False
    End If
    
    strSql = ""
    If tipo = 0 Or tipo = 1 Then
        strSql = strSql & "SELECT 'C' AS TIPO,b.ID_EQUIPO,b.PERIODICIDAD_CALIBRACION_ID, DATE_SUB(b.FECHA_PROX_CALIBRACION, INTERVAL p.DIAS_PREAVISO DAY) AS FECHA_PREAVISO " & _
                 "       ,b.FECHA_PROX_CALIBRACION as FECHA_PREVISTA,b.NOMBRE,b.RESPONSABLE_ID as ID_EMPLEADO, concat(coalesce(c.nombre,''),' ', coalesce(c.apellidos,'N/A')) as RESPONSABLE_INT " & _
                 "       ,IF(b.ESTADO_ID NOT IN ('B','F/S'),0,1) as FS, b.ESTADO_ID " & _
                 "  FROM equipos b " & _
                 "  LEFT JOIN usuarios c ON b.RESPONSABLE_ID = c.ID_EMPLEADO " & _
                 "  LEFT JOIN eq_periodicidad p ON b.PERIODICIDAD_CALIBRACION_ID = p.ID_PERIODICIDAD " & _
                 " WHERE b.ALTA_BAJA = 0 And b.CON_CALIBRACION = 1 AND CURRENT_DATE >  DATE_SUB(b.FECHA_PROX_CALIBRACION, INTERVAL p.DIAS_PREAVISO DAY) AND b.CLIENTE_ID = 0 "
        If Not blnUsuarioAutorizado Then strSql = strSql & " and b.RESPONSABLE_ID = " & CStr(prmIdUsuario)
        If FUERA_SERVICIO = False Then
            strSql = strSql & " AND b.ESTADO_ID NOT IN ('B','F/S','CAU','E','I','R') "
        End If
    End If
    
    If tipo = 0 Or tipo = 2 Then
        If strSql <> "" Then
            strSql = strSql & " UNION "
        End If
        strSql = strSql & "SELECT 'V' AS TIPO,b.ID_EQUIPO,b.PERIODICIDAD_VERIFICACION_ID, DATE_SUB(b.FECHA_PROX_VERIFICACION, INTERVAL p.DIAS_PREAVISO DAY) AS FECHA_PREAVISO " & _
                 "       ,b.FECHA_PROX_VERIFICACION as FECHA_PREVISTA,b.NOMBRE,b.RESPONSABLE_ID as ID_EMPLEADO, concat(coalesce(c.nombre,''),' ', coalesce(c.apellidos,'N/A')) as RESPONSABLE_INT " & _
                 "       ,IF(b.ESTADO_ID NOT IN ('B','F/S'),0,1) as FS, b.ESTADO_ID " & _
                 "  FROM equipos b " & _
                 "  LEFT JOIN usuarios c ON b.RESPONSABLE_ID = c.ID_EMPLEADO " & _
                 "  LEFT JOIN eq_periodicidad p ON b.PERIODICIDAD_VERIFICACION_ID = p.ID_PERIODICIDAD " & _
                 " WHERE b.ALTA_BAJA = 0 And b.CON_VERIFICACION = 1 AND CURRENT_DATE >  DATE_SUB(b.FECHA_PROX_VERIFICACION, INTERVAL p.DIAS_PREAVISO DAY) AND b.CLIENTE_ID = 0 "
        If Not blnUsuarioAutorizado Then strSql = strSql & " and b.RESPONSABLE_ID = " & CStr(prmIdUsuario)
        If FUERA_SERVICIO = False Then
            strSql = strSql & " AND b.ESTADO_ID NOT IN ('B','F/S','CAU','E','I','R') "
        End If
    End If
    
    If tipo = 0 Or tipo = 3 Then
        If strSql <> "" Then
            strSql = strSql & " UNION "
        End If
        strSql = strSql & "SELECT 'M' AS TIPO,b.ID_EQUIPO,b.PERIODICIDAD_MANTENIMIENTO_ID, DATE_SUB(b.FECHA_PROX_MANTENIMIENTO, INTERVAL IF(isnull(p.DIAS_PREAVISO),1,p.DIAS_PREAVISO) DAY) AS FECHA_PREAVISO " & _
                 "       ,b.FECHA_PROX_MANTENIMIENTO as FECHA_PREVISTA,b.NOMBRE,b.RESPONSABLE_ID as ID_EMPLEADO, concat(coalesce(c.nombre,''),' ', coalesce(c.apellidos,'N/A')) as RESPONSABLE_INT " & _
                 "       ,IF(b.ESTADO_ID NOT IN ('B','F/S'),0,1) as FS, b.ESTADO_ID " & _
                 "  FROM equipos b " & _
                 "  LEFT JOIN usuarios c ON b.RESPONSABLE_ID = c.ID_EMPLEADO " & _
                 "  LEFT JOIN eq_periodicidad p ON b.PERIODICIDAD_MANTENIMIENTO_ID = p.ID_PERIODICIDAD " & _
                 " WHERE b.ALTA_BAJA = 0 And b.CON_MANTENIMIENTO = 1 AND CURRENT_DATE >  DATE_SUB(b.FECHA_PROX_MANTENIMIENTO, INTERVAL IF(isnull(p.DIAS_PREAVISO),1,p.DIAS_PREAVISO) DAY) AND b.CLIENTE_ID = 0 "
        If Not blnUsuarioAutorizado Then strSql = strSql & " and b.RESPONSABLE_ID = " & CStr(prmIdUsuario)
        If FUERA_SERVICIO = False Then
            strSql = strSql & " AND b.ESTADO_ID NOT IN ('B','F/S') "
        End If
    End If
    
    strSql = strSql & " order by FECHA_PREVISTA asc"

'    strSql = "select op.EQUIPO_ID as id_equipo, op.PERIODICIDAD_ID, op.FECHA_PREVISTA, op.ID_CVM as id, op.TIPO_CVM_ID as id_tipo, cvm.descripcion as tipo, op.ID_CVM as id_evento "
'    strSql = strSql & "  , e.nombre , op.RESPONSABLE_ID as id_empleado, concat(coalesce(u.nombre,''),' ', coalesce(u.apellidos,'N/A')) as responsable_int, 0 dias_preaviso, op.FECHA_PREAVISO, e.fuera_servicio as FS "
'    strSql = strSql & "  from eq_operaciones_pendientes op"
'    strSql = strSql & "  inner join decodificadora cvm on op.TIPO_CVM_ID = cvm.VALOR and cvm.CODIGO = 123"
'    strSql = strSql & "  inner join equipos e on op.EQUIPO_ID = e.id_equipo"
'    strSql = strSql & "  left outer join usuarios u on op.RESPONSABLE_ID = u.ID_EMPLEADO"
'    strSql = strSql & "  Where op.fecha_preaviso <= localtimestamp "
'    strSql = strSql & "    and e.alta_baja = 0 "
'    strSql = strSql & sfuera
'    strSql = strSql & sTIPO
'    If Not blnUsuarioAutorizado Then strSql = strSql & " and op.RESPONSABLE_ID = " & CStr(prmIdUsuario)
'    strSql = strSql & "  order by fecha_prevista asc"

        
    Set ListadoFueraFecha = datos_bd(strSql)

   On Error GoTo 0
   Exit Function

ListadoFueraFecha_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ListadoFueraFecha of Módulo de clase clsEquipos"
    
End Function

Public Function ok_rapido_equipos_pendientes_diario(ByVal prmid_equipo As Long, ByVal prmFecha As Date, Optional ByVal prm_id_mantenedor As Long = 0)

'Da el ok Rápido a todos los equipos que tengan diario, hoy pendiente
Dim strCad As String, rs As ADODB.Recordset, strListaExcepciones As String
Dim arrSQL() As String, lngIndice As Long
Dim arrEventoCierre() As String, lngIndiceEventos As Long
Dim arrEquipos() As String, arrMto() As String, arrPlanes() As String
Dim oM As clsEquipoMantenimiento

    ok_rapido_equipos_pendientes_diario = False

    strCad = "SELECT eq_mantenimiento_equipos.*, eq_planes_mantenimiento.FRECUENCIA_ID "
    strCad = strCad & ", equipos.NOMBRE AS EQUIPO "
    strCad = strCad & " from eq_mantenimiento_equipos "
    strCad = strCad & " INNER JOIN eq_planes_mantenimiento ON eq_mantenimiento_equipos.PLANMTO_ID = eq_planes_mantenimiento.ID_PLAN_MTO"
    strCad = strCad & " INNER JOIN equipos ON eq_mantenimiento_equipos.EQUIPO_ID = equipos.ID_EQUIPO"
    strCad = strCad & " Where equipo_id = " & CStr(prmid_equipo)
    'solo busca las que sean de la misma fecha o anterior, independiente de la periodicidad
    'strCad = strCad & " and eq_planes_mantenimiento.FRECUENCIA_ID in(1, 14)"
    strCad = strCad & " AND fecha_actual <= '" & Format(prmFecha, "yyyy-mm-dd") & "'"
    strCad = strCad & " AND ESTADO = 0"
    If prm_id_mantenedor = 0 Then
        strCad = strCad & " AND eq_mantenimiento_equipos.MANTENEDOR_ID = " & USUARIO.getID_EMPLEADO
    Else
        strCad = strCad & " AND eq_mantenimiento_equipos.MANTENEDOR_ID = " & CStr(prm_id_mantenedor)
    End If
    strCad = strCad & " order by eq_mantenimiento_equipos.fecha_actual asc"

    Set rs = datos_bd(strCad)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        Exit Function
    End If
    
    
    Set frmEquiposEdicionMtoListaConformeRapido.rs = rs
    frmEquiposEdicionMtoListaConformeRapido.Show vbModal
    If frmEquiposEdicionMtoListaConformeRapido.RESULTADO Then
        strListaExcepciones = frmEquiposEdicionMtoListaConformeRapido.Excepciones
        Unload frmEquiposEdicionMtoListaConformeRapido
    Else
        Set rs = Nothing
        Exit Function
    End If
    
    lngIndice = 0
    lngIndiceEventos = 0
    
    
    rs.MoveFirst
    While Not rs.EOF
        
        
        If InStr(1, strListaExcepciones, ";" & CStr(rs!id_mantenimiento) & ";") <= 0 Then
            ' Lo primero es cerrarlo conforme, siempre que no esté en las excepciones
            ReDim Preserve arrSQL(lngIndice)
                arrSQL(lngIndice) = "UPDATE eq_mantenimiento_equipos SET estado = 1, FECHA_ACTUAL = '" & Format(prmFecha, "yyyy-mm-dd") & "' WHERE ID_MANTENIMIENTO = " & rs("ID_MANTENIMIENTO")
                lngIndice = lngIndice + 1
                
            
            ReDim Preserve arrSQL(lngIndice)
            ReDim Preserve arrEquipos(lngIndice)
            ReDim Preserve arrPlanes(lngIndice)
            ReDim Preserve arrMto(lngIndice)
                                                
                arrEquipos(lngIndice) = rs("EQUIPO_ID")
                arrPlanes(lngIndice) = rs("PLANMTO_ID")
                arrMto(lngIndice) = rs("ID_MANTENIMIENTO")
                
                arrSQL(lngIndice) = "insert into eq_mantenimiento_equipos_acciones (PLAN_MTO_ID, ACCION_ID, MANTENIMIENTO_ID) "
                arrSQL(lngIndice) = arrSQL(lngIndice) & " SELECT eq_planes_mantenimiento_detalle.PLAN_MTO_ID"
                arrSQL(lngIndice) = arrSQL(lngIndice) & " , eq_planes_mantenimiento_detalle.ACCION_ID AS ID_ACCION"
                arrSQL(lngIndice) = arrSQL(lngIndice) & " , " & rs("ID_MANTENIMIENTO") & " AS MANTENIMIENTO_ID"
                arrSQL(lngIndice) = arrSQL(lngIndice) & " FROM eq_planes_mantenimiento_detalle WHERE eq_planes_mantenimiento_detalle.PLAN_MTO_ID = " & rs("PLANMTO_ID")
                
                lngIndice = lngIndice + 1
        End If
        
        rs.MoveNext
    Wend
    
    Set oM = New clsEquipoMantenimiento
    
    For lngIndice = 0 To UBound(arrSQL)
        ok_rapido_equipos_pendientes_diario = True
        execute_bd arrSQL(lngIndice)
        If lngIndice Mod 2 <> 0 Then
            ' revisa las tareas pendientes
            oM.revisar_mantenimiento_pendiente_cuaderno CLng(arrEquipos(lngIndice)), CLng(arrPlanes(lngIndice)), CLng(arrMto(lngIndice))
        End If
        
    Next lngIndice
    


End Function

'Public Function sqlSelAccesorios() As String
'On Error GoTo sqlSelAccesorios_Error
'
'    sqlSelAccesorios = "SELECT ID_ACCESORIO AS ID, NOMBRE AS 'Accesorio' from eq_accesorios where 1=1 "
'
'On Error GoTo 0
'    Exit Function
'sqlSelAccesorios_Error:
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sqlSelAccesorios of Módulo de clase clsEquipos"
'    sqlSelAccesorios = ""
'End Function

'Public Function sqlSelProcedimientos() As String
'On Error GoTo sqlSelProcedimientos_Error
'
'    sqlSelProcedimientos = "select id_norma as id, concat(codigo, ': ', nombre) as 'Procedimiento' from ca_normas where 1=1"
'
'On Error GoTo 0
'    Exit Function
'sqlSelProcedimientos_Error:
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sqlSelAccesorios of Módulo de clase clsEquipos"
'    sqlSelProcedimientos = ""
'End Function


'Private Sub Class_Terminate()

'    Set mvarobjAdjuntosDocumentacion = Nothing

'J5    Set mvarobjAdjuntosInstruccionesTecnicas = Nothing
'    Set mvarobjCalibraciones = Nothing
'    Set mvarobjVerificaciones = Nothing
'    Set mvarobjMantenimientos = Nothing
'    Set mvarobjNormas = Nothing
    
'End Sub
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Property Let setFUERA_SERVICIO(ByVal dato As Integer)
    FUERA_SERVICIO = dato
End Property

Public Property Let setMOSTRAR_EN_PLANO(ByVal dato As Integer)
    MOSTRAR_EN_PLANO = dato
End Property
Public Property Get getMOSTRAR_EN_PLANO() As Integer
    getMOSTRAR_EN_PLANO = MOSTRAR_EN_PLANO
End Property
Public Property Let setPRIORITARIO(ByVal dato As Integer)
    PRIORITARIO = dato
End Property
Public Property Get getPRIORITARIO() As Integer
    getPRIORITARIO = PRIORITARIO
End Property
Public Property Let setINSITU(ByVal dato As Integer)
    INSITU = dato
End Property
Public Property Get getINSITU() As Integer
    getINSITU = INSITU
End Property
Public Property Let setID_EQUIPO(ByVal dato As Long)
        ID_EQUIPO = dato
End Property
Public Property Let setPARTICULAR(ByVal dato As Long)
        particular = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setCONDICIONES_AMBIENTALES(ByVal dato As Long)
        CONDICIONES_AMBIENTALES = dato
End Property
Public Property Let setPROVEEDOR_ID(ByVal dato As Integer)
        PROVEEDOR_ID = dato
End Property
Public Property Let setFECHA_RECEPCION(ByVal dato As String)
        FECHA_RECEPCION = dato
End Property
Public Property Let setFECHA_SERVICIO(ByVal dato As String)
        FECHA_SERVICIO = dato
End Property
Public Property Let setALTA_BAJA(ByVal dato As Integer)
        ALTA_BAJA = dato
End Property
'M1124-I
Public Property Let setESTADO_ID(ByVal dato As String)
        ESTADO_ID = dato
End Property
'M1124-F
Public Property Let setSERIE(ByVal dato As String)
        SERIE = dato
End Property
Public Property Let setNOTAS(ByVal dato As String)
        NOTAS = dato
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Long)
        FAMILIA_ID = dato
End Property
Public Property Let setMODELO(ByVal dato As String)
        MODELO = dato
End Property
Public Property Let setFABRICANTE(ByVal dato As String)
        fabricante = dato
End Property
Public Property Let setES_NADCAP(ByVal dato As String)
        ES_NADCAP = dato
End Property
Public Property Let setES_MTL(ByVal dato As String)
        ES_MTL = dato
End Property
Public Property Let setES_CP(ByVal dato As String)
        ES_CP = dato
End Property
Public Property Let setES_ENAC(ByVal dato As String)
        ES_ENAC = dato
End Property
Public Property Let setES_ACCESORIO(ByVal dato As Integer)
        es_accesorio = dato
End Property
Public Property Let setNUMERO_USOS_MAXIMO(ByVal dato As Integer)
        NUMERO_USOS_MAXIMO = dato
End Property
Public Property Let setNUMERO_USOS_CONTADOR(ByVal dato As Integer)
        NUMERO_USOS_CONTADOR = dato
End Property
'M1050-I
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setTRAZMET_ID(ByVal dato As Long)
        TRAZMET_ID = dato
End Property
Public Property Let setNUMERO_EQUIPO_CLIENTE(ByVal dato As String)
        NUMERO_EQUIPO_CLIENTE = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'M1050-F
Public Property Let setCRITICO(ByVal dato As String)
        CRITICO = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Integer)
        CENTRO_ID = dato
End Property

Public Property Let setRANGO_MEDIDA_MIN(ByVal dato As String)
        RANGO_MEDIDA_MIN = dato
End Property
Public Property Let setRANGO_MEDIDA_MAX(ByVal dato As String)
        RANGO_MEDIDA_MAX = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As String)
        UNIDAD_ID = dato
End Property
Public Property Let setTOLERANCIA_MAXIMA(ByVal dato As String)
        TOLERANCIA_MAXIMA = dato
End Property
Public Property Let setINCERTIDUMBRE_MAXIMA(ByVal dato As String)
        INCERTIDUMBRE_MAXIMA = dato
End Property
Public Property Let setPRECISIONN(ByVal dato As String)
        PRECISIONN = dato
End Property
Public Property Let setTEMPERATURA_MIN(ByVal dato As String)
        TEMPERATURA_MIN = dato
End Property
Public Property Let setTEMPERATURA_MAX(ByVal dato As String)
        TEMPERATURA_MAX = dato
End Property
Public Property Let setHUMEDAD_MIN(ByVal dato As String)
        HUMEDAD_MIN = dato
End Property
Public Property Let setHUMEDAD_MAX(ByVal dato As String)
        HUMEDAD_MAX = dato
End Property
Public Property Let setCOND_AMBIENTALES_OTRAS(ByVal dato As String)
        COND_AMBIENTALES_OTRAS = dato
End Property
Public Property Let setRANGO_TRABAJO_MIN(ByVal dato As String)
        RANGO_TRABAJO_MIN = dato
End Property
Public Property Let setRANGO_TRABAJO_MAX(ByVal dato As String)
        RANGO_TRABAJO_MAX = dato
End Property
Public Property Let setLIMITACIONES_USO(ByVal dato As String)
        LIMITACIONES_USO = dato
End Property
Public Property Set setLIMITACIONES_USO_COL(ByVal objCol As clsGenericCollection)
        Set LIMITACIONES_USO_COL = objCol
End Property

Public Property Let setPERIODICIDAD_CALIBRACION_ID(ByVal dato As String)
        PERIODICIDAD_CALIBRACION_ID = dato
End Property
Public Property Let setFECHA_ULT_CALIBRACION(ByVal dato As String)
        FECHA_ULT_CALIBRACION = dato
End Property
Public Property Let setFECHA_PROX_CALIBRACION(ByVal dato As String)
        FECHA_PROX_CALIBRACION = dato
End Property
Public Property Let setPERIODICIDAD_VERIFICACION_ID(ByVal dato As String)
        PERIODICIDAD_VERIFICACION_ID = dato
End Property
Public Property Let setFECHA_ULT_VERIFICACION(ByVal dato As String)
        FECHA_ULT_VERIFICACION = dato
End Property
Public Property Let setFECHA_PROX_VERIFICACION(ByVal dato As String)
        FECHA_PROX_VERIFICACION = dato
End Property
Public Property Let setPERIODICIDAD_MANTENIMIENTO_ID(ByVal dato As String)
        PERIODICIDAD_MANTENIMIENTO_ID = dato
End Property
Public Property Let setFECHA_PROX_MANTENIMIENTO(ByVal dato As String)
        FECHA_PROX_MANTENIMIENTO = dato
End Property
Public Property Let setFECHA_ULT_MANTENIMIENTO(ByVal dato As String)
        FECHA_ULT_MANTENIMIENTO = dato
End Property
Public Property Let setTIPO_CALIBRACION_ID(ByVal dato As Integer)
        TIPO_CALIBRACION_ID = dato
End Property

Public Property Let setTIPO_VERIFICACION_ID(ByVal dato As Integer)
        TIPO_VERIFICACION_ID = dato
End Property

Public Property Let setCALIBRADOR_ID(ByVal dato As Integer)
        CALIBRADOR_ID = dato
End Property
Public Property Let setCALIBRADOR_INTERNO_ID(ByVal dato As Integer)
        CALIBRADOR_INTERNO_ID = dato
End Property

Public Property Let setVERIFICADOR_ID(ByVal dato As Integer)
        VERIFICADOR_ID = dato
End Property
Public Property Let setVERIFICADOR_INTERNO_ID(ByVal dato As Integer)
        VERIFICADOR_INTERNO_ID = dato
End Property

Public Property Let setSITUACION_ID(ByVal dato As Long)
        SITUACION_ID = dato
End Property
Public Property Let setSITUACION_DESCRIPCION(ByVal dato As String)
        SITUACION_DESCRIPCION = dato
End Property


Public Property Let setCON_VERIFICACION(ByVal dato As Integer)
        CON_VERIFICACION = dato
End Property
Public Property Let setCON_CALIBRACION(ByVal dato As Integer)
        CON_CALIBRACION = dato
End Property
Public Property Let setCON_MANTENIMIENTO(ByVal dato As Integer)
        CON_MANTENIMIENTO = dato
End Property
Public Property Let setREVISADO(ByVal dato As String)
        REVISADO = dato
End Property
Public Property Let setTIPO_EQUIPO(ByVal dato As String)
        TIPO_EQUIPO = dato
End Property
Public Property Let setTIPO_EQUIPO_ID(ByVal dato As Integer)
        TIPO_EQUIPO_ID = dato
End Property

Public Property Let setCoordX(ByVal dato As Long)
        COORDX = dato
End Property
Public Property Let setCoordY(ByVal dato As Long)
        COORDY = dato
End Property
Public Property Let setID_AUX(ByVal dato As Integer)
        ID_AUX = dato
End Property
Public Property Set setFAMILIA(ByVal obj As clsFamiliasEquipos)
        Set familia = obj
End Property '***************************************************************
Public Property Let setPROCEDENCIA_ID(ByVal dato As Long)
        PROCEDENCIA_ID = dato
End Property '***************************************************************

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Property Get getFUERA_SERVICIO() As Integer
    getFUERA_SERVICIO = FUERA_SERVICIO
End Property

Public Property Get getID_EQUIPO() As Long
        getID_EQUIPO = ID_EQUIPO
End Property
Public Property Get getPARTICULAR() As Long
        getPARTICULAR = particular
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property

Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property

Public Property Get getTIPO_EQUIPO_ID() As Integer
        getTIPO_EQUIPO_ID = TIPO_EQUIPO_ID
End Property
Public Property Get getTIPO_EQUIPO() As String
        getTIPO_EQUIPO = TIPO_EQUIPO
End Property


Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property
Public Property Get getPROCEDENCIA_ID() As Long
        getPROCEDENCIA_ID = PROCEDENCIA_ID
End Property

Public Property Get getCONDICIONES_AMBIENTALES() As Long
        getCONDICIONES_AMBIENTALES = CONDICIONES_AMBIENTALES
End Property
Public Property Get getPROVEEDOR_ID() As Integer
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getFECHA_RECEPCION() As String
        getFECHA_RECEPCION = FECHA_RECEPCION
End Property
Public Property Get getFECHA_SERVICIO() As String
        getFECHA_SERVICIO = FECHA_SERVICIO
End Property
Public Property Get getALTA_BAJA() As Integer
        getALTA_BAJA = ALTA_BAJA
End Property
'M1124-I
Public Property Get getESTADO_ID() As String
        getESTADO_ID = ESTADO_ID
End Property
'M1124-F
Public Property Get getSERIE() As String
        getSERIE = SERIE
End Property
Public Property Get getNOTAS() As String
        getNOTAS = NOTAS
End Property
Public Property Get getFAMILIA_ID() As Integer
        getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getMODELO() As String
        getMODELO = MODELO
End Property
Public Property Get getFABRICANTE() As String
        getFABRICANTE = fabricante
End Property
Public Property Get getES_NADCAP() As Integer
        getES_NADCAP = ES_NADCAP
End Property
Public Property Get getES_MTL() As Integer
        getES_MTL = ES_MTL
End Property
Public Property Get getES_CP() As Integer
        getES_CP = ES_CP
End Property
Public Property Get getES_ENAC() As Integer
        getES_ENAC = ES_ENAC
End Property
Public Property Get getES_ACCESORIO() As Integer
        getES_ACCESORIO = es_accesorio
End Property
Public Property Get getNUMERO_USOS_MAXIMO() As Integer
        getNUMERO_USOS_MAXIMO = NUMERO_USOS_MAXIMO
End Property
Public Property Get getNUMERO_USOS_CONTADOR() As Integer
        getNUMERO_USOS_CONTADOR = NUMERO_USOS_CONTADOR
End Property
'M1050-I
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getTRAZMET_ID() As Long
        getTRAZMET_ID = TRAZMET_ID
End Property
Public Property Get getNUMERO_EQUIPO_CLIENTE() As String
        getNUMERO_EQUIPO_CLIENTE = NUMERO_EQUIPO_CLIENTE
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'M1050-F
Public Property Get getCRITICO() As Integer
        getCRITICO = CRITICO
End Property
Public Property Get getCENTRO_ID() As Integer
        getCENTRO_ID = CENTRO_ID
End Property

Public Property Get getRANGO_MEDIDA_MIN() As String
        getRANGO_MEDIDA_MIN = RANGO_MEDIDA_MIN
End Property
Public Property Get getRANGO_MEDIDA_MAX() As String
        getRANGO_MEDIDA_MAX = RANGO_MEDIDA_MAX
End Property
Public Property Get getUNIDAD_ID() As Integer
        getUNIDAD_ID = UNIDAD_ID
End Property
Public Property Get getTOLERANCIA_MAXIMA() As String
        getTOLERANCIA_MAXIMA = TOLERANCIA_MAXIMA
End Property
Public Property Get getINCERTIDUMBRE_MAXIMA() As String
        getINCERTIDUMBRE_MAXIMA = INCERTIDUMBRE_MAXIMA
End Property
Public Property Get getPRECISIONN() As String
        getPRECISIONN = PRECISIONN
End Property
Public Property Get getTEMPERATURA_MIN() As String
        getTEMPERATURA_MIN = TEMPERATURA_MIN
End Property
Public Property Get getTEMPERATURA_MAX() As String
        getTEMPERATURA_MAX = TEMPERATURA_MAX
End Property
Public Property Get getHUMEDAD_MIN() As String
        getHUMEDAD_MIN = HUMEDAD_MIN
End Property
Public Property Get getHUMEDAD_MAX() As String
        getHUMEDAD_MAX = HUMEDAD_MAX
End Property
Public Property Get getCOND_AMBIENTALES_OTRAS() As String
        getCOND_AMBIENTALES_OTRAS = COND_AMBIENTALES_OTRAS
End Property
Public Property Get getRANGO_TRABAJO_MIN() As String
        getRANGO_TRABAJO_MIN = RANGO_TRABAJO_MIN
End Property
Public Property Get getRANGO_TRABAJO_MAX() As String
        getRANGO_TRABAJO_MAX = RANGO_TRABAJO_MAX
End Property
Public Property Get getLIMITACIONES_USO() As String
        getLIMITACIONES_USO = LIMITACIONES_USO
End Property
Public Property Get getLIMITACIONES_USO_COL() As clsGenericCollection
        Set getLIMITACIONES_USO_COL = LIMITACIONES_USO_COL
End Property


Public Property Get getPERIODICIDAD_CALIBRACION_ID() As Integer
        getPERIODICIDAD_CALIBRACION_ID = PERIODICIDAD_CALIBRACION_ID
End Property
Public Property Get getFECHA_ULT_CALIBRACION() As String
        getFECHA_ULT_CALIBRACION = FECHA_ULT_CALIBRACION
End Property
Public Property Get getFECHA_PROX_CALIBRACION() As String
        getFECHA_PROX_CALIBRACION = FECHA_PROX_CALIBRACION
End Property
Public Property Get getPROCEDIMIENTO_CALIBRACION_ID() As Long
        getPROCEDIMIENTO_CALIBRACION_ID = PROCEDIMIENTO_CALIBRACION_ID
End Property
Public Property Get getPROCEDIMIENTO_VERIFICACION_ID() As Long
        getPROCEDIMIENTO_VERIFICACION_ID = PROCEDIMIENTO_VERIFICACION_ID
End Property
Public Property Let setPROCEDIMIENTO_VERIFICACION_ID(ByVal dato As Long)
        PROCEDIMIENTO_VERIFICACION_ID = dato
End Property
Public Property Let setPROCEDIMIENTO_CALIBRACION_ID(ByVal dato As Long)
        PROCEDIMIENTO_CALIBRACION_ID = dato
End Property

Public Property Get getPERIODICIDAD_VERIFICACION_ID() As Integer
        getPERIODICIDAD_VERIFICACION_ID = PERIODICIDAD_VERIFICACION_ID
End Property
Public Property Get getFECHA_PROX_VERIFICACION() As String
        getFECHA_PROX_VERIFICACION = FECHA_PROX_VERIFICACION
End Property
Public Property Get getFECHA_ULT_VERIFICACION() As String
        getFECHA_ULT_VERIFICACION = FECHA_ULT_VERIFICACION
End Property
Public Property Get getPERIODICIDAD_MANTENIMIENTO_ID() As Integer
        getPERIODICIDAD_MANTENIMIENTO_ID = PERIODICIDAD_MANTENIMIENTO_ID
End Property
Public Property Get getFECHA_PROX_MANTENIMIENTO() As String
        getFECHA_PROX_MANTENIMIENTO = FECHA_PROX_MANTENIMIENTO
End Property
Public Property Get getFECHA_ULT_MANTENIMIENTO() As String
        getFECHA_ULT_MANTENIMIENTO = FECHA_ULT_MANTENIMIENTO
End Property
Public Property Get getTIPO_CALIBRACION_ID() As Integer
        getTIPO_CALIBRACION_ID = TIPO_CALIBRACION_ID
End Property
Public Property Get getTIPO_VERIFICACION_ID() As Integer
        getTIPO_VERIFICACION_ID = TIPO_VERIFICACION_ID
End Property

Public Property Get getCALIBRADOR_ID() As Integer
        getCALIBRADOR_ID = CALIBRADOR_ID
End Property
Public Property Get getCALIBRADOR_INTERNO_ID() As Integer
        getCALIBRADOR_INTERNO_ID = CALIBRADOR_INTERNO_ID
End Property

Public Property Get getVERIFICADOR_ID() As Integer
        getVERIFICADOR_ID = VERIFICADOR_ID
End Property
Public Property Get getVERIFICADOR_INTERNO_ID() As Integer
        getVERIFICADOR_INTERNO_ID = VERIFICADOR_INTERNO_ID
End Property

Public Property Get getSITUACION_ID() As Integer
        getSITUACION_ID = SITUACION_ID
End Property
Public Property Get getSITUACION_DESCRIPCION() As String
        getSITUACION_DESCRIPCION = SITUACION_DESCRIPCION
End Property
Public Property Get getCON_CALIBRACION() As Integer
        getCON_CALIBRACION = CON_CALIBRACION
End Property
Public Property Get getCON_VERIFICACION() As Integer
        getCON_VERIFICACION = CON_VERIFICACION
End Property
Public Property Get getCON_MANTENIMIENTO() As Integer
        getCON_MANTENIMIENTO = CON_MANTENIMIENTO
End Property
Public Property Get getREVISADO() As Integer
        getREVISADO = REVISADO
End Property

Public Property Get getCoordx() As Long
        getCoordx = COORDX
End Property
Public Property Get getCoordY() As Long
        getCoordY = COORDY
End Property

Public Property Get getID_AUX() As Integer
        getID_AUX = ID_AUX
End Property

Public Property Get getFAMILIA() As clsFamiliasEquipos
    Set getFAMILIA = familia
End Property


Public Property Get getPLAN_MANTENIMIENTO_ID() As Long
    getPLAN_MANTENIMIENTO_ID = PLAN_MANTENIMIENTO_ID
End Property

Public Property Let setPLAN_MANTENIMIENTO_ID(ByVal dato As Long)
    PLAN_MANTENIMIENTO_ID = dato
End Property

Public Property Get getMANTENEDOR_ID() As Long
    getMANTENEDOR_ID = MANTENEDOR_ID
End Property

Public Property Let setMANTENEDOR_ID(ByVal dato As Long)
    MANTENEDOR_ID = dato
End Property


Public Property Get getACCESORIOS_COL() As clsGenericCollection
    Set getACCESORIOS_COL = ACCESORIOS_COL
End Property

Public Property Set setACCESORIOS_COL(ByVal obj As clsGenericCollection)
    Set ACCESORIOS_COL = obj
End Property
'M0459-I
Public Property Get getUNIDADES_LOTE() As Integer
        getUNIDADES_LOTE = UNIDADES_LOTE
End Property
Public Property Let setUNIDADES_LOTE(ByVal dato As Integer)
    UNIDADES_LOTE = dato
End Property

'M0459-F

'Public Property Get Normas() As clsGenericCollection
'
'    Set Normas = mvarobjNormas
'
'End Property

Public Function listado_normas_equipos() As ADODB.Recordset

    Dim consulta As String
    
    
    consulta = "SELECT en.documento_id, ca.nombre,ca.ruta,en.tipo,ca.anulado " & _
               "  from eq_normas_equipos en " & _
               " inner join ca_documentos ca on en.documento_id = ca.id_documento " & _
               " where en.equipo_id = " & CStr(ID_EQUIPO) & " and en.TIPO = 0 " & _
               " union " & _
               "SELECT en.documento_id, ca.nombre,ca.ruta,en.tipo,0 " & _
               "  from eq_normas_equipos en " & _
               " inner join ca_normas ca on en.documento_id = ca.id_norma " & _
               " where en.equipo_id = " & CStr(ID_EQUIPO) & " and en.TIPO = 1 "
        
    Set listado_normas_equipos = datos_bd(consulta)
    
End Function

'Public Property Set Normas(objNormas As clsGenericCollection)
'
'    Set mvarobjNormas = objNormas
'
'End Property


'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Function carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQUIPOS WHERE ID_EQUIPO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                carga = False
        Else
                ID_EQUIPO = rs("ID_EQUIPO")
                particular = rs("PARTICULAR")
                nombre = rs("NOMBRE")
                'M0459-I
                UNIDADES_LOTE = rs("UNIDADES_LOTE")
                'M0459-F
                CONDICIONES_AMBIENTALES = rs("CONDICIONES_AMBIENTALES")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                FECHA_RECEPCION = rs("FECHA_RECEPCION")
                FECHA_SERVICIO = rs("FECHA_SERVICIO")
                ALTA_BAJA = rs("ALTA_BAJA")
                'M1124-I
                ESTADO_ID = rs("ESTADO_ID")
                'M1124-F
                SERIE = rs("SERIE")
                FAMILIA_ID = rs("FAMILIA_ID")
                MODELO = rs("MODELO")
                fabricante = rs("FABRICANTE")
                ES_NADCAP = rs("ES_NADCAP")
                ES_MTL = rs("ES_MTL")
                ES_CP = rs("ES_CP")
                ES_ENAC = rs("ES_ENAC")
                RANGO_MEDIDA_MIN = rs("RANGO_MEDIDA_MIN")
                RANGO_MEDIDA_MAX = rs("RANGO_MEDIDA_MAX")
                UNIDAD_ID = rs("UNIDAD_ID")
                TOLERANCIA_MAXIMA = rs("TOLERANCIA_MAXIMA")
                INCERTIDUMBRE_MAXIMA = rs("INCERTIDUMBRE_MAXIMA")
                PRECISIONN = rs("PRECISIONN")
                TEMPERATURA_MIN = rs("TEMPERATURA_MIN")
                TEMPERATURA_MAX = rs("TEMPERATURA_MAX")
                HUMEDAD_MIN = rs("HUMEDAD_MIN")
                HUMEDAD_MAX = rs("HUMEDAD_MAX")
                COND_AMBIENTALES_OTRAS = rs("COND_AMBIENTALES_OTRAS")
                RANGO_TRABAJO_MIN = rs("RANGO_TRABAJO_MIN")
                RANGO_TRABAJO_MAX = rs("RANGO_TRABAJO_MAX")
                LIMITACIONES_USO = rs("LIMITACIONES_USO")
                PERIODICIDAD_CALIBRACION_ID = rs("PERIODICIDAD_CALIBRACION_ID")
                If Not IsNull(rs("FECHA_ULT_CALIBRACION")) Then
                    FECHA_ULT_CALIBRACION = rs("FECHA_ULT_CALIBRACION")
                End If
                If Not IsNull(rs("FECHA_PROX_CALIBRACION")) Then
                    FECHA_PROX_CALIBRACION = rs("FECHA_PROX_CALIBRACION")
                End If
                PERIODICIDAD_VERIFICACION_ID = rs("PERIODICIDAD_VERIFICACION_ID")
                If Not IsNull(rs("FECHA_ULT_VERIFICACION")) Then
                    FECHA_ULT_VERIFICACION = rs("FECHA_ULT_VERIFICACION")
                End If
                If Not IsNull(rs("FECHA_PROX_VERIFICACION")) Then
                    FECHA_PROX_VERIFICACION = rs("FECHA_PROX_VERIFICACION")
                End If
                PERIODICIDAD_MANTENIMIENTO_ID = rs("PERIODICIDAD_MANTENIMIENTO_ID")
                If Not IsNull(rs("FECHA_ULT_MANTENIMIENTO")) Then
                    FECHA_ULT_MANTENIMIENTO = rs("FECHA_ULT_MANTENIMIENTO")
                End If
                If Not IsNull(rs("FECHA_PROX_MANTENIMIENTO")) Then
                    FECHA_PROX_MANTENIMIENTO = rs("FECHA_PROX_MANTENIMIENTO")
                End If
                SITUACION_ID = rs("SITUACION_ID")
                If Not IsNull(rs("SITUACION_DESCRIPCION")) Then
                    SITUACION_DESCRIPCION = rs("SITUACION_DESCRIPCION")
                End If
                TIPO_CALIBRACION_ID = rs("TIPO_CALIBRACION_ID")
                CALIBRADOR_ID = rs("CALIBRADOR_ID")
                CALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
                CON_CALIBRACION = rs("CON_CALIBRACION")
                CON_VERIFICACION = rs("CON_VERIFICACION")
                CON_MANTENIMIENTO = rs("CON_MANTENIMIENTO")
                REVISADO = rs("REVISADO")
                If Not IsNull(rs("NOTAS")) Then
                    NOTAS = rs("NOTAS")
                End If
                es_accesorio = 0
                NUMERO_USOS_MAXIMO = rs("NUMERO_USOS_MAXIMO")
                NUMERO_USOS_CONTADOR = rs("NUMERO_USOS_CONTADOR")
                'M1050-I
                CLIENTE_ID = rs("CLIENTE_ID")
                
                TRAZMET_ID = rs("TRAZMET_ID")
                If Not IsNull(rs("NUMERO_EQUIPO_CLIENTE")) Then
                    NUMERO_EQUIPO_CLIENTE = rs("NUMERO_EQUIPO_CLIENTE")
                End If
                If Not IsNull(rs("OBSERVACIONES")) Then
                    OBSERVACIONES = rs("OBSERVACIONES")
                End If
                'M1050-F
                CRITICO = rs("CRITICO")
                CENTRO_ID = rs("CENTRO_ID")
                COORDX = 0
                COORDY = 0
                DESCRIPCION = ""
                TIPO_EQUIPO_ID = 0
                RESPONSABLE_ID = 0
                PROCEDENCIA_ID = 0
                TIPO_VERIFICACION_ID = 0
                VERIFICADOR_ID = 0
                VERIFICADOR_INTERNO_ID = 0
                PLAN_MANTENIMIENTO_ID = 0
                MANTENEDOR_ID = 0
                MOSTRAR_EN_PLANO = 0
                PRIORITARIO = 0
                INSITU = 0
                FUERA_SERVICIO = 0
                
                If Not IsNull(rs("COORDX")) Then COORDX = rs("COORDX")
                If Not IsNull(rs("COORDY")) Then COORDY = rs("COORDY")
                If Not IsNull(rs("DESCRIPCION")) Then DESCRIPCION = rs("DESCRIPCION")
                If Not IsNull(rs("TIPO_EQUIPO_ID")) Then TIPO_EQUIPO_ID = rs("TIPO_EQUIPO_ID")
                If Not IsNull(rs("RESPONSABLE_ID")) Then RESPONSABLE_ID = rs("RESPONSABLE_ID")
                If Not IsNull(rs("PROCEDENCIA_ID")) Then PROCEDENCIA_ID = rs("PROCEDENCIA_ID")
                If Not IsNull(rs("TIPO_VERIFICACION_ID")) Then TIPO_VERIFICACION_ID = rs("TIPO_VERIFICACION_ID")
                If Not IsNull(rs("VERIFICADOR_ID")) Then VERIFICADOR_ID = rs("VERIFICADOR_ID")
                If Not IsNull(rs("VERIFICADOR_INTERNO_ID")) Then VERIFICADOR_INTERNO_ID = rs("VERIFICADOR_INTERNO_ID")
                If Not IsNull(rs("PLAN_MANTENIMIENTO_ID")) Then PLAN_MANTENIMIENTO_ID = rs("PLAN_MANTENIMIENTO_ID")
                If Not IsNull(rs("MANTENEDOR_ID")) Then MANTENEDOR_ID = rs("MANTENEDOR_ID")
                If Not IsNull(rs("MOSTRAR_EN_PLANO")) Then MOSTRAR_EN_PLANO = rs("MOSTRAR_EN_PLANO")
                If Not IsNull(rs("PRIORITARIO")) Then PRIORITARIO = rs("PRIORITARIO")
                If Not IsNull(rs("INSITU")) Then INSITU = rs("INSITU")
                If Not IsNull(rs("FUERA_SERVICIO")) Then FUERA_SERVICIO = rs("FUERA_SERVICIO")
                If Not IsNull(rs("ES_ACCESORIO")) Then es_accesorio = CInt(rs("ES_ACCESORIO"))
                
                If IsNull(rs("PROCEDIMIENTO_CALIBRACION_ID")) Then
                    PROCEDIMIENTO_CALIBRACION_ID = 0
                Else
                    PROCEDIMIENTO_CALIBRACION_ID = rs("PROCEDIMIENTO_CALIBRACION_ID")
                End If
                If IsNull(rs("PROCEDIMIENTO_VERIFICACION_ID")) Then
                    PROCEDIMIENTO_VERIFICACION_ID = 0
                Else
                    PROCEDIMIENTO_VERIFICACION_ID = rs("PROCEDIMIENTO_VERIFICACION_ID")
                End If
                
                Set familia = New clsFamiliasEquipos
                Call familia.carga(FAMILIA_ID)
                
                Call Cargar_LimitacionesUso
'M1050                Call cargar_accesorios
'M1050                Call Carga_Normas
                
                carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        carga = False
        'Resume Next
        MsgBox "Error al cargar los datos(clsEquipos)", vbCritical, Err.Description
End Function

Public Sub Cargar_LimitacionesUso()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim objLU As clsGenericClass
    
    Set LIMITACIONES_USO_COL = New clsGenericCollection
    LIMITACIONES_USO_COL.KeyName = "setID"
    
    consulta = "SELECT * FROM eq_limitaciones_uso WHERE EQUIPO_ID = " & ID_EQUIPO
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount = 0 Then Exit Sub
    
    Set LIMITACIONES_USO_COL = New clsGenericCollection
    LIMITACIONES_USO_COL.KeyName = "setID"
    rs.MoveFirst
    While Not rs.EOF
    
        Set objLU = New clsGenericClass
        objLU.setID = rs("ID_LIMITACION_USO")
        objLU.setNOMBRE = rs("DESCRIPCION")
        
        Call LIMITACIONES_USO_COL.Add(objLU, objLU.getID)
        
        rs.MoveNext
    Wend

End Sub
'M1050-I
'Private Sub cargar_accesorios()
'    Dim rs As ADODB.RecordSet
'    Dim consulta As String
'    Dim objAcc As clsEquipoAccesorios
'
'    Set ACCESORIOS_COL = New clsGenericCollection
'    ACCESORIOS_COL.KeyName = "setID_ACCESORIO"
'
'    consulta = "SELECT eq_accesorios.*, Equipos.NOMBRE AS NOMBRE_ACCESORIO, Equipos.DESCRIPCION AS DESCRIPCION_ACCESORIO "
'    consulta = consulta & "FROM eq_accesorios "
'    consulta = consulta & " INNER JOIN EQuipos ON  eq_accesorios.ID_ACCESORIO = EQUIPOS.ID_EQUIPO "
'    consulta = consulta & "WHERE eq_accesorios.EQUIPO_ID = " & ID_EQUIPO
'    Set rs = datos_bd(consulta)
'
'    If rs.RecordCount = 0 Then Exit Sub
'
'    Set ACCESORIOS_COL = New clsGenericCollection
'    ACCESORIOS_COL.KeyName = "setID_ACCESORIO"
'
'    rs.MoveFirst
'    While Not rs.EOF
'
'        Set objAcc = New clsEquipoAccesorios
'        objAcc.Carga rs("ID_ACCESORIO"), rs("EQUIPO_ID")
'        'objAcc.setID = rs("ID_ACCESORIO")
'        'objAcc.setNOMBRE = rs("NOMBRE_ACCESORIO")
'
'        Call ACCESORIOS_COL.Add(objAcc, objAcc.getID_ACCESORIO)
'
'        rs.MoveNext
'    Wend
'
'
'End Sub
'M1050-F

Public Function CrearID() As Long
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_EQUIPO) FROM EQUIPOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_EQUIPO = 1
        Else
                ID_EQUIPO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        CrearID = ID_EQUIPO
End Function
Public Function CrearIDParticular(FAMILIA_ID As Long) As Long
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(PARTICULAR) FROM EQUIPOS WHERE FAMILIA_ID = " & FAMILIA_ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                CrearIDParticular = 1
        Else
                CrearIDParticular = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        If ID_EQUIPO = 0 Then
            Me.CrearID
        End If
        
        consulta = "INSERT INTO EQUIPOS (ID_EQUIPO, NOTAS) VALUES (" & CStr(ID_EQUIPO) & ", '')"
        
        execute_bd consulta
        Insertar = ID_EQUIPO
        Call Modificar(ID_EQUIPO)
        
        If ALTA_BAJA = 0 Then
            Dim oEV As New clsEquipoEventos
            With oEV
                .setEQUIPO_ID = ID_EQUIPO
                .setEVENTO_ID = EQUIPOS_EVENTOS.EVT_CAMBIO_ESTADO_EQUIPO
                .setRAZON_ID = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_ALTA_DE_EQUIPO
                .setOBSERVACIONES = ""
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .Insertar
            End With
            Set oEV = Nothing
        End If
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
'        Dim oEq As New clsEquipos
        Dim p As Long
'        oEq.Carga ID
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("SELECT FAMILIA_ID FROM EQUIPOS WHERE ID_EQUIPO = " & ID)
        If rs.RecordCount > 0 Then
            If rs(0) <> FAMILIA_ID Then
                p = Me.CrearIDParticular(CLng(FAMILIA_ID))
                execute_bd ("UPDATE EQUIPOS SET PARTICULAR = " & p & " WHERE ID_EQUIPO = " & ID)
            End If
        End If
'        Set oEq = Nothing
        ' M0459 : UNIDADES DE LOTE
        ' M1124 : ESTADO_ ID
        consulta = "UPDATE EQUIPOS SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " UNIDADES_LOTE = " & UNIDADES_LOTE & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " CONDICIONES_AMBIENTALES = " & CONDICIONES_AMBIENTALES & ", " & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " FECHA_RECEPCION = '" & FECHA_RECEPCION & "'," & _
                        " FECHA_SERVICIO = '" & FECHA_SERVICIO & "'," & _
                        " SERIE = '" & SERIE & "'," & _
                        " ALTA_BAJA = " & ALTA_BAJA & "," & _
                        " ESTADO_ID = '" & ESTADO_ID & "'," & _
                        " NOTAS = '" & NOTAS & "', " & _
                        " RESPONSABLE_ID = " & RESPONSABLE_ID & ", " & _
                        " PROCEDENCIA_ID = " & PROCEDENCIA_ID & ", "
                consulta = consulta & _
                        " FAMILIA_ID = " & FAMILIA_ID & ", " & _
                        " MODELO = '" & MODELO & "', " & " FABRICANTE = '" & fabricante & "', " & _
                        " ES_NADCAP = " & ES_NADCAP & ", " & " ES_ENAC = " & ES_ENAC & ", " & _
                        " ES_MTL = " & ES_MTL & ", " & " ES_CP = " & ES_CP & ", " & _
                        " RANGO_MEDIDA_MIN = '" & CStr(RANGO_MEDIDA_MIN) & "', " & _
                        " RANGO_MEDIDA_MAX = '" & CStr(RANGO_MEDIDA_MAX) & "', " & _
                        " UNIDAD_ID = " & UNIDAD_ID & ", " & _
                        " TOLERANCIA_MAXIMA = '" & CStr(TOLERANCIA_MAXIMA) & "', " & _
                        " INCERTIDUMBRE_MAXIMA = '" & CStr(INCERTIDUMBRE_MAXIMA) & "', " & _
                        " PRECISIONN = '" & CStr(PRECISIONN) & "', " & _
                        " TEMPERATURA_MIN = '" & CStr(TEMPERATURA_MIN) & "', " & _
                        " TEMPERATURA_MAX = '" & CStr(TEMPERATURA_MAX) & "', " & _
                        " HUMEDAD_MIN = '" & CStr(HUMEDAD_MIN) & "', " & _
                        " HUMEDAD_MAX = '" & CStr(HUMEDAD_MAX) & "', " & _
                        " COND_AMBIENTALES_OTRAS = '" & COND_AMBIENTALES_OTRAS & "', " & _
                        " RANGO_TRABAJO_MIN = '" & CStr(RANGO_TRABAJO_MIN) & "', " & _
                        " RANGO_TRABAJO_MAX = '" & CStr(RANGO_TRABAJO_MAX) & "', " & _
                        " SITUACION_DESCRIPCION = '" & SITUACION_DESCRIPCION & "', " & _
                        " VERIFICADOR_ID = '" & VERIFICADOR_ID & "', " & _
                        " VERIFICADOR_INTERNO_ID = " & VERIFICADOR_INTERNO_ID & ", " & _
                        " TIPO_VERIFICACION_ID = " & TIPO_VERIFICACION_ID & ", " & _
                        " MOSTRAR_EN_PLANO = " & MOSTRAR_EN_PLANO & ", " & _
                        " PRIORITARIO = " & PRIORITARIO & ", " & _
                        " INSITU = " & INSITU & ", "

                consulta = consulta & _
                        " CON_CALIBRACION = " & CON_CALIBRACION & ", " & _
                        " CON_VERIFICACION = " & CON_VERIFICACION & ", " & _
                        " CON_MANTENIMIENTO = " & CON_MANTENIMIENTO & ", " & _
                        " PERIODICIDAD_CALIBRACION_ID = " & PERIODICIDAD_CALIBRACION_ID & ", " & _
                        " FECHA_ULT_CALIBRACION = " & FECHA_ULT_CALIBRACION & ", " & _
                        " FECHA_PROX_CALIBRACION = " & FECHA_PROX_CALIBRACION & ", " & _
                        " PERIODICIDAD_VERIFICACION_ID = " & PERIODICIDAD_VERIFICACION_ID & ", " & _
                        " FECHA_ULT_VERIFICACION = " & FECHA_ULT_VERIFICACION & ", " & _
                        " FECHA_PROX_VERIFICACION = " & FECHA_PROX_VERIFICACION & ", " & _
                        " PERIODICIDAD_MANTENIMIENTO_ID = " & PERIODICIDAD_MANTENIMIENTO_ID & ", " & _
                        " FECHA_ULT_MANTENIMIENTO = " & FECHA_ULT_MANTENIMIENTO & ", " & _
                        " FECHA_PROX_MANTENIMIENTO = " & FECHA_PROX_MANTENIMIENTO & ", " & _
                        " SITUACION_ID = " & SITUACION_ID & ", " & _
                        " TIPO_CALIBRACION_ID = " & TIPO_CALIBRACION_ID & ", " & _
                        " CALIBRADOR_ID = " & CALIBRADOR_ID & ", " & _
                        " CALIBRADOR_INTERNO_ID = " & CALIBRADOR_INTERNO_ID & ", " & _
                        " REVISADO = " & REVISADO & ", " & _
                        " COORDX = " & COORDX & ", " & _
                        " COORDY = " & COORDY & ", " & _
                        " NUMERO_USOS_MAXIMO = " & NUMERO_USOS_MAXIMO & ", " & _
                        " NUMERO_USOS_CONTADOR = " & NUMERO_USOS_CONTADOR & ", " & _
                        " MANTENEDOR_ID = " & MANTENEDOR_ID & ", " & _
                        " CUSERID = '" & USUARIO.getUSUARIO & "', " & _
                        " FTIMESTP = CURRENT_TIMESTAMP, "
                'M1050-I
                consulta = consulta & _
                        " PLAN_MANTENIMIENTO_ID = '" & PLAN_MANTENIMIENTO_ID & "', " & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " TRAZMET_ID = " & TRAZMET_ID & "," & _
                        " NUMERO_EQUIPO_CLIENTE = '" & NUMERO_EQUIPO_CLIENTE & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " CRITICO = " & CRITICO & ", CENTRO_ID = " & CENTRO_ID & ","
                consulta = consulta & " PROCEDIMIENTO_CALIBRACION_ID = " & PROCEDIMIENTO_CALIBRACION_ID & ", "
                consulta = consulta & " PROCEDIMIENTO_VERIFICACION_ID = " & PROCEDIMIENTO_VERIFICACION_ID & ", "
                'consulta = consulta & " FUERA_SERVICIO = " & FUERA_SERVICIO & ", "
                
                
                consulta = consulta & _
                        " TIPO_EQUIPO_ID = " & TIPO_EQUIPO_ID & ", " & _
                        " ES_ACCESORIO = " & CStr(es_accesorio)
                
                consulta = consulta & _
                                " WHERE ID_EQUIPO = " & ID
                        
        execute_bd consulta
        ' si el equipo se da de baja genera un evento
        If ALTA_BAJA = 1 Then
            Dim oEV As New clsEquipoEventos
            With oEV
                .setEQUIPO_ID = ID_EQUIPO
                .setEVENTO_ID = EQUIPOS_EVENTOS.EVT_CAMBIO_ESTADO_EQUIPO
                .setRAZON_ID = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_BAJA_DE_EQUIPO
                .setOBSERVACIONES = ""
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .Insertar
            End With
            Set oEV = Nothing
        End If
        
        
        Modificar = True
        Exit Function
        
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        execute_bd "DELETE FROM eq_accesorios WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_CALIBRACION_EQUIPOS WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_verificaCION_EQUIPOS WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_mantenimiento_EQUIPOS WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_planes_mantenimiento_EQUIPOS WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_normas_EQUIPOS WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_especificaciones_tecnicas WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_recepcion WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQ_eventos WHERE EQUIPO_ID = " & ID
        ' finalmente se borra el equipo
        execute_bd "DELETE FROM EQUIPOS WHERE ID_EQUIPO = " & ID
        Eliminar = True
        Exit Function
        
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos)", vbCritical, Err.Description
End Function

Public Function Listado(Optional ByVal fdes As String = "", Optional ByVal fserie As String = "", Optional ByVal fequipo As String = "", Optional ByVal fnombre As String = "", Optional ByVal ffamilia As String = "", Optional ByVal fsituacion As String = "", Optional ByVal fproveedor As String = "", Optional ByVal baja As Integer = 0, Optional ByVal con_calib As Integer = 1, Optional ByVal con_verf As Integer = 1, Optional ByVal con_mto As Integer = 1, Optional ByVal estado_calibracion As String = "", Optional ByVal estado_verificacion As String = "", Optional ByVal estado_mantenimiento As String = "", Optional ByVal es_accesorio As String = "", Optional ByVal TIPO_EQUIPO As String = "", _
                        Optional ByRef criterio_listado As String = "", Optional ByRef proveedor As Long, Optional FUERA_SERVICIO As Boolean, Optional ByRef responsable As Long, _
                        Optional PRIORITARIO As Boolean, Optional CLIENTE_ID As Long, Optional NUMERO_EQUIPO_CLIENTE As String, Optional SOLO_CLIENTES As Boolean, Optional SOLO_CANAGROSA As Boolean, Optional ESTADO As String, Optional CAL_INTERNAS As Boolean, Optional CAL_EXTERNAS As Boolean, Optional CENTRO_ID As String, Optional chkENAC As Boolean, Optional chkNADCAP As Boolean, Optional chkMTL As Boolean, Optional chkCP As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim eq As String
        Dim filtro As String
        Dim strRevision As String
        Dim sq_familia As String, sq_localizacion As String, sq_tipo_equipo As String
        Dim inner_cal As String, inner_ver As String, inner_man As String
        Dim campos_cal As String, campos_ver As String, campos_man As String
        
        devolver_filtros_listado_con_fechas_preavisos inner_cal, inner_ver, inner_man, campos_cal, campos_ver, campos_man
        
        filtro = inner_cal & inner_ver & inner_man & " WHERE 1=1"
        
        If fdes <> "" Then
            filtro = filtro & " AND EQUIPOS.DESCRIPCION LIKE '%" & fdes & "%'"
            criterio_listado = criterio_listado & " and {equipos.DESCRIPCION} like " & Chr(34) & "*" & fdes & "*" & Chr(34)
        End If
        If fequipo <> "" Then
            If IsNumeric(fequipo) Then
                filtro = filtro & " AND EQUIPOS.ID_EQUIPO = " & fequipo
                criterio_listado = criterio_listado & " and {equipos.ID_EQUIPO} = " & fequipo
            End If
        End If
        If fnombre <> "" Then
            filtro = filtro & " AND EQUIPOS.NOMBRE LIKE '%" & fnombre & "%'"
            criterio_listado = criterio_listado & " and {equipos.NOMBRE} like " & Chr(34) & "*" & fnombre & "*" & Chr(34)
        End If
        If fserie <> "" Then
            filtro = filtro & " AND EQUIPOS.SERIE LIKE '%" & fserie & "%'"
            criterio_listado = criterio_listado & " and {equipos.SERIE} like " & Chr(34) & "*" & fserie & "*" & Chr(34)
        End If
        If ffamilia <> "" And ffamilia <> "0" Then
            filtro = filtro & " AND EQUIPOS.familia_id = " & ffamilia
            criterio_listado = criterio_listado & " and {equipos.FAMILIA_ID} = " & ffamilia
        End If
        If fsituacion <> "0" And fsituacion <> "" Then
            filtro = filtro & " AND EQUIPOS.situacion_id = " & fsituacion
            criterio_listado = criterio_listado & " and {equipos.SITUACION_ID} = " & fsituacion
        End If
        If fproveedor <> "" Then
            filtro = filtro & " AND EQUIPOS.fabricante LIKE '%" & fproveedor & "%'"
            criterio_listado = criterio_listado & " and {EQUIPOS.fabricante} like " & Chr(34) & "*" & fproveedor & "*" & Chr(34)
        End If
        If TIPO_EQUIPO <> "" Then
            filtro = filtro & " AND EQUIPOS.TIPO_EQUIPO_ID = " & TIPO_EQUIPO
            criterio_listado = criterio_listado & " and {EQUIPOS.TIPO_EQUIPO_ID} = " & TIPO_EQUIPO
        End If
        If es_accesorio <> "" Then
            filtro = filtro & " AND EQUIPOS.ES_ACCESORIO = " & es_accesorio
            criterio_listado = criterio_listado & " and {EQUIPOS.ES_ACCESORIO} = " & es_accesorio
        End If
        Dim joinCAL As String
        If con_calib = 1 Then
            filtro = filtro & " AND EQUIPOS.CON_CALIBRACION = 1"
            criterio_listado = criterio_listado & " and {EQUIPOS.CON_CALIBRACION} = 1"
            If CAL_INTERNAS = True Or CAL_EXTERNAS = True Then
                Dim t As String
                If CAL_INTERNAS = True Then
                    t = t & "1"
                End If
                If CAL_EXTERNAS = True Then
                    If t <> "" Then
                        t = t & ","
                    End If
                    t = t & "2"
                End If
                joinCAL = " INNER JOIN ( " & _
                          " SELECT AA.EQUIPO_ID,IF(AA.TIPO_ID=2,P.NOMBRE,CONCAT(U.APELLIDOS,',',U.NOMBRE)) AS CALIBRADOR " & _
                          "   FROM EQ_CALIBRACION_EQUIPOS AA " & _
                          "   LEFT JOIN PROVEEDORES P ON AA.CALIBRADOR_EXTERNO_ID = P.ID_PROVEEDOR " & _
                          "   LEFT JOIN usuarios U ON AA.CALIBRADOR_INTERNO_ID = U.ID_EMPLEADO " & _
                          "  WHERE AA.TIPO_ID in (" & t & ") And AA.ESTADO = 0 " & _
                          "  GROUP BY AA.EQUIPO_ID ORDER BY AA.ID_CALIBRACION DESC ) AS CALI ON EQUIPOS.ID_EQUIPO = CALI.EQUIPO_ID"
'                criterio_listado = criterio_listado & " and {CAL.TIPO_ID} = 1"
            End If
        End If
        If joinCAL = "" Then
                joinCAL = " LEFT JOIN ( " & _
                          " SELECT AA.EQUIPO_ID,IF(AA.TIPO_ID=2,P.NOMBRE,CONCAT(U.APELLIDOS,',',U.NOMBRE)) AS CALIBRADOR " & _
                          "   FROM EQ_CALIBRACION_EQUIPOS AA " & _
                          "   LEFT JOIN PROVEEDORES P ON AA.CALIBRADOR_EXTERNO_ID = P.ID_PROVEEDOR " & _
                          "   LEFT JOIN usuarios U ON AA.CALIBRADOR_INTERNO_ID = U.ID_EMPLEADO " & _
                          "  WHERE AA.ESTADO = 0 " & _
                          "  GROUP BY AA.EQUIPO_ID ORDER BY AA.ID_CALIBRACION DESC ) AS CALI ON EQUIPOS.ID_EQUIPO = CALI.EQUIPO_ID"
        End If
        If con_verf = 1 Then
            filtro = filtro & " AND EQUIPOS.CON_VERIFICACION = 1"
            criterio_listado = criterio_listado & " and {EQUIPOS.CON_VERIFICACION} = 1"
        End If
        If con_mto = 1 Then
            filtro = filtro & " AND EQUIPOS.CON_MANTENIMIENTO = 1"
            criterio_listado = criterio_listado & " and {EQUIPOS.CON_MANTENIMIENTO} = 1"
        End If
        If proveedor <> 0 Then
            filtro = filtro & " AND EQUIPOS.PROVEEDOR_ID = " & proveedor
            criterio_listado = criterio_listado & " and {EQUIPOS.PROVEEDOR_ID} = " & proveedor
        End If
        If responsable <> 0 Then
            filtro = filtro & " AND EQUIPOS.RESPONSABLE_ID = " & responsable
            criterio_listado = criterio_listado & " and {EQUIPOS.RESPONSABLE_ID} = " & responsable
        End If
        If FUERA_SERVICIO = False Then
            filtro = filtro & " AND EQUIPOS.FUERA_SERVICIO = 0"
            criterio_listado = criterio_listado & " and {EQUIPOS.FUERA_SERVICIO} = 0"
        End If
'        If PRIORITARIO = True Then
'            filtro = filtro & " AND EQUIPOS.PRIORITARIO = 1"
'            criterio_listado = criterio_listado & " and {EQUIPOS.PRIORITARIO} = 1"
'        End If
        'M1050-I
        If CLIENTE_ID <> 0 Then
            filtro = filtro & " AND EQUIPOS.CLIENTE_ID = " & CLIENTE_ID
            criterio_listado = criterio_listado & " and {EQUIPOS.CLIENTE_ID} = " & CLIENTE_ID
        End If
        If NUMERO_EQUIPO_CLIENTE <> "" Then
            filtro = filtro & " AND EQUIPOS.NUMERO_EQUIPO_CLIENTE LIKE '%" & NUMERO_EQUIPO_CLIENTE & "%'"
            criterio_listado = criterio_listado & " and {EQUIPOS.NUMERO_EQUIPO_CLIENTE} LIKE '*" & NUMERO_EQUIPO_CLIENTE & "*'"
        End If
        If SOLO_CLIENTES = True Then
            filtro = filtro & " AND EQUIPOS.CLIENTE_ID > 0"
            criterio_listado = criterio_listado & " and {EQUIPOS.CLIENTE_ID} > 0"
        End If
        If SOLO_CANAGROSA = True Then
            filtro = filtro & " AND EQUIPOS.CLIENTE_ID = 0"
            criterio_listado = criterio_listado & " and {EQUIPOS.CLIENTE_ID} = 0"
        End If
        If chkENAC = True Then
            filtro = filtro & " AND EQUIPOS.ES_ENAC = 1"
            criterio_listado = criterio_listado & " and {EQUIPOS.ES_ENAC} = 1"
        End If
        If chkNADCAP = True Then
            filtro = filtro & " AND EQUIPOS.ES_NADCAP = 1"
            criterio_listado = criterio_listado & " and {EQUIPOS.ES_NADCAP} = 1"
        End If
        If chkMTL = True Then
            filtro = filtro & " AND EQUIPOS.ES_MTL = 1"
            criterio_listado = criterio_listado & " and {EQUIPOS.ES_MTL} = 1"
        End If
        If chkCP = True Then
            filtro = filtro & " AND EQUIPOS.ES_CP = 1"
            criterio_listado = criterio_listado & " and {EQUIPOS.ES_CP} = 1"
        End If

        'M1050-F
        ' Estados de Mantenimientos
        If estado_calibracion <> "" Then
            Select Case CInt(estado_calibracion)
                Case 0
'                    filtro = filtro & " AND coalesce(CAL.FECHA_PREAVISO,'') <> '' AND CAL.FECHA_PREAVISO > localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_calibraciones.fecha_preaviso}) and {inner_calibraciones.fecha_preaviso} > CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_CALIBRACION > LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_CALIBRACION > CurrentDate"
                Case 1
'                    filtro = filtro & " AND coalesce(CAL.FECHA_PREAVISO,'') <> '' AND CAL.FECHA_PREAVISO <= localtimestamp and CAL.FECHA_PREVISTA > localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_calibraciones.fecha_preaviso}) and {inner_calibraciones.fecha_preaviso} <= CurrentDate and {inner_calibraciones.fecha_proxima} > CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_CALIBRACION < LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_CALIBRACION < CurrentDate"
                Case 2
'                    filtro = filtro & " AND coalesce(CAL.FECHA_PREAVISO,'') <> '' AND CAL.FECHA_PREAVISO <= localtimestamp and CAL.FECHA_PREVISTA <= localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_calibraciones.fecha_preaviso}) and {inner_calibraciones.fecha_preaviso} <= CurrentDate and {inner_calibraciones.fecha_proxima} <= CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_CALIBRACION < LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_CALIBRACION < CurrentDate"
            End Select
        End If

        If estado_verificacion <> "" Then
            Select Case CInt(estado_verificacion)
                Case 0
'                    filtro = filtro & " AND coalesce(VER.FECHA_PREAVISO, '') <> ''  AND VER.FECHA_PREAVISO > localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_verificaciones.fecha_preaviso}) and {inner_verificaciones.fecha_preaviso} > CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_VERIFICACION > LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_VERIFICACION > CurrentDate"
                Case 1
'                    filtro = filtro & " AND coalesce(VER.FECHA_PREAVISO, '') <> ''  AND VER.FECHA_PREAVISO < localtimestamp and VER.FECHA_PREVISTA > localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_verificaciones.fecha_preaviso}) and {inner_verificaciones.fecha_preaviso} <= CurrentDate and {inner_verificaciones.fecha_proxima} > CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_VERIFICACION < LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_VERIFICACION < CurrentDate"
                Case 2
'                    filtro = filtro & " AND coalesce(VER.FECHA_PREAVISO, '') <> ''  AND VER.FECHA_PREAVISO < localtimestamp and VER.FECHA_PREVISTA <= localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_verificaciones.fecha_preaviso}) and {inner_verificaciones.fecha_preaviso} <= CurrentDate and {inner_verificaciones.fecha_proxima} <= CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_VERIFICACION < LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_VERIFICACION < CurrentDate"
            End Select
        End If

        If estado_mantenimiento <> "" Then
            Select Case CInt(estado_mantenimiento)
                Case 0
'                    filtro = filtro & " AND coalesce(MAN.FECHA_PREAVISO, '') <> '' AND MAN.FECHA_PREAVISO > localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_mantenimiento.fecha_preaviso}) and {inner_mantenimiento.fecha_preaviso} > CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_MANTENIMIENTO > LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_MANTENIMIENTO > CurrentDate"
                Case 1
'                    filtro = filtro & " AND coalesce(MAN.FECHA_PREAVISO, '') <> '' AND MAN.FECHA_PREAVISO < localtimestamp and MAN.FECHA_PREVISTA > localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_mantenimiento.fecha_preaviso}) and {inner_mantenimiento.fecha_preaviso} <= CurrentDate and {inner_mantenimiento.fecha_proxima} > CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_MANTENIMIENTO < LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_MANTENIMIENTO < CurrentDate"
                Case 2
'                    filtro = filtro & " AND coalesce(MAN.FECHA_PREAVISO, '') <> '' AND MAN.FECHA_PREAVISO < localtimestamp and MAN.FECHA_PREVISTA <= localtimestamp"
'                    criterio_listado = criterio_listado & " and not isnull({inner_mantenimiento.fecha_preaviso}) and {inner_mantenimiento.fecha_preaviso} <= CurrentDate and {inner_mantenimiento.fecha_proxima} <= CurrentDate"
                    filtro = filtro & " AND EQUIPOS.FECHA_PROX_MANTENIMIENTO < LOCALTIMESTAMP "
                    criterio_listado = criterio_listado & " AND EQUIPOS.FECHA_PROX_MANTENIMIENTO < CurrentDate"
            End Select
        End If
        'M1124-I
        If ESTADO <> "" Then
            filtro = filtro & " AND EQUIPOS.ESTADO_ID = '" & ESTADO & "'"
        End If
        If CENTRO_ID <> "" Then
            filtro = filtro & " AND EQUIPOS.CENTRO_ID = '" & CENTRO_ID & "'"
        End If
        'M1124-F
        filtro = filtro & " AND EQUIPOS.alta_baja = " & baja
        criterio_listado = "{equipos.alta_baja} = " & baja & criterio_listado
        
        consulta = "SELECT EQUIPOS.ID_EQUIPO, EQUIPOS.NUMERO_EQUIPO_CLIENTE, EQUIPOS.ESTADO_ID,EQUIPOS.NOMBRE,EQUIPOS.SERIE, "
        consulta = consulta & campos_cal & campos_ver & campos_man
        consulta = consulta & "   COALESCE(SQ_LOCALIZACIONES.DESCRIPCION, 'N/A') AS NOMBRE_LOCALIZACION, "
        consulta = consulta & "   CONCAT(USU.NOMBRE,' ',USU.APELLIDOS) AS RESPONSABLE, EQUIPOS.FUERA_SERVICIO,EQUIPOS.PRIORITARIO,EQUIPOS.AV,EQUIPOS.SPR,EQUIPOS.RPR,IFNULL(CALI.CALIBRADOR,'') AS CALIBRADOR, CLI.NOMBRE AS CLIENTE,CEN.NOMBRE AS CENTRO "
        consulta = consulta & " FROM EQUIPOS "
        consulta = consulta & " LEFT JOIN CLIENTES CLI on EQUIPOS.CLIENTE_ID = CLI.ID_CLIENTE "
        consulta = consulta & " LEFT OUTER JOIN usuarios AS USU ON EQUIPOS.RESPONSABLE_ID = USU.ID_EMPLEADO "
        consulta = consulta & " LEFT JOIN decodificadora AS SQ_LOCALIZACIONES ON EQUIPOS.SITUACION_ID = SQ_LOCALIZACIONES.VALOR AND SQ_LOCALIZACIONES.CODIGO = " & DECODIFICADORA.EQ_SITUACIONES
        consulta = consulta & " LEFT JOIN CENTROS AS CEN  ON EQUIPOS.CENTRO_ID = CEN.ID_CENTRO "
        consulta = consulta & joinCAL
        Dim ORDEN As String
        ORDEN = " order by id_equipo "
        consulta = consulta & filtro & ORDEN
'        Set Listado = datos_bd(consulta & filtro & " order by id_equipo")
        Set Listado = datos_bd(consulta)
       
End Function
Public Function Listado_equipo_unico(ID As Long) As ADODB.Recordset
        Dim consulta As String
        Dim eq As String
        Dim filtro As String
        Dim strRevision As String
        Dim sq_familia As String, sq_localizacion As String, sq_tipo_equipo As String
        Dim inner_cal As String, inner_ver As String, inner_man As String
        Dim campos_cal As String, campos_ver As String, campos_man As String
        
        devolver_filtros_listado_con_fechas_preavisos inner_cal, inner_ver, inner_man, campos_cal, campos_ver, campos_man
        
        filtro = inner_cal & inner_ver & inner_man & " where equipos.id_equipo = " & CStr(ID)
                
        consulta = "SELECT EQUIPOS.ID_EQUIPO, EQUIPOS.NUMERO_EQUIPO_CLIENTE, EQUIPOS.ESTADO_ID,EQUIPOS.NOMBRE,EQUIPOS.SERIE, "
        consulta = consulta & campos_cal & campos_ver & campos_man
        consulta = consulta & "   COALESCE(SQ_LOCALIZACIONES.DESCRIPCION, 'N/A') AS NOMBRE_LOCALIZACION, "
        consulta = consulta & "   CONCAT(USU.NOMBRE,' ',USU.APELLIDOS) AS RESPONSABLE, EQUIPOS.FUERA_SERVICIO,EQUIPOS.PRIORITARIO,EQUIPOS.AV,EQUIPOS.SPR,EQUIPOS.RPR, CLI.NOMBRE AS CLIENTE,CEN.NOMBRE AS CENTRO "
        consulta = consulta & " FROM EQUIPOS "
        consulta = consulta & " LEFT JOIN CLIENTES CLI on EQUIPOS.CLIENTE_ID = CLI.ID_CLIENTE "
        consulta = consulta & " LEFT OUTER JOIN usuarios AS USU ON EQUIPOS.RESPONSABLE_ID = USU.ID_EMPLEADO "
        consulta = consulta & " LEFT JOIN decodificadora AS SQ_LOCALIZACIONES ON EQUIPOS.SITUACION_ID = SQ_LOCALIZACIONES.VALOR AND SQ_LOCALIZACIONES.CODIGO = " & DECODIFICADORA.EQ_SITUACIONES
        consulta = consulta & " LEFT JOIN CENTROS AS CEN  ON EQUIPOS.CENTRO_ID = CEN.ID_CENTRO "
        consulta = consulta & filtro
        
        Set Listado_equipo_unico = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT ID_EQUIPO, NOMBRE FROM EQUIPOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "EQUIPOS"
        .setDESCRIPCION = "Equipos"
        .setPK = "ID_EQUIPO"
        .setCAMPO = " concat(Nombre,' (Nº ',CAST(ID_EQUIPO AS CHAR),') -> Estado : ',ESTADO_ID)"
        .setMUESTRA_DETALLE = True
        .setFILTRO = " ALTA_BAJA=0 AND FUERA_SERVICIO = 0 " & filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Public Function ListadoObjetos() As clsGenericCollection
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Dim objEq As clsEquipos
        Dim objCol As clsGenericCollection
        
        Set rs = Listado()
        
        Set objCol = New clsGenericCollection
        objCol.KeyName = "setID_EQUIPO"
        
        If rs.RecordCount = 0 Then
            Set ListadoObjetos = objCol
            Exit Function
        End If
        
        
        On Error Resume Next
        
        rs.MoveFirst
        While Not rs.EOF
            Set objEq = New clsEquipos
            objEq.setID_EQUIPO = rs("ID_EQUIPO")
            objEq.setNOMBRE = rs("NOMBRE")
            objEq.setCONDICIONES_AMBIENTALES = rs("CONDICIONES_AMBIENTALES")
            objEq.setPROVEEDOR_ID = rs("PROVEEDOR_ID")
            objEq.setFECHA_RECEPCION = rs("FECHA_RECEPCION")
            objEq.setFECHA_SERVICIO = rs("FECHA_SERVICIO")
            objEq.setALTA_BAJA = rs("ALTA_BAJA")
            'M1124-I
            objEq.setESTADO_ID = rs("ESTADO_ID")
            'M1124-F
            objEq.setSERIE = rs("SERIE")
            objEq.setFAMILIA_ID = rs("FAMILIA_ID")
            objEq.setPROVEEDOR_ID = rs("PROVEEDOR_ID")
            objEq.setMODELO = rs("MODELO")
            objEq.setFABRICANTE = rs("FABRICANTE")
            objEq.setES_NADCAP = rs("ES_NADCAP")
            objEq.setES_MTL = rs("ES_MTL")
            objEq.setES_CP = rs("ES_CP")
            objEq.setES_ENAC = rs("ES_ENAC")
            objEq.setRANGO_MEDIDA_MIN = rs("RANGO_MEDIDA_MIN")
            objEq.setRANGO_MEDIDA_MAX = rs("RANGO_MEDIDA_MAX")
            objEq.setUNIDAD_ID = rs("UNIDAD_ID")
            objEq.setTOLERANCIA_MAXIMA = rs("TOLERANCIA_MAXIMA")
            objEq.setINCERTIDUMBRE_MAXIMA = rs("INCERTIDUMBRE_MAXIMA")
            objEq.setPRECISIONN = rs("PRECISIONN")
            objEq.setTEMPERATURA_MIN = rs("TEMPERATURA_MIN")
            objEq.setTEMPERATURA_MAX = rs("TEMPERATURA_MAX")
            objEq.setHUMEDAD_MIN = rs("HUMEDAD_MIN")
            objEq.setHUMEDAD_MAX = rs("HUMEDAD_MAX")
            objEq.setCOND_AMBIENTALES_OTRAS = rs("COND_AMBIENTALES_OTRAS")
            objEq.setRANGO_TRABAJO_MIN = rs("RANGO_TRABAJO_MIN")
            objEq.setRANGO_TRABAJO_MAX = rs("RANGO_TRABAJO_MAX")
            objEq.setLIMITACIONES_USO = rs("LIMITACIONES_USO")
            objEq.setPERIODICIDAD_CALIBRACION_ID = rs("PERIODICIDAD_CALIBRACION_ID")
            objEq.setFECHA_ULT_CALIBRACION = rs("FECHA_ULT_CALIBRACION")
            objEq.setFECHA_PROX_CALIBRACION = rs("FECHA_PROX_CALIBRACION")
            objEq.setPERIODICIDAD_VERIFICACION_ID = rs("PERIODICIDAD_VERIFICACION_ID")
            objEq.setFECHA_ULT_VERIFICACION = rs("FECHA_ULT_VERIFICACION")
            objEq.setFECHA_PROX_VERIFICACION = rs("FECHA_PROX_VERIFICACION")
            objEq.setPERIODICIDAD_MANTENIMIENTO_ID = rs("PERIODICIDAD_MANTENIMIENTO_ID")
            objEq.setFECHA_ULT_MANTENIMIENTO = rs("FECHA_ULT_MANTENIMIENTO")
            objEq.setFECHA_PROX_MANTENIMIENTO = rs("FECHA_PROX_MANTENIMIENTO")
            objEq.setSITUACION_ID = rs("SITUACION_ID")
            objEq.setTIPO_CALIBRACION_ID = rs("TIPO_CALIBRACION_ID")
            objEq.setCALIBRADOR_ID = rs("CALIBRADOR_ID")
            objEq.setCALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
            objEq.setCON_CALIBRACION = rs("CON_CALIBRACION")
            objEq.setCON_VERIFICACION = rs("CON_VERIFICACION")
            objEq.setCON_MANTENIMIENTO = rs("CON_MANTENIMIENTO")
            objEq.setTIPO_EQUIPO_ID = rs("TIPO_EQUIPO_ID")
            objEq.setTIPO_EQUIPO = rs("TIPO_EQUIPO")
            objEq.setREVISADO = rs("REVISADO")
            If IsNull(rs("PROCEDIMIENTO_VERIFICACION_ID")) Then
                objEq.setPROCEDIMIENTO_VERIFICACION_ID = 0
            Else
                objEq.setPROCEDIMIENTO_VERIFICACION_ID = rs("PROCEDIMIENTO_VERIFICACION_ID")
            End If
            If IsNull(rs("PROCEDIMIENTO_CALIBRACION_ID")) Then
                objEq.setPROCEDIMIENTO_CALIBRACION_ID = 0
            Else
                objEq.setPROCEDIMIENTO_CALIBRACION_ID = rs("PROCEDIMIENTO_CALIBRACION_ID")
            End If
            
            objEq.setNOTAS = rs("NOTAS")
            objEq.setNOTAS = rs("NOTAS")
            If IsNull(rs("COORDX")) Then
                objEq.setCoordX = 0
            Else
                objEq.setCoordX = rs("COORDX")
            End If
            If IsNull(rs("COORDY")) Then
                objEq.setCoordY = 0
            Else
                objEq.setCoordY = rs("COORDY")
            End If
            
            Call objCol.Add(objEq, objEq.getID_EQUIPO, 1)
            
            rs.MoveNext
        Wend
        
        On Error GoTo 0
        
        Set ListadoObjetos = objCol
        
End Function

Public Function ModificarPosicionEnPlano() As Boolean
Dim consulta As String

On Error GoTo ModificarPosicionEnPlano_Error

    If ID_EQUIPO <= 0 Then Exit Function
    
    consulta = "update equipos set coordx = " & CStr(COORDX)
    consulta = consulta & ", coordy = " & CStr(COORDY)
    consulta = consulta & " where id_equipo = " & ID_EQUIPO
    
    execute_bd consulta
    
    ModificarPosicionEnPlano = True

On Error GoTo 0
    Exit Function
ModificarPosicionEnPlano_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarPosicionEnPlano of Módulo de clase clsEquipo"
    ModificarPosicionEnPlano = False
End Function

Private Sub Class_Initialize()
    ACCESORIOS_COL.KeyName = "setID_ACCESORIO"
    LIMITACIONES_USO_COL.KeyName = "setID"
End Sub

Public Function DevolverParametrosResultados(ByVal ID_EQUIPO As Long, ByVal ID_TIPO_RESULTADO As Integer) As ADODB.Recordset

    Dim sql As String
    
    ' El tipo_resultado = 1 es calibracion
    sql = "SELECT eq_parametros_resultados.*, unidades.nombre as Unidad "
    sql = sql & " FROM eq_parametros_resultados"
    sql = sql & " INNER JOIN unidades on eq_parametros_resultados.unidad_id = unidades.id_unidad "
    sql = sql & " Where tipo_resultado = " & CStr(ID_TIPO_RESULTADO) & " And EQUIPO_ID = " & CStr(ID_EQUIPO)
    
    Set DevolverParametrosResultados = datos_bd(sql)


End Function

'Public Function AnadirParametroResultado(ByVal prmid_equipo, ByVal prmid_tipo_parametro As Integer, ByVal prmDescripcion As String, ByVal prmid_unidad As Long, prmRangoMin As String, prmRangoMax As String, prmTolerancia As String, prmCorrecion As String, prmIncertidumbre As String) As Boolean
'
'Dim sql As String
'
'    sql = "INSERT INTO eq_parametros_resultados (id_parametro, equipo_id, unidad_id, rango_min, rango_max, tolerancia_max, correccion, incertidumbre, descripcion, tipo_resultado) "
'    sql = sql & " SELECT COALESCE(MAX(id_parametro), 0) + 1 AS id_parametro"
'    sql = sql & " ," & CStr(prmid_equipo)
'    sql = sql & " ," & CStr(prmid_unidad)
'    If IsNumeric(prmRangoMin) Then
'        sql = sql & " ," & Replace(CStr(prmRangoMin), ",", ".")
'    Else
'        sql = sql & " ,0"
'    End If
'    If IsNumeric(prmRangoMax) Then
'        sql = sql & " ," & Replace(CStr(prmRangoMax), ",", ".")
'    Else
'        sql = sql & " ,0"
'    End If
'    If IsNumeric(prmTolerancia) Then
'        sql = sql & " ," & Replace(CStr(prmTolerancia), ",", ".")
'    Else
'        sql = sql & " ,0"
'    End If
'    If IsNumeric(prmCorrecion) Then
'        sql = sql & " ," & Replace(CStr(prmCorrecion), ",", ".")
'    Else
'        sql = sql & " ,0"
'    End If
'    If IsNumeric(prmIncertidumbre) Then
'        sql = sql & " ," & Replace(CStr(prmIncertidumbre), ",", ".")
'    Else
'        sql = sql & " ,0"
'    End If
'
'    sql = sql & " ,'" & CStr(prmDescripcion) & "'"
'    sql = sql & " ," & CStr(prmid_tipo_parametro)
'    sql = sql & " from eq_parametros_resultados"
'
'
'    execute_bd sql
'End Function
Public Function DevolverParametrosResultadosEquipoCalibracion(ByVal prmid_equipo As String) As ADODB.Recordset
    Dim sql As String

    sql = "SELECT -1 as id_resultado, 0 as resultado, equipo_id, unidad_id, unidades.nombre as unidad, rango_min, rango_max, tolerancia_max, correccion, incertidumbre, descripcion from eq_parametros_resultados, unidades where unidades.id_unidad = unidad_id and tipo_resultado = 1 and equipo_id = " & prmid_equipo
    
    Set DevolverParametrosResultadosEquipoCalibracion = datos_bd(sql)
End Function

'Public Function DevolverParametrosResultadosEquipoVerificacion(ByVal prmid_equipo As String) As ADODB.Recordset
'Dim sql As String
'
'
'    sql = "SELECT -1 as id_resultado, 0 as resultado, equipo_id, unidad_id, unidades.nombre as unidad, rango_min, rango_max, tolerancia_max, correccion, incertidumbre, descripcion from eq_parametros_resultados, unidades where unidades.id_unidad = unidad_id and tipo_resultado = 2 and equipo_id = " & prmid_equipo
'
'    Set DevolverParametrosResultadosEquipoVerificacion = datos_bd(sql)
'End Function

Public Function PonerEquipoFueraServicio(ByVal prmid_equipo As String) As Boolean
'Public Function PonerEquipoFueraServicio(ByVal prmid_equipo As String, Optional ByVal prmTipo As Integer = 0, Optional ByVal prmIdCVM As Long = 0) As Boolean
Dim sql As String

    ' prmTipo   = 1 -> Calibracion
    '           = 2 -> Verificacion
    '           = 3 -> Mantenimiento
    ' prmIdCVM = ID_CALIBRACION, ID_VERIFICACION, ID_MANTENIMIENTO, segun el Tipo prmTipo
        
    'M1124-I
    'sql = "UPDATE EQUIPOS SET FUERA_SERVICIO = 1 where ID_EQUIPO = " & prmid_equipo
    sql = "UPDATE EQUIPOS SET FUERA_SERVICIO = 1 ,ESTADO_ID='" & ENUM_EQ_ESTADOS_FUERA_SERVICIO & "' where ID_EQUIPO = " & prmid_equipo
    'M1124-F
    
    execute_bd sql
    
    PonerEquipoFueraServicio = True
    
End Function

Public Function PonerEquipoEnServicio(ByVal prmid_equipo As String) As Boolean
Dim sql As String

    ' prmTipo   = 1 -> Calibracion
    '           = 2 -> Verificacion
    '           = 3 -> Mantenimiento
    ' prmIdCVM = ID_CALIBRACION, ID_VERIFICACION, ID_MANTENIMIENTO, segun el Tipo prmTipo
        
    'M1124-I
    'sql = "UPDATE EQUIPOS SET ALTA_BAJA = 0, FUERA_SERVICIO = 0 where ID_EQUIPO = " & prmid_equipo
    sql = "UPDATE EQUIPOS SET ALTA_BAJA = 0, FUERA_SERVICIO = 0,ESTADO_ID='" & ENUM_EQ_ESTADOS_ACTIVO & "' where ID_EQUIPO = " & prmid_equipo
    'M1124-F
    
    execute_bd sql
    
    PonerEquipoEnServicio = True
    
End Function
Public Function PonerEquipoEnM1(ByVal prmid_equipo As String) As Boolean
    Dim sql As String
    sql = "UPDATE EQUIPOS SET ESTADO_ID='" & ENUM_EQ_ESTADOS_METROLOGIA & "' where ID_EQUIPO = " & prmid_equipo
    execute_bd sql
    PonerEquipoEnM1 = True
End Function
Public Function PonerEquipoEnM2(ByVal prmid_equipo As String) As Boolean
    Dim sql As String
    sql = "UPDATE EQUIPOS SET ESTADO_ID='" & ENUM_EQ_ESTADOS_EVALUACION_CERTIFICADO & "' where ID_EQUIPO = " & prmid_equipo
    execute_bd sql
    PonerEquipoEnM2 = True
End Function


Public Function devolver_lista_planes_mantenimiento() As ADODB.Recordset
Dim strCad As String


    strCad = "SELECT eq_planes_mantenimiento_equipos.* " & _
            ", concat(eq_planes_mantenimiento.NOMBRE, ': ', eq_planes_mantenimiento.DESCRIPCION) as NOMBRE " & _
            ", eq_planes_mantenimiento.PROTOCOLO_ID as PROTOCOLO_ID " & _
            ", COALESCE(ca_documentos.nombre, 'N/A') as PROTOCOLO " & _
            " From eq_planes_mantenimiento_equipos " & _
            " inner join eq_planes_mantenimiento on eq_planes_mantenimiento_equipos.PLAN_MANTENIMIENTO_ID = eq_planes_mantenimiento.ID_PLAN_MTO" & _
            " left outer join ca_documentos on eq_planes_mantenimiento.protocolo_id = ca_documentos.ID_documento" & _
            " WHERE eq_planes_mantenimiento_equipos.EQUIPO_ID = " & CStr(ID_EQUIPO)

    
    Set devolver_lista_planes_mantenimiento = datos_bd(strCad)
    
End Function


Public Function devolver_lista_mantenimientos(Optional ByVal prm_filtro As String = "") As ADODB.Recordset

    Dim strCad As String

    strCad = "select eq_mantenimiento_equipos.* "
    strCad = strCad & ", coalesce(ca_documentos.nombre,'') as PROTOCOLO"
    strCad = strCad & ", coalesce(eq_planes_mantenimiento.nombre,'') as PLAN_MANTENIMIENTO"
    strCad = strCad & ", concat(coalesce(usuarios.nombre,''), ' ', coalesce(usuarios.apellidos,'')) as RESPONSABLE"
    strCad = strCad & ", coalesce(decodificadora.descripcion,'') as PERIODICIDAD"
    strCad = strCad & " from eq_mantenimiento_equipos"
    strCad = strCad & " INNER join usuarios on eq_mantenimiento_equipos.MANTENEDOR_id = usuarios.ID_empleado"
    strCad = strCad & " INNER join eq_planes_mantenimiento on eq_mantenimiento_equipos.planmto_id = eq_planes_mantenimiento.ID_PLAN_MTO"
    strCad = strCad & " INNER join decodificadora ON eq_planes_mantenimiento.FRECUENCIA_ID = decodificadora.valor and decodificadora.codigo = " & DECODIFICADORA.EQ_periodicidad
    strCad = strCad & " left outer join ca_documentos on eq_mantenimiento_equipos.PROCEDIMIENTO_ID = ca_documentos.ID_DOCUMENTO"
    strCad = strCad & " where equipo_id = " & ID_EQUIPO
    If Trim(prm_filtro) <> "" Then
        strCad = strCad & prm_filtro
    End If
    
    If LCase(prm_filtro) = " and estado = 0" Then
        strCad = strCad & " order by fecha_actual asc, id_mantenimiento"
    Else
        strCad = strCad & " order by fecha_actual desc, id_mantenimiento"
    End If
    
    'strCad = strCad & " order by fecha_actual desc, id_mantenimiento"
    
    Set devolver_lista_mantenimientos = datos_bd(strCad)
    
End Function


Public Function devolver_lista_mantenimientos_historicos() As ADODB.Recordset
    
    Set devolver_lista_mantenimientos_historicos = devolver_lista_mantenimientos(" and estado > 0")
    
End Function
Public Function devolver_lista_mantenimientos_pendientes() As ADODB.Recordset
    
    Set devolver_lista_mantenimientos_pendientes = devolver_lista_mantenimientos(" and estado = 0")
    
End Function





Public Function Carga_Datos_Basicos(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_EQUIPO,NOMBRE,SERIE,NUMERO_USOS_MAXIMO FROM EQUIPOS WHERE ID_EQUIPO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_Datos_Basicos = False
        Else
                ID_EQUIPO = rs("ID_EQUIPO")
                nombre = rs("NOMBRE")
                SERIE = rs("SERIE")
                NUMERO_USOS_MAXIMO = rs("NUMERO_USOS_MAXIMO")
                Carga_Datos_Basicos = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_Datos_Basicos = False
        'Resume Next
        MsgBox "Error al cargar los datos basicos(clsEquipos)", vbCritical, Err.Description
End Function

Private Sub avisar_dpto_equipo_fuera_servicio(ByVal fuera_servicio_definitivo As Boolean, ByVal tipo_cvm As Integer)
Dim objMensaje As New clsMensajes, objDpto As clsGenericClass
Dim strSql As String, strCad As String, strArr() As String
Dim idRespCalidad() As Long, indexResp As Long, intCont As Integer
Dim rs As ADODB.Recordset
Dim ASUNTO As String
Dim texto As String
Dim mens As Integer, x As Integer
Dim blnRespYaEnLista As Boolean
Dim tipo As String


    Call devolverListaResponsablesDpto(idRespCalidad, enumDPTO.Calidad)
    Call devolverListaResponsablesDpto(idRespCalidad, enumDPTO.METROLOGIA)
    Call devolverListaResponsablesDpto(idRespCalidad, enumDPTO.LABORATORIO_AERONAUTICO)
    Call devolverListaResponsablesDpto(idRespCalidad, enumDPTO.LABORATORIO_AGROALIMENTARIO)
    Call devolverListaResponsablesDpto(idRespCalidad, enumDPTO.GERENCIA)
    Call devolverListaResponsablesDpto(idRespCalidad, enumDPTO.I_D)
    
    
    
    If tipo_cvm = EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION Then
        tipo = "Calibración"
    ElseIf tipo_cvm = EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION Then
        tipo = "Verificación"
    Else
        tipo = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
    End If
    
    On Error Resume Next
    indexResp = UBound(idRespCalidad)
    If Err.Number <> 0 Then indexResp = 0
    
    On Error GoTo 0
    
    If indexResp = 0 Then
        ' cuando no hay responsables de calidad, se lo envia a Gerencia
        ReDim Preserve idRespCalidad(0)
        idRespCalidad(0) = 1 ' El ID 1 es el Gerente.
        indexResp = 1
    End If
    
    If tipo_cvm = EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION Or tipo_cvm = EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION Then
    
        If Not fuera_servicio_definitivo Then
            ASUNTO = tipo & " NO CONFORME Equipo Nº " & Format(ID_EQUIPO, "#,#00")
            texto = "Con fecha " & Format(Now, "dd/mm/yyyy hh:nn:ss") & ", se ha realizado una " & tipo & " No Conforme. Si la próxima " & tipo & " vuelve a resultar no conforme el equipo quedará fuera de servicio."
        Else
            ASUNTO = "FUERA DE SERVICIO Equipo Nº " & Format(ID_EQUIPO, "#,#00")
            texto = "Con fecha " & Format(Now, "dd/mm/yyyy hh:nn:ss") & ", se ha realizado una segunda " & tipo & " No Conforme Consecutiva. El equipo ha quedado fuera de servicio hasta que se realice una nueva " & tipo & " con resultado Conforme."
        End If
    Else
        ASUNTO = "FUERA DE SERVICIO Equipo Nº " & Format(ID_EQUIPO, "#,#00")
        texto = "Con fecha " & Format(Now, "dd/mm/yyyy") & ", el usuario " & tipo & " ha puesto Fuera de Servicio el Equipo Nº " & Format(ID_EQUIPO, "#,#00")
    End If
    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmEquipoEdicion;" & CStr(ID_EQUIPO), Now)

End Sub

Public Sub devolverListaResponsablesDpto(ByRef idRespCalidad() As Long, prm_id_departamento As Long)

Dim indexResp As Integer, blnRespYaEnLista As Boolean
Dim rs As ADODB.Recordset
Dim strSql As String, strCad As String
Dim strArr() As String
Dim intCont As Integer, x As Integer
On Error GoTo devolverListaResponsablesDpto_Error

indexResp = 0

    strSql = "SELECT parametros From decodificadora where valor = " & CStr(prm_id_departamento) & " and codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & ""
    
    Set rs = datos_bd(strSql)
    
    On Error Resume Next
    
    indexResp = UBound(idRespCalidad)
    If Err.Number <> 0 Then indexResp = -1
    
    On Error GoTo devolverListaResponsablesDpto_Error
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            strCad = rs("parametros")
            If Trim(strCad) <> "" And InStr(1, strCad, "=") > 0 Then
                strArr = Split(strCad, ";")
                For intCont = 0 To UBound(strArr) - 1
                    strCad = Split(strArr(intCont), "=")(1)
                    If Trim(strCad) <> "0" Then
                        blnRespYaEnLista = False
                        If indexResp > 0 Then
                            For x = 0 To indexResp - 1
                                If idRespCalidad(x) = CLng(strCad) Then
                                    blnRespYaEnLista = True
                                End If
                            Next x
                        End If
                        If Not blnRespYaEnLista Then
                            indexResp = indexResp + 1
                            ReDim Preserve idRespCalidad(indexResp)
                            idRespCalidad(indexResp) = CLng(strCad)
                        End If
                    End If
                Next intCont
            End If
            rs.MoveNext
        Wend
    End If

    

On Error GoTo 0
Exit Sub
devolverListaResponsablesDpto_Error:
End Sub


Private Sub EnviarMensajeInterno(ByRef prmListaUsuarioDestinatario() As Long, ByVal prmAsunto As String, ByVal prmCuerpo As String, ByVal prmEnlaceForm As String, prmFechaEnvio As Date)
Dim objMensaje As New clsMensajes
Dim mens As Long
Dim omu As New clsMensajes_usuarios
Dim x As Integer
Dim lista_enviados As String
On Error GoTo EnviarMensajeInterno_Error

With objMensaje
        .setASUNTO = prmAsunto
        .setTEXTO = prmCuerpo
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA_INICIO = Format(prmFechaEnvio, "yyyy-mm-dd")
        .setFECHA_FIN = Format(prmFechaEnvio + 7, "yyyy-mm-dd")
        .setACCION = prmEnlaceForm
        
        .setHORA_INICIO = Format(Time, "hh:mm:ss")
        .setHORA_FIN = Format(Time, "hh:mm:ss")
        .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_EQUIPOS
        .setDURACION = 0
        
        mens = .Insertar
        If mens > 0 Then
            For x = 0 To UBound(prmListaUsuarioDestinatario())
                If InStr(lista_enviados, ":" & prmListaUsuarioDestinatario(x) & ":") = 0 Then
                    lista_enviados = lista_enviados & ":" & prmListaUsuarioDestinatario(x) & ":"
                    omu.setEMPLEADO_ID = prmListaUsuarioDestinatario(x)
                    omu.setMENSAJE_ID = mens
                    
                    omu.Insertar
                End If
            Next x
            frmCalendario.cargar_eventos
        End If
End With

On Error GoTo 0
Exit Sub
EnviarMensajeInterno_Error:
End Sub

Public Function volver_dar_alta_equipo() As Boolean
On Error GoTo err_alta

    Dim cad As String
    'M1124-I
    'cad = "update equipos set alta_baja = 0 where id_equipo = " & ID_EQUIPO
    cad = "update equipos set alta_baja = 0,ESTADO_ID = '" & ENUM_EQ_ESTADOS_ACTIVO & "' where id_equipo = " & ID_EQUIPO
    'M1124-F
    execute_bd cad

    Dim oEV As New clsEquipoEventos
    With oEV
        .setEQUIPO_ID = ID_EQUIPO
        .setEVENTO_ID = EQUIPOS_EVENTOS.EVT_CAMBIO_ESTADO_EQUIPO
        .setRAZON_ID = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_ALTA_DE_EQUIPO
        .setOBSERVACIONES = ""
        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
        .Insertar
    End With
    Set oEV = Nothing

    volver_dar_alta_equipo = True

Exit Function
err_alta:
End Function

'M1326-I
Public Function abrir_cerrar_equipo(ESTADO As Integer) As Boolean
On Error GoTo err_abrir_cerrar_equipo
    Dim strmotivo As String
        
    If ESTADO = 1 Then
       strmotivo = EQUIPOS_EVENTOS_MOTIVOS.EVTR_APERTURA
    Else
       strmotivo = EQUIPOS_EVENTOS_MOTIVOS.EVTR_CIERRE
    End If
        
    generar_evento ID_EQUIPO, EQUIPOS_EVENTOS.EVT_APERTURA_CIERRE_EQUIPO, strmotivo, "", 0
    abrir_cerrar_equipo = True
Exit Function
err_abrir_cerrar_equipo:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure abrir_cerrar_equipo of Módulo de clase clsEquipo"
    abrir_cerrar_equipo = False
End Function

Public Function ultimo_movimiento_apertura() As ADODB.Recordset

    Dim consulta As String
    Dim strMotivos As String
    
    strMotivos = " AND ev.razon_id in (" & EQUIPOS_EVENTOS_MOTIVOS.EVTR_APERTURA & "," & EQUIPOS_EVENTOS_MOTIVOS.EVTR_CIERRE & ")"
    consulta = "SELECT ev.* " & _
               "  from eq_eventos ev " & _
               " where ev.equipo_id = " & CStr(ID_EQUIPO) & " and ev.evento_id = " & EQUIPOS_EVENTOS.EVT_APERTURA_CIERRE_EQUIPO & strMotivos & _
               " order by ev.id_EventoEquipo desc " & _
               " Limit 1"
  
    Set ultimo_movimiento_apertura = datos_bd(consulta)
End Function
'M1326-F
