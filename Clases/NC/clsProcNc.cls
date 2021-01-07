VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcNc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSNC
'***************************************************************
Private id_procnc As Long
Private ID_PARTICULAR As Long
Private TIPO_ID As Long
Private tipo As String
Private RESPONSABLE_ID As Long
Private RESPONSABLE_ID_DEPARTAMENTO As Long
Private RESPONSABLE_NOMBRE_APELLIDOS As String
Private RESPONSABLE_DEPARTAMENTO As String
Private FECHA_ALTA As Date
Private DESCRIPCION_INCIDENCIA As String
Private fecha_cierre As Date
Private FECHA_ULT_MOVIMIENTO As Date
Private ESTADO_ID As Long
Private ORIGEN_ID As Long
Private AUDITORIA_ID As Long
Private ID_AUX As Integer
Private ORIGEN_OTROS As String
Private OBSERVACIONES_RESULTADO As String
Private es_no_conformidad As Integer
Private es_solucion_aceptable As Integer
Private EVIDENCIAS_EN_PLAZO As Integer
Private EVIDENCIAS_EFECTIVAS As Integer
Private EVIDENCIAS_EVIDENCIAS As Integer
Private EVIDENCIAS_COMUNICADO_MODIFICACIONES As Integer
Private RESUMEN As String
Private ANALISIS_ESCENA_LOCALIZACION As String
Private ANALISIS_ESCENA_FECHA As Date
Private ANALISIS_ESCENA_HORA As Integer
Private ANALISIS_ESCENA_MINUTOS As Integer
Private ANALISIS_ESCENA_CAMBIOS_RECIENTES As String
Private ANALISIS_ESCENA_CONDICIONES_AMBIENTALES As String
Private ANALISIS_ESCENA_CONDICIONES_OPERACION As String
Private ANALISIS_ESCENA_EQUIPOSIMPLICADOS As String
Private ANALISIS_ESCENA_EXPERIENCIA As String
Private ANALISIS_ESCENA_FORMACION As String
Private ANALISIS_ESCENA_SECUENCIA As String
Private ANALISIS_ESCENA_COMUNICACION As String
Private RECOLECCION_DATOS As String

Private PROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO As Boolean
Private PROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS As Boolean
Private PROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS As Boolean
Private PROBLEMAS_HUMANOS_SUFICIENTE_FORMACION As Boolean
Private PROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE As Boolean
Private PROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO As Boolean

Private PROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA As String
Private PROBLEMAS_CLASIFICACION_AFECTADOS As String
Private PROBLEMAS_CLASIFICACION_ALCANCE As String

' Causas

Private RESUMEN_CAUSAS As String
Private causa_directa As String
Private causa_raiz As String
Private CAUSA_CONTRIBUTIVA_1 As String
Private CAUSA_CONTRIBUTIVA_1_DESCRIPCION As String
Private CAUSA_CONTRIBUTIVA_2 As String
Private CAUSA_CONTRIBUTIVA_2_DESCRIPCION As String
Private CAUSA_CONTRIBUTIVA_3 As String
Private CAUSA_CONTRIBUTIVA_3_DESCRIPCION As String
Private CAUSA_CONTRIBUTIVA_4 As String
Private CAUSA_CONTRIBUTIVA_4_DESCRIPCION As String
Private CAUSA_CONTRIBUTIVA_5 As String
Private CAUSA_CONTRIBUTIVA_5_DESCRIPCION As String
Private DESVIACION_ID As Integer

'M1315-I
Private N_ACCIONES As Long
Private REVISADA_USUARIO_ID As Long
Private REVISADA_FECHA As String
'M1315-F

' Colecciones de datos
Private mvarobjColAccionesCorrectoras As New clsGenericCollection
Private mvarobjColAccionesInmediatas As New clsGenericCollection
Private mvarobjColPreguntasRespuestas As New clsGenericCollection
Private mvarobjColOrigenesIncidencia As New clsGenericCollection
Private mvarobjColDepartamentosImplicados As New clsGenericCollection
Private mvarobjColTrazabilidad As New clsGenericCollection
Private mvarobjColEquipoHumano As New clsGenericCollection
Private mvarobjColPersonalImplicado As New clsGenericCollection
Private mvarobjColCausas As New clsGenericCollection
'De la pestaña 2
Private mvarobjColAdjuntosIncidencia As New clsGenericCollection
Private mvarobjColAdjuntosIndentificacionProblemas As New clsGenericCollection
Private mvarobjColAdjuntosRecoleccionDatos As New clsGenericCollection

Private mvarstrRutaInicial As String


Public Function anadir_personal_implicado(ByVal prmId As Long) As Boolean

    Dim strCad As String
    Dim rs As ADODB.Recordset

On Error GoTo anadir_personal_implicado_Error
    
    anadir_personal_implicado = False
    
    strCad = "SELECT * FROM procnc_personalimplicado WHERE id_usuario = " & prmId & " AND id_procnc = " & id_procnc
    Set rs = datos_bd(strCad)
    If rs.RecordCount <> 0 Then
        Set rs = Nothing
        anadir_personal_implicado = True
        Exit Function
    End If
    
        
    strCad = "insert into procnc_personalimplicado(id_usuario, id_procnc) VALUES (" & prmId & ", " & id_procnc & ")"

    execute_bd strCad
    
    anadir_personal_implicado = True
    
On Error GoTo 0
    Exit Function
anadir_personal_implicado_Error:
    anadir_personal_implicado = False
End Function


Public Function anadir_equipamiento_implicado(ByVal prmId As Long) As Boolean

    Dim strCad As String
    Dim rs As ADODB.Recordset
On Error GoTo anadir_equipamiento_implicado_Error
    
    anadir_equipamiento_implicado = False
    
    strCad = "SELECT * FROM procnc_equipamientoimplicado WHERE id_equipo = " & prmId & " AND id_procnc = " & id_procnc
    Set rs = datos_bd(strCad)
    If rs.RecordCount <> 0 Then
        Set rs = Nothing
        anadir_equipamiento_implicado = True
        Exit Function
    End If
       
    strCad = "insert into procnc_equipamientoimplicado(id_equipo, id_procnc) VALUES (" & prmId & ", " & id_procnc & ")"
    execute_bd strCad
    anadir_equipamiento_implicado = True
    
On Error GoTo 0
    Exit Function
anadir_equipamiento_implicado_Error:
    anadir_equipamiento_implicado = False
End Function
Public Function anadir_departamento_origen(ByVal strId As String) As Boolean

    Dim strCad As String, rs As ADODB.Recordset

On Error GoTo anadir_departamento_origen_Error
    
    anadir_departamento_origen = False
    
    If Trim(strId = "") Then
        MsgBox "No ha señalado un Departamento de Origen correctamente.", vbInformation, "Añádir Departamento de Origen"
        Exit Function
    End If
    
    strCad = " select * from procnc_departamentosimplicados where id_procnc = " & CStr(id_procnc) & " AND id_departamento = " & strId
    
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount <> 0 Then
        anadir_departamento_origen = True
        Set rs = Nothing
        Exit Function
    End If
    
    strCad = "insert into procnc_departamentosimplicados(id_departamento, id_procnc) VALUES (" & strId & ", " & CStr(id_procnc) & ")"

    execute_bd strCad
    
    anadir_departamento_origen = True
    
On Error GoTo 0
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.anadir_accion_inmediata"
    Exit Function
anadir_departamento_origen_Error:
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.anadir_accion_inmediata"
    'error_grave Err.Number & " (" & Err.Description & ") in procedure anadir_accion_inmediata of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    'G_TRAZABILIDAD_ERROR = ""

End Function



Public Function anadir_equipo_investigacion_departamento(ByVal prmId As Long) As Boolean

    Dim strCad As String, rs As ADODB.Recordset
    Dim strParam As String, lngid_jefe As Long

On Error GoTo anadir_equipo_investigacion_departamento_Error
    
    ' Busca si existe un jefe de equipo ya en el equipo
    lngid_jefe = -1
    strCad = " select * from procnc_equipohumano where jefe_equipo = 1 and id_procnc = " & id_procnc
    Set rs = datos_bd(strCad)
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        ' lo pone en negativo para saber despues si es un nuevo jefe o es antiguo
        lngid_jefe = (-1) * CInt(rs("id_usuario"))
    End If
    
    ' carga los jefes de departamento
    
    ' carga los jefes de departamento
    strCad = " select parametros from decodificadora where codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & "  and valor = " & CStr(prmId)
    Set rs = datos_bd(strCad)
    
    rs.MoveFirst
    strParam = rs("parametros")
    
    ' carga los usuarios del departamento
    strCad = "Select * from usuarios_departamentos where id_departamento = " & CStr(prmId) & " and id_usuario not in (select id_usuario from procnc_equipohumano where id_procnc = " & CStr(id_procnc) & ")"
    Set rs = datos_bd(strCad)
    
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        'lngid_jefe = rs("id_usuario")
        While Not rs.EOF
            If InStr(1, strParam, "idresp=" & rs("id_usuario") & ";") > 0 And (lngid_jefe = -1) Then
                lngid_jefe = rs("id_usuario")
                strCad = "insert into procnc_equipohumano (id_procnc, id_usuario, jefe_equipo) values (" & CStr(id_procnc) & ", " & rs("id_usuario") & ", 1)"
            Else
                strCad = "insert into procnc_equipohumano (id_procnc, id_usuario, jefe_equipo) values (" & CStr(id_procnc) & ", " & rs("id_usuario") & ", 0)"
            End If
            execute_bd strCad
            
            rs.MoveNext
        Wend
    End If
        
    ' Si aún no se ha designado a un jefe de equipo, se estima el primer usuario encontrado
    If lngid_jefe > 0 Then
        strCad = "update procnc_equipohumano set jefe_equipo = 0 where id_procnc = " & CStr(id_procnc)
        execute_bd strCad
        
        strCad = "update procnc_equipohumano set jefe_equipo = 1 where id_procnc = " & CStr(id_procnc) & " and id_usuario = " & lngid_jefe
        execute_bd strCad
    End If
    
    anadir_equipo_investigacion_departamento = True
    
On Error GoTo 0
    Exit Function
anadir_equipo_investigacion_departamento_Error:

End Function

Public Function anadir_equipo_investigacion_usuario(ByVal prmId As Long) As Boolean

    Dim strCad As String, rs As ADODB.Recordset
    Dim strParam As String

On Error GoTo anadir_equipo_investigacion_usuario_Error
    
    
    ' busca si el usuario ya existe en el equipo humano
    strCad = "select * from procnc_equipohumano where id_usuario = " & prmId & " AND id_procnc = " & id_procnc
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount <> 0 Then
        ' El usuario ya existe, sale de la funcion
        Set rs = Nothing
        anadir_equipo_investigacion_usuario = True
        Exit Function
    End If
        
    
    strCad = "insert into procnc_equipohumano (id_procnc, id_usuario, jefe_equipo) values (" & CStr(id_procnc) & ", " & prmId & ", 0)"
    execute_bd strCad
    
    anadir_equipo_investigacion_usuario = True
    
On Error GoTo 0
    Exit Function
anadir_equipo_investigacion_usuario_Error:

End Function





Public Function Carga_desde_correctiva(ByVal id_correctiva As Long) As Boolean
Dim rs As ADODB.Recordset
Dim lng_id As Long

Set rs = datos_bd("SELECT id_procnc from procnc_accionescorrectivas where id_accion_correctiva = " & CStr(id_correctiva))

rs.MoveFirst
lng_id = rs("id_procnc")

Carga_desde_correctiva = Carga(lng_id)

End Function



Public Function eliminar_equipo_investigacion_usuario(ByVal prmId As Long) As Boolean

    Dim strCad As String, rs As ADODB.Recordset
    
On Error GoTo eliminar_equipo_investigacion_usuario_Error
    
    
    ' busca si el usuario ya existe en el equipo humano
    strCad = "delete from procnc_equipohumano where id_usuario = " & prmId & " AND id_procnc = " & id_procnc
    execute_bd strCad
        
    eliminar_equipo_investigacion_usuario = True
    
On Error GoTo 0
    Exit Function
eliminar_equipo_investigacion_usuario_Error:

End Function

Public Function comprobar_equipo_investigacion_tiene_jefe_equipo() As Boolean

    Dim strCad As String, rs As ADODB.Recordset
    
On Error GoTo eliminar_equipo_investigacion_usuario_Error
    
    
    ' busca si el usuario ya existe en el equipo humano
    strCad = "select * from procnc_equipohumano where jefe_equipo = 1 AND id_procnc = " & id_procnc
    Set rs = datos_bd(strCad)
        
    If rs.RecordCount <> 0 Then
        comprobar_equipo_investigacion_tiene_jefe_equipo = True
    Else
        comprobar_equipo_investigacion_tiene_jefe_equipo = False
    End If
    
    Set rs = Nothing
    
On Error GoTo 0
    Exit Function
eliminar_equipo_investigacion_usuario_Error:
    comprobar_equipo_investigacion_tiene_jefe_equipo = False
End Function








Public Function anadir_origen_incidencia(ByVal strId As String, nombre As String) As Boolean

    Dim strCad As String, rs As ADODB.Recordset

On Error GoTo anadir_origen_incidencia_Error
    
    anadir_origen_incidencia = False
    
    If Trim(strId = "") Then
        MsgBox "No ha señalado un Origen de la Incidencia correctamente.", vbInformation, "Señalar Origen de la Incidencia"
        Exit Function
    End If
    
    strCad = " SELECT * FROM procnc_origenesincidencia WHERE id_origen_incidencia = " & strId & " AND id_procnc = " & id_procnc
    
    Set rs = datos_bd(strCad)
    If rs.RecordCount = 0 Then
    
        strCad = "insert into procnc_origenesincidencia(id_origen_incidencia, descripcion, id_procnc) SELECT " & strId & ", '" & nombre & "', " & CStr(id_procnc) & " "

        execute_bd strCad
    End If
    
    Set rs = Nothing
    anadir_origen_incidencia = True
    
On Error GoTo 0
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.anadir_accion_inmediata"
    Exit Function
anadir_origen_incidencia_Error:
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.anadir_accion_inmediata"
    'error_grave Err.Number & " (" & Err.Description & ") in procedure anadir_accion_inmediata of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    'G_TRAZABILIDAD_ERROR = ""

End Function



Public Function anadir_clasificacion_problema_incidencia(ByVal prmId As Long, Optional ByVal prmid_procnc As Long = 0) As Boolean

    Dim strCad As String, rs As ADODB.Recordset

On Error GoTo anadir_clasificacion_problema_incidencia_Error
    
    anadir_clasificacion_problema_incidencia = False
    
    If prmid_procnc = 0 Then
        prmid_procnc = id_procnc
    End If
    
    Set rs = datos_bd("select id_tipocausa from procnc_causas where id_causa = " & CStr(prmId) & " AND id_procnc = " & prmid_procnc)
    If rs.RecordCount <> 0 Then Exit Function
    
    
    Set rs = datos_bd("select id_tipocausa from procnc_causasproblemas where id_causa = " & CStr(prmId))
    If rs.RecordCount = 0 Then Exit Function
    
    rs.MoveFirst
        
    strCad = "insert into procnc_causas(id_causa, id_tipocausa, id_procnc) VALUES (" & prmId & ", " & rs("id_tipocausa") & ", " & CStr(id_procnc) & ")"

    Set rs = Nothing
    
    execute_bd strCad
    
    anadir_clasificacion_problema_incidencia = True
    
On Error GoTo 0
    Exit Function
anadir_clasificacion_problema_incidencia_Error:

Set rs = Nothing
End Function

Public Function devolver_listado_origenes(ByVal indice As Integer) As ADODB.Recordset

    Dim strCad As String
    
    
    strCad = "SELECT * FROM decodificadora where codigo = " & DECODIFICADORA.PROCNC_ORIGENES & " and parametros like 'idraiz=" & CStr(indice) & ";'"
    
    Set devolver_listado_origenes = datos_bd(strCad)

End Function

Public Function devolver_listado_clasificacion_problemas(ByVal indice As Integer) As ADODB.Recordset

    Dim strCad As String
    
    
    strCad = "SELECT * FROM procnc_causasproblemas where autorizado = 1"
    
    Set devolver_listado_clasificacion_problemas = datos_bd(strCad)

End Function

Public Function devolver_origenes_incidencia() As ADODB.Recordset
On Error GoTo devolver_origenes_incidencia_Error
    
    Dim strSql As String
    
    
    strSql = "select * from procnc_origenesincidencia Where id_procnc = " & id_procnc
    
    
    Set devolver_origenes_incidencia = datos_bd(strSql)
    
On Error GoTo 0
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolver_origenes_incidenica"
    Exit Function
devolver_origenes_incidencia_Error:
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolver_origenes_incidenica"
    'error_grave Err.Number & " (" & Err.Description & ") in procedure devolver_origenes_incidenica of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    'G_TRAZABILIDAD_ERROR = ""
End Function
Public Function devolver_clasificacion_problemas_incidencia() As ADODB.Recordset
On Error GoTo devolver_clasificacion_problemas_incidencia_Error
    
    Dim strSql As String
    
    
    strSql = "select * from procnc_causas Where id_procnc = " & id_procnc & " order by id_tipocausa, id_causa"
    
    
    Set devolver_clasificacion_problemas_incidencia = datos_bd(strSql)
    
On Error GoTo 0
    Exit Function
devolver_clasificacion_problemas_incidencia_Error:

End Function
Public Function devolver_departamentos_origen() As ADODB.Recordset
On Error GoTo devolver_departamentos_origen_Error
    
    Dim strSql As String
        
    strSql = "select procnc_departamentosimplicados.*, decodificadora.descripcion as departamento" & _
    " From procnc_departamentosimplicados" & _
    " inner join decodificadora on procnc_departamentosimplicados.id_departamento = decodificadora.valor and decodificadora.CODIGO = " & CStr(DECODIFICADORA.PROCNC_DEPARTAMENTOS) & _
    " Where procnc_departamentosimplicados.id_procnc = " & CStr(id_procnc)

    Set devolver_departamentos_origen = datos_bd(strSql)
    
On Error GoTo 0
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolver_origenes_incidenica"
    Exit Function
devolver_departamentos_origen_Error:
    'G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolver_origenes_incidenica"
    'error_grave Err.Number & " (" & Err.Description & ") in procedure devolver_origenes_incidenica of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    'G_TRAZABILIDAD_ERROR = ""
End Function
Public Function devolver_equipo_investigacion() As ADODB.Recordset
On Error GoTo devolver_equipo_investigacion_Error
    
    Dim strSql As String
        
    strSql = "select procnc_equipohumano.*, concat(usuarios.nombre, ' ', usuarios.apellidos) as usuario" & _
    "  from procnc_equipohumano" & _
    " inner join usuarios on  procnc_equipohumano.id_usuario = usuarios.id_empleado" & _
    " Where procnc_equipohumano.id_procnc = " & CStr(id_procnc)

    Set devolver_equipo_investigacion = datos_bd(strSql)
    
On Error GoTo 0
    Exit Function
devolver_equipo_investigacion_Error:

End Function

Public Function devolver_listado_personal_implicado() As ADODB.Recordset
On Error GoTo devolver_listado_personal_implicado_Error
    
    Dim strSql As String
        
    strSql = "select procnc_personalimplicado.*, concat(usuarios.nombre, ' ', usuarios.apellidos) as usuario" & _
    "  from procnc_personalimplicado" & _
    " inner join usuarios on  procnc_personalimplicado.id_usuario = usuarios.id_empleado" & _
    " Where procnc_personalimplicado.id_procnc = " & CStr(id_procnc)

    Set devolver_listado_personal_implicado = datos_bd(strSql)
    
On Error GoTo 0
    Exit Function
devolver_listado_personal_implicado_Error:

End Function

Public Function devolver_listado_equipos_implicados() As ADODB.Recordset
On Error GoTo devolver_listado_equipos_implicados_Error
    
    Dim strSql As String
        
    strSql = "select procnc_equipamientoimplicado.*, concat(equipos.Nombre,' (Nº ',CAST(equipos.ID_EQUIPO AS CHAR),')') as equipo" & _
    "  from procnc_equipamientoimplicado" & _
    " inner join equipos on  procnc_equipamientoimplicado.id_equipo = equipos.id_equipo" & _
    " Where procnc_equipamientoimplicado.id_procnc = " & CStr(id_procnc)

    Set devolver_listado_equipos_implicados = datos_bd(strSql)
    
On Error GoTo 0
    Exit Function
devolver_listado_equipos_implicados_Error:

End Function


Public Function devolver_listado_secuencia_investigacion() As ADODB.Recordset
On Error GoTo devolver_listado_secuencia_investigacion_Error
    
    Dim strSql As String
        
    strSql = "select procnc_secuenciainvestigacion.*" & _
    " from procnc_secuenciainvestigacion" & _
    " Where procnc_secuenciainvestigacion.id_procnc = " & CStr(id_procnc) & _
    " order by procnc_secuenciainvestigacion.orden asc"
    

    Set devolver_listado_secuencia_investigacion = datos_bd(strSql)
    
On Error GoTo 0
    Exit Function
devolver_listado_secuencia_investigacion_Error:

End Function

Public Function devolver_listado_acciones_correctivas() As ADODB.Recordset
On Error GoTo devolver_listado_acciones_correctivas_Error
    
    Dim strSql As String
        
    strSql = "SELECT procnc_accionescorrectivas.*, concat(usuarios.nombre, ' ', usuarios.apellidos) as RESPONSABLE, d2.descripcion as ESTADO, d3.descripcion as tipo" & _
    " FROM procnc_accionescorrectivas " & _
    " join usuarios on procnc_accionescorrectivas.responsable_id = usuarios.id_empleado" & _
    " join decodificadora as d2 on procnc_accionescorrectivas.estado_id = d2.valor and d2.codigo=111" & _
    " join decodificadora as d3 on procnc_accionescorrectivas.tipo_id = d3.valor and d3.codigo=" & DECODIFICADORA.PROCNC_ACCIONES_TIPOS & _
    " WHERE procnc_accionescorrectivas.id_procnc = " & CStr(id_procnc)
    
    
    Set devolver_listado_acciones_correctivas = datos_bd(strSql)
    
On Error GoTo 0
    Exit Function
devolver_listado_acciones_correctivas_Error:

End Function


Public Function devolverListaResponsablesCalidad(ByRef idRespCalidad() As Long)

Dim indexResp As Integer, blnRespYaEnLista As Boolean
Dim rs As ADODB.Recordset
Dim strSql As String, strCad As String
Dim strArr() As String
Dim intCont As Integer, x As Integer
On Error GoTo devolverListaResponsablesCalidad_Error

indexResp = 0

    strSql = "SELECT parametros From decodificadora where valor = " & enumDPTO.Calidad & " and codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & ""
    
    Set rs = datos_bd(strSql)
    indexResp = 0
    
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
                            idRespCalidad(indexResp - 1) = CLng(strCad)
                        End If
                    End If
                Next intCont
            End If
            rs.MoveNext
        Wend
    End If

    devolverListaResponsablesCalidad = True

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolverListaResponsablesCalidad"
    Exit Function
devolverListaResponsablesCalidad_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolverListaResponsablesCalidad"
    error_grave Err.Number & " (" & Err.Description & ") in procedure devolverListaResponsablesCalidad of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Public Function devolverListaResponsablesDepartamentos(ByRef idRespCalidad() As Long)
Dim indexResp As Integer, blnRespYaEnLista As Boolean
Dim rs As ADODB.Recordset
Dim strSql As String, strCad As String
Dim strArr() As String
Dim intCont As Integer, x As Integer

On Error GoTo devolverListaResponsablesDepartamentos_Error

indexResp = 0

strSql = "SELECT parametros From decodificadora where codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & ""
    
    Set rs = datos_bd(strSql)
    indexResp = 0
    
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
                            idRespCalidad(indexResp - 1) = CLng(strCad)
                        End If
                    End If
                Next intCont
            End If
            rs.MoveNext
        Wend
    End If

    devolverListaResponsablesDepartamentos = True

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolverListaResponsablesDepartamentos"
    Exit Function
devolverListaResponsablesDepartamentos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolverListaResponsablesDepartamentos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure devolverListaResponsablesDepartamentos of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function




Public Function eliminar_accion_correctiva(ByVal ID_ACCION As Long) As Boolean

    Dim strCad As String
    Dim rs As ADODB.Recordset
    
    strCad = "select * from procnc_accionescorrectivas where id_accion_correctiva = " & ID_ACCION
    
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount = 0 Then Exit Function
    
    rs.MoveFirst
    If rs("estado_id") > 1 Then
        Set rs = Nothing
        MsgBox "No es posible eliminar una Accion Correctiva una vez que ha pasado al estado Tramitación", vbInformation, "Eliminar Accion Correctiva"
        Exit Function
    End If
'JGM-I
'    strCad = "delete from procnc_accionescorrectivas_avisos_pendiente where accion_id = " & ID_ACCION
'    execute_bd strCad
'JGM-F
    strCad = "delete from procnc_accionescorrectivas where id_accion_correctiva = " & ID_ACCION
    execute_bd strCad
    
End Function

Private Sub establecer_id_particular()

On Error GoTo establecer_id_particular_Error
    
    Dim strCad As String, rs  As ADODB.Recordset, lngid As Long
    
    strCad = "SELECT coalesce(max(ID_PARTICULAR), 0) AS ID FROM PROCNC WHERE TIPO_ID = " & CStr(TIPO_ID)
    Set rs = datos_bd(strCad)
    
    rs.MoveFirst
    lngid = CLng(rs!ID) + 1
    
    strCad = "UPDATE PROCNC SET ID_PARTICULAR = " & CStr(lngid)
    execute_bd strCad
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.establecer_id_particular"
    Exit Sub
establecer_id_particular_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.establecer_id_particular"
    error_grave Err.Number & " (" & Err.Description & ") in procedure establecer_id_particular of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Public Function establecer_nivel_acceso() As Integer
Dim x As Integer
Dim blnRespCalidad As Boolean, blnRespOtros As Boolean


    ' si son resp. de Calidad o Gerencia, acceso total
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.GERENCIA) = vbChecked Or USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.Calidad) = vbChecked Then
        blnRespCalidad = True
    End If
    
    ' Si es Responsable de algun equipo
    For x = 3 To enumDPTO.TOTAL_DEPARTAMENTOS
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(x) = vbChecked Then
            blnRespOtros = True
            Exit For
        End If
    Next x

    ' Por Defecto, solo lectura
    establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.SOLO_LECTURA

    ' Si es un alta, al menos debe ser Responsable de Departamento
    If id_procnc = 0 Then
        If USUARIO.getPER_NC Then
            If blnRespCalidad Then
                establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL
            ElseIf blnRespOtros Then
                establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.RESPONSABLE_DEPARTAMENTO
            Else
                establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.SOLO_LECTURA
            End If
        Else
            
            MsgBox "No posee Privilegios , aun siendo Responsable de Departamento, para crear Procedimientos de No Conformidad", vbInformation, "Privilegios insuficientes"
            establecer_nivel_acceso = -1
            Exit Function
            'Unload Me
        End If
        Exit Function
    Else
        ' si son resp. de Calidad o Gerencia, acceso total
        If blnRespCalidad Then
            establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.ACCESO_TOTAL
            Exit Function
        End If
        
        ' Si es Responsable de algun equipo
        If blnRespOtros Then
            establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.RESPONSABLE_DEPARTAMENTO
            If Es_Jefe_Equipo(USUARIO.getID_EMPLEADO) Then
                ' si tambien es Jefe de investigacion, se indica
                establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.JEFE_EQUIPO_INVESTIGACION_RESPONSABLE_DEPARTAMENTO
            End If
            Exit Function
        End If
        
        If Es_Jefe_Equipo(USUARIO.getID_EMPLEADO) Then
            establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.JEFE_EQUIPO_INVESTIGACION
            Exit Function
        End If
        
        ' Cuando está en acciones correctivas
        If ESTADO_ID >= C_PROCNC_ESTADOS.pte_cierre Then
            If Tiene_Asignada_Correctiva(USUARIO.getID_EMPLEADO) Then
                establecer_nivel_acceso = C_PROCNC_NIVELES_ACCESO.SOLO_CORRECTIVAS_ASIGNADAS
                Exit Function
            End If
        End If
    End If
End Function
Public Function eliminar_secuencia_investigacion(ByVal prmOrden As Long)

    Dim strCad As String
    
    strCad = "delete from procnc_secuenciainvestigacion where orden = " & prmOrden & " and id_procnc = " & id_procnc
    execute_bd strCad
    
    ' reordena todos los elementos que están despues en la secuencia
    strCad = "update procnc_secuenciainvestigacion set orden = orden -1 where orden > " & prmOrden & " and id_procnc = " & id_procnc
    execute_bd strCad
    
    
    
End Function

Public Function eliminar_personal_implicado(ByVal prmId As Long)

    Dim strCad As String
    
    strCad = "delete from procnc_personalimplicado where id_usuario = " & prmId & " and id_procnc = " & id_procnc

    execute_bd strCad
    
End Function

Public Function eliminar_equipamiento_implicado(ByVal prmId As Long)

    Dim strCad As String
    
    strCad = "delete from procnc_equipamientoimplicado where id_equipo = " & prmId & " and id_procnc = " & id_procnc

    execute_bd strCad
    
End Function
Public Function eliminar_departamento_origen(ByVal ID_DPTO As Long)

    Dim strCad As String
    
    strCad = "delete from procnc_departamentosimplicados where id_departamento = " & CStr(ID_DPTO) & " and id_procnc = " & id_procnc

    execute_bd strCad
    
End Function

Public Function eliminar_origen_incidencia(ByVal ID As Long)

    Dim strCad As String
    
    strCad = "delete from procnc_origenesincidencia where id_origen_incidencia = " & CStr(ID) & " and id_procnc = " & id_procnc

    execute_bd strCad
    
End Function

Public Function eliminar_clasificacion_problema_incidencia(ByVal ID As Long)

    Dim strCad As String
    
    strCad = "delete from procnc_causas where id_causa = " & CStr(ID) & " and id_procnc = " & id_procnc

    execute_bd strCad
    
End Function
Public Function establecer_jefe_equipo(ByVal prmId As Long) As Boolean
On Error GoTo establecer_jefe_equipo_Error
    Dim strCad As String
    
    strCad = "update procnc_equipohumano set jefe_equipo = 0 where id_procnc = " & CStr(id_procnc)
    execute_bd strCad
    
    strCad = "update procnc_equipohumano set jefe_equipo = 1 where id_procnc = " & CStr(id_procnc) & " and id_usuario = " & prmId
    execute_bd strCad
        
    
On Error GoTo 0
    Exit Function
establecer_jefe_equipo_Error:

    establecer_jefe_equipo = False

End Function

Public Property Get getID_PARTICULAR() As Long
    getID_PARTICULAR = ID_PARTICULAR
End Property

Public Function guardar_datos_investigacion_escena(ByVal RecoleccionDatos As String, ByVal FECHAHORA As Date, ByVal localizacion As String, ByVal CondicionesOperacion As String, ByVal CondicionesAmbientales As String, ByVal Comunicacion As String, ByVal CambiosRecientes As String, ByVal FORMACION As String, ByVal Experiencia As String) As Boolean

    Dim strSql As String

    strSql = "UPDATE procnc SET RECOLECCION_DATOS = '" & Trim(RecoleccionDatos) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_FECHA = '" & Format(FECHAHORA, "yyyy-mm-dd") & "' "
    strSql = strSql & ", ANALISIS_ESCENA_HORA = " & DatePart("h", FECHAHORA) & " "
    strSql = strSql & ", ANALISIS_ESCENA_MINUTOS = " & DatePart("n", FECHAHORA) & " "
    strSql = strSql & ", ANALISIS_ESCENA_LOCALIZACION = '" & Trim(localizacion) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_CONDICIONES_OPERACION = '" & Trim(CondicionesOperacion) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_CONDICIONES_AMBIENTALES = '" & Trim(CondicionesAmbientales) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_COMUNICACION = '" & Trim(Comunicacion) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_CAMBIOS_RECIENTES = '" & Trim(CambiosRecientes) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_FORMACION = '" & Trim(FORMACION) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_EXPERIENCIA = '" & Trim(Experiencia) & "' "
    strSql = strSql & ", ANALISIS_ESCENA_EQUIPOSIMPLICADOS = '' "
    strSql = strSql & ", ANALISIS_ESCENA_SECUENCIA = '' "
    strSql = strSql & " Where ID_PROCNC = " & CStr(id_procnc)
    execute_bd strSql

End Function
Public Function guardar_datos_causas(ByVal prmcausa_raiz As String, ByVal prmcausa_directa As String, ByVal cc_1 As String, ByVal cc_desc_1 As String, ByVal cc_2 As String, ByVal cc_desc_2 As String, ByVal cc_3 As String, ByVal cc_desc_3 As String, ByVal cc_4 As String, ByVal cc_desc_4 As String, ByVal cc_5 As String, ByVal cc_desc_5 As String, ByVal DESVIACION_ID As Integer) As Boolean

    Dim strSql As String

    strSql = "UPDATE procnc SET RESUMEN_CAUSAS = '' "
    strSql = strSql & ", CAUSA_DIRECTA = '" & Trim(prmcausa_directa) & "' "
    strSql = strSql & ", CAUSA_RAIZ = '" & Trim(prmcausa_raiz) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_1 = '" & Trim(cc_1) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_1_DESCRIPCION = '" & Trim(cc_desc_1) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_2 = '" & Trim(cc_2) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_2_DESCRIPCION = '" & Trim(cc_desc_2) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_3 = '" & Trim(cc_3) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_3_DESCRIPCION = '" & Trim(cc_desc_3) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_4 = '" & Trim(cc_4) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_4_DESCRIPCION = '" & Trim(cc_desc_4) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_5 = '" & Trim(cc_5) & "' "
    strSql = strSql & ", CAUSA_CONTRIBUTIVA_5_DESCRIPCION = '" & Trim(cc_desc_5) & "' "
    strSql = strSql & ", DESVIACION_ID = '" & DESVIACION_ID & "' "
    strSql = strSql & " Where ID_PROCNC = " & CStr(id_procnc)
    
    execute_bd strSql

End Function

Public Function guardar_datos_evaluacion(ByVal prm_es_no_confommidad As Integer, ByVal prm_es_solucion_aceptable As Integer, ByVal prm_evidencia_en_plazo As Integer, ByVal prm_evidencia_efectiva As Integer, ByVal prm_evidencia_evidencia As Integer, ByVal prm_evidencia_comunicado_modificaciones As Integer, ByVal prm_observaciones As String) As Boolean
                                                                                                                                   
    Dim strSql As String

    strSql = "UPDATE procnc SET ES_NO_CONFORMIDAD = " & CStr(prm_es_no_confommidad)
    strSql = strSql & ", ES_SOLUCION_ACEPTABLE = " & CStr(prm_es_solucion_aceptable)
    strSql = strSql & ", EVIDENCIAS_EN_PLAZO = " & CStr(prm_evidencia_en_plazo)
    strSql = strSql & ", EVIDENCIAS_EFECTIVAS = " & CStr(prm_evidencia_efectiva)
    strSql = strSql & ", EVIDENCIAS_EVIDENCIAS = " & CStr(prm_evidencia_evidencia)
    strSql = strSql & ", EVIDENCIAS_COMUNICADO_MODIFICACIONES = " & CStr(prm_evidencia_comunicado_modificaciones)
    strSql = strSql & ", OBSERVACIONES_RESULTADO = '" & Trim(prm_observaciones) & "' "
    strSql = strSql & " Where ID_PROCNC = " & CStr(id_procnc)
    
    execute_bd strSql

End Function
Public Function guardar_datos_clasificacion_problemas(ByVal TotalProblemas As String, ByVal Afectados As String, ByVal ALCANCE As String, ByVal operador_sustituido As Integer, ByVal instrucciones_incompletas As Integer, ByVal objetivos_marcados_claramente As Integer, ByVal formacion_suficiente As Integer, ByVal herramientas_adecuadas As Integer, ByVal proceso_inusual_complejo As Integer) As Boolean

Dim strSql As String

    strSql = "UPDATE procnc SET PROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA = '" & Trim(TotalProblemas) & "' "
    strSql = strSql & ", PROBLEMAS_CLASIFICACION_AFECTADOS = '" & Trim(Afectados) & "' "
    strSql = strSql & ", PROBLEMAS_CLASIFICACION_ALCANCE = '" & Trim(ALCANCE) & "' "
    strSql = strSql & ", PROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO = " & operador_sustituido & " "
    strSql = strSql & ", PROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS = " & instrucciones_incompletas & " "
    strSql = strSql & ", PROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE = " & objetivos_marcados_claramente & " "
    strSql = strSql & ", PROBLEMAS_HUMANOS_SUFICIENTE_FORMACION = " & formacion_suficiente & " "
    strSql = strSql & ", PROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS = " & herramientas_adecuadas & " "
    strSql = strSql & ", PROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO = " & proceso_inusual_complejo & " "
    strSql = strSql & " Where ID_PROCNC = " & CStr(id_procnc)
    execute_bd strSql

 
End Function
Public Function anadir_secuencia_investigacion(ByVal strSecuencia As String) As Boolean

    Dim strCad As String

On Error GoTo anadir_secuencia_investigacion_Error
    
    anadir_secuencia_investigacion = False
    
    If Trim(strSecuencia) = "" Then
        MsgBox "La Secuencia no puede estar vacía.", vbInformation, "Añádir Secuencia"
        Exit Function
    End If
    
    strCad = "insert into procnc_secuenciainvestigacion(orden, id_procnc, descripcion) SELECT cast(coalesce(max(orden), 0) as unsigned)+1, " & id_procnc & ", '" & strSecuencia & "' from procnc_secuenciainvestigacion WHERE id_procnc = " & id_procnc

    execute_bd strCad
    
    anadir_secuencia_investigacion = True
    
On Error GoTo 0
    Exit Function
anadir_secuencia_investigacion_Error:
    anadir_secuencia_investigacion = False
End Function

Public Function guardar_datos_origen_incidencia(ByVal datos As String)
Dim strSql As String

    strSql = "UPDATE procnc SET origen_otros = '" & datos & "' WHERE ID_PROCNC = " & CStr(id_procnc)
    execute_bd strSql
    

End Function

Public Function secuencia_cambiar_orden(orden_ant, orden_nuevo) As Boolean
On Error GoTo secuencia_cambiar_orden_Error
    Dim strCad As String
    Dim rs As ADODB.Recordset
    Dim lngOrden As Long
    ' comprueba que no sea el primero
    If orden_nuevo <= 0 Then
        secuencia_cambiar_orden = True
        Exit Function
    End If
    
    ' comprueba que no sea el último
    
    strCad = "SELECT max(orden) as orden from procnc_secuenciainvestigacion where id_procnc = " & CStr(id_procnc)
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        secuencia_cambiar_orden = True
        Exit Function
    End If
    
    rs.MoveFirst
    lngOrden = CLng(rs("orden"))
        
    ' si se pretende bajar el orden a más de lo que hay, nada
    If orden_nuevo > lngOrden Then
        secuencia_cambiar_orden = True
        Exit Function
    End If
    
    
    strCad = "update procnc_secuenciainvestigacion set orden = " & CStr((-1) * orden_nuevo) & " where id_procnc = " & CStr(id_procnc) & " AND orden = " & CStr(orden_nuevo)
    execute_bd strCad
    strCad = "update procnc_secuenciainvestigacion set orden = " & CStr(orden_nuevo) & " where id_procnc = " & CStr(id_procnc) & " AND orden = " & CStr(orden_ant)
    execute_bd strCad
    strCad = "update procnc_secuenciainvestigacion set orden = " & CStr(orden_ant) & " where id_procnc = " & CStr(id_procnc) & " AND orden = " & CStr((-1) * orden_nuevo)
    execute_bd strCad
        
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.secuencia_cambiar_orden"
    Exit Function
secuencia_cambiar_orden_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.secuencia_cambiar_orden"
    error_grave Err.Number & " (" & Err.Description & ") in procedure secuencia_cambiar_orden of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function

Public Property Let setID_PARTICULAR(ByVal dato As Long)
    ID_PARTICULAR = dato
End Property
'M1315-I
Public Property Let setN_ACCIONES(ByVal dato As Long)
    N_ACCIONES = dato
End Property
Public Property Let setREVISADA_USUARIO_ID(ByVal dato As Long)
    REVISADA_USUARIO_ID = dato
End Property
Public Property Let setREVISADA_FECHA(ByVal dato As String)
    REVISADA_FECHA = dato
End Property
'M1315-F

Public Property Let setOBSERVACIONES_RESULTADO(ByVal dato As String)
    OBSERVACIONES_RESULTADO = dato
End Property
Public Property Get getOBSERVACIONES_RESULTADO() As String
    getOBSERVACIONES_RESULTADO = OBSERVACIONES_RESULTADO
End Property

Public Property Let setES_NO_CONFORMIDAD(ByVal dato As Integer)
    es_no_conformidad = dato
End Property
Public Property Let setES_SOLUCION_ACEPTABLE(ByVal dato As Integer)
    es_solucion_aceptable = dato
End Property
Public Property Let setEVIDENCIAS_EN_PLAZO(ByVal dato As Integer)
    EVIDENCIAS_EN_PLAZO = dato
End Property
Public Property Let setEVIDENCIAS_EFECTIVAS(ByVal dato As Integer)
    EVIDENCIAS_EFECTIVAS = dato
End Property
Public Property Let setEVIDENCIAS_EVIDENCIAS(ByVal dato As Integer)
    EVIDENCIAS_EVIDENCIAS = dato
End Property
Public Property Let setEVIDENCIAS_COMUNICADO_MODIFICACIONES(ByVal dato As Integer)
    EVIDENCIAS_COMUNICADO_MODIFICACIONES = dato
End Property

Public Property Get getES_NO_CONFORMIDAD() As Integer
    getES_NO_CONFORMIDAD = es_no_conformidad
End Property
Public Property Get getES_SOLUCION_ACEPTABLE() As Integer
    getES_SOLUCION_ACEPTABLE = es_solucion_aceptable
End Property
Public Property Get getEVIDENCIAS_EN_PLAZO() As Integer
    getEVIDENCIAS_EN_PLAZO = EVIDENCIAS_EN_PLAZO
End Property
Public Property Get getEVIDENCIAS_EFECTIVAS() As Integer
    getEVIDENCIAS_EFECTIVAS = EVIDENCIAS_EFECTIVAS
End Property
Public Property Get getEVIDENCIAS_EVIDENCIAS() As Integer
    getEVIDENCIAS_EVIDENCIAS = EVIDENCIAS_EVIDENCIAS
End Property
Public Property Get getEVIDENCIAS_COMUNICADO_MODIFICACIONES() As Integer
    getEVIDENCIAS_COMUNICADO_MODIFICACIONES = EVIDENCIAS_COMUNICADO_MODIFICACIONES
End Property
Public Property Let setRESUMEN(ByVal VALOR As String)
    RESUMEN = VALOR
End Property
Public Property Let setRESUMEN_CAUSAS(ByVal VALOR As String)
RESUMEN_CAUSAS = VALOR
End Property
Public Property Let setCAUSA_DIRECTA(ByVal VALOR As String)
causa_directa = VALOR
End Property
Public Property Let setCAUSA_RAIZ(ByVal VALOR As String)
causa_raiz = VALOR
End Property
Public Property Let setCAUSA_CONTRIBUTIVA_1(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_1 = VALOR
End Property
Public Property Let setCAUSA_CONTRIBUTIVA_1_DESCRIPCION(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_1_DESCRIPCION = VALOR
End Property
Public Property Let setCAUSA_CONTRIBUTIVA_2(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_2 = VALOR
End Property
Public Property Let setCAUSA_CONTRIBUTIVA_2_DESCRIPCION(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_2_DESCRIPCION = VALOR
End Property

Public Property Let setTIPO_ID(ByVal dato As Long)
    TIPO_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
    tipo = dato
End Property

Public Property Get getTIPO_ID() As Long
    getTIPO_ID = TIPO_ID
End Property
'M1315-I
Public Property Get getN_ACCIONES() As Long
    getN_ACCIONES = N_ACCIONES
End Property
Public Property Get getREVISADA_USUARIO_ID() As Long
    getREVISADA_USUARIO_ID = REVISADA_USUARIO_ID
End Property
Public Property Get getREVISADA_FECHA() As String
    getREVISADA_FECHA = REVISADA_FECHA
End Property
'M1315-F
Public Property Get getTIPO() As String
    getTIPO = tipo
End Property

Public Property Let setCAUSA_CONTRIBUTIVA_3(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_3 = VALOR
End Property
Public Property Let setCAUSA_CONTRIBUTIVA_3_DESCRIPCION(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_3_DESCRIPCION = VALOR
End Property

Public Property Let setCAUSA_CONTRIBUTIVA_4(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_4 = VALOR
End Property
Public Property Let setCAUSA_CONTRIBUTIVA_4_DESCRIPCION(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_4_DESCRIPCION = VALOR
End Property

Public Property Let setCAUSA_CONTRIBUTIVA_5(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_5 = VALOR
End Property
Public Property Let setCAUSA_CONTRIBUTIVA_5_DESCRIPCION(ByVal VALOR As String)
CAUSA_CONTRIBUTIVA_5_DESCRIPCION = VALOR
End Property
Public Property Let setDESVIACION_ID(ByVal VALOR As Integer)
    DESVIACION_ID = VALOR
End Property

Public Property Get getRESUMEN_CAUSAS() As String
    getRESUMEN_CAUSAS = RESUMEN_CAUSAS
End Property
Public Property Get getRESUMEN() As String
    getRESUMEN = RESUMEN
End Property
Public Property Get getCAUSA_DIRECTA() As String
    getCAUSA_DIRECTA = causa_directa
End Property
Public Property Get getCAUSA_RAIZ() As String
    getCAUSA_RAIZ = causa_raiz
End Property
Public Property Get getCAUSA_CONTRIBUTIVA_1() As String
    getCAUSA_CONTRIBUTIVA_1 = CAUSA_CONTRIBUTIVA_1
End Property
Public Property Get getCAUSA_CONTRIBUTIVA_1_DESCRIPCION() As String
    getCAUSA_CONTRIBUTIVA_1_DESCRIPCION = CAUSA_CONTRIBUTIVA_1_DESCRIPCION
End Property

Public Property Get getCAUSA_CONTRIBUTIVA_2() As String
    getCAUSA_CONTRIBUTIVA_2 = CAUSA_CONTRIBUTIVA_2
End Property
Public Property Get getCAUSA_CONTRIBUTIVA_2_DESCRIPCION() As String
    getCAUSA_CONTRIBUTIVA_2_DESCRIPCION = CAUSA_CONTRIBUTIVA_2_DESCRIPCION
End Property

Public Property Get getCAUSA_CONTRIBUTIVA_3() As String
    getCAUSA_CONTRIBUTIVA_3 = CAUSA_CONTRIBUTIVA_3
End Property
Public Property Get getCAUSA_CONTRIBUTIVA_3_DESCRIPCION() As String
    getCAUSA_CONTRIBUTIVA_3_DESCRIPCION = CAUSA_CONTRIBUTIVA_3_DESCRIPCION
End Property

Public Property Get getCAUSA_CONTRIBUTIVA_4() As String
    getCAUSA_CONTRIBUTIVA_4 = CAUSA_CONTRIBUTIVA_4
End Property
Public Property Get getCAUSA_CONTRIBUTIVA_4_DESCRIPCION() As String
    getCAUSA_CONTRIBUTIVA_4_DESCRIPCION = CAUSA_CONTRIBUTIVA_4_DESCRIPCION
End Property

Public Property Get getCAUSA_CONTRIBUTIVA_5() As String
    getCAUSA_CONTRIBUTIVA_5 = CAUSA_CONTRIBUTIVA_5
End Property
Public Property Get getCAUSA_CONTRIBUTIVA_5_DESCRIPCION() As String
    getCAUSA_CONTRIBUTIVA_5_DESCRIPCION = CAUSA_CONTRIBUTIVA_5_DESCRIPCION
End Property
Public Property Get getDESVIACION_ID() As Integer
    getDESVIACION_ID = DESVIACION_ID
End Property
Public Property Get getPROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO() As Boolean
    getPROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO = PROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO
End Property
Public Property Let setPROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO(ByVal VALOR As Boolean)
    PROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO = VALOR
End Property
Public Property Get getPROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS() As Boolean
    getPROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS = PROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS
End Property
Public Property Let setPROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS(ByVal VALOR As Boolean)
     PROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS = VALOR
End Property
Public Property Get getPROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS() As Boolean
    getPROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS = PROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS
End Property
Public Property Let setPROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS(ByVal VALOR As Boolean)
    PROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS = VALOR
End Property
Public Property Get getPROBLEMAS_HUMANOS_SUFICIENTE_FORMACION() As Boolean
    getPROBLEMAS_HUMANOS_SUFICIENTE_FORMACION = PROBLEMAS_HUMANOS_SUFICIENTE_FORMACION
End Property
Public Property Let setPROBLEMAS_HUMANOS_SUFICIENTE_FORMACION(ByVal VALOR As Boolean)
    PROBLEMAS_HUMANOS_SUFICIENTE_FORMACION = VALOR
End Property
Public Property Get getPROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE() As Boolean
    getPROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE = PROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE
End Property
Public Property Let setPROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE(ByVal VALOR As Boolean)
    PROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE = VALOR
End Property
Public Property Get getPROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO() As Boolean
    getPROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO = PROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO
End Property
Public Property Let setPROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO(ByVal VALOR As Boolean)
    PROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO = VALOR
End Property

Public Function cambiarEstado(ByVal prmintNuevoEstado As Integer, Optional ByVal prmstrMotivoCambioEstado As String = "") As Boolean
    Dim strSql As String
    

On Error GoTo cambiarEstado_Error

    strSql = "UPDATE procnc SET FECHA_ULT_MOVIMIENTO=CURRENT_TIMESTAMP, estado_id = " & prmintNuevoEstado & " WHERE ID_PROCNC = " & CStr(id_procnc)
    execute_bd strSql
    
    ESTADO_ID = prmintNuevoEstado
    
    If ESTADO_ID = C_PROCNC_ESTADOS.CERRADA Then
        establecer_id_particular
    End If
    
    Call RegistrarMovimientoRegistro("STATE", prmstrMotivoCambioEstado)
    
    Select Case prmintNuevoEstado
'        Case C_PROCNC_ESTADOS.PTE_VISTO_BUENO
'            Call EnviarMensaje_RespCalidadVoBo
'        Case C_PROCNC_ESTADOS.ABIERTA
'            Call EnviarMensaje_RechazadoVoBo(prmstrMotivoCambioEstado)
        Case C_PROCNC_ESTADOS.EN_TRAMITACION
'            If Trim(prmstrMotivoCambioEstado) <> "" Then
'                Call EnviarMensaje_VolverATramitacionPorRechazo(prmstrMotivoCambioEstado)
'            Else
                Call EnviarMensaje_RespCalidadVoBo
'                Call EnviarMensaje_VoBo_Ok
'            End If
'        Case C_PROCNC_ESTADOS.PTE_CONFIRMACION_CALIDAD
'            Call EnviarMensaje_PteConfirmacionCalidad
        Case C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS
            If Trim(prmstrMotivoCambioEstado) <> "" Then
                Call EnviarMensaje_VolverATramitacionPorRechazo(prmstrMotivoCambioEstado)
            Else
                Call EnviarMensaje_PtePlanAccionesCorrectivas
            End If
        Case C_PROCNC_ESTADOS.CERRADA
        
            strSql = "UPDATE procnc SET FECHA_CIERRE=CURRENT_TIMESTAMP WHERE ID_PROCNC = " & CStr(id_procnc)
            execute_bd strSql
        
            Call EnviarMensaje_CerradaTotal
'        Case C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL
'            Call EnviarMensaje_CerradaParcial
        Case C_PROCNC_ESTADOS.pte_cierre
            Call EnviarMensaje_PteCierre
        Case Else
    End Select
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> ID_PROCNC : " & CStr(id_procnc)
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.cambiarEstado"
    Exit Function
cambiarEstado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> ID_PROCNC : " & CStr(id_procnc)
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.cambiarEstado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cambiarEstado of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Private Sub EnviarMensajeInterno(ByRef prmListaUsuarioDestinatario() As Long, ByVal prmAsunto As String, ByVal prmCuerpo As String, ByVal prmEnlaceForm As String, prmFechaEnvio As Date)
Dim objMensaje As New clsMensajes
Dim mens As Long
Dim omu As New clsMensajes_usuarios
Dim x As Integer

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
        .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_PROCN
        .setDURACION = 0
        
        
        mens = .Insertar
        If mens > 0 Then
            For x = 0 To UBound(prmListaUsuarioDestinatario()) - 1
                omu.setEMPLEADO_ID = prmListaUsuarioDestinatario(x)
                omu.setMENSAJE_ID = mens
                omu.Insertar
            Next x
            frmCalendario.cargar_eventos
        End If
End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensajeInterno"
    Exit Sub
EnviarMensajeInterno_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensajeInterno"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensajeInterno of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Public Property Get EquipoHumano() As clsGenericCollection
    Set EquipoHumano = mvarobjColEquipoHumano
End Property

Public Property Set EquipoHumano(ByRef Value As clsGenericCollection)
    Set mvarobjColEquipoHumano = Value
End Property

Public Property Get PersonalImplicado() As clsGenericCollection
    Set PersonalImplicado = mvarobjColPersonalImplicado
End Property

Public Property Set PersonalImplicado(ByRef Value As clsGenericCollection)
    Set mvarobjColPersonalImplicado = Value
End Property

Public Property Let setRESPONSABLE_NOMBRE_APELLIDOS(ByVal VALOR As String)
    RESPONSABLE_NOMBRE_APELLIDOS = VALOR
End Property
Public Property Let setORIGEN_OTROS(ByVal VALOR As String)
    ORIGEN_OTROS = VALOR
End Property

Public Property Get getRESPONSABLE_NOMBRE_APELLIDOS() As String
    getRESPONSABLE_NOMBRE_APELLIDOS = RESPONSABLE_NOMBRE_APELLIDOS
End Property
Public Property Let setRESPONSABLE_DEPARTAMENTO(ByVal VALOR As String)
    RESPONSABLE_DEPARTAMENTO = VALOR
End Property
Public Property Get getRESPONSABLE_DEPARTAMENTO() As String
    getRESPONSABLE_DEPARTAMENTO = RESPONSABLE_DEPARTAMENTO
End Property
Public Property Let setRESPONSABLE_ID_DEPARTAMENTO(ByVal VALOR As Long)
    RESPONSABLE_ID_DEPARTAMENTO = VALOR
End Property
Public Property Get getRESPONSABLE_ID_DEPARTAMENTO() As Long
    getRESPONSABLE_ID_DEPARTAMENTO = RESPONSABLE_ID_DEPARTAMENTO
End Property

Public Property Set AccionesCorrectoras(ByRef Value As clsGenericCollection)
    Set mvarobjColAccionesCorrectoras = Value
End Property

Public Property Get AccionesCorrectoras() As clsGenericCollection
    Set AccionesCorrectoras = mvarobjColAccionesCorrectoras
End Property
Public Property Set AccionesInmediatas(ByRef Value As clsGenericCollection)
    Set mvarobjColAccionesInmediatas = Value
End Property

Public Property Get AccionesInmediatas() As clsGenericCollection
    Set AccionesInmediatas = mvarobjColAccionesInmediatas
End Property


Public Property Set Trazabilidad(ByRef Value As clsGenericCollection)
    Set mvarobjColTrazabilidad = Value
End Property

Public Property Get Trazabilidad() As clsGenericCollection
    Set Trazabilidad = mvarobjColTrazabilidad
End Property

Public Property Set PreguntasRespuestas(ByRef Value As clsGenericCollection)
    Set mvarobjColPreguntasRespuestas = Value
End Property

Public Property Get PreguntasRespuestas() As clsGenericCollection
    Set PreguntasRespuestas = mvarobjColPreguntasRespuestas
End Property
Public Property Set OrigenesIncidencia(ByRef Value As clsGenericCollection)
    Set mvarobjColOrigenesIncidencia = Value
End Property

Public Property Get AdjuntosIncidencia() As clsGenericCollection
    Set AdjuntosIncidencia = mvarobjColAdjuntosIncidencia
End Property
Public Property Set AdjuntosIncidencia(ByRef Value As clsGenericCollection)
    Set mvarobjColAdjuntosIncidencia = Value
End Property

Public Property Get AdjuntosIndentificacionProblemas() As clsGenericCollection
    Set AdjuntosIndentificacionProblemas = mvarobjColAdjuntosIndentificacionProblemas
End Property
Public Property Set AdjuntosIndentificacionProblemas(ByRef Value As clsGenericCollection)
    Set mvarobjColAdjuntosIndentificacionProblemas = Value
End Property

Public Property Get AdjuntosRecoleccionDatos() As clsGenericCollection
    Set AdjuntosRecoleccionDatos = mvarobjColAdjuntosRecoleccionDatos
End Property
Public Property Set AdjuntosRecoleccionDatos(ByRef Value As clsGenericCollection)
    Set mvarobjColAdjuntosRecoleccionDatos = Value
End Property


Public Property Get OrigenesIncidencia() As clsGenericCollection
    Set OrigenesIncidencia = mvarobjColOrigenesIncidencia
End Property

Public Property Set DepartamentosImplicados(ByRef Value As clsGenericCollection)
    Set mvarobjColDepartamentosImplicados = Value
End Property

Public Property Get DepartamentosImplicados() As clsGenericCollection
    Set DepartamentosImplicados = mvarobjColDepartamentosImplicados
End Property

Public Property Get Causas() As clsGenericCollection
    Set Causas = mvarobjColCausas
End Property
Public Property Let setID_AUX(ByVal VALOR As Integer)
    ID_AUX = VALOR
End Property

Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSNC
'***************************************************************
Public Property Let setID_PROCNC(ByVal dato As Long)
        id_procnc = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
    RESPONSABLE_ID = dato
End Property
Public Property Let setFECHA_ALTA(ByVal dato As Date)
    FECHA_ALTA = dato
End Property
Public Property Let setDESCRIPCION_INCIDENCIA(ByVal dato As String)
    DESCRIPCION_INCIDENCIA = dato
End Property
Public Property Let setFECHA_CIERRE(ByVal dato As Date)
    fecha_cierre = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
Public Property Let setAUDITORIA_ID(ByVal dato As Long)
        AUDITORIA_ID = dato
End Property
Public Property Let setORIGEN_ID(ByVal dato As Long)
        ORIGEN_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSNC
'***************************************************************
Public Property Get getID_PROCNC() As Long
        getID_PROCNC = id_procnc
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getDESCRIPCION_INCIDENCIA() As String
        getDESCRIPCION_INCIDENCIA = DESCRIPCION_INCIDENCIA
End Property
Public Property Get getORIGEN_OTROS() As String
        getORIGEN_OTROS = ORIGEN_OTROS
End Property
Public Property Get getFECHA_ALTA() As Date
        getFECHA_ALTA = FECHA_ALTA
End Property
Public Property Get getFECHA_CIERRE() As Date
        getFECHA_CIERRE = fecha_cierre
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getORIGEN_ID() As Long
        getORIGEN_ID = ORIGEN_ID
End Property
Public Property Get getAUDITORIA_ID() As Long
        getAUDITORIA_ID = AUDITORIA_ID
End Property

Public Property Get getFECHA_ULT_MOVIMIENTO() As Date
    getFECHA_ULT_MOVIMIENTO = FECHA_ULT_MOVIMIENTO
End Property
Public Property Let setFECHA_ULT_MOVIMIENTO(ByVal VALOR As Date)
    FECHA_ULT_MOVIMIENTO = VALOR
End Property

Public Property Get getTOTAL_MOVIMIENTOS() As Integer
    getTOTAL_MOVIMIENTOS = mvarobjColTrazabilidad.Count
End Property



'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSNC
'***************************************************************
'
'
Public Function Carga(ID As Long) As Boolean
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
On Error GoTo Carga_Error

        consulta = "SELECT procnc.*, concat(usuarios.nombre, ' ', usuarios.apellidos) as RESPONSABLE_NOMBRE_APELLIDOS" & _
        ", deco_tipo.descripcion as TIPO_NO_CONFORMIDAD" & _
        ", decodificadora.descripcion as RESPONSABLE_DEPARTAMENTO" & _
        " From procnc " & _
        " join usuarios on procnc.responsable_id = usuarios.id_empleado " & _
        " join decodificadora on procnc.responsable_id_departamento = decodificadora.valor AND decodificadora.codigo=" & DECODIFICADORA.PROCNC_DEPARTAMENTOS & _
        " join decodificadora as deco_tipo on procnc.tipo_id = deco_tipo.valor AND deco_tipo.codigo=" & DECODIFICADORA.PROCNC_TIPOS_NO_CONFORMIDAD & _
        " WHERE procnc.ID_PROCNC = " & ID
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount <> 0 Then
            id_procnc = rs("ID_PROCNC")
            RESPONSABLE_ID = rs("RESPONSABLE_ID")
            RESPONSABLE_DEPARTAMENTO = rs("RESPONSABLE_DEPARTAMENTO")
            RESPONSABLE_ID_DEPARTAMENTO = rs("RESPONSABLE_ID_DEPARTAMENTO")
            RESPONSABLE_NOMBRE_APELLIDOS = rs("RESPONSABLE_NOMBRE_APELLIDOS")
            DESCRIPCION_INCIDENCIA = rs("DESCRIPCION_INCIDENCIA")
            RESUMEN = rs("RESUMEN")
            TIPO_ID = rs("TIPO_ID")
            tipo = rs("TIPO_NO_CONFORMIDAD")
            
            ORIGEN_OTROS = rs("ORIGEN_OTROS")
            If Not IsNull(rs("FECHA_ALTA")) Then
                FECHA_ALTA = rs("FECHA_ALTA")
            End If
            If Not IsNull(rs("FECHA_CIERRE")) Then
                fecha_cierre = rs("FECHA_CIERRE")
            End If
            If Not IsNull(rs("FECHA_ULT_MOVIMIENTO")) Then
                FECHA_ULT_MOVIMIENTO = Format(rs("FECHA_ULT_MOVIMIENTO"), "dd-mm-yyyy")
            End If
            ESTADO_ID = rs("ESTADO_ID")
            ORIGEN_ID = rs("ORIGEN_ID")
            AUDITORIA_ID = rs("AUDITORIA_ID")
            
            ANALISIS_ESCENA_CAMBIOS_RECIENTES = rs("ANALISIS_ESCENA_CAMBIOS_RECIENTES")
            ANALISIS_ESCENA_COMUNICACION = rs("ANALISIS_ESCENA_COMUNICACION")
            ANALISIS_ESCENA_CONDICIONES_AMBIENTALES = rs("ANALISIS_ESCENA_CONDICIONES_AMBIENTALES")
            ANALISIS_ESCENA_CONDICIONES_OPERACION = rs("ANALISIS_ESCENA_CONDICIONES_OPERACION")
            ANALISIS_ESCENA_EQUIPOSIMPLICADOS = rs("ANALISIS_ESCENA_EQUIPOSIMPLICADOS")
            ANALISIS_ESCENA_EXPERIENCIA = rs("ANALISIS_ESCENA_EXPERIENCIA")
            If Not IsNull(rs("ANALISIS_ESCENA_FECHA")) Then
                ANALISIS_ESCENA_FECHA = rs("ANALISIS_ESCENA_FECHA")
            End If
            ANALISIS_ESCENA_FORMACION = rs("ANALISIS_ESCENA_FORMACION")
            ANALISIS_ESCENA_HORA = rs("ANALISIS_ESCENA_HORA")
            ANALISIS_ESCENA_LOCALIZACION = rs("ANALISIS_ESCENA_LOCALIZACION")
            ANALISIS_ESCENA_MINUTOS = rs("ANALISIS_ESCENA_MINUTOS")
            ANALISIS_ESCENA_SECUENCIA = rs("ANALISIS_ESCENA_SECUENCIA")
            RECOLECCION_DATOS = rs("RECOLECCION_DATOS")
            es_no_conformidad = rs("ES_NO_CONFORMIDAD")
            es_solucion_aceptable = rs("ES_SOLUCION_ACEPTABLE")
            EVIDENCIAS_EN_PLAZO = rs("EVIDENCIAS_EN_PLAZO")
            EVIDENCIAS_EFECTIVAS = rs("EVIDENCIAS_EFECTIVAS")
            EVIDENCIAS_EVIDENCIAS = rs("EVIDENCIAS_EVIDENCIAS")
            EVIDENCIAS_COMUNICADO_MODIFICACIONES = rs("EVIDENCIAS_COMUNICADO_MODIFICACIONES")

            PROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO = rs("PROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO")
            PROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS = rs("PROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS")
            PROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS = rs("PROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS")
            PROBLEMAS_HUMANOS_SUFICIENTE_FORMACION = rs("PROBLEMAS_HUMANOS_SUFICIENTE_FORMACION")
            PROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE = rs("PROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE")
            PROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO = rs("PROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO")
            RESUMEN_CAUSAS = rs("RESUMEN_CAUSAS")
            causa_directa = rs("CAUSA_DIRECTA")
            causa_raiz = rs("CAUSA_RAIZ")
            CAUSA_CONTRIBUTIVA_1 = rs("CAUSA_CONTRIBUTIVA_1")
            CAUSA_CONTRIBUTIVA_1_DESCRIPCION = rs("CAUSA_CONTRIBUTIVA_1_DESCRIPCION")
            CAUSA_CONTRIBUTIVA_2 = rs("CAUSA_CONTRIBUTIVA_2")
            CAUSA_CONTRIBUTIVA_2_DESCRIPCION = rs("CAUSA_CONTRIBUTIVA_2_DESCRIPCION")
            CAUSA_CONTRIBUTIVA_3 = rs("CAUSA_CONTRIBUTIVA_3")
            CAUSA_CONTRIBUTIVA_3_DESCRIPCION = rs("CAUSA_CONTRIBUTIVA_3_DESCRIPCION")
            CAUSA_CONTRIBUTIVA_4 = rs("CAUSA_CONTRIBUTIVA_4")
            CAUSA_CONTRIBUTIVA_4_DESCRIPCION = rs("CAUSA_CONTRIBUTIVA_4_DESCRIPCION")
            CAUSA_CONTRIBUTIVA_5 = rs("CAUSA_CONTRIBUTIVA_5")
            CAUSA_CONTRIBUTIVA_5_DESCRIPCION = rs("CAUSA_CONTRIBUTIVA_5_DESCRIPCION")
            DESVIACION_ID = rs("DESVIACION_ID")
            OBSERVACIONES_RESULTADO = rs("OBSERVACIONES_RESULTADO")
            
            PROBLEMAS_CLASIFICACION_AFECTADOS = rs("PROBLEMAS_CLASIFICACION_AFECTADOS")
            PROBLEMAS_CLASIFICACION_ALCANCE = rs("PROBLEMAS_CLASIFICACION_ALCANCE")
            PROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA = rs("PROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA")
            'M1315-I
            N_ACCIONES = rs("N_ACCIONES")
            REVISADA_USUARIO_ID = rs("REVISADA_USUARIO_ID")
            If Not IsNull(rs("REVISADA_FECHA")) Then
                REVISADA_FECHA = rs("REVISADA_FECHA")
            End If
            'M1315-F
            Carga = True
        Else
            Carga = False
        End If
        
        Set rs = Nothing
        Exit Function

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Carga"
    Exit Function
Carga_Error:
    Carga = False
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Carga"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Carga of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function
Public Function CrearID() As Long
        Dim rs As ADODB.Recordset
        Dim consulta As String
        
        '' Jonathan. Temporal pruebas
        'ID_PROCNC = 1
        'Exit Function
        '' Fin Temporal Pruebas
        
On Error GoTo CrearID_Error

        consulta = "SELECT COALESCE(MAX(ID_PROCNC),0) as id_procnc FROM PROCNC"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                id_procnc = 1
        Else
                id_procnc = CLng(rs.Fields("ID_PROCNC")) + 1
        End If
        
        Set rs = Nothing
        
        CrearID = id_procnc

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.CrearID"
    Exit Function
CrearID_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.CrearID"
    error_grave Err.Number & " (" & Err.Description & ") in procedure CrearID of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
        
End Function
Public Function Insertar() As Long
On Error GoTo Insertar_Error

    Dim consulta As String
    Call CrearID
    
    ID_PARTICULAR = 0
    
    ' Ahora lo inserta directamente
    consulta = "INSERT INTO procnc (id_procnc, ID_PARTICULAR, TIPO_ID, estado_id,ORIGEN_ID,AUDITORIA_ID, PROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA, PROBLEMAS_CLASIFICACION_AFECTADOS, PROBLEMAS_CLASIFICACION_ALCANCE"
    consulta = consulta & ", ANALISIS_ESCENA_CAMBIOS_RECIENTES"
    consulta = consulta & ", ANALISIS_ESCENA_COMUNICACION"
    consulta = consulta & ", ANALISIS_ESCENA_CONDICIONES_AMBIENTALES"
    consulta = consulta & ", ANALISIS_ESCENA_CONDICIONES_OPERACION"
    consulta = consulta & ", ANALISIS_ESCENA_EQUIPOSIMPLICADOS"
    consulta = consulta & ", ANALISIS_ESCENA_EXPERIENCIA"
    consulta = consulta & ", ANALISIS_ESCENA_FORMACION"
    consulta = consulta & ", ANALISIS_ESCENA_FECHA"
    consulta = consulta & ", ANALISIS_ESCENA_HORA"
    consulta = consulta & ", ANALISIS_ESCENA_MINUTOS"
    consulta = consulta & ", ANALISIS_ESCENA_LOCALIZACION"
    consulta = consulta & ", ANALISIS_ESCENA_SECUENCIA"
    consulta = consulta & ", RECOLECCION_DATOS"
    consulta = consulta & ", ES_NO_CONFORMIDAD"
    consulta = consulta & ", ES_SOLUCION_ACEPTABLE"
    consulta = consulta & ", EVIDENCIAS_EN_PLAZO"
    consulta = consulta & ", EVIDENCIAS_EFECTIVAS"
    consulta = consulta & ", EVIDENCIAS_EVIDENCIAS"
    consulta = consulta & ", EVIDENCIAS_COMUNICADO_MODIFICACIONES"
    consulta = consulta & ", PROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO"
    consulta = consulta & ", PROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS"
    consulta = consulta & ", PROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS"
    consulta = consulta & ", PROBLEMAS_HUMANOS_SUFICIENTE_FORMACION"
    consulta = consulta & ", PROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE"
    consulta = consulta & ", PROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO"
    consulta = consulta & ", OBSERVACIONES_RESULTADO"
    consulta = consulta & ", ORIGEN_OTROS"
    consulta = consulta & ", RESUMEN_CAUSAS"
    consulta = consulta & ", CAUSA_DIRECTA"
    consulta = consulta & ", CAUSA_RAIZ"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_1"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_1_DESCRIPCION"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_2"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_2_DESCRIPCION"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_3"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_3_DESCRIPCION"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_4"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_4_DESCRIPCION"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_5"
    consulta = consulta & ", CAUSA_CONTRIBUTIVA_5_DESCRIPCION"
    consulta = consulta & ", DESVIACION_ID"
    consulta = consulta & ", N_ACCIONES"
    consulta = consulta & ") VALUES (" & CStr(id_procnc) & ", " & CStr(ID_PARTICULAR) & ", " & CStr(TIPO_ID) & ", " & C_PROCNC_ESTADOS.ABIERTA & "," & ORIGEN_ID & "," & AUDITORIA_ID & ", '', '', ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", '1900-01-01'"
    consulta = consulta & ", 0"
    consulta = consulta & ", 0"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", -1"
    consulta = consulta & ", -1"
    consulta = consulta & ", -1"
    consulta = consulta & ", -1"
    consulta = consulta & ", -1"
    consulta = consulta & ", -1"
    consulta = consulta & ", 0"
    consulta = consulta & ", 0"
    consulta = consulta & ", 0"
    consulta = consulta & ", 0"
    consulta = consulta & ", 0"
    consulta = consulta & ", 0"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", ''"
    consulta = consulta & ", 0" ' DESVIACION_ID
    consulta = consulta & "," & N_ACCIONES
    consulta = consulta & ") "
    
    
    
    execute_bd consulta
    

    
    Call Modificar(True)
    
On Error GoTo 0
        
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Insertar"
    Exit Function
Insertar_Error:
    Insertar = 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Insertar"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Insertar of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Public Function Modificar(Optional prmEsAltaRegistro As Boolean = False) As Boolean
On Error GoTo Modificar_Error


    Dim consulta As String, lngTipo As Long, rs As ADODB.Recordset, blnReportarCambioTipo As Boolean

        
        
    ' Toma el tipo de la consulta para comparar
    consulta = "select tipo_id from procnc where id_procnc = " & id_procnc
    Set rs = datos_bd(consulta)
    
    blnReportarCambioTipo = (rs!TIPO_ID <> TIPO_ID)
                
    consulta = "UPDATE PROCNC SET "
    consulta = consulta & " RESPONSABLE_ID = " & CStr(RESPONSABLE_ID) & ", "
    consulta = consulta & " ID_PARTICULAR = " & CStr(ID_PARTICULAR) & ", "
    consulta = consulta & " TIPO_ID = " & CStr(TIPO_ID) & ", "
    If CLng(FECHA_ALTA) > 0 Then _
        consulta = consulta & " FECHA_ALTA = '" & Format(FECHA_ALTA, "yyyy-mm-dd Hh:Nn:Ss") & "', "
    If CLng(fecha_cierre) > 0 Then _
        consulta = consulta & " FECHA_CIERRE = '" & Format(fecha_cierre, "yyyy-mm-dd Hh:Nn:Ss") & "', "
    If CLng(FECHA_ULT_MOVIMIENTO) > 0 Then _
        consulta = consulta & " FECHA_ULT_MOVIMIENTO = '" & Format(FECHA_ULT_MOVIMIENTO, "yyyy-mm-dd Hh:Nn:Ss") & "', "
    consulta = consulta & " RESPONSABLE_ID_DEPARTAMENTO = " & RESPONSABLE_ID_DEPARTAMENTO & ", "
    consulta = consulta & " DESCRIPCION_INCIDENCIA = '" & DESCRIPCION_INCIDENCIA & "', "
    consulta = consulta & "  RESUMEN = '" & Trim(RESUMEN) & "', "
    consulta = consulta & " ESTADO_ID = " & ESTADO_ID & ", "
    consulta = consulta & " AUDITORIA_ID = " & AUDITORIA_ID & ", "
    consulta = consulta & " ORIGEN_ID = " & ORIGEN_ID & ", "
    consulta = consulta & " N_ACCIONES = " & N_ACCIONES & " "
    consulta = consulta & " WHERE ID_PROCNC = " & id_procnc
    
    execute_bd consulta
    
    Modificar = True
            
    If blnReportarCambioTipo Then
        Call RegistrarMovimientoRegistro("CAMBIO_TIPO", CStr(rs!TIPO_ID) & ";" & CStr(TIPO_ID))
    End If
    Set rs = Nothing
    
    If prmEsAltaRegistro Then
        Call RegistrarMovimientoRegistro("INSERT")
    Else
        Call RegistrarMovimientoRegistro("UPDATE")
    End If
    
    Exit Function

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Modificar"
    Exit Function
Modificar_Error:
    Modificar = False
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Modificar"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Modificar of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function
Public Function Eliminar(ID As Long) As Boolean
        Dim consulta As String
        
        
On Error GoTo Eliminar_Error

        consulta = "DELETE FROM procnc_accionesinmediatas WHERE ID_PROCNC = " & ID
        execute_bd consulta
        
        consulta = "DELETE FROM procnc_accionescorrectivas WHERE ID_PROCNC = " & ID
        execute_bd consulta
        
        consulta = "DELETE FROM procnc_causas WHERE ID_PROCNC = " & ID
        execute_bd consulta
        
        consulta = "DELETE FROM procnc_departamentosimplicados WHERE ID_PROCNC = " & ID
        execute_bd consulta
        
        consulta = "DELETE FROM procnc_equipohumano WHERE ID_PROCNC = " & ID
        execute_bd consulta
        consulta = "DELETE FROM procnc_origenesincidencia WHERE ID_PROCNC = " & ID
        execute_bd consulta
        consulta = "DELETE FROM procnc_personalimplicado WHERE ID_PROCNC = " & ID
        execute_bd consulta
        consulta = "DELETE FROM procnc_preguntasrespuestas WHERE ID_PROCNC = " & ID
        execute_bd consulta
        consulta = "DELETE FROM procnc_trazabilidad WHERE ID_PROCNC = " & ID
        execute_bd consulta
        
        consulta = "DELETE FROM PROCNC WHERE ID_PROCNC = " & ID
        execute_bd consulta
        
        Eliminar = True
        
        Exit Function

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Eliminar"
    Exit Function
Eliminar_Error:
    Eliminar = False
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Eliminar"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Eliminar of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Function
Public Function Listado(strOrigen As String, strAuditoria As String, strTipo As String, strEstado As String, strDesc As String, fdesde As String, fhasta As String, strNumero As String, chkFechaCierre As Boolean, fcdesde As String, fchasta As String, DESVIACION_ID As String) As ADODB.Recordset
       
        
        Dim consulta As String, filtro As String, strJoins As String
        
On Error GoTo Listado_Error

        filtro = ""
        
        If Trim(strOrigen) <> "" And Trim(strOrigen) <> "0" Then
            filtro = filtro & " AND ORIGEN_ID = " & strOrigen
        End If
        If Trim(strAuditoria) <> "" And Trim(strAuditoria) <> "0" Then
            filtro = filtro & " AND AUDITORIA_ID = " & strAuditoria
        End If
        If Trim(DESVIACION_ID) <> "" And Trim(DESVIACION_ID) <> "0" Then
            filtro = filtro & " AND DESVIACION_ID = " & DESVIACION_ID
        End If
        If Trim(strTipo) <> "" Then
            filtro = filtro & " AND TIPO_ID = " & strTipo
        End If
        If Trim(strEstado) <> "" Then
            filtro = filtro & " AND ESTADO_ID = " & strEstado
        End If
        If Trim(strNumero) <> "" Then
            If IsNumeric(strNumero) Then
                filtro = filtro & " AND procnc.ID_PROCNC = " & strNumero
            End If
        End If
        filtro = filtro & " AND FECHA_ALTA >= '" & Format(fdesde, "yyyy-mm-dd hh:mm:ss") & "' and fecha_alta <= '" & Format(fhasta, "yyyy-mm-dd hh:mm:ss") & "'"
        If chkFechaCierre Then
            If fcdesde <> "" Then
                filtro = filtro & " AND FECHA_CIERRE >= '" & Format(fcdesde, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
            If fchasta <> "" Then
                filtro = filtro & " AND FECHA_CIERRE <= '" & Format(fchasta, "yyyy-mm-dd hh:mm:ss") & "'"
            End If
        End If
        strJoins = " LEFT JOIN clientes as clientes ON procnc.auditoria_id = clientes.ID_CLIENTE " & _
                   " LEFT JOIN proveedores as prov ON procnc.auditoria_id = prov.ID_PROVEEDOR " & _
                   " LEFT JOIN decodificadora AS dAuditoria2 ON procnc.auditoria_id = dAuditoria2.valor AND dAuditoria2.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_INTERNA & _
                   " LEFT JOIN decodificadora AS dAuditoria3 ON procnc.auditoria_id = dAuditoria3.valor AND dAuditoria3.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_EXTERNA & _
                   " LEFT JOIN decodificadora AS dAuditoria4 ON procnc.auditoria_id = dAuditoria4.valor AND dAuditoria4.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_DETECCION

        If Trim(strDesc) <> "" Then filtro = filtro & " AND lower(concat(DESCRIPCION_INCIDENCIA, RESUMEN)) LIKE '%" & ClrStr(strDesc) & "%'"
        
        consulta = "" & _
        "SELECT procnc.*" & _
        " ,coalesce(procnc.fecha_cierre,'') as fecha_cierre_vista" & _
        " ,usuarios.usuario as RESPONSABLE" & _
        " ,d1.descripcion AS RESPONSABLE_DEPARTAMENTO " & _
        " ,d2.descripcion as estado" & _
        " ,dOrigen.descripcion as origen" & _
        " ,dTipo.descripcion as tipo" & _
        " ,clientes.NOMBRE as auditoria1,dAuditoria2.DESCRIPCION as auditoria2, dAuditoria3.DESCRIPCION as auditoria3, dAuditoria4.DESCRIPCION  as auditoria4" & _
        " ,prov.NOMBRE as auditoria5 " & _
        " ,if (current_date > date_add(procnc.FECHA_CIERRE,interval 3 MONTH),1,0) as flimite " & _
        " FROM procnc " & _
        " join usuarios on procnc.responsable_id = usuarios.id_empleado" & _
        " join decodificadora as d1 on procnc.responsable_id_departamento = d1.valor and d1.codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & "" & _
        " join decodificadora as d2 on procnc.estado_id = d2.valor and d2.codigo =110" & _
        " join decodificadora as dOrigen on procnc.origen_id = dOrigen.valor and dOrigen.codigo =" & DECODIFICADORA.PROCNC_ORIGEN & _
        " join decodificadora as dTipo on procnc.tipo_id = dTipo.valor and dTipo.codigo =" & DECODIFICADORA.PROCNC_TIPOS_NO_CONFORMIDAD & _
        strJoins & _
        " WHERE 1=1 " & filtro & _
        " ORDER BY procnc.FECHA_ALTA DESC"

        Set Listado = datos_bd(consulta)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Listado"
    Exit Function
Listado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Listado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Listado of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
        
End Function
Public Function ListadoId(ID As Long) As ADODB.Recordset
       
        
        Dim consulta As String, filtro As String, strJoins As String
        
On Error GoTo Listado_Error

        filtro = " and procnc.id_procnc = " & ID
        
        strJoins = " LEFT JOIN clientes as clientes ON procnc.auditoria_id = clientes.ID_CLIENTE " & _
                   " LEFT JOIN proveedores as prov ON procnc.auditoria_id = prov.ID_PROVEEDOR " & _
                   " LEFT JOIN decodificadora AS dAuditoria2 ON procnc.auditoria_id = dAuditoria2.valor AND dAuditoria2.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_INTERNA & _
                   " LEFT JOIN decodificadora AS dAuditoria3 ON procnc.auditoria_id = dAuditoria3.valor AND dAuditoria3.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_EXTERNA & _
                   " LEFT JOIN decodificadora AS dAuditoria4 ON procnc.auditoria_id = dAuditoria4.valor AND dAuditoria4.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_DETECCION

        consulta = "" & _
        "SELECT procnc.*" & _
        " ,coalesce(procnc.fecha_cierre,'') as fecha_cierre_vista" & _
        " ,usuarios.usuario as RESPONSABLE" & _
        " ,d1.descripcion AS RESPONSABLE_DEPARTAMENTO " & _
        " ,d2.descripcion as estado" & _
        " ,dOrigen.descripcion as origen" & _
        " ,dTipo.descripcion as tipo" & _
        " ,clientes.NOMBRE as auditoria1,dAuditoria2.DESCRIPCION as auditoria2, dAuditoria3.DESCRIPCION as auditoria3, dAuditoria4.DESCRIPCION  as auditoria4" & _
        " ,prov.NOMBRE as auditoria5 " & _
        " ,if (current_date > date_add(procnc.FECHA_CIERRE,interval 3 MONTH),1,0) as flimite " & _
        " FROM procnc " & _
        " join usuarios on procnc.responsable_id = usuarios.id_empleado" & _
        " join decodificadora as d1 on procnc.responsable_id_departamento = d1.valor and d1.codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & "" & _
        " join decodificadora as d2 on procnc.estado_id = d2.valor and d2.codigo =110" & _
        " join decodificadora as dOrigen on procnc.origen_id = dOrigen.valor and dOrigen.codigo =" & DECODIFICADORA.PROCNC_ORIGEN & _
        " join decodificadora as dTipo on procnc.tipo_id = dTipo.valor and dTipo.codigo =" & DECODIFICADORA.PROCNC_TIPOS_NO_CONFORMIDAD & _
        strJoins & _
        " WHERE 1=1 " & filtro & _
        " ORDER BY procnc.FECHA_ALTA DESC"

        Set ListadoId = datos_bd(consulta)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Listado"
    Exit Function
Listado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Listado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Listado of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
        
End Function

'Public Function Listado_Combo() As ADODB.Recordset
'        Dim consulta As String
'On Error GoTo Listado_Combo_Error
'
'        consulta = "SELECT ID_PROCNC,NUMERO FROM PROCNC ORDER BY NUMERO"
'        Set Listado_Combo = datos_bd(consulta)
'
'On Error GoTo 0
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Listado_Combo"
'    Exit Function
'Listado_Combo_Error:
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Listado_Combo"
'    error_grave Err.Number & " (" & Err.Description & ") in procedure Listado_Combo of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
'    G_TRAZABILIDAD_ERROR = ""
'End Function

Private Sub Class_Initialize()
On Error GoTo Class_Initialize_Error

mvarstrRutaInicial = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta") & "\ADJUNTOS\PROCNC\"

id_procnc = -1
ID_AUX = -2
mvarobjColAccionesCorrectoras.KeyName = "setID_ACCION"
mvarobjColPreguntasRespuestas.KeyName = "setID_PREGUNTA_RESPUESTA"
mvarobjColAdjuntosIncidencia.KeyName = "setID_ADJUNTO"
mvarobjColAdjuntosIndentificacionProblemas.KeyName = "setID_ADJUNTO"
mvarobjColAdjuntosRecoleccionDatos.KeyName = "setID_ADJUNTO"
mvarobjColDepartamentosImplicados.KeyName = "setID"
mvarobjColOrigenesIncidencia.KeyName = "setID"
mvarobjColTrazabilidad.KeyName = "setID_TRAZABILIDAD"
mvarobjColEquipoHumano.KeyName = "setID"
mvarobjColPersonalImplicado.KeyName = "setID"
mvarobjColCausas.KeyName = "setID"

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Class_Initialize"
    Exit Sub
Class_Initialize_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.Class_Initialize"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Public Function cargar_AccionesInmediatas() As String
    Dim oPAI As New clsProcnc_accionesinmediatas
    Dim rs As ADODB.Recordset
    Dim salida As String
    Dim intCont As Integer
    intCont = 1
    Set rs = oPAI.Listado(id_procnc)
    If rs.RecordCount = 0 Then
        salida = "Acc.Inmediata : El Responsable No se ha determinado ninguna Accion inmediata aún." & vbNewLine & vbNewLine
    Else
        Do
            salida = salida & "Acc.Inmediata " & CStr(intCont) & " : " & rs("DESCRIPCION") & vbNewLine & vbNewLine
            intCont = intCont + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    cargar_AccionesInmediatas = salida
End Function

Private Sub cargar_DepartamentosImplicados()
Dim objItem As clsGenericClass
Dim strSql As String
Dim rs As New ADODB.Recordset

On Error GoTo cargar_DepartamentosImplicados_Error

    strSql = "SELECT procnc_departamentosimplicados.*, d1.descripcion"
    strSql = strSql & " FROM procnc_departamentosimplicados "
    strSql = strSql & " join decodificadora as d1 on procnc_departamentosimplicados.id_departamento = d1.valor and d1.codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & " "
    strSql = strSql & " WHERE procnc_departamentosimplicados.id_procnc = " & CStr(id_procnc)
    
    Set rs = datos_bd(strSql)
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            Set objItem = New clsGenericClass
            With objItem
                .setID = rs("id_departamento")
                .setDESCRIPCION = rs("descripcion")
                .setID_AUX = 0
            End With
            rs.MoveNext
            
            Call mvarobjColDepartamentosImplicados.Add(objItem, objItem.getID)
        Wend
    End If
    
    Set rs = Nothing


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.cargar_DepartamentosImplicados"
    Exit Sub
cargar_DepartamentosImplicados_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.cargar_DepartamentosImplicados"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_DepartamentosImplicados of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Public Sub RegistrarMovimientoRegistro(ByVal strTipoMovimiento As String, Optional ByVal prmstrMotivoCambioEstado As String = "")
Dim objwf As New clsProcNcTrazabilidad

On Error GoTo RegistrarMovimientoRegistro_Error

    Dim arrTipo(0 To 4) As String
    
    arrTipo(0) = "N/A"
    arrTipo(1) = "PROPUESTA DE MEJORA"
    arrTipo(2) = "OBSERVACION"
    arrTipo(3) = "INCIDENCIA"
    arrTipo(4) = "NO CONFORMIDAD"

    objwf.setID_PROCNC = id_procnc
    objwf.setFECHAHORA = Now
    objwf.setID_USUARIO = USUARIO.getID_EMPLEADO
    objwf.setUSUARIO = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
    
    Select Case strTipoMovimiento
        Case "INSERT"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_ALTA_REGISTRO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha dado de alta de Registro de Incidencia. El estado de la incidencia es : " & getDenominacionEstado()
        Case "UPDATE"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_MODIFICACION_REGISTRO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha modificado el Registro de Incidencia"
        Case "STATE"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_CAMBIO_ESTADO_REGISTRO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha modificado el estado de la indicencia. El nuevo estado es : " & getDenominacionEstado()
            If Trim(prmstrMotivoCambioEstado) <> "" Then
                objwf.setDESCRIPCION = objwf.getDESCRIPCION & vbCrLf & "El Motivo expuesto para dicho cambio de estado es el siguente: " & vbCrLf & prmstrMotivoCambioEstado
            End If
        Case "HISTORY"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_ENVIAR_REGISTRO_HISTORICO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha archivado en el Historico la indicencia."
        Case "QUERY"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_MODIFICACION_REGISTRO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha consultado la indicencia."
        Case "AC_CORR_STATE_TRAMIT"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_CAMBIO_ESTADO_REGISTRO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha dado enviado a Tramitación la Acción Correctiva: " & prmstrMotivoCambioEstado & ". El estado de la Accion Correctiva es : En Tramtación"
        Case "AC_CORR_STATE_CLOSE_SUCCESS"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_CAMBIO_ESTADO_REGISTRO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha Cerrado con Éxito la Acción Correctiva: " & prmstrMotivoCambioEstado & ". El estado de la Accion Correctiva es : Cerrada"
        Case "AC_CORR_STATE_CLOSE_NEW_PNC"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_CAMBIO_ESTADO_REGISTRO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha Cerrado Sin Éxito (Se ha generado un nueva Incidencia a partir de esta) la Acción Correctiva: " & prmstrMotivoCambioEstado & ". El estado de la Accion Correctiva es : Cerrada (Genera Nueva Incidencia)"
        Case "CAMBIO_TIPO"
            objwf.setID_ACCION = enumTipoAccionTrazabilidad.WF_CAMBIO_TIPO
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha cambiado el Tipo de la No Conformidad de " & arrTipo(CInt(Split(prmstrMotivoCambioEstado, ";")(0))) & " a " & arrTipo(CInt(Split(prmstrMotivoCambioEstado, ";")(1))) & "."
        Case Else
            objwf.setID_ACCION = -1
            objwf.setDESCRIPCION = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha realizado cambios sin especificar en la Incidencia."
    End Select
    
    
    objwf.Insertar

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.RegistrarMovimientoRegistro"
    Exit Sub
RegistrarMovimientoRegistro_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.RegistrarMovimientoRegistro"
    error_grave Err.Number & " (" & Err.Description & ") in procedure RegistrarMovimientoRegistro of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Function getDenominacionEstado() As String

On Error GoTo getDenominacionEstado_Error

Select Case ESTADO_ID
    Case C_PROCNC_ESTADOS.ABIERTA
        getDenominacionEstado = "Incidencia Abierta"
    Case C_PROCNC_ESTADOS.CERRADA
        getDenominacionEstado = "Incidencia Cerrada"
    Case C_PROCNC_ESTADOS.CERRADA_PARCIAL_EVAL
        getDenominacionEstado = "Incidencia Parcialmente Cerrada (pendiente de finalización de Acciones Correctivas)"
    Case C_PROCNC_ESTADOS.EN_TRAMITACION
        getDenominacionEstado = "Tramitación de la Incidencia"
    Case C_PROCNC_ESTADOS.pte_cierre
        getDenominacionEstado = "Pendiente de Cierre de Incidencia"
    Case C_PROCNC_ESTADOS.PTE_PLAN_ACCIONES_CORRECTIVAS
        getDenominacionEstado = "Pendiente del Plan de Acciones Correctivas"
    Case C_PROCNC_ESTADOS.PTE_VISTO_BUENO
        getDenominacionEstado = "Pendiente de Validación (Vº Bº) por parte de Responsable de Calidad"
    Case Else
        getDenominacionEstado = "Estado No Especificado"
End Select

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.getDenominacionEstado"
    Exit Function
getDenominacionEstado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.getDenominacionEstado"
    error_grave Err.Number & " (" & Err.Description & ") in procedure getDenominacionEstado of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Function

Private Sub EnviarMensaje_PtePlanAccionesCorrectivas()
Dim objMensaje As New clsMensajes, objDpto As clsGenericClass
Dim strSql As String, strCad As String, strArr() As String
Dim idRespCalidad() As Long, indexResp As Long, intCont As Integer
Dim rs As ADODB.Recordset
Dim ASUNTO As String
Dim texto As String
Dim mens As Integer, x As Integer
Dim blnRespYaEnLista As Boolean

On Error GoTo EnviarMensaje_PtePlanAccionesCorrectivas_Error

    strSql = "SELECT * From procnc_equipohumano where id_procnc = " & id_procnc & " and jefe_equipo = 1"
    Set rs = datos_bd(strSql)
    
    'Avisa al responsable del equipo
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        indexResp = 1
        ReDim Preserve idRespCalidad(indexResp)
        idRespCalidad(indexResp - 1) = CLng(rs("id_usuario"))
    End If
    
    If indexResp = 0 Then
        ' cuando no hay responsables de calidad, se lo envia a Gerencia
        ReDim Preserve idRespCalidad(indexResp + 1)
        idRespCalidad(indexResp) = 1 ' El ID 1 es el Gerente.
        indexResp = indexResp + 1
    End If
    
    ASUNTO = "Solicitud Plan Acciones Correctivas Incidencia/No Conformidad/Propuesta Mejora, Nº: " & Format(id_procnc, "00000")
    texto = texto & "El usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha supervisado la Tramitación de la Incidencia por parte de la persona designada como Jefe de Equipo Humano. Se ha aprobado la tramitación de la misma y se procede a solicitar al Jefe de Equipo la confección del Plan de Acciones Correctivas Correspondientes." & vbNewLine & vbNewLine
    texto = texto & "Dpto. Responsable: " & RESPONSABLE_DEPARTAMENTO & ". " & vbNewLine & vbNewLine
    texto = texto & "Fecha de Alta : " & Format(FECHA_ALTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Fecha de Ult. Modificacion: " & Format(FECHA_ULT_MOVIMIENTO, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Sinopsis: " & RESUMEN & vbNewLine & vbNewLine
    texto = texto & "Descripción : " & DESCRIPCION_INCIDENCIA & vbNewLine & vbNewLine
    
    texto = texto & cargar_AccionesInmediatas
    
    cargar_DepartamentosImplicados
    If mvarobjColDepartamentosImplicados.Count > 0 Then
    
        strCad = ""
        For Each objDpto In mvarobjColDepartamentosImplicados.Iterator
             strCad = strCad & ", " & objDpto.getDESCRIPCION
        Next objDpto
        texto = texto & "Dpto/s Implicado/s :" & Right(strCad, Len(strCad) - 2) & vbNewLine & vbNewLine
    End If
    
    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion;" & CStr(id_procnc), Now)


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_PtePlanAccionesCorrectivas"
    Exit Sub
EnviarMensaje_PtePlanAccionesCorrectivas_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_PtePlanAccionesCorrectivas"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_PtePlanAccionesCorrectivas of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub EnviarMensaje_RespCalidadVoBo()
Dim objMensaje As New clsMensajes, objDpto As clsGenericClass
Dim strSql As String, strCad As String, strArr() As String
Dim idRespCalidad() As Long, indexResp As Long, intCont As Integer
Dim rs As ADODB.Recordset
Dim ASUNTO As String
Dim texto As String
Dim mens As Integer, x As Integer
Dim blnRespYaEnLista As Boolean

On Error GoTo EnviarMensaje_RespCalidadVoBo_Error

    Call devolverListaResponsablesCalidad(idRespCalidad)
    indexResp = UBound(idRespCalidad)
    
    If indexResp = 0 Then
        ' cuando no hay responsables de calidad, se lo envia a Gerencia
        ReDim Preserve idRespCalidad(indexResp + 1)
        idRespCalidad(indexResp) = 1 ' El ID 1 es el Gerente.
        indexResp = indexResp + 1
    End If
    
    ASUNTO = "VºBº Pendiente Incidencia/No Conformidad/Propuesta Mejora, Nº: " & Format(id_procnc, "00000")
    texto = texto & "El usuario " & RESPONSABLE_NOMBRE_APELLIDOS & " ha dado de alta una Incidencia/No conformidad. " & vbNewLine & vbNewLine
    texto = texto & "Dpto. Responsable: " & RESPONSABLE_DEPARTAMENTO & ". " & vbNewLine & vbNewLine
    texto = texto & "Fecha de Alta : " & Format(FECHA_ALTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Fecha de Ult. Modificacion: " & Format(FECHA_ULT_MOVIMIENTO, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Sinopsis: " & RESUMEN & vbNewLine & vbNewLine
    texto = texto & "Descripción : " & DESCRIPCION_INCIDENCIA & vbNewLine & vbNewLine
    
    ' Carga en la coleccion las acciones inmediatas
    texto = texto & cargar_AccionesInmediatas
    
    'strCad = ""
    'For Each objDpto In mvarobjColDepartamentosImplicados.Iterator
    '     strCad = strCad & ", " & objDpto.getDESCRIPCION
    'Next objDpto
    'texto = texto & "Dpto/s Implicado/s :" & Right(strCad, Len(strCad) - 2) & vbNewLine & vbNewLine
    
    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion;" & CStr(id_procnc), Now)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_RespCalidadVoBo"
    Exit Sub
EnviarMensaje_RespCalidadVoBo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_RespCalidadVoBo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_RespCalidadVoBo of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Private Sub EnviarMensaje_VolverATramitacionPorRechazo(ByVal prmstrMotivo As String)
Dim objMensaje As New clsMensajes, objDpto As clsGenericClass
Dim strSql As String, strCad As String, strArr() As String
Dim idRespCalidad() As Long, indexResp As Long, intCont As Integer
Dim rs As ADODB.Recordset
Dim ASUNTO As String
Dim texto As String
Dim mens As Integer, x As Integer
Dim blnRespYaEnLista As Boolean

    
On Error GoTo EnviarMensaje_VolverATramitacionPorRechazo_Error

    strSql = "SELECT * From procnc_equipohumano where id_procnc = " & id_procnc & " and jefe_equipo = 1"
    Set rs = datos_bd(strSql)
    
    'Avisa al responsable del equipo
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        indexResp = indexResp + 1
        ReDim Preserve idRespCalidad(indexResp)
        idRespCalidad(indexResp - 1) = CLng(rs("id_usuario"))
    End If
    
    
    ASUNTO = "Rechazada Confirmación Calidad Incidencia/No Conformidad/Propuesta Mejora, Nº: " & Format(id_procnc, "00000")
    texto = texto & "El usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & ", como Responsable de Calidad, ha rechazado la Confirmación de la Incidencia/No conformidad. " & vbNewLine & vbNewLine
    texto = texto & "Situación: La incidencia/No Confirmidad vuelve a estado En Tramitación. " & vbNewLine & vbNewLine
    texto = texto & "Motivo Rechazo: " & prmstrMotivo & ". " & vbNewLine & vbNewLine
    texto = texto & "Fecha de Alta : " & Format(FECHA_ALTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Fecha de Ult. Modificacion: " & Format(FECHA_ULT_MOVIMIENTO, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Sinopsis: " & RESUMEN & vbNewLine & vbNewLine
    texto = texto & "Descripción : " & DESCRIPCION_INCIDENCIA & vbNewLine & vbNewLine
    
    texto = texto & cargar_AccionesInmediatas
    
    Call cargar_DepartamentosImplicados
    
    If mvarobjColDepartamentosImplicados.Count > 0 Then
        strCad = ""
        For Each objDpto In mvarobjColDepartamentosImplicados.Iterator
             strCad = strCad & ", " & objDpto.getDESCRIPCION
        Next objDpto
        texto = texto & "Dpto/s Implicado/s :" & Right(strCad, Len(strCad) - 2) & vbNewLine & vbNewLine
    End If
    
    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion;" & CStr(id_procnc), Now)


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_VolverATramitacionPorRechazo"
    Exit Sub
EnviarMensaje_VolverATramitacionPorRechazo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_VolverATramitacionPorRechazo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_VolverATramitacionPorRechazo of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub
Private Sub EnviarMensaje_CerradaTotal()
Dim objDpto As clsGenericClass
Dim strSql As String, strCad As String, strArr() As String
Dim idRespCalidad() As Long, indexResp As Long, intCont As Integer
Dim arrEquipo() As Long, indexEquipo As Integer
Dim rs As ADODB.Recordset
Dim ASUNTO As String
Dim texto As String
Dim mens As Integer
Dim blnYaExistente As Boolean
On Error GoTo EnviarMensaje_CerradaTotal_Error

blnYaExistente = False
Dim x As Integer, y As Integer, z As Integer
Dim arrFinal() As Long, indexFinal As Integer
    
    Call devolverListaResponsablesDepartamentos(idRespCalidad)
    indexResp = UBound(idRespCalidad)
'    Call devolverListaUsuariosEquipoHumano(arrEquipo)
'    indexEquipo = UBound(arrEquipo)
    Call devolverListaUsuariosEquipoHumano(arrEquipo, indexEquipo)
    
    ' Avisa al responsable de la apertura
    For x = 0 To indexResp - 1
        If idRespCalidad(x) = RESPONSABLE_ID Then
            blnYaExistente = True
        End If
    Next x
    If Not blnYaExistente Then
        indexResp = indexResp + 1
        ReDim Preserve idRespCalidad(indexResp)
        idRespCalidad(indexResp - 1) = RESPONSABLE_ID
    End If
    
    indexFinal = 0
    
    ' incluye a los componentes del equipo humano
    For y = 0 To indexEquipo - 1
        blnYaExistente = False
        For x = 0 To indexResp - 1
            If idRespCalidad(x) = arrEquipo(y) Then
                blnYaExistente = True
            End If
        Next x
        If Not blnYaExistente Then
            indexFinal = indexFinal + 1
            ReDim Preserve arrFinal(indexFinal)
            arrFinal(indexFinal - 1) = arrEquipo(y)
        End If
    Next y
    
    For x = 0 To indexResp - 1
        indexFinal = indexFinal + 1
        ReDim Preserve arrFinal(indexFinal)
        arrFinal(indexFinal - 1) = idRespCalidad(x)
    Next x
    
    
    ASUNTO = "Incidencia/No Conformidad/Propuesta Mejora CERRADA Nº: " & Format(id_procnc, "00000")
    texto = texto & "El Responsable del Dpto. Calidad " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & " ha dado la Incidencia/No Confirmidad por CERRADA y FINALIZADA TOTALMENTE." & vbNewLine & vbNewLine
    texto = texto & "Dpto. Responsable: " & RESPONSABLE_DEPARTAMENTO & ". " & vbNewLine & vbNewLine
    texto = texto & "Fecha de Alta : " & Format(FECHA_ALTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Fecha de Ult. Modificacion: " & Format(FECHA_ULT_MOVIMIENTO, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Sinopsis: " & RESUMEN & vbNewLine & vbNewLine
    texto = texto & "Descripción : " & DESCRIPCION_INCIDENCIA & vbNewLine & vbNewLine
    
    texto = texto & cargar_AccionesInmediatas
    
    cargar_DepartamentosImplicados
    If mvarobjColDepartamentosImplicados.Count > 0 Then
    
        strCad = ""
        For Each objDpto In mvarobjColDepartamentosImplicados.Iterator
             strCad = strCad & ", " & objDpto.getDESCRIPCION
        Next objDpto
        texto = texto & "Dpto/s Implicado/s :" & Right(strCad, Len(strCad) - 2) & vbNewLine & vbNewLine
    End If
    Call EnviarMensajeInterno(arrFinal, ASUNTO, texto, "frmProcNCEdicion;" & CStr(id_procnc), Now)


On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_VoBo_Ok"
    Exit Sub
EnviarMensaje_CerradaTotal_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.EnviarMensaje_CerradaTotal"
    error_grave Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_CerradaTotal of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Public Property Get getANALISIS_ESCENA_LOCALIZACION() As String
    getANALISIS_ESCENA_LOCALIZACION = ANALISIS_ESCENA_LOCALIZACION
End Property
Public Property Let setANALISIS_ESCENA_LOCALIZACION(ByRef Value As String)
        ANALISIS_ESCENA_LOCALIZACION = Value

End Property
Public Property Get getANALISIS_ESCENA_FECHA() As Date
getANALISIS_ESCENA_FECHA = ANALISIS_ESCENA_FECHA

End Property
Public Property Let setANALISIS_ESCENA_FECHA(ByRef Value As Date)
            ANALISIS_ESCENA_FECHA = Value

End Property
Public Property Get getANALISIS_ESCENA_HORA() As Integer
getANALISIS_ESCENA_HORA = ANALISIS_ESCENA_HORA

End Property
Public Property Let setANALISIS_ESCENA_HORA(ByRef Value As Integer)
        ANALISIS_ESCENA_HORA = Value

End Property
Public Property Get getANALISIS_ESCENA_MINUTOS() As Integer
getANALISIS_ESCENA_MINUTOS = ANALISIS_ESCENA_MINUTOS

End Property
Public Property Let setANALISIS_ESCENA_MINUTOS(ByRef Value As Integer)
        ANALISIS_ESCENA_MINUTOS = Value

End Property
Public Property Get getANALISIS_ESCENA_CAMBIOS_RECIENTES() As String

getANALISIS_ESCENA_CAMBIOS_RECIENTES = ANALISIS_ESCENA_CAMBIOS_RECIENTES
End Property
Public Property Let setANALISIS_ESCENA_CAMBIOS_RECIENTES(ByRef Value As String)
        ANALISIS_ESCENA_CAMBIOS_RECIENTES = Value

End Property
Public Property Get getANALISIS_ESCENA_CONDICIONES_AMBIENTALES() As String
getANALISIS_ESCENA_CONDICIONES_AMBIENTALES = ANALISIS_ESCENA_CONDICIONES_AMBIENTALES

End Property
Public Property Let setANALISIS_ESCENA_CONDICIONES_AMBIENTALES(ByRef Value As String)
        ANALISIS_ESCENA_CONDICIONES_AMBIENTALES = Value

End Property

Public Property Get getANALISIS_ESCENA_CONDICIONES_OPERACION() As String
getANALISIS_ESCENA_CONDICIONES_OPERACION = ANALISIS_ESCENA_CONDICIONES_OPERACION

End Property
Public Property Let setANALISIS_ESCENA_CONDICIONES_OPERACION(ByRef Value As String)
        ANALISIS_ESCENA_CONDICIONES_OPERACION = Value

End Property

Public Property Get getPROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA() As String
    getPROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA = PROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA
End Property
Public Property Get getPROBLEMAS_CLASIFICACION_AFECTADOS() As String
    getPROBLEMAS_CLASIFICACION_AFECTADOS = PROBLEMAS_CLASIFICACION_AFECTADOS
End Property
Public Property Get getPROBLEMAS_CLASIFICACION_ALCANCE() As String
    getPROBLEMAS_CLASIFICACION_ALCANCE = PROBLEMAS_CLASIFICACION_ALCANCE
End Property

Public Property Let setPROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA(ByRef Value As String)
       PROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA = Value
End Property
Public Property Let setPROBLEMAS_CLASIFICACION_AFECTADOS(ByRef Value As String)
       PROBLEMAS_CLASIFICACION_AFECTADOS = Value
End Property
Public Property Let setPROBLEMAS_CLASIFICACION_ALCANCE(ByRef Value As String)
       PROBLEMAS_CLASIFICACION_ALCANCE = Value
End Property

Public Property Get getANALISIS_ESCENA_EQUIPOSIMPLICADOS() As String
getANALISIS_ESCENA_EQUIPOSIMPLICADOS = ANALISIS_ESCENA_EQUIPOSIMPLICADOS

End Property
Public Property Let setANALISIS_ESCENA_EQUIPOSIMPLICADOS(ByRef Value As String)
        ANALISIS_ESCENA_EQUIPOSIMPLICADOS = Value

End Property
Public Property Get getANALISIS_ESCENA_EXPERIENCIA() As String
getANALISIS_ESCENA_EXPERIENCIA = ANALISIS_ESCENA_EXPERIENCIA

End Property
Public Property Let setANALISIS_ESCENA_EXPERIENCIA(ByRef Value As String)
        ANALISIS_ESCENA_EXPERIENCIA = Value

End Property
Public Property Get getANALISIS_ESCENA_FORMACION() As String
getANALISIS_ESCENA_FORMACION = ANALISIS_ESCENA_FORMACION

End Property
Public Property Let setANALISIS_ESCENA_FORMACION(ByRef Value As String)
        ANALISIS_ESCENA_FORMACION = Value

End Property
Public Property Get getANALISIS_ESCENA_SECUENCIA() As String
getANALISIS_ESCENA_SECUENCIA = ANALISIS_ESCENA_SECUENCIA

End Property
Public Property Let setANALISIS_ESCENA_SECUENCIA(ByRef Value As String)
        ANALISIS_ESCENA_SECUENCIA = Value

End Property
Public Property Get getANALISIS_ESCENA_COMUNICACION() As String
    getANALISIS_ESCENA_COMUNICACION = ANALISIS_ESCENA_COMUNICACION

End Property
Public Property Let setANALISIS_ESCENA_COMUNICACION(ByVal Value As String)
    ANALISIS_ESCENA_COMUNICACION = Value
End Property


Public Property Get getRECOLECCION_DATOS() As String
    getRECOLECCION_DATOS = RECOLECCION_DATOS
End Property
Public Property Let setRECOLECCION_DATOS(ByVal Value As String)
    RECOLECCION_DATOS = Value
End Property
Private Sub devolverListaUsuariosEquipoHumano(ByRef arr() As Long, ByRef idx As Integer)
Dim strSql As String, rs As ADODB.Recordset
'Dim idx As Integer

On Error GoTo devolverListaUsuariosEquipoHumano_Error

    strSql = "SELECT * From procnc_equipohumano where id_procnc = " & id_procnc & " and jefe_equipo = 1"
    Set rs = datos_bd(strSql)
    
    idx = 0
    'Avisa al responsable del equipo
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            idx = idx + 1
            ReDim Preserve arr(idx)
            arr(idx - 1) = CLng(rs("id_usuario"))
            rs.MoveNext
        Wend
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolverListaUsuariosEquipoHumano"
    Exit Sub
devolverListaUsuariosEquipoHumano_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.devolverListaUsuariosEquipoHumano"
    error_grave Err.Number & " (" & Err.Description & ") in procedure devolverListaUsuariosEquipoHumano of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub EnviarMensaje_PteCierre()

On Error GoTo EnviarMensaje_PteCierre_Error
    
Dim objMensaje As New clsMensajes, objDpto As clsGenericClass
Dim strSql As String, strCad As String, strArr() As String
Dim idRespCalidad() As Long, indexResp As Long, intCont As Integer
Dim rs As ADODB.Recordset
Dim ASUNTO As String
Dim texto As String
Dim mens As Integer, x As Integer
Dim blnRespYaEnLista As Boolean


    Call devolverListaResponsablesCalidad(idRespCalidad)
    indexResp = UBound(idRespCalidad)
    
    If indexResp = 0 Then
        ' cuando no hay responsables de calidad, se lo envia a Gerencia
        ReDim Preserve idRespCalidad(indexResp + 1)
        idRespCalidad(indexResp) = 1 ' El ID 1 es el Gerente.
        indexResp = indexResp + 1
    End If
    
    ASUNTO = "Plan Acciones Creado: Pte. Desarrollo Plan/Cierre Pendiente Incidencia/No Conformidad/Propuesta Mejora, Nº: " & Format(id_procnc, "00000")
    texto = texto & "El usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & ", como Responsable de Equipo de gestión, ha creado un Plan de Acciones Correctivas y solicita el Desarrollo/Supervisión/Cierre de la incidencia por parte del departamento de Calidad. " & vbNewLine & vbNewLine
    texto = texto & "Dpto. Responsable: " & RESPONSABLE_DEPARTAMENTO & ". " & vbNewLine & vbNewLine
    texto = texto & "Fecha de Alta : " & Format(FECHA_ALTA, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Fecha de Ult. Modificacion: " & Format(FECHA_ULT_MOVIMIENTO, "dd-mm-yyyy") & vbNewLine & vbNewLine
    texto = texto & "Sinopsis: " & RESUMEN & vbNewLine & vbNewLine
    texto = texto & "Descripción : " & DESCRIPCION_INCIDENCIA & vbNewLine & vbNewLine
    
    texto = texto & cargar_AccionesInmediatas
    
    cargar_DepartamentosImplicados
    If mvarobjColDepartamentosImplicados.Count > 0 Then
    
        strCad = ""
        For Each objDpto In mvarobjColDepartamentosImplicados.Iterator
             strCad = strCad & ", " & objDpto.getDESCRIPCION
        Next objDpto
        texto = texto & "Dpto/s Implicado/s :" & Right(strCad, Len(strCad) - 2) & vbNewLine & vbNewLine
    End If
    
    Call EnviarMensajeInterno(idRespCalidad, ASUNTO, texto, "frmProcNCEdicion;" & CStr(id_procnc), Now)
On Error GoTo 0
    Exit Sub
EnviarMensaje_PteCierre_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EnviarMensaje_PteCierre of Módulo de clase clsProcNc"
End Sub

Public Function Es_Jefe_Equipo(ByVal ID_EMPLEADO As Long) As Boolean

    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = "SELECT * From procnc_equipohumano where id_procnc = " & id_procnc & " and jefe_equipo = 1"
    Set rs = datos_bd(strSql)
    
    
    Es_Jefe_Equipo = False
    
    'Avisa al responsable del equipo
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            If CLng(rs("id_usuario")) = ID_EMPLEADO Then
                Es_Jefe_Equipo = True
                Set rs = Nothing
                Exit Function
            End If
            rs.MoveNext
        Wend
    End If

End Function

Public Function Tiene_Asignada_Correctiva(ByVal ID_EMPLEADO As Long) As Boolean
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = "SELECT * From procnc_accionescorrectivas where id_procnc = " & id_procnc & " and responsable_id = " & ID_EMPLEADO
    Set rs = datos_bd(strSql)
    
    
    Tiene_Asignada_Correctiva = False
    
    'Avisa al responsable del equipo
    If rs.RecordCount <> 0 Then
        Tiene_Asignada_Correctiva = True
    End If

    Set rs = Nothing
        
End Function
Public Function marcarRevisada(ID As Long) As Boolean
    Dim strSql As String

   On Error GoTo marcarRevisada_Error

    strSql = "UPDATE procnc SET REVISADA_USUARIO_ID=" & USUARIO.getID_EMPLEADO & ",REVISADA_FECHA=CURRENT_TIMESTAMP WHERE ID_PROCNC = " & ID
    execute_bd strSql
    

   On Error GoTo 0
   Exit Function

marcarRevisada_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure marcarRevisada of Módulo de clase clsProcNc"

End Function

Public Function habilitarRevision(ID As Long) As Integer
    Dim strSql As String
    Dim rs As ADODB.Recordset
    
    strSql = "SELECT ID_PROCNC,if (current_date > date_add(FECHA_CIERRE,interval 3 MONTH),1,0) as F_LIMITE FROM procnc WHERE 1=1 and ID_PROCNC=" & ID
    Set rs = datos_bd(strSql)
    If rs.RecordCount <> 0 Then
        habilitarRevision = rs(1)
    Else
        habilitarRevision = 0
    End If

    Set rs = Nothing
        
End Function
Public Function ListadoRevision() As ADODB.Recordset
        Dim consulta As String, filtro As String, strJoins As String
        filtro = filtro & " and RESPONSABLE_ID = " & USUARIO.getID_EMPLEADO
        filtro = filtro & " and CURRENT_DATE > DATE_ADD(procnc.FECHA_CIERRE, INTERVAL 3 MONTH) "
        filtro = filtro & " and revisada_usuario_id = 0 "
        strJoins = " LEFT JOIN clientes as clientes ON procnc.auditoria_id = clientes.ID_CLIENTE " & _
                   " LEFT JOIN decodificadora AS dAuditoria2 ON procnc.auditoria_id = dAuditoria2.valor AND dAuditoria2.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_INTERNA & _
                   " LEFT JOIN decodificadora AS dAuditoria3 ON procnc.auditoria_id = dAuditoria3.valor AND dAuditoria3.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_EXTERNA & _
                   " LEFT JOIN decodificadora AS dAuditoria4 ON procnc.auditoria_id = dAuditoria4.valor AND dAuditoria4.codigo = " & DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_DETECCION

        consulta = "" & _
        "SELECT procnc.*" & _
        " ,coalesce(procnc.fecha_cierre,'') as fecha_cierre_vista" & _
        " ,usuarios.usuario as RESPONSABLE" & _
        " ,d1.descripcion AS RESPONSABLE_DEPARTAMENTO " & _
        " ,d2.descripcion as estado" & _
        " ,dOrigen.descripcion as origen" & _
        " ,dTipo.descripcion as tipo" & _
        " ,clientes.NOMBRE as auditoria1,dAuditoria2.DESCRIPCION as auditoria2, dAuditoria3.DESCRIPCION as auditoria3, dAuditoria4.DESCRIPCION  as auditoria4" & _
        " ,if (current_date > date_add(procnc.FECHA_CIERRE,interval 3 MONTH),1,0) as flimite " & _
        " FROM procnc " & _
        " join usuarios on procnc.responsable_id = usuarios.id_empleado" & _
        " join decodificadora as d1 on procnc.responsable_id_departamento = d1.valor and d1.codigo = " & DECODIFICADORA.PROCNC_DEPARTAMENTOS & "" & _
        " join decodificadora as d2 on procnc.estado_id = d2.valor and d2.codigo =110" & _
        " join decodificadora as dOrigen on procnc.origen_id = dOrigen.valor and dOrigen.codigo =" & DECODIFICADORA.PROCNC_ORIGEN & _
        " join decodificadora as dTipo on procnc.tipo_id = dTipo.valor and dTipo.codigo =" & DECODIFICADORA.PROCNC_TIPOS_NO_CONFORMIDAD & _
        strJoins & _
        " WHERE 1=1 " & filtro & _
        " ORDER BY procnc.FECHA_ALTA DESC"

        Set ListadoRevision = datos_bd(consulta)

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.ListadoRevision"
    Exit Function
Listado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.ListadoRevision"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Listado of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
        
End Function
Public Function ListadoRevisionCantidad() As Integer
        Dim consulta As String, filtro As String, strJoins As String
        filtro = filtro & " and RESPONSABLE_ID = " & USUARIO.getID_EMPLEADO
        filtro = filtro & " and CURRENT_DATE > DATE_ADD(procnc.FECHA_CIERRE, INTERVAL 3 MONTH) "
        filtro = filtro & " and revisada_usuario_id = 0 "
        
        consulta = "" & _
        "SELECT count(*) " & _
        " FROM procnc " & _
        " WHERE 1=1 " & filtro
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        
        ListadoRevisionCantidad = rs(0)
        Exit Function
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.ListadoRevisionCantidad"
    Exit Function
Listado_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsProcNc.ListadoRevisionCantidad"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Listado of Módulo de clase clsProcNc" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
        
End Function


