VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquiposOperacionesPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsEquiposOperacionesPendientes
'***************************************************************
Private ID_CVM As Long
Private EQUIPO_ID As Long
Private TIPO_CVM_ID As Long
Private ID_CVM_ORIGEN As Long
Private RESPONSABLE_ID As String
Private PERIODICIDAD_ID As Long
Private FECHA_PREVISTA As String
Private fecha_preaviso As String
Public Function crear_calibracion_pendiente(ByVal prm_id_calibracion As Long, ByVal prm_id_calibracion_prevista As Long, ByVal prm_id_equipo As Long, ByRef prm_oPer As clsEquiposPeriodicidad) As Boolean

On Error GoTo crear_calibracion_pendiente_err

    Dim rs As ADODB.Recordset
    Dim consulta As String, consulta_alt As String
    
    Set rs = datos_bd("select * from eq_calibracion_equipos where id_calibracion = " & CStr(prm_id_calibracion_prevista))
    rs.MoveFirst
        
    
        EQUIPO_ID = prm_id_equipo
        fecha_preaviso = prm_oPer.calcular_fecha_preaviso(rs!FECHA_PREVISTA, rs!PERIODICIDAD_ID)
        FECHA_PREVISTA = rs!FECHA_PREVISTA
        ID_CVM = prm_id_calibracion_prevista
        ID_CVM_ORIGEN = prm_id_calibracion
        PERIODICIDAD_ID = rs!PERIODICIDAD_ID
        RESPONSABLE_ID = rs!CALIBRADOR_INTERNO_ID
        TIPO_CVM_ID = EQUIPOS_CVM_TIPOS.CVM_TIPO_CALIBRACION
        If Insertar <> 0 Then
            crear_calibracion_pendiente = True
        Else
            crear_calibracion_pendiente = False
        End If
    
    'crear_calibracion_pendiente = True
    
Exit Function

crear_calibracion_pendiente_err:

crear_calibracion_pendiente = False
End Function

Public Function crear_mantenimientos_pendiente_para_nuevas_fechas_planes(ByVal ID_EQUIPO As Long, ByVal lista_planes As String) As Boolean

Dim arrPM() As String, rs As ADODB.Recordset, rs2 As ADODB.Recordset
Dim consulta As String, x As Long

' revisa las tareas pendientes para los nuevos mantenimientos generados

lista_planes = Replace(lista_planes, "::", ",")
lista_planes = Replace(lista_planes, ":", "")

arrPM = Split(lista_planes, ",")

For x = 0 To UBound(arrPM)
    
    'recogemos los mantenimientos creados y previstos segun este plan de mantenimiento
    consulta = "SELECT * from eq_mantenimiento_equipos where estado = 0 and equipo_id = " & CStr(ID_EQUIPO) & " and planmto_id = " & arrPM(x) & " order by fecha_actual asc"
    Set rs = datos_bd(consulta)
    
    If rs.RecordCount <> 0 Then
        ' Tomamos, de este, el de menor fecha
        Do
'        rs.MoveFirst
        
        ' recogemos las tareas pendientes
'            Set rs2 = devolver_registros_mantenimiento_segun_plan(ID_EQUIPO, CLng(arrPM(x)))
            
'            If rs2.RecordCount = 0 Then
                ' no existe, por lo que habrá que darlo de alta
                crear_mantenimiento_pendiente 0, rs!id_mantenimiento, rs!EQUIPO_ID, New clsEquiposPeriodicidad
'            Else
                ' existe, por lo que habrá que comprobar que si la fecha de la tarea pendiente es posterior a la nueva
'                If CDate(rs2!FECHA_PREVISTA) < CDate(rs!fecha_actual) Then
                    ' al ser el nuevo mantenimiento menor que el que existia, se reemplaza
'                    crear_mantenimiento_pendiente rs2!id_mantenimiento, rs!id_mantenimiento, rs!EQUIPO_ID, New clsEquiposPeriodicidad
'                End If
'            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
Next x




End Function

Public Function crear_verificacion_pendiente(ByVal prm_ID_VERIFICACION As Long, ByVal prm_id_verificacion_prevista As Long, ByVal prm_id_equipo As Long, ByRef prm_oPer As clsEquiposPeriodicidad) As Boolean

On Error GoTo crear_verificacion_pendiente_err

    Dim rs As ADODB.Recordset
    Dim consulta As String, consulta_alt As String
    
    Set rs = datos_bd("select * from eq_verificacion_equipos where id_verificacion = " & CStr(prm_id_verificacion_prevista))
    rs.MoveFirst
        
    
        EQUIPO_ID = prm_id_equipo
        fecha_preaviso = prm_oPer.calcular_fecha_preaviso(rs!FECHA_PREVISTA, rs!PERIODICIDAD_ID)
        FECHA_PREVISTA = rs!FECHA_PREVISTA
        ID_CVM = prm_id_verificacion_prevista
        ID_CVM_ORIGEN = prm_ID_VERIFICACION
        PERIODICIDAD_ID = rs!PERIODICIDAD_ID
        RESPONSABLE_ID = rs!VERIFICADOR_INTERNO_ID
        TIPO_CVM_ID = EQUIPOS_CVM_TIPOS.CVM_TIPO_VERIFICACION
        If Insertar <> 0 Then
            crear_verificacion_pendiente = True
        Else
            crear_verificacion_pendiente = False
        End If
    
Exit Function

crear_verificacion_pendiente_err:

crear_verificacion_pendiente = False
End Function


Public Function crear_mantenimiento_pendiente(ByVal prm_ID_MANTENIMIENTO As Long, ByVal prm_ID_MANTENIMIENTO_prevista As Long, ByVal prm_id_equipo As Long, ByRef prm_oPer As clsEquiposPeriodicidad) As Boolean

On Error GoTo crear_mantenimiento_pendiente_err

    Dim rs As ADODB.Recordset
    Dim consulta As String, consulta_alt As String
    
    consulta = "select em.*, p.frecuencia_id from eq_MANTENIMIENTO_equipos em"
    consulta = consulta & " inner join eq_planes_mantenimiento p on em.planmto_id = p.id_plan_mto"
    consulta = consulta & " where em.id_MANTENIMIENTO = " & CStr(prm_ID_MANTENIMIENTO_prevista)
    Set rs = datos_bd(consulta)
    rs.MoveFirst
        
        EQUIPO_ID = prm_id_equipo
        fecha_preaviso = prm_oPer.calcular_fecha_preaviso(rs!fecha_actual, rs!FRECUENCIA_ID)
        FECHA_PREVISTA = rs!fecha_actual
        ID_CVM = prm_ID_MANTENIMIENTO_prevista
        ID_CVM_ORIGEN = prm_ID_MANTENIMIENTO
        PERIODICIDAD_ID = rs!FRECUENCIA_ID
        RESPONSABLE_ID = rs!MANTENEDOR_ID
        TIPO_CVM_ID = EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
        Insertar
    
'    Me.revisar_tarea prm_ID_MANTENIMIENTO, prm_ID_MANTENIMIENTO_prevista, EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO
    
    
    crear_mantenimiento_pendiente = True
    
Exit Function

crear_mantenimiento_pendiente_err:

crear_mantenimiento_pendiente = False
End Function



Public Function devolver_ultimo_registro(ByVal ID_EQUIPO As Long, ByVal tipo As Integer) As ADODB.Recordset
    Dim strCad As String
    
    
    strCad = "select * from eq_operaciones_pendientes where equipo_id = " & CStr(ID_EQUIPO) & " AND tipo_cvm_id = " & CStr(tipo)
    
    Set devolver_ultimo_registro = datos_bd(strCad)
    
    
End Function


Public Function devolver_registros_mantenimiento_segun_plan(ByVal ID_EQUIPO As Long, ByVal plan_mto_id As Long) As ADODB.Recordset
    Dim strCad As String
    
    
    strCad = "select * from eq_operaciones_pendientes where equipo_id = " & CStr(ID_EQUIPO) & " AND tipo_cvm_id = " & CStr(EQUIPOS_CVM_TIPOS.CVM_TIPO_MANTENIMIENTO)
    strCad = strCad & " and id_cvm in (select id_mantenimiento from eq_mantenimiento_equipos em where em.planmto_id = " & CStr(plan_mto_id) & " and em.equipo_id = " & CStr(ID_EQUIPO) & " and em.estado=0)"
    
    Set devolver_registros_mantenimiento_segun_plan = datos_bd(strCad)
    
    
End Function



Public Function Listado_calibraciones_pendientes() As ADODB.Recordset

    Dim consulta As String, tipo As String
    tipo = "0"
    
    consulta = "select * from "
    consulta = consulta & "      (select op.*, cast(2 as unsigned) as estado from eq_operaciones_pendientes op where fecha_preaviso > LOCALTIMESTAMP AND fecha_prevista > localtimestamp and tipo_cvm_id = " & tipo
    consulta = consulta & " union select op.*, cast(1 as unsigned) as estado from eq_operaciones_pendientes op where fecha_preaviso <= localtimestamp and fecha_prevista > localtimestamp and tipo_cvm_id = " & tipo
    consulta = consulta & " union select op.*, cast(0 as unsigned) as estado from eq_operaciones_pendientes op where fecha_prevista <= localtimestamp and tipo_cvm_id = " & tipo & ") as t_estado"
    
    Set Listado_calibraciones_pendientes = datos_bd(consulta)

End Function

Public Function Listado_verificaciones_pendientes() As ADODB.Recordset

    Dim consulta As String, tipo As String
    tipo = "1"
    
    consulta = "select * from "
    consulta = consulta & "      (select op.*, cast(2 as unsigned) as estado from eq_operaciones_pendientes op where fecha_preaviso  > LOCALTIMESTAMP AND fecha_prevista > localtimestamp and tipo_cvm_id = " & tipo
    consulta = consulta & " union select op.*, cast(1 as unsigned) as estado from eq_operaciones_pendientes op where fecha_preaviso <= localtimestamp and fecha_prevista > localtimestamp and tipo_cvm_id = " & tipo
    consulta = consulta & " union select op.*, cast(0 as unsigned) as estado from eq_operaciones_pendientes op where fecha_prevista <= localtimestamp and tipo_cvm_id = " & tipo & ") as t_estado"
    
    Set Listado_verificaciones_pendientes = datos_bd(consulta)

End Function


Public Function Listado_mantenimientos_pendientes() As ADODB.Recordset

    Dim consulta As String, tipo As String
    tipo = "2"
    
    consulta = "select * from "
    consulta = consulta & "      (select op.*, cast(2 as unsigned) as estado from eq_operaciones_pendientes op where fecha_preaviso  > LOCALTIMESTAMP AND fecha_prevista > localtimestamp and tipo_cvm_id = " & tipo
    consulta = consulta & " union select op.*, cast(1 as unsigned) as estado from eq_operaciones_pendientes op where fecha_preaviso <= localtimestamp and fecha_prevista > localtimestamp and tipo_cvm_id = " & tipo
    consulta = consulta & " union select op.*, cast(0 as unsigned) as estado from eq_operaciones_pendientes op where fecha_prevista <= localtimestamp and tipo_cvm_id = " & tipo & ") as t_estado"
    
    Set Listado_mantenimientos_pendientes = datos_bd(consulta)

End Function



Public Function revisar_tarea(ByVal prm_id As Long, ByVal prm_id_equipo As Long, ByVal prm_id_tipo As Long) As String
    
    Dim oCVM As Object
    Dim oPer As New clsEquiposPeriodicidad
    
'    If Carga(prm_id, prm_id_tipo) Then
        
        If prm_id_tipo = 0 Then ' Calibracion
          Set oCVM = New clsEquipoCalibracion
          oCVM.Carga prm_id
          RESPONSABLE_ID = oCVM.getCALIBRADOR_INTERNO_ID
        ElseIf prm_id_tipo = 1 Then ' Verificacion
            Set oCVM = New clsEquipoVerificacion
            oCVM.Carga prm_id
            RESPONSABLE_ID = oCVM.getVERIFICADOR_INTERNO_ID
        Else ' Mantenimiento
            Set oCVM = New clsEquipoMantenimiento
            oCVM.Carga prm_id
            RESPONSABLE_ID = oCVM.getMANTENEDOR_ID
        End If
        
        Set oPer = New clsEquiposPeriodicidad
        EQUIPO_ID = prm_id_equipo
        fecha_preaviso = oPer.calcular_fecha_preaviso(oCVM.getFECHA_ACTUAL, oCVM.getPERIODICIDAD_ID)
        FECHA_PREVISTA = oCVM.getFECHA_ACTUAL
        PERIODICIDAD_ID = oCVM.getPERIODICIDAD_ID
        
        Select Case oCVM.getESTADO
            Case CAL_ESTADOS.CAL_ESTADO_PREVISTA
                If Carga(prm_id, prm_id_tipo) Then
                    Modificar
                Else
                    If prm_id_tipo = 0 Then ' Calibracion
                        crear_calibracion_pendiente 0, prm_id, prm_id_equipo, oPer
                    ElseIf prm_id_tipo = 1 Then ' Verificacion
                        crear_verificacion_pendiente 0, prm_id, prm_id_equipo, oPer
                    Else ' Mantenimiento
                        crear_mantenimiento_pendiente 0, prm_id, prm_id_equipo, oPer
                    End If
                End If
            Case Else
                Eliminar prm_id, prm_id_equipo, prm_id_tipo
        End Select
        
'    Else
'        If prm_id_tipo = 0 Then ' Calibracion
'            crear_calibracion_pendiente 0, prm_id, prm_id_equipo, oPer
'        ElseIf prm_id_tipo = 1 Then ' Verificacion
'            crear_verificacion_pendiente 0, prm_id, prm_id_equipo, oPer
'        Else ' Mantenimiento
'            crear_mantenimiento_pendiente 0, prm_id, prm_id_equipo, oPer
'        End If
'    End If

    Set oCVM = Nothing
    Set oPer = Nothing
    
End Function

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsEquiposOperacionesPendientes
'***************************************************************
Public Property Let setID_CVM(ByVal dato As Long)
        ID_CVM = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setTIPO_CVM_ID(ByVal dato As Long)
        TIPO_CVM_ID = dato
End Property
Public Property Let setID_CVM_ORIGEN(ByVal dato As Long)
        ID_CVM_ORIGEN = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As String)
        RESPONSABLE_ID = dato
End Property
Public Property Let setPERIODICIDAD_ID(ByVal dato As Long)
        PERIODICIDAD_ID = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        FECHA_PREVISTA = dato
End Property
Public Property Let setFECHA_PREAVISO(ByVal dato As String)
        fecha_preaviso = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsEquiposOperacionesPendientes
'***************************************************************
Public Property Get getID_CVM() As Long
        getID_CVM = ID_CVM
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getTIPO_CVM_ID() As Long
        getTIPO_CVM_ID = TIPO_CVM_ID
End Property
Public Property Get getID_CVM_ORIGEN() As Long
        getID_CVM_ORIGEN = ID_CVM_ORIGEN
End Property
Public Property Get getRESPONSABLE_ID() As String
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getPERIODICIDAD_ID() As Long
        getPERIODICIDAD_ID = PERIODICIDAD_ID
End Property
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = FECHA_PREVISTA
End Property
Public Property Get getFECHA_PREAVISO() As String
        getFECHA_PREAVISO = fecha_preaviso
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsEquiposOperacionesPendientes
'***************************************************************
Public Function Carga(ID As Long, TIPO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_OPERACIONES_PENDIENTES WHERE ID_CVM = " & ID & " AND TIPO_CVM_ID = " & TIPO_ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CVM = rs("ID_CVM")
                EQUIPO_ID = rs("EQUIPO_ID")
                TIPO_CVM_ID = rs("TIPO_CVM_ID")
                ID_CVM_ORIGEN = rs("ID_CVM_ORIGEN")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                PERIODICIDAD_ID = rs("PERIODICIDAD_ID")
                FECHA_PREVISTA = rs("FECHA_PREVISTA")
                fecha_preaviso = rs("FECHA_PREAVISO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquiposOperacionesPendientes)", vbCritical, Err.Description
End Function


Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        ' Elimina el aviso del equipo que le precede, si existiera
'OOP        Eliminar ID_CVM_ORIGEN, EQUIPO_ID, TIPO_CVM_ID
        
        consulta = "REPLACE INTO EQ_OPERACIONES_PENDIENTES " & _
                   "  VALUES (" & _
                        ID_CVM & "," & EQUIPO_ID & "," & TIPO_CVM_ID & "," & _
                        ID_CVM_ORIGEN & "," & "'" & RESPONSABLE_ID & "'" & "," & PERIODICIDAD_ID & "," & _
                        "'" & Format(FECHA_PREVISTA, "yyyy-mm-dd") & "'" & "," & "'" & Format(fecha_preaviso, "yyyy-mm-dd") & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_CVM
        Exit Function
fallo:
        Insertar = 0
        ' puede que la tarea ya haya sido dado de alta por una calibracion anterior
        'MsgBox "Error al insertar (clsEquiposOperacionesPendientes)", vbCritical, Err.Description
End Function
Private Function Modificar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_OPERACIONES_PENDIENTES SET " & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " TIPO_CVM_ID = " & TIPO_CVM_ID & "," & _
                        " ID_CVM_ORIGEN = " & ID_CVM_ORIGEN & "," & _
                        " RESPONSABLE_ID = '" & RESPONSABLE_ID & "'," & _
                        " PERIODICIDAD_ID = " & PERIODICIDAD_ID & "," & _
                        " FECHA_PREVISTA = '" & Format(FECHA_PREVISTA, "yyyy-mm-dd") & "'," & _
                        " FECHA_PREAVISO = '" & Format(fecha_preaviso, "yyyy-mm-dd") & "'" & _
                " WHERE ID_CVM = " & ID_CVM & " AND EQUIPO_ID = " & EQUIPO_ID & " AND TIPO_CVM_ID = " & TIPO_CVM_ID
                
                
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquiposOperacionesPendientes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, prm_id_equipo As Long, prm_id_tipo As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EQ_OPERACIONES_PENDIENTES " & _
                "    WHERE ID_CVM = " & ID & " AND EQUIPO_ID = " & prm_id_equipo & " AND TIPO_CVM_ID = " & prm_id_tipo
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquiposOperacionesPendientes)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_OPERACIONES_PENDIENTES ORDER BY ID_CVM"
        Set Listado = datos_bd(consulta)
End Function


