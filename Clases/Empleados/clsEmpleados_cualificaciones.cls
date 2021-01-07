VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_cualificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES
'***************************************************************
Private ID_CUALIFICACION As Long
Private EMPLEADO_ID As Long
Private DOCUMENTO_ID As Long
Private EMPLEADO_ID_FORMADOR As Long
Private FECHA_FORMACION_TEORICA As String
Private TEXTO_FORMACION_TEORICA As String
Private FECHA_FIRMA_TECNICO As String
Private FECHA_FIRMA_FORMADOR As String
Private FECHA_FIRMA_DIRECTOR As String
Private ESTADO As Long
Private MODALIDAD As Long
Private FECHA_ULTIMA_RECUALIFICACION As String
Private EN_HISTORICO As Integer
Private ES_FORMADOR As Integer
Private FORMADOR_NO_CUALIFICADO As Integer
Private mtl As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES
'***************************************************************
Public Property Let setID_CUALIFICACION(ByVal dato As Long)
        ID_CUALIFICACION = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setEMPLEADO_ID_FORMADOR(ByVal dato As Long)
        EMPLEADO_ID_FORMADOR = dato
End Property
Public Property Let setFECHA_FORMACION_TEORICA(ByVal dato As String)
        FECHA_FORMACION_TEORICA = dato
End Property
Public Property Let setTEXTO_FORMACION_TEORICA(ByVal dato As String)
        TEXTO_FORMACION_TEORICA = dato
End Property
Public Property Let setFECHA_FIRMA_TECNICO(ByVal dato As String)
        FECHA_FIRMA_TECNICO = dato
End Property
Public Property Let setFECHA_FIRMA_FORMADOR(ByVal dato As String)
        FECHA_FIRMA_FORMADOR = dato
End Property
Public Property Let setFECHA_FIRMA_DIRECTOR(ByVal dato As String)
        FECHA_FIRMA_DIRECTOR = dato
End Property
Public Property Let setESTADO(ByVal dato As Long)
        ESTADO = dato
End Property
Public Property Let setMODALIDAD(ByVal dato As Long)
        MODALIDAD = dato
End Property
Public Property Let setFECHA_ULTIMA_RECUALIFICACION(ByVal dato As String)
        FECHA_ULTIMA_RECUALIFICACION = dato
End Property
Public Property Let setEN_HISTORICO(ByVal dato As Integer)
        EN_HISTORICO = dato
End Property
Public Property Let setES_FORMADOR(ByVal dato As Integer)
        ES_FORMADOR = dato
End Property
Public Property Let setFORMADOR_NO_CUALIFICADO(ByVal dato As Integer)
        FORMADOR_NO_CUALIFICADO = dato
End Property
Public Property Let setMTL(ByVal dato As Integer)
        mtl = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES
'***************************************************************
Public Property Get getID_CUALIFICACION() As Long
        getID_CUALIFICACION = ID_CUALIFICACION
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getEMPLEADO_ID_FORMADOR() As Long
        getEMPLEADO_ID_FORMADOR = EMPLEADO_ID_FORMADOR
End Property
Public Property Get getFECHA_FORMACION_TEORICA() As String
        getFECHA_FORMACION_TEORICA = FECHA_FORMACION_TEORICA
End Property
Public Property Get getTEXTO_FORMACION_TEORICA() As String
        getTEXTO_FORMACION_TEORICA = TEXTO_FORMACION_TEORICA
End Property
Public Property Get getFECHA_FIRMA_TECNICO() As String
        getFECHA_FIRMA_TECNICO = FECHA_FIRMA_TECNICO
End Property
Public Property Get getFECHA_FIRMA_FORMADOR() As String
        getFECHA_FIRMA_FORMADOR = FECHA_FIRMA_FORMADOR
End Property
Public Property Get getFECHA_FIRMA_DIRECTOR() As String
        getFECHA_FIRMA_DIRECTOR = FECHA_FIRMA_DIRECTOR
End Property
Public Property Get getESTADO() As Long
        getESTADO = ESTADO
End Property
Public Property Get getMODALIDAD() As Long
        getMODALIDAD = MODALIDAD
End Property
Public Property Get getFECHA_ULTIMA_RECUALIFICACION() As String
        getFECHA_ULTIMA_RECUALIFICACION = FECHA_ULTIMA_RECUALIFICACION
End Property
Public Property Get getEN_HISTORICO() As Integer
        getEN_HISTORICO = EN_HISTORICO
End Property
Public Property Get getES_FORMADOR() As Integer
        getES_FORMADOR = ES_FORMADOR
End Property
Public Property Get getFORMADOR_NO_CUALIFICADO() As Integer
        getFORMADOR_NO_CUALIFICADO = FORMADOR_NO_CUALIFICADO
End Property
Public Property Get getMTL() As Integer
        getMTL = mtl
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_CUALIFICACIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM empleados_cualificaciones WHERE ID_CUALIFICACION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CUALIFICACION = rs("ID_CUALIFICACION")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                EMPLEADO_ID_FORMADOR = rs("EMPLEADO_ID_FORMADOR")
                FECHA_FORMACION_TEORICA = rs("FECHA_FORMACION_TEORICA")
                TEXTO_FORMACION_TEORICA = rs("TEXTO_FORMACION_TEORICA")
                FECHA_FIRMA_TECNICO = rs("FECHA_FIRMA_TECNICO")
                FECHA_FIRMA_FORMADOR = rs("FECHA_FIRMA_FORMADOR")
                FECHA_FIRMA_DIRECTOR = rs("FECHA_FIRMA_DIRECTOR")
                ESTADO = rs("ESTADO")
                MODALIDAD = rs("MODALIDAD")
                FECHA_ULTIMA_RECUALIFICACION = rs("FECHA_ULTIMA_RECUALIFICACION")
                EN_HISTORICO = rs("EN_HISTORICO")
                ES_FORMADOR = rs("ES_FORMADOR")
                FORMADOR_NO_CUALIFICADO = rs("FORMADOR_NO_CUALIFICADO")
                mtl = rs("MTL")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_cualificaciones)", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_CUALIFICACION) FROM empleados_cualificaciones"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_CUALIFICACION = 1
    Else
        ID_CUALIFICACION = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO empleados_cualificaciones " & _
                   "  VALUES (" & _
                        ID_CUALIFICACION & "," & EMPLEADO_ID & "," & DOCUMENTO_ID & "," & EMPLEADO_ID_FORMADOR & "," & _
                        "'" & FECHA_FORMACION_TEORICA & "'" & "," & "'" & TEXTO_FORMACION_TEORICA & "'" & "," & "'" & FECHA_FIRMA_TECNICO & "'" & "," & _
                        "'" & FECHA_FIRMA_FORMADOR & "'" & "," & "'" & FECHA_FIRMA_DIRECTOR & "'" & "," & ESTADO & "," & _
                        MODALIDAD & "," & "'" & FECHA_ULTIMA_RECUALIFICACION & "'," & EN_HISTORICO & "," & ES_FORMADOR & "," & FORMADOR_NO_CUALIFICADO & "," & mtl & _
                ")"
        execute_bd consulta
        Insertar = ID_CUALIFICACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_cualificaciones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE empleados_cualificaciones SET " & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " EMPLEADO_ID_FORMADOR = " & EMPLEADO_ID_FORMADOR & "," & _
                        " FECHA_FORMACION_TEORICA = '" & FECHA_FORMACION_TEORICA & "'," & _
                        " TEXTO_FORMACION_TEORICA = '" & TEXTO_FORMACION_TEORICA & "'," & _
                        " FECHA_FIRMA_TECNICO = '" & FECHA_FIRMA_TECNICO & "'," & _
                        " FECHA_FIRMA_FORMADOR = '" & FECHA_FIRMA_FORMADOR & "'," & _
                        " FECHA_FIRMA_DIRECTOR = '" & FECHA_FIRMA_DIRECTOR & "'," & _
                        " ESTADO = " & ESTADO & "," & _
                        " MODALIDAD = " & MODALIDAD & "," & _
                        " EN_HISTORICO = " & EN_HISTORICO & "," & _
                        " ES_FORMADOR = " & ES_FORMADOR & "," & _
                        " FORMADOR_NO_CUALIFICADO = " & FORMADOR_NO_CUALIFICADO & "," & _
                        " MTL = " & mtl & "," & _
                        " FECHA_ULTIMA_RECUALIFICACION = '" & FECHA_ULTIMA_RECUALIFICACION & "'" & _
                " WHERE ID_CUALIFICACION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEmpleados_cualificaciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "DELETE FROM empleados_cualificaciones WHERE ID_CUALIFICACION = " & ID
        execute_bd "DELETE FROM empleados_cualificaciones_muestras WHERE CUALIFICACION_ID = " & ID
        execute_bd "DELETE FROM empleados_cualificaciones_evidencias WHERE CUALIFICACION_ID = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_cualificaciones)", vbCritical, Err.Description
End Function
Public Function Listado(ID_EMPLEADO As Long, PNT As String, INTERNA As Boolean, EXTERNA As Boolean, formador As Long) As ADODB.Recordset
        Dim consulta As String
        Dim MODALIDAD As String
        If INTERNA And Not EXTERNA Then
            MODALIDAD = " AND EC.MODALIDAD = 0 "
        ElseIf EXTERNA And Not INTERNA Then
            MODALIDAD = " AND EC.MODALIDAD = 1 "
        End If
        consulta = "SELECT EC.ID_CUALIFICACION, CD.NOMBRE, EC.MODALIDAD, E.NOMBRE, " & _
                   "       EC.FECHA_FORMACION_TEORICA, EC.FECHA_FIRMA_DIRECTOR, " & _
                   "       EC.FECHA_ULTIMA_RECUALIFICACION, EFE.NOMBRE, EC.EN_HISTORICO, EC.DOCUMENTO_ID " & _
                   "  FROM empleados_cualificaciones EC " & _
                   "  LEFT JOIN empleados E ON EC.EMPLEADO_ID_FORMADOR = E.ID_EMPLEADO " & _
                   "  LEFT JOIN ca_documentos CD ON EC.DOCUMENTO_ID = CD.ID_DOCUMENTO " & _
                   "  LEFT JOIN proveedores EFE ON EC.EMPLEADO_ID_FORMADOR = EFE.ID_PROVEEDOR " & _
                   " WHERE EC.EMPLEADO_ID = " & ID_EMPLEADO & _
                   "   AND CD.NOMBRE LIKE '%" & PNT & "%'" & _
                   IIf(formador <> 0, " AND EC.EMPLEADO_ID_FORMADOR = " & formador, "") & _
                   MODALIDAD & _
                   " ORDER BY ID_CUALIFICACION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Formacion(EMPLEADO As Long, PNT As String, formado As Long) As ADODB.Recordset
        Dim where As String
        If PNT <> "" Then
            where = where & " AND ca.NOMBRE like '%" & PNT & "%' "
        End If
        If formado <> 0 Then
            where = where & " AND ec.EMPLEADO_ID = " & formado
        End If
        Dim consulta As String
        consulta = "select ec.EMPLEADO_ID,ec.ID_CUALIFICACION, ca.NOMBRE , ec.FECHA_FIRMA_TECNICO,'', ec.empleado_id_formador,ec.DOCUMENTO_ID " & _
                   " from empleados_cualificaciones ec inner join ca_documentos ca on ec.DOCUMENTO_ID = ca.ID_DOCUMENTO " & _
                   " WHERE ec.EMPLEADO_ID = " & EMPLEADO & _
                   "   And ec.ES_FORMADOR = 1 And ec.modalidad = 0 " & _
                   where & _
                   " Union " & _
                   " select ec.EMPLEADO_ID,ec.ID_CUALIFICACION, ca.NOMBRE, ec.FECHA_FIRMA_TECNICO , emp.NOMBRE, ec.empleado_id_formador,ec.DOCUMENTO_ID " & _
                   " from empleados_cualificaciones ec inner join empleados emp on ec.EMPLEADO_ID = emp.ID_EMPLEADO " & _
                   "                                   inner join ca_documentos ca on ec.DOCUMENTO_ID = ca.ID_DOCUMENTO " & _
                   " WHERE ec.EMPLEADO_ID_FORMADOR = " & EMPLEADO & _
                   "   And ec.modalidad = 0 " & _
                   where & _
                   " order by 3,5"
        Set Listado_Formacion = datos_bd(consulta)
End Function
Public Function Listado_Cualificacion(ID_CUALIFICACION As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT EC.ID_CUALIFICACION, CD.NOMBRE, EC.MODALIDAD, E.NOMBRE, EC.FECHA_FORMACION_TEORICA, " & _
                   "       EC.FECHA_FIRMA_DIRECTOR, EC.FECHA_ULTIMA_RECUALIFICACION, EC.EMPLEADO_ID_FORMADOR, " & _
                   "       EC.EN_HISTORICO, EC.DOCUMENTO_ID " & _
                   "  FROM empleados_cualificaciones EC, empleados E, ca_documentos CD " & _
                   " WHERE EC.EMPLEADO_ID_FORMADOR = E.ID_EMPLEADO " & _
                   "   AND EC.DOCUMENTO_ID = CD.ID_DOCUMENTO " & _
                   "   AND EC.ID_CUALIFICACION = " & ID_CUALIFICACION
        Set Listado_Cualificacion = datos_bd(consulta)
End Function
'M1110-I
Public Function Listado_Empleado_DOC(ID_EMPLEADO As Long, DOC_ID As Long, formador As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM empleados_cualificaciones WHERE EMPLEADO_ID = " & ID_EMPLEADO & " AND DOCUMENTO_ID = " & DOC_ID & _
                   " AND FECHA_FORMACION_TEORICA <> ''"
        If formador > 0 Then
            consulta = consulta & " AND ES_FORMADOR = 1"
        End If
        Set Listado_Empleado_DOC = datos_bd(consulta)
End Function
'M1110-F

Public Function MatrizCualificaciones(oFamilia As String, oSUBFAMILIA As String, oPNT As Long, oPNTTEXTO As String, oENAC As Integer, oNADCAP As Integer, oEQA As Integer, oTIPO As Integer, empleados As String, empresa As String, oMTL As Integer, oBaja As Integer) As ADODB.Recordset
        Dim consulta As String
        Dim where As String
        where = " WHERE 1 = 1 "
        If oFamilia <> "" And oFamilia <> "0" Then
            where = where & " AND ca.familia_id = " & CInt(oFamilia)
        End If
        If oSUBFAMILIA <> "" And oSUBFAMILIA <> "0" Then
            where = where & " AND ca.subfamilia_id = " & CInt(oSUBFAMILIA)
        End If
        If oPNT <> 0 Then
            where = where & " AND ca.id_documento = " & oPNT
        End If
        If oPNTTEXTO <> "" Then
            where = where & " AND ca.nombre like '%" & oPNTTEXTO & "%' "
        End If
        If oENAC = 1 Then
            where = where & " AND ca.ENAC = 1 "
        End If
        If oNADCAP = 1 Then
            where = where & " AND ca.NADCAP = 1 "
        End If
        If oEQA = 1 Then
            where = where & " AND ca.EQA = 1 "
        End If
        If empresa <> "" And empresa <> "0" Then
            where = where & " and e.EMPRESA_ID = " & CInt(empresa)
        End If
        If oMTL = 1 Then
            where = where & " AND ec.MTL = 1 "
        End If
        
        If empleados <> "" Then
            where = where & " AND ec.EMPLEADO_ID IN (" & empleados & ") "
        End If
        If oBaja = 0 Then
            where = where & " AND e.ESTADO_ID = 0 "
        End If
        where = where & " AND ca.anulado = 0 "
        ' PNT 2 años
        ' PNT C 3 años
        If oTIPO <> 0 Then
            Select Case oTIPO
            Case 1 ' Recualificaciones Todas
                where = where & " AND ec.fecha_ultima_recualificacion <> '1900-01-01' "
            Case 2 ' Recualificaciones pendientes
                where = where & " AND (( " & _
                      " CURRENT_DATE > DATE_ADD(ec.FECHA_ULTIMA_RECUALIFICACION, INTERVAL IF (INSTR(ca.NOMBRE,'PNT C'),3,2) YEAR) " & _
                      " AND ec.FECHA_ULTIMA_RECUALIFICACION <> '1900-01-01' " & _
                      " ) OR ( " & _
                      " CURRENT_DATE > DATE_ADD(ec.FECHA_FIRMA_TECNICO , INTERVAL IF (INSTR(ca.NOMBRE,'PNT C'),3,2) YEAR) " & _
                      " AND ec.FECHA_ULTIMA_RECUALIFICACION = '1900-01-01' " & _
                      " ) " & _
                      " ) AND ec.EN_HISTORICO = 0 "
            Case 3 ' Formador
                where = where & " AND ec.ES_FORMADOR = 1 "
            Case 4 ' En formacion
                where = where & " AND ec.FECHA_FIRMA_TECNICO = '1900-01-01' "
            End Select
        End If
'        WHERE = WHERE & " AND ec.ID_CUALIFICACION IN (158,196) "
        consulta = "select ec.DOCUMENTO_ID,ec.EMPLEADO_ID, ca.NOMBRE, e.NOMBRE, " & _
                   "       ec.FECHA_ULTIMA_RECUALIFICACION, ec.EN_HISTORICO, ec.FECHA_FIRMA_TECNICO, " & _
                   "       ec.ID_CUALIFICACION, ec.ES_FORMADOR, usu.imagen, CA.CODIGO,ec.FORMADOR_NO_CUALIFICADO " & _
                   "  from empleados_cualificaciones ec " & _
                   " inner join empleados e on ec.EMPLEADO_ID = e.ID_EMPLEADO " & _
                   " inner join ca_documentos ca on ec.DOCUMENTO_ID = ca.ID_DOCUMENTO " & _
                   " inner join usuarios usu on e.USUARIO_ID = usu.ID_EMPLEADO " & _
                   where & _
                   " order by ca.NOMBRE,e.nombre "
        Set MatrizCualificaciones = datos_bd(consulta)
End Function

Public Function RecualificacionesPendientes(USUARIO As Long, todas As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim where As String
        where = " WHERE 1 = 1 "
        If todas = False Then
            If USUARIO <> 0 Then
                Dim oE As New clsEmpleados
                oE.CARGAR_POR_USUARIO USUARIO
                where = where & " AND EC.EMPLEADO_ID = " & CInt(oE.getID_EMPLEADO)
                Set oE = Nothing
            End If
        End If
        consulta = "SELECT EC.ID_CUALIFICACION,D.NOMBRE,E1.NOMBRE,E2.NOMBRE,EC.FECHA_ULTIMA_RECUALIFICACION," & _
                    "      E1.ID_EMPLEADO,DATE_ADD(EC.FECHA_ULTIMA_RECUALIFICACION, INTERVAL IF (INSTR(D.NOMBRE,'PNT C'),3,2) YEAR) " & _
                    " FROM empleados_cualificaciones EC " & _
                    " INNER JOIN ca_documentos D ON EC.DOCUMENTO_ID = D.ID_DOCUMENTO " & _
                    " INNER JOIN empleados E1 ON EC.EMPLEADO_ID = E1.ID_EMPLEADO " & _
                    " INNER JOIN empleados E2 ON EC.EMPLEADO_ID_FORMADOR = E2.ID_EMPLEADO " & _
                    where & _
                    " AND E1.ESTADO_ID = 0 AND EC.EN_HISTORICO = 0 AND EC.FECHA_FIRMA_TECNICO <> '1900-01-01' " & _
                    " AND D.USO = 1 " & _
                    " AND ( " & _
                    " (CURRENT_DATE > DATE_ADD(EC.FECHA_ULTIMA_RECUALIFICACION, INTERVAL IF (INSTR(D.NOMBRE,'PNT C'),35,23) MONTH) AND EC.FECHA_ULTIMA_RECUALIFICACION <> '1900-01-01') OR " & _
                    " (CURRENT_DATE > DATE_ADD(EC.FECHA_FIRMA_TECNICO, INTERVAL IF (INSTR(D.NOMBRE,'PNT C'),35,23) MONTH) AND EC.FECHA_ULTIMA_RECUALIFICACION = '1900-01-01') " & _
                    " )"
        'Todos los PNT caducidad dos años (23 meses) excepto PNT C que son tres años (35 meses) : 1 mes menos para avisar con antelación
        Set RecualificacionesPendientes = datos_bd(consulta)
End Function

Public Function RecualificacionesPendientesUsuario(USUARIO As Long) As Integer
        Dim consulta As String
        Dim where As String
        where = " WHERE 1 = 1 "
        If USUARIO <> 0 Then
            Dim oE As New clsEmpleados
            oE.CARGAR_POR_USUARIO USUARIO
            where = where & " AND EC.EMPLEADO_ID = " & CInt(oE.getID_EMPLEADO)
            Set oE = Nothing
        End If
        consulta = "SELECT COUNT(*) " & _
                    " FROM empleados_cualificaciones EC " & _
                    " INNER JOIN ca_documentos D ON EC.DOCUMENTO_ID = D.ID_DOCUMENTO " & _
                    " INNER JOIN empleados E1 ON EC.EMPLEADO_ID = E1.ID_EMPLEADO " & _
                    " INNER JOIN empleados E2 ON EC.EMPLEADO_ID_FORMADOR = E2.ID_EMPLEADO " & _
                    where & _
                    " AND E1.ESTADO_ID = 0 AND EC.EN_HISTORICO = 0 AND EC.FECHA_FIRMA_TECNICO <> '1900-01-01' " & _
                    " AND D.USO = 1 " & _
                    " AND ( " & _
                    " (CURRENT_DATE > DATE_ADD(EC.FECHA_ULTIMA_RECUALIFICACION, INTERVAL IF (INSTR(D.NOMBRE,'PNT C'),35,23) MONTH) AND EC.FECHA_ULTIMA_RECUALIFICACION <> '1900-01-01') OR " & _
                    " (CURRENT_DATE > DATE_ADD(EC.FECHA_FIRMA_TECNICO, INTERVAL IF (INSTR(D.NOMBRE,'PNT C'),35,23) MONTH) AND EC.FECHA_ULTIMA_RECUALIFICACION = '1900-01-01') " & _
                    " )"
        'Todos los PNT caducidad dos años (23 meses) excepto PNT C que son tres años (35 meses) : 1 mes menos para avisar con antelación
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            RecualificacionesPendientesUsuario = rs(0)
        Else
            RecualificacionesPendientesUsuario = 0
        End If
End Function
Public Function estaCualificadoPNT(USUARIO As Long, DOCUMENTO_ID As Long) As Boolean
        Dim consulta As String
        Dim where As String
        where = " WHERE 1 = 1 "
        If USUARIO <> 0 Then
            Dim oE As New clsEmpleados
            oE.CARGAR_POR_USUARIO USUARIO
            where = where & " AND EC.EMPLEADO_ID = " & CInt(oE.getID_EMPLEADO)
            Set oE = Nothing
        End If
        ' Verificar si tiene registro en la tabla de cualificaciones
        Dim rs As ADODB.Recordset
        consulta = "SELECT EC.ID_CUALIFICACION " & _
                    " FROM empleados_cualificaciones EC " & _
                    where & _
                    " AND EC.DOCUMENTO_ID = " & DOCUMENTO_ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            estaCualificadoPNT = False
            Exit Function
        End If
        ' Verificar si la fecha de cualificación esta correcta
        consulta = "SELECT EC.ID_CUALIFICACION " & _
                    " FROM empleados_cualificaciones EC " & _
                    " INNER JOIN ca_documentos D ON EC.DOCUMENTO_ID = D.ID_DOCUMENTO " & _
                    " INNER JOIN empleados E1 ON EC.EMPLEADO_ID = E1.ID_EMPLEADO " & _
                    " INNER JOIN empleados E2 ON EC.EMPLEADO_ID_FORMADOR = E2.ID_EMPLEADO " & _
                    where & _
                    " AND EC.DOCUMENTO_ID = " & DOCUMENTO_ID & _
                    " AND E1.ESTADO_ID = 0 AND EC.EN_HISTORICO = 0 AND EC.FECHA_FIRMA_TECNICO <> '1900-01-01' " & _
                    " AND D.USO = 1 " & _
                    " AND ( " & _
                    " (CURRENT_DATE > DATE_ADD(EC.FECHA_ULTIMA_RECUALIFICACION, INTERVAL IF (INSTR(D.NOMBRE,'PNT C'),35,23) MONTH) AND EC.FECHA_ULTIMA_RECUALIFICACION <> '1900-01-01') OR " & _
                    " (CURRENT_DATE > DATE_ADD(EC.FECHA_FIRMA_TECNICO, INTERVAL IF (INSTR(D.NOMBRE,'PNT C'),35,23) MONTH) AND EC.FECHA_ULTIMA_RECUALIFICACION = '1900-01-01') " & _
                    " )"
        'Todos los PNT caducidad dos años (23 meses) excepto PNT C que son tres años (35 meses) : 1 mes menos para avisar con antelación
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            estaCualificadoPNT = False
        Else
            estaCualificadoPNT = True
        End If
End Function


Public Function enviarCorreoNoCualificado(DOCUMENTO_ID As Long, MOTIVO As String) As Boolean
    Dim destinatario As String
    Dim mensaje As String
    Dim ASUNTO As String
    Dim oParametro As New clsParametros
    Dim oca_documentos As New clsCa_documentos
   On Error GoTo enviarCorreoNoCualificado_Error

    oca_documentos.Carga DOCUMENTO_ID
    oParametro.Carga parametros.ENVIO_CORREO_NO_CUALIFICADO, ""
    destinatario = oParametro.getVALOR
    If destinatario <> "" Then
        ASUNTO = "El Usuario " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS & ", no esta cualificado en el procedimiento : " & oca_documentos.getCODIGO
        mensaje = "Usuario No cualificado: " & vbNewLine & vbNewLine
        mensaje = mensaje & vbNewLine & " Motivo : " & MOTIVO
        mensaje = mensaje & vbNewLine & " Código : " & oca_documentos.getCODIGO
        mensaje = mensaje & vbNewLine & " Usuario : " & USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        mensaje = mensaje & vbNewLine & " Fecha : " & USUARIO.getUSUARIO
        mensaje = mensaje & vbNewLine & vbNewLine
        mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
        ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
    End If
    Set oParametro = Nothing

   On Error GoTo 0
   Exit Function

enviarCorreoNoCualificado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure enviarCorreoNoCualificado of Módulo de clase clsEmpleados_cualificaciones"
End Function


