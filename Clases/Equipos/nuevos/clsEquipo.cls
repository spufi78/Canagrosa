VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPO
'***************************************************************
Private mvarlngID_EQUIPO As Long
Private mvarstrNOMBRE As String
Private mvarlngCONDICIONES_AMBIENTALES As Long
Private mvarintPROVEEDOR_ID As Integer
Private mvarstrFECHA_RECEPCION As String
Private mvarstrFECHA_SERVICIO As String
Private mvarintALTA_BAJA As Integer
Private mvarstrSERIE As String
Private mvarstrNOTAS As String
Private mvarintFAMILIA_ID As Integer
Private mvarstrMODELO As String
Private mvarstrFABRICANTE As String
Private mvarintES_NADCAP As Integer
Private mvarstrRANGO_MEDIDA_MIN As String
Private mvarstrRANGO_MEDIDA_MAX As String
Private mvarintUNIDAD_ID As Integer
Private mvarstrTOLERANCIA_MAXIMA As String
Private mvarstrINCERTIDUMBRE_MAXIMA As String
Private mvarstrPRECISIONN As String
Private mvarstrTEMPERATURA_MIN As String
Private mvarstrTEMPERATURA_MAX As String
Private mvarstrHUMEDAD_MIN As String
Private mvarstrHUMEDAD_MAX As String
Private mvarstrCOND_AMBIENTALES_OTRAS As String
Private mvarstrRANGO_TRABAJO_MIN As String
Private mvarstrRANGO_TRABAJO_MAX As String
Private mvarstrLIMITACIONES_USO As String
Private mvarintPERIODICIDAD_CALIBRACION_ID As Integer
Private mvarstrFECHA_PROX_CALIBRACION As String
Private mvarintPERIODICIDAD_VERIFICACION_ID As Integer
Private mvarstrFECHA_PROX_VERIFICACION As String
Private mvarintPERIODICIDAD_MANTENIMIENTO_ID As Integer
Private mvarstrFECHA_PROX_MANTENIMIENTO As String
Private mvarintTIPO_CALIBRACION_ID As Integer
Private mvarintCALIBRADOR_ID As Integer
Private mvarintCALIBRADOR_INTERNO_ID As Integer
Private mvarintSITUACION_ID As Integer
Private mvarintCON_CALIBRACION As Integer
Private mvarintCON_VERIFICACION As Integer
Private mvarintCON_MANTENIMIENTO As Integer
Private mvarintREVISADO As Integer
Private mvarlngSUBFAMILIA_ID As Long
Private mvarlngCoordX As Long
Private mvarlngCoordY As Long

Public Property Get CoordX() As Long

On Error GoTo CoordX_Error

    CoordX = mvarlngCoordX

On Error GoTo 0
    Exit Property
CoordX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CoordX of Módulo de clase clsEquipo"

End Property

Public Property Let CoordX(ByVal lngCoordX As Long)

On Error GoTo CoordX_Error

    mvarlngCoordX = lngCoordX

On Error GoTo 0
    Exit Property
CoordX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CoordX of Módulo de clase clsEquipo"

End Property



'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM EQUIPOS WHERE ID_EQUIPO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_EQUIPO = rs("ID_EQUIPO")
                NOMBRE = rs("NOMBRE")
                CONDICIONES_AMBIENTALES = rs("CONDICIONES_AMBIENTALES")
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                FECHA_RECEPCION = rs("FECHA_RECEPCION")
                FECHA_SERVICIO = rs("FECHA_SERVICIO")
                ALTA_BAJA = rs("ALTA_BAJA")
                SERIE = rs("SERIE")
                FAMILIA_ID = rs("FAMILIA_ID")
                MODELO = rs("MODELO")
                FABRICANTE = rs("FABRICANTE")
                ES_NADCAP = rs("ES_NADCAP")
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
                FECHA_PROX_CALIBRACION = rs("FECHA_PROX_CALIBRACION")
                PERIODICIDAD_VERIFICACION_ID = rs("PERIODICIDAD_VERIFICACION_ID")
                FECHA_PROX_VERIFICACION = rs("FECHA_PROX_VERIFICACION")
                PERIODICIDAD_MANTENIMIENTO_ID = rs("PERIODICIDAD_MANTENIMIENTO_ID")
                FECHA_PROX_MANTENIMIENTO = rs("FECHA_PROX_MANTENIMIENTO")
                SITUACION_ID = rs("SITUACION_ID")
                TIPO_CALIBRACION_ID = rs("TIPO_CALIBRACION_ID")
                CALIBRADOR_ID = rs("CALIBRADOR_ID")
                CALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
                CON_CALIBRACION = rs("CON_CALIBRACION")
                CON_VERIFICACION = rs("CON_VERIFICACION")
                CON_MANTENIMIENTO = rs("CON_MANTENIMIENTO")
                REVISADO = rs("REVISADO")
                NOTAS = rs("NOTAS")
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_EQUIPO) FROM EQUIPOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_EQUIPO = 1
        Else
                ID_EQUIPO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        If ID_EQUIPO = 0 Then
            Me.CrearID
        End If
        
        consulta = "INSERT INTO EQUIPOS (" & _
                    "ID_EQUIPO, NOMBRE, CONDICIONES_AMBIENTALES, " & _
                    "PROVEEDOR_ID, FECHA_RECEPCION, FECHA_SERVICIO, ALTA_BAJA, " & _
                    "SERIE, FAMILIA_ID, MODELO, FABRICANTE, ES_NADCAP, RANGO_MEDIDA_MIN, " & _
                    "RANGO_MEDIDA_MAX, UNIDAD_ID, TOLERANCIA_MAXIMA, INCERTIDUMBRE_MAXIMA, PRECISIONN, " & _
                    "TEMPERATURA_MIN, TEMPERATURA_MAX, HUMEDAD_MIN, HUMEDAD_MAX, COND_AMBIENTALES_OTRAS, " & _
                    "RANGO_TRABAJO_MIN, RANGO_TRABAJO_MAX, LIMITACIONES_USO, PERIODICIDAD_CALIBRACION_ID, " & _
                    "FECHA_PROX_CALIBRACION, PERIODICIDAD_VERIFICACION_ID, FECHA_PROX_VERIFICACION, " & _
                    "PERIODICIDAD_MANTENIMIENTO_ID, FECHA_PROX_MANTENIMIENTO, SITUACION_ID, " & _
                    "CON_CALIBRACION, CON_VERIFICACION, CON_MANTENIMIENTO, " & _
                    "TIPO_CALIBRACION_ID, CALIBRADOR_ID, CALIBRADOR_INTERNO_ID, NOTAS, " & _
                    "REVISADO, " & _
                    "CUSERID, FTIMESTP "
        
        consulta = consulta & _
                    ") VALUES (" & _
                    ID_EQUIPO & "," & "'" & NOMBRE & "'" & "," & CONDICIONES_AMBIENTALES & "," & _
                    PROVEEDOR_ID & "," & "'" & FECHA_RECEPCION & "'" & ", " & "'" & FECHA_SERVICIO & "'" & "," & ALTA_BAJA & "," & _
                    "'" & SERIE & "'" & "," & _
                    FAMILIA_ID & ", " & "'" & MODELO & "'" & ", " & "'" & FABRICANTE & "'" & ", " & ES_NADCAP & ", " & "'" & RANGO_MEDIDA_MIN & "'" & ", " & _
                    "'" & RANGO_MEDIDA_MAX & "'" & ", " & UNIDAD_ID & ", " & "'" & TOLERANCIA_MAXIMA & "'" & ", " & "'" & INCERTIDUMBRE_MAXIMA & "'" & ", " & "'" & PRECISIONN & "'" & ", " & _
                    "'" & TEMPERATURA_MIN & "'" & ", " & "'" & TEMPERATURA_MAX & "'" & ", " & "'" & HUMEDAD_MIN & "'" & ", " & "'" & HUMEDAD_MAX & "'" & ", " & "'" & COND_AMBIENTALES_OTRAS & "'" & ", " & _
                    "'" & RANGO_TRABAJO_MIN & "'" & ", " & "'" & RANGO_TRABAJO_MAX & "'" & ", " & "'" & LIMITACIONES_USO & "'" & ", " & PERIODICIDAD_CALIBRACION_ID & ", " & _
                    "'" & FECHA_PROX_CALIBRACION & "'" & ", " & PERIODICIDAD_VERIFICACION_ID & ", " & "'" & FECHA_PROX_VERIFICACION & "'" & ", " & _
                    PERIODICIDAD_MANTENIMIENTO_ID & ", " & "'" & FECHA_PROX_MANTENIMIENTO & "'" & ", " & SITUACION_ID & ", " & _
                    CON_CALIBRACION & ", " & CON_VERIFICACION & ", " & CON_MANTENIMIENTO & ", " & _
                    TIPO_CALIBRACION_ID & ", " & CALIBRADOR_ID & ", " & CALIBRADOR_INTERNO_ID & ", " & _
                    "'" & NOTAS & "'" & ", " & REVISADO & ", " & _
                    "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP)"
        
        execute_bd consulta
        Insertar = ID_EQUIPO
        Exit Function
        
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String

        consulta = "UPDATE EQUIPOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " CONDICIONES_AMBIENTALES = " & CONDICIONES_AMBIENTALES & ", " & _
                        " PROVEEDOR_ID = " & PROVEEDOR_ID & "," & _
                        " FECHA_RECEPCION = '" & FECHA_RECEPCION & "'," & _
                        " FECHA_SERVICIO = '" & FECHA_SERVICIO & "'," & _
                        " SERIE = '" & SERIE & "'," & _
                        " ALTA_BAJA = " & ALTA_BAJA & "," & _
                        " NOTAS = '" & NOTAS & "', "

                consulta = consulta & _
                        " FAMILIA_ID = " & FAMILIA_ID & ", " & _
                        " MODELO = '" & MODELO & "', " & _
                        " FABRICANTE = '" & FABRICANTE & "', " & _
                        " ES_NADCAP = " & ES_NADCAP & ", " & _
                        " RANGO_MEDIDA_MIN = '" & RANGO_MEDIDA_MIN & "', " & _
                        " RANGO_MEDIDA_MAX = '" & RANGO_MEDIDA_MAX & "', " & _
                        " UNIDAD_ID = " & UNIDAD_ID & ", " & _
                        " TOLERANCIA_MAXIMA = '" & TOLERANCIA_MAXIMA & "', " & _
                        " INCERTIDUMBRE_MAXIMA = '" & INCERTIDUMBRE_MAXIMA & "', " & _
                        " PRECISIONN = '" & PRECISIONN & "', " & _
                        " TEMPERATURA_MIN = '" & TEMPERATURA_MIN & "', " & _
                        " TEMPERATURA_MAX = '" & TEMPERATURA_MAX & "', " & _
                        " HUMEDAD_MIN = '" & HUMEDAD_MIN & "', " & _
                        " HUMEDAD_MAX = '" & HUMEDAD_MAX & "', " & _
                        " COND_AMBIENTALES_OTRAS = '" & COND_AMBIENTALES_OTRAS & "', " & _
                        " RANGO_TRABAJO_MIN = '" & RANGO_TRABAJO_MIN & "', " & _
                        " RANGO_TRABAJO_MAX = '" & RANGO_TRABAJO_MAX & "', " & _
                        " LIMITACIONES_USO = '" & LIMITACIONES_USO & "', "
                consulta = consulta & _
                        " CON_CALIBRACION = " & CON_CALIBRACION & ", " & _
                        " CON_VERIFICACION = " & CON_VERIFICACION & ", " & _
                        " CON_MANTENIMIENTO = " & CON_MANTENIMIENTO & ", " & _
                        " PERIODICIDAD_CALIBRACION_ID = " & PERIODICIDAD_CALIBRACION_ID & ", " & _
                        " FECHA_PROX_CALIBRACION = '" & FECHA_PROX_CALIBRACION & "', " & _
                        " PERIODICIDAD_VERIFICACION_ID = " & PERIODICIDAD_VERIFICACION_ID & ", " & _
                        " FECHA_PROX_VERIFICACION = '" & FECHA_PROX_VERIFICACION & "', " & _
                        " PERIODICIDAD_MANTENIMIENTO_ID = " & PERIODICIDAD_MANTENIMIENTO_ID & ", " & _
                        " FECHA_PROX_MANTENIMIENTO = '" & FECHA_PROX_MANTENIMIENTO & "', " & _
                        " SITUACION_ID = " & SITUACION_ID & ", " & _
                        " TIPO_CALIBRACION_ID = " & TIPO_CALIBRACION_ID & ", " & _
                        " CALIBRADOR_ID = " & CALIBRADOR_ID & ", " & _
                        " CALIBRADOR_INTERNO_ID = " & CALIBRADOR_INTERNO_ID & ", " & _
                        " REVISADO = " & REVISADO & ", " & _
                        " CUSERID = '" & USUARIO.getUSUARIO & "', " & _
                        " FTIMESTP = CURRENT_TIMESTAMP "
                consulta = consulta & _
                        " WHERE ID_EQUIPO = " & ID
        
        execute_bd consulta
        Modificar = True
        Exit Function
        
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        ' se eliminan calibración, verificación y mantenimiento
        execute_bd "DELETE FROM EQUIPOS_CALIBRACION WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQUIPOS_VERIFICACION WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQUIPOS_MANTENIMIENTO WHERE EQUIPO_ID = " & ID
        
        ' se eliminan los planes de mantenimiento que tenga asignados
        execute_bd "DELETE FROM EQUIPOS_PLANES_MTO_ASIGNADOS WHERE EQUIPO_ID = " & ID
        
        ' se eliminan los documentos, accesorios y normas
        execute_bd "DELETE FROM EQUIPOS_DOCUMENTACION WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQUIPOS_ACCESORIOS WHERE EQUIPO_ID = " & ID
        execute_bd "DELETE FROM EQUIPOS_NORMAS WHERE EQUIPO_ID = " & ID
        
        ' finalmente se borra el equipo
        execute_bd "DELETE FROM EQUIPOS WHERE ID_EQUIPO = " & ID
        Eliminar = True
        Exit Function
        
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos)", vbCritical, Err.Description
End Function

Public Function Listado(fdes As String, fserie As String, fequipo As String, fnombre As String, ffamilia As String, fsituacion As String, fproveedor As String, baja As Integer, op1 As Integer, op2 As Integer, op3 As Integer, op4 As Integer, opRev0 As Integer, opRev1 As Integer, opRev2 As Integer) As ADODB.RecordSet
        Dim consulta As String
        Dim eq As String
        Dim FILTRO As String
        Dim strRevision As String
        
        If fequipo <> "" And IsNumeric(fequipo) Then
            eq = " AND A.ID_EQUIPO = " & fequipo
        End If
        If op1 <> 0 Then ' Todos
            FILTRO = ""
        End If
        If op2 <> 0 Then ' Calibracion
            FILTRO = " AND A.CON_CALIBRACION = 1"
        End If
        If op3 <> 0 Then  ' Verificacion
            FILTRO = " AND A.CON_VERIFICACION = 1"
        End If
        If op4 <> 0 Then  ' Verificacion
            FILTRO = " AND A.CON_MANTENIMIENTO = 1"
        End If
        
        If opRev0 <> 0 Then ' todos
            strRevision = ""
        End If
        If opRev1 <> 0 Then ' revisados
            strRevision = " AND A.REVISADO = 1 "
        End If
        If opRev2 <> 0 Then  ' sin revisar
            strRevision = " AND A.REVISADO = 0 "
        End If
        
        consulta = "SELECT A.ID_EQUIPO, A.NOMBRE, A.SERIE, C.NOMBRE, " & _
                   "       B.DESCRIPCION, D.DESCRIPCION, E.DESCRIPCION " & _
                   "  FROM EQUIPOS A, DECODIFICADORA B, PROVEEDORES C, DECODIFICADORA D, DECODIFICADORA E " & _
                   " WHERE (A.SITUACION_ID = B.VALOR AND B.CODIGO = " & decodificadora.EQ_SITUACIONES & ")" & _
                   "   AND (A.PROVEEDOR_ID = C.ID_PROVEEDOR) " & _
                   "   AND (A.FAMILIA_ID = D.VALOR AND D.CODIGO = " & decodificadora.EQ_FAMILIAS & ")" & _
                   "   AND (A.CONDICIONES_AMBIENTALES = E.VALOR AND E.CODIGO = " & decodificadora.EQ_SINO & ")" & _
                   "   AND A.SERIE LIKE '%" & fserie & "%' " & _
                   "   AND A.NOMBRE LIKE '%" & fnombre & "%' " & _
                   "   AND D.DESCRIPCION LIKE '%" & ffamilia & "%' " & _
                   "   AND B.DESCRIPCION LIKE '%" & fsituacion & "%' " & _
                   "   AND C.NOMBRE LIKE '%" & fproveedor & "%' " & _
                   eq & FILTRO & strRevision & _
                   "   AND A.ALTA_BAJA = " & CInt(baja) & _
                   " ORDER BY A.NOMBRE"
        
        Set Listado = datos_bd(consulta)
End Function


Public Function ListadoObjetos() As clsGenericCollection
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        Dim objEq As clsEquipo
        Dim objCol As clsGenericCollection
        
        consulta = "SELECT A.ID_EQUIPO, A.NOMBRE, A.SERIE, C.NOMBRE, " & _
                   "       B.DESCRIPCION, D.DESCRIPCION, E.DESCRIPCION " & _
                   "  FROM EQUIPOS A, DECODIFICADORA B, PROVEEDORES C, DECODIFICADORA D, DECODIFICADORA E " & _
                   " WHERE (A.SITUACION_ID = B.VALOR AND B.CODIGO = " & decodificadora.EQ_SITUACIONES & ")" & _
                   "   AND (A.PROVEEDOR_ID = C.ID_PROVEEDOR) " & _
                   "   AND (A.FAMILIA_ID = D.VALOR AND D.CODIGO = " & decodificadora.EQ_FAMILIAS & ")" & _
                   "   AND (A.CONDICIONES_AMBIENTALES = E.VALOR AND E.CODIGO = " & decodificadora.EQ_SINO & ")" & _
                   "   AND A.SERIE LIKE '%" & "" & "%' " & _
                   "   AND A.NOMBRE LIKE '%" & "" & "%' " & _
                   "   AND D.DESCRIPCION LIKE '%" & "" & "%' " & _
                   "   AND B.DESCRIPCION LIKE '%" & "" & "%' " & _
                   "   AND C.NOMBRE LIKE '%" & "" & "%' " & _
                   "" & "" & "" & _
                   "   AND A.ALTA_BAJA = " & CInt(0) & _
                   " ORDER BY A.NOMBRE"
        
        Set rs = datos_bd(consulta)
        
        Set objCol = New clsGenericCollection
        
        If rs.RecordCount = 0 Then
            Set ListadoObjetos = objCol
            Exit Function
        End If
        
        
        rs.MoveFirst
        While Not rs.EOF
            Set objEq = New clsEquipo
            objEq.ID_EQUIPO = rs("ID_EQUIPO")
            objEq.NOMBRE = rs("NOMBRE")
            objEq.CONDICIONES_AMBIENTALES = rs("CONDICIONES_AMBIENTALES")
            objEq.PROVEEDOR_ID = rs("PROVEEDOR_ID")
            objEq.FECHA_RECEPCION = rs("FECHA_RECEPCION")
            objEq.FECHA_SERVICIO = rs("FECHA_SERVICIO")
            objEq.ALTA_BAJA = rs("ALTA_BAJA")
            objEq.SERIE = rs("SERIE")
            objEq.FAMILIA_ID = rs("FAMILIA_ID")
            objEq.MODELO = rs("MODELO")
            objEq.FABRICANTE = rs("FABRICANTE")
            objEq.ES_NADCAP = rs("ES_NADCAP")
            objEq.RANGO_MEDIDA_MIN = rs("RANGO_MEDIDA_MIN")
            objEq.RANGO_MEDIDA_MAX = rs("RANGO_MEDIDA_MAX")
            objEq.UNIDAD_ID = rs("UNIDAD_ID")
            objEq.TOLERANCIA_MAXIMA = rs("TOLERANCIA_MAXIMA")
            objEq.INCERTIDUMBRE_MAXIMA = rs("INCERTIDUMBRE_MAXIMA")
            objEq.PRECISIONN = rs("PRECISIONN")
            objEq.TEMPERATURA_MIN = rs("TEMPERATURA_MIN")
            objEq.TEMPERATURA_MAX = rs("TEMPERATURA_MAX")
            objEq.HUMEDAD_MIN = rs("HUMEDAD_MIN")
            objEq.HUMEDAD_MAX = rs("HUMEDAD_MAX")
            objEq.COND_AMBIENTALES_OTRAS = rs("COND_AMBIENTALES_OTRAS")
            objEq.RANGO_TRABAJO_MIN = rs("RANGO_TRABAJO_MIN")
            objEq.RANGO_TRABAJO_MAX = rs("RANGO_TRABAJO_MAX")
            objEq.LIMITACIONES_USO = rs("LIMITACIONES_USO")
            objEq.PERIODICIDAD_CALIBRACION_ID = rs("PERIODICIDAD_CALIBRACION_ID")
            objEq.FECHA_PROX_CALIBRACION = rs("FECHA_PROX_CALIBRACION")
            objEq.PERIODICIDAD_VERIFICACION_ID = rs("PERIODICIDAD_VERIFICACION_ID")
            objEq.FECHA_PROX_VERIFICACION = rs("FECHA_PROX_VERIFICACION")
            objEq.PERIODICIDAD_MANTENIMIENTO_ID = rs("PERIODICIDAD_MANTENIMIENTO_ID")
            objEq.FECHA_PROX_MANTENIMIENTO = rs("FECHA_PROX_MANTENIMIENTO")
            objEq.SITUACION_ID = rs("SITUACION_ID")
            objEq.TIPO_CALIBRACION_ID = rs("TIPO_CALIBRACION_ID")
            objEq.CALIBRADOR_ID = rs("CALIBRADOR_ID")
            objEq.CALIBRADOR_INTERNO_ID = rs("CALIBRADOR_INTERNO_ID")
            objEq.CON_CALIBRACION = rs("CON_CALIBRACION")
            objEq.CON_VERIFICACION = rs("CON_VERIFICACION")
            objEq.CON_MANTENIMIENTO = rs("CON_MANTENIMIENTO")
            objEq.REVISADO = rs("REVISADO")
            objEq.NOTAS = rs("NOTAS")
            
            Call objCol.Add(objEq, objEq.ID_EQUIPO)
            
            rs.MoveNext
        Wend
        
        Set ListadoObjetos = objCol
        
End Function


Public Function Listado_completo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT A.ID_EQUIPO,A.NOMBRE,A.SERIE,B.DESCRIPCION,A.SITUACION " & _
                   "  FROM EQUIPOS A, EQUIPOS_PROVEEDORES B " & _
                   " WHERE A.PROVEEDOR_ID = B.ID_PROVEEDOR " & _
                   " ORDER BY A.NOMBRE"
        Set Listado_completo = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_EQUIPO, NOMBRE FROM EQUIPOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function imprimir(Equipo As Long) As Boolean
        With frmReport
            .iniciar
            .informe = "rptequipo_ficha"
            .criterio = "{dec_fam_eq.CODIGO}=" & decodificadora.EQ_FAMILIAS & " and {dec_situaciones_eq.CODIGO}=" & decodificadora.EQ_SITUACIONES & _
                        " and {dec_sino.CODIGO}=" & decodificadora.EQ_SINO & " and {dec_periodicidad_mto.CODIGO}= " & decodificadora.EQ_periodicidad & _
                        " and {dec_periodicidad_ver.CODIGO}=" & decodificadora.EQ_periodicidad & " and {dec_periodicidad_cal.CODIGO}= " & decodificadora.EQ_periodicidad & _
                        " and {dec_sino_esnadcap.CODIGO}=" & decodificadora.EQ_SINO & " and {dec_tipo_calibracion.CODIGO}= " & decodificadora.EQ_TIPO_CALIBRACION & _
                        " AND {equipos.ID_EQUIPO}=" & Equipo
            
            .imprimir = False
            .generar
            .Visible = True
        End With
End Function

Public Function Listado_calibracion_avisos(fdes As String, fserie As String, fequipo As String, fnombre As String, ffamilia As String, fsituacion As String, fproveedor As String, baja As Integer, opRev0 As Integer, opRev1 As Integer, opRev2 As Integer) As ADODB.RecordSet
        Dim consulta As String
        Dim eq As String
        Dim strRevision As String
        Dim strCondicionEmpleado As String
        
        If fequipo <> "" And IsNumeric(fequipo) Then
            eq = " AND E.ID_EQUIPO = " & fequipo
        End If
        
        If opRev0 <> 0 Then ' todos
            strRevision = ""
        End If
        If opRev1 <> 0 Then ' revisados
            strRevision = " AND E.REVISADO = 1 "
        End If
        If opRev2 <> 0 Then  ' sin revisar
            strRevision = " AND E.REVISADO = 0 "
        End If
        
        If usuario_logado_ve_externos() Then
            strCondicionEmpleado = "   AND (EC.CALIBRADOR_INTERNO_ID = " & USUARIO.getID_EMPLEADO & " " & _
                                   "     OR EC.MODALIDAD_ID = 2) "
        Else
            strCondicionEmpleado = "   AND (EC.CALIBRADOR_INTERNO_ID = " & USUARIO.getID_EMPLEADO & ") "
        End If
        
        consulta = "SELECT E.ID_EQUIPO, E.NOMBRE, E.SERIE, PROVEEDORES.NOMBRE " & _
                   "     , SITUACIONES.DESCRIPCION SITUACION, FAMILIAS.DESCRIPCION FAMILIA, SINO.DESCRIPCION CONDICIONES_AMBIENTALES" & _
                   "  FROM EQUIPOS E, DECODIFICADORA SITUACIONES, PROVEEDORES " & _
                   "     , DECODIFICADORA FAMILIAS, DECODIFICADORA SINO " & _
                   "     , EQUIPOS_CALIBRACION EC " & _
                   "     , DECODIFICADORA PERIODICIDAD " & _
                   " WHERE (E.SITUACION_ID = SITUACIONES.VALOR AND SITUACIONES.codigo = " & decodificadora.EQ_SITUACIONES & ") " & _
                   "   AND (E.PROVEEDOR_ID = PROVEEDORES.ID_PROVEEDOR) " & _
                   "   AND (E.FAMILIA_ID = FAMILIAS.VALOR AND FAMILIAS.CODIGO = " & decodificadora.EQ_FAMILIAS & ") " & _
                   "   AND (E.CONDICIONES_AMBIENTALES = SINO.VALOR AND SINO.CODIGO = " & decodificadora.EQ_SINO & ") " & _
                   "   AND E.NOMBRE LIKE '%" & fnombre & "%' " & _
                   "   AND E.SERIE LIKE '%" & fserie & "%' " & _
                   "   AND PROVEEDORES.NOMBRE LIKE '%" & fproveedor & "%' " & _
                   "   AND SITUACIONES.DESCRIPCION LIKE '%" & fsituacion & "%' " & _
                   "   AND FAMILIAS.DESCRIPCION LIKE '%" & ffamilia & "%' " & _
                   "   AND E.ID_EQUIPO = EC.EQUIPO_ID " & _
                   "   AND E.ALTA_BAJA = " & CInt(baja) & _
                   eq & strRevision & _
                   "   AND E.CON_CALIBRACION = 1 " & _
                   "   AND (EC.PERIODICIDAD_ID = PERIODICIDAD.VALOR AND PERIODICIDAD.CODIGO = " & decodificadora.EQ_periodicidad & ") " & _
                   "   AND PERIODICIDAD.DESCRIPCION NOT LIKE '%*%' " & _
                   "   AND EC.FECHA_PROXIMA < '" & Format(Date, "yyyy-mm-dd") & "' " & _
                   strCondicionEmpleado & _
                   " ORDER BY E.ID_EQUIPO "
        
        Set Listado_calibracion_avisos = datos_bd(consulta)
End Function

Public Function Listado_verificacion_avisos(fdes As String, fserie As String, fequipo As String, fnombre As String, ffamilia As String, fsituacion As String, fproveedor As String, baja As Integer, opRev0 As Integer, opRev1 As Integer, opRev2 As Integer) As ADODB.RecordSet
        Dim consulta As String
        Dim eq As String
        Dim strRevision As String
        Dim strCondicionEmpleado As String
        
        If fequipo <> "" And IsNumeric(fequipo) Then
            eq = " AND E.ID_EQUIPO = " & fequipo
        End If
        
        If opRev0 <> 0 Then ' todos
            strRevision = ""
        End If
        If opRev1 <> 0 Then ' revisados
            strRevision = " AND E.REVISADO = 1 "
        End If
        If opRev2 <> 0 Then  ' sin revisar
            strRevision = " AND E.REVISADO = 0 "
        End If
        
        If usuario_logado_ve_externos() Then
            strCondicionEmpleado = "   AND (EV.VERIFICADOR_INTERNO_ID = " & USUARIO.getID_EMPLEADO & " " & _
                                   "     OR EV.MODALIDAD_ID = 2) "
        Else
            strCondicionEmpleado = "   AND (EV.VERIFICADOR_INTERNO_ID = " & USUARIO.getID_EMPLEADO & ") "
        End If
        
        consulta = "SELECT E.ID_EQUIPO, E.NOMBRE, E.SERIE, PROVEEDORES.NOMBRE " & _
                   "     , SITUACIONES.DESCRIPCION SITUACION, FAMILIAS.DESCRIPCION FAMILIA, SINO.DESCRIPCION CONDICIONES_AMBIENTALES " & _
                   "  FROM EQUIPOS E, DECODIFICADORA SITUACIONES, PROVEEDORES " & _
                   "     , DECODIFICADORA FAMILIAS, DECODIFICADORA SINO " & _
                   "     , EQUIPOS_VERIFICACION EV " & _
                   "     , DECODIFICADORA PERIODICIDAD " & _
                   " WHERE (E.SITUACION_ID = SITUACIONES.VALOR AND SITUACIONES.codigo = " & decodificadora.EQ_SITUACIONES & ") " & _
                   "   AND (E.PROVEEDOR_ID = PROVEEDORES.ID_PROVEEDOR) " & _
                   "   AND (E.FAMILIA_ID = FAMILIAS.VALOR AND FAMILIAS.CODIGO = " & decodificadora.EQ_FAMILIAS & ") " & _
                   "   AND (E.CONDICIONES_AMBIENTALES = SINO.VALOR AND SINO.CODIGO = " & decodificadora.EQ_SINO & ") " & _
                   "   AND E.NOMBRE LIKE '%" & fnombre & "%' " & _
                   "   AND PROVEEDORES.NOMBRE LIKE '%" & fproveedor & "%' " & _
                   "   AND E.SERIE LIKE '%" & fserie & "%' " & _
                   "   AND SITUACIONES.DESCRIPCION LIKE '%" & fsituacion & "%' " & _
                   "   AND FAMILIAS.DESCRIPCION LIKE '%" & ffamilia & "%' " & _
                   "   AND E.ID_EQUIPO = EV.EQUIPO_ID " & _
                   "   AND E.ALTA_BAJA = " & CInt(baja) & _
                   eq & strRevision & _
                   "   AND E.CON_VERIFICACION = 1 " & _
                   "   AND (EV.PERIODICIDAD_ID = PERIODICIDAD.VALOR AND PERIODICIDAD.CODIGO = " & decodificadora.EQ_periodicidad & ") " & _
                   "   AND PERIODICIDAD.DESCRIPCION NOT LIKE '%*%' " & _
                   "   AND EV.FECHA_PROXIMA < '" & Format(Date, "yyyy-mm-dd") & "' " & _
                   strCondicionEmpleado & _
                   "ORDER BY E.ID_EQUIPO "
        
        Set Listado_verificacion_avisos = datos_bd(consulta)
End Function

Public Function Listado_mantenimiento_avisos(fdes As String, fserie As String, fequipo As String, fnombre As String, ffamilia As String, fsituacion As String, fproveedor As String, baja As Integer, opRev0 As Integer, opRev1 As Integer, opRev2 As Integer) As ADODB.RecordSet
        Dim consulta As String
        Dim eq As String
        Dim strRevision As String
        Dim strCondicionEmpleado As String
        
        If fequipo <> "" And IsNumeric(fequipo) Then
            eq = " AND E.ID_EQUIPO = " & fequipo
        End If
        
        If opRev0 <> 0 Then ' todos
            strRevision = ""
        End If
        If opRev1 <> 0 Then ' revisados
            strRevision = " AND E.REVISADO = 1 "
        End If
        If opRev2 <> 0 Then  ' sin revisar
            strRevision = " AND E.REVISADO = 0 "
        End If
        
        If usuario_logado_ve_externos() Then
            strCondicionEmpleado = "   AND ((EM.SEMANAL_MANTENIMIENTO = 1 " & _
                                   "        AND EM.SEMANAL_FECHA < '" & Format(Date, "yyyy-mm-dd") & "' " & _
                                   "        AND (EM.SEMANAL_RESPONSABLE_INTERNO_ID = " & USUARIO.getID_EMPLEADO & _
                                   "               OR EM.SEMANAL_MODALIDAD_ID = 2)) " & _
                                   "    OR (EM.MENSUAL_MANTENIMIENTO = 1 " & _
                                   "        AND EM.MENSUAL_FECHA < '" & Format(Date, "yyyy-mm-dd") & "' " & _
                                   "        AND (EM.MENSUAL_RESPONSABLE_INTERNO_ID = " & USUARIO.getID_EMPLEADO & _
                                   "               OR EM.MENSUAL_MODALIDAD_ID = 2))) "
        Else
            strCondicionEmpleado = "   AND ((EM.SEMANAL_MANTENIMIENTO = 1 " & _
                                   "        AND EM.SEMANAL_FECHA < '" & Format(Date, "yyyy-mm-dd") & "' " & _
                                   "        AND EM.SEMANAL_RESPONSABLE_INTERNO_ID = " & USUARIO.getID_EMPLEADO & ") " & _
                                   "    OR (EM.MENSUAL_MANTENIMIENTO = 1 " & _
                                   "        AND EM.MENSUAL_FECHA < '" & Format(Date, "yyyy-mm-dd") & "' " & _
                                   "        AND EM.MENSUAL_RESPONSABLE_INTERNO_ID = " & USUARIO.getID_EMPLEADO & ")) "
        End If
        
        consulta = "SELECT E.ID_EQUIPO, E.NOMBRE, E.SERIE, PROVEEDORES.NOMBRE " & _
                   "     , SITUACIONES.DESCRIPCION SITUACION, FAMILIAS.DESCRIPCION FAMILIA " & _
                   "     , SINO.DESCRIPCION CONDICIONES_AMBIENTALES " & _
                   "  FROM EQUIPOS E, DECODIFICADORA SITUACIONES, PROVEEDORES " & _
                   "     , DECODIFICADORA FAMILIAS, DECODIFICADORA SINO " & _
                   "     , EQUIPOS_MANTENIMIENTO EM " & _
                   " WHERE E.ID_EQUIPO = EM.EQUIPO_ID " & _
                   "   AND (E.SITUACION_ID = SITUACIONES.VALOR AND SITUACIONES.codigo = " & decodificadora.EQ_SITUACIONES & ") " & _
                   "   AND (E.PROVEEDOR_ID = PROVEEDORES.ID_PROVEEDOR) " & _
                   "   AND (E.FAMILIA_ID = FAMILIAS.VALOR AND FAMILIAS.CODIGO = " & decodificadora.EQ_FAMILIAS & ") " & _
                   "   AND (E.CONDICIONES_AMBIENTALES = SINO.VALOR AND SINO.CODIGO = " & decodificadora.EQ_SINO & ") " & _
                   "   AND E.ALTA_BAJA = " & CInt(baja) & _
                   eq & strRevision & _
                   "   AND E.CON_MANTENIMIENTO = 1 " & _
                   "   AND PROVEEDORES.NOMBRE LIKE '%" & fproveedor & "%' " & _
                   "   AND E.NOMBRE LIKE '%" & fnombre & "%' " & _
                   "   AND E.SERIE LIKE '%" & fserie & "%' " & _
                   "   AND SITUACIONES.DESCRIPCION LIKE '%" & fsituacion & "%' " & _
                   "   AND FAMILIAS.DESCRIPCION LIKE '%" & ffamilia & "%' " & _
                   strCondicionEmpleado & _
                   " ORDER BY E.ID_EQUIPO "
        
        Set Listado_mantenimiento_avisos = datos_bd(consulta)
End Function

' función que devuelve si el usuario logado está entre los
' que verán los equipos con cvm externo (se encuentran en parametros)
Private Function usuario_logado_ve_externos() As Boolean
    Dim oParametro As New clsParametros
    Dim strUsuarios() As String
    Dim i As Long
    
    usuario_logado_ve_externos = False
    
    oParametro.Carga parametros.USUARIOS_VEN_EQUIPOS_CVM_EXTERNO, ""
    strUsuarios = Split(oParametro.getVALOR, ",")
    For i = 0 To UBound(strUsuarios)
        If USUARIO.getID_EMPLEADO = Trim(strUsuarios(i)) Then
            usuario_logado_ve_externos = True
        End If
    Next i
    
    Set oParametro = Nothing
End Function
Public Function llenar_combo(combo As miCombo, FK As Long, FORMULARIO As Form, FILTRO As String)
    With combo
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "EQUIPOS"
        .setDESCRIPCION = "Equipos"
        .setPK = "ID_EQUIPO"
        .setCAMPO = "concat(Nombre,' (Nº ',CAST(ID_EQUIPO AS CHAR),')')"
        .setMUESTRA_DETALLE = True
        .setFILTRO = " ALTA_BAJA=0 "
        Set .FORMULARIO = FORMULARIO
    End With
End Function


Public Property Get ID_EQUIPO() As Long

On Error GoTo ID_EQUIPO_Error

    ID_EQUIPO = mvarlngID_EQUIPO

On Error GoTo 0
    Exit Property
ID_EQUIPO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ID_EQUIPO of Módulo de clase clsEquipo"

End Property

Public Property Let ID_EQUIPO(ByVal lngID_EQUIPO As Long)

On Error GoTo ID_EQUIPO_Error

    mvarlngID_EQUIPO = lngID_EQUIPO

On Error GoTo 0
    Exit Property
ID_EQUIPO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ID_EQUIPO of Módulo de clase clsEquipo"

End Property

Public Property Get NOMBRE() As String

On Error GoTo NOMBRE_Error

    NOMBRE = mvarstrNOMBRE

On Error GoTo 0
    Exit Property
NOMBRE_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NOMBRE of Módulo de clase clsEquipo"

End Property

Public Property Let NOMBRE(ByVal strNOMBRE As String)

On Error GoTo NOMBRE_Error

    mvarstrNOMBRE = strNOMBRE

On Error GoTo 0
    Exit Property
NOMBRE_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NOMBRE of Módulo de clase clsEquipo"

End Property

Public Property Get CONDICIONES_AMBIENTALES() As Long

On Error GoTo CONDICIONES_AMBIENTALES_Error

    CONDICIONES_AMBIENTALES = mvarlngCONDICIONES_AMBIENTALES

On Error GoTo 0
    Exit Property
CONDICIONES_AMBIENTALES_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CONDICIONES_AMBIENTALES of Módulo de clase clsEquipo"

End Property

Public Property Let CONDICIONES_AMBIENTALES(ByVal lngCONDICIONES_AMBIENTALES As Long)

On Error GoTo CONDICIONES_AMBIENTALES_Error

    mvarlngCONDICIONES_AMBIENTALES = lngCONDICIONES_AMBIENTALES

On Error GoTo 0
    Exit Property
CONDICIONES_AMBIENTALES_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CONDICIONES_AMBIENTALES of Módulo de clase clsEquipo"

End Property

Public Property Get PROVEEDOR_ID() As Integer

On Error GoTo PROVEEDOR_ID_Error

    PROVEEDOR_ID = mvarintPROVEEDOR_ID

On Error GoTo 0
    Exit Property
PROVEEDOR_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PROVEEDOR_ID of Módulo de clase clsEquipo"

End Property

Public Property Let PROVEEDOR_ID(ByVal intPROVEEDOR_ID As Integer)

On Error GoTo PROVEEDOR_ID_Error

    mvarintPROVEEDOR_ID = intPROVEEDOR_ID

On Error GoTo 0
    Exit Property
PROVEEDOR_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PROVEEDOR_ID of Módulo de clase clsEquipo"

End Property

Public Property Get FECHA_RECEPCION() As String

On Error GoTo FECHA_RECEPCION_Error

    FECHA_RECEPCION = mvarstrFECHA_RECEPCION

On Error GoTo 0
    Exit Property
FECHA_RECEPCION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_RECEPCION of Módulo de clase clsEquipo"

End Property

Public Property Let FECHA_RECEPCION(ByVal strFECHA_RECEPCION As String)

On Error GoTo FECHA_RECEPCION_Error

    mvarstrFECHA_RECEPCION = strFECHA_RECEPCION

On Error GoTo 0
    Exit Property
FECHA_RECEPCION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_RECEPCION of Módulo de clase clsEquipo"

End Property

Public Property Get FECHA_SERVICIO() As String

On Error GoTo FECHA_SERVICIO_Error

    FECHA_SERVICIO = mvarstrFECHA_SERVICIO

On Error GoTo 0
    Exit Property
FECHA_SERVICIO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_SERVICIO of Módulo de clase clsEquipo"

End Property

Public Property Let FECHA_SERVICIO(ByVal strFECHA_SERVICIO As String)

On Error GoTo FECHA_SERVICIO_Error

    mvarstrFECHA_SERVICIO = strFECHA_SERVICIO

On Error GoTo 0
    Exit Property
FECHA_SERVICIO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_SERVICIO of Módulo de clase clsEquipo"

End Property

Public Property Get ALTA_BAJA() As Integer

On Error GoTo ALTA_BAJA_Error

    ALTA_BAJA = mvarintALTA_BAJA

On Error GoTo 0
    Exit Property
ALTA_BAJA_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ALTA_BAJA of Módulo de clase clsEquipo"

End Property

Public Property Let ALTA_BAJA(ByVal intALTA_BAJA As Integer)

On Error GoTo ALTA_BAJA_Error

    mvarintALTA_BAJA = intALTA_BAJA

On Error GoTo 0
    Exit Property
ALTA_BAJA_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ALTA_BAJA of Módulo de clase clsEquipo"

End Property

Public Property Get SERIE() As String

On Error GoTo SERIE_Error

    SERIE = mvarstrSERIE

On Error GoTo 0
    Exit Property
SERIE_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SERIE of Módulo de clase clsEquipo"

End Property

Public Property Let SERIE(ByVal strSERIE As String)

On Error GoTo SERIE_Error

    mvarstrSERIE = strSERIE

On Error GoTo 0
    Exit Property
SERIE_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SERIE of Módulo de clase clsEquipo"

End Property

Public Property Get NOTAS() As String

On Error GoTo NOTAS_Error

    NOTAS = mvarstrNOTAS

On Error GoTo 0
    Exit Property
NOTAS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NOTAS of Módulo de clase clsEquipo"

End Property

Public Property Let NOTAS(ByVal strNOTAS As String)

On Error GoTo NOTAS_Error

    mvarstrNOTAS = strNOTAS

On Error GoTo 0
    Exit Property
NOTAS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure NOTAS of Módulo de clase clsEquipo"

End Property

Public Property Get FAMILIA_ID() As Integer

On Error GoTo FAMILIA_ID_Error

    FAMILIA_ID = mvarintFAMILIA_ID

On Error GoTo 0
    Exit Property
FAMILIA_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FAMILIA_ID of Módulo de clase clsEquipo"

End Property

Public Property Let FAMILIA_ID(ByVal intFAMILIA_ID As Integer)

On Error GoTo FAMILIA_ID_Error

    mvarintFAMILIA_ID = intFAMILIA_ID

On Error GoTo 0
    Exit Property
FAMILIA_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FAMILIA_ID of Módulo de clase clsEquipo"

End Property

Public Property Get MODELO() As String

On Error GoTo MODELO_Error

    MODELO = mvarstrMODELO

On Error GoTo 0
    Exit Property
MODELO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MODELO of Módulo de clase clsEquipo"

End Property

Public Property Let MODELO(ByVal strMODELO As String)

On Error GoTo MODELO_Error

    mvarstrMODELO = strMODELO

On Error GoTo 0
    Exit Property
MODELO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure MODELO of Módulo de clase clsEquipo"

End Property

Public Property Get FABRICANTE() As String

On Error GoTo FABRICANTE_Error

    FABRICANTE = mvarstrFABRICANTE

On Error GoTo 0
    Exit Property
FABRICANTE_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FABRICANTE of Módulo de clase clsEquipo"

End Property

Public Property Let FABRICANTE(ByVal strFABRICANTE As String)

On Error GoTo FABRICANTE_Error

    mvarstrFABRICANTE = strFABRICANTE

On Error GoTo 0
    Exit Property
FABRICANTE_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FABRICANTE of Módulo de clase clsEquipo"

End Property

Public Property Get ES_NADCAP() As Integer

On Error GoTo ES_NADCAP_Error

    ES_NADCAP = mvarintES_NADCAP

On Error GoTo 0
    Exit Property
ES_NADCAP_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ES_NADCAP of Módulo de clase clsEquipo"

End Property

Public Property Let ES_NADCAP(ByVal intES_NADCAP As Integer)

On Error GoTo ES_NADCAP_Error

    mvarintES_NADCAP = intES_NADCAP

On Error GoTo 0
    Exit Property
ES_NADCAP_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ES_NADCAP of Módulo de clase clsEquipo"

End Property

Public Property Get RANGO_MEDIDA_MIN() As String

On Error GoTo RANGO_MEDIDA_MIN_Error

    RANGO_MEDIDA_MIN = mvarstrRANGO_MEDIDA_MIN

On Error GoTo 0
    Exit Property
RANGO_MEDIDA_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_MEDIDA_MIN of Módulo de clase clsEquipo"

End Property

Public Property Let RANGO_MEDIDA_MIN(ByVal strRANGO_MEDIDA_MIN As String)

On Error GoTo RANGO_MEDIDA_MIN_Error

    mvarstrRANGO_MEDIDA_MIN = strRANGO_MEDIDA_MIN

On Error GoTo 0
    Exit Property
RANGO_MEDIDA_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_MEDIDA_MIN of Módulo de clase clsEquipo"

End Property

Public Property Get RANGO_MEDIDA_MAX() As String

On Error GoTo RANGO_MEDIDA_MAX_Error

    RANGO_MEDIDA_MAX = mvarstrRANGO_MEDIDA_MAX

On Error GoTo 0
    Exit Property
RANGO_MEDIDA_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_MEDIDA_MAX of Módulo de clase clsEquipo"

End Property

Public Property Let RANGO_MEDIDA_MAX(ByVal strRANGO_MEDIDA_MAX As String)

On Error GoTo RANGO_MEDIDA_MAX_Error

    mvarstrRANGO_MEDIDA_MAX = strRANGO_MEDIDA_MAX

On Error GoTo 0
    Exit Property
RANGO_MEDIDA_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_MEDIDA_MAX of Módulo de clase clsEquipo"

End Property

Public Property Get UNIDAD_ID() As Integer

On Error GoTo UNIDAD_ID_Error

    UNIDAD_ID = mvarintUNIDAD_ID

On Error GoTo 0
    Exit Property
UNIDAD_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UNIDAD_ID of Módulo de clase clsEquipo"

End Property

Public Property Let UNIDAD_ID(ByVal intUNIDAD_ID As Integer)

On Error GoTo UNIDAD_ID_Error

    mvarintUNIDAD_ID = intUNIDAD_ID

On Error GoTo 0
    Exit Property
UNIDAD_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure UNIDAD_ID of Módulo de clase clsEquipo"

End Property

Public Property Get TOLERANCIA_MAXIMA() As String

On Error GoTo TOLERANCIA_MAXIMA_Error

    TOLERANCIA_MAXIMA = mvarstrTOLERANCIA_MAXIMA

On Error GoTo 0
    Exit Property
TOLERANCIA_MAXIMA_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TOLERANCIA_MAXIMA of Módulo de clase clsEquipo"

End Property

Public Property Let TOLERANCIA_MAXIMA(ByVal strTOLERANCIA_MAXIMA As String)

On Error GoTo TOLERANCIA_MAXIMA_Error

    mvarstrTOLERANCIA_MAXIMA = strTOLERANCIA_MAXIMA

On Error GoTo 0
    Exit Property
TOLERANCIA_MAXIMA_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TOLERANCIA_MAXIMA of Módulo de clase clsEquipo"

End Property

Public Property Get INCERTIDUMBRE_MAXIMA() As String

On Error GoTo INCERTIDUMBRE_MAXIMA_Error

    INCERTIDUMBRE_MAXIMA = mvarstrINCERTIDUMBRE_MAXIMA

On Error GoTo 0
    Exit Property
INCERTIDUMBRE_MAXIMA_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure INCERTIDUMBRE_MAXIMA of Módulo de clase clsEquipo"

End Property

Public Property Let INCERTIDUMBRE_MAXIMA(ByVal strINCERTIDUMBRE_MAXIMA As String)

On Error GoTo INCERTIDUMBRE_MAXIMA_Error

    mvarstrINCERTIDUMBRE_MAXIMA = strINCERTIDUMBRE_MAXIMA

On Error GoTo 0
    Exit Property
INCERTIDUMBRE_MAXIMA_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure INCERTIDUMBRE_MAXIMA of Módulo de clase clsEquipo"

End Property

Public Property Get PRECISIONN() As String

On Error GoTo PRECISIONN_Error

    PRECISIONN = mvarstrPRECISIONN

On Error GoTo 0
    Exit Property
PRECISIONN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PRECISIONN of Módulo de clase clsEquipo"

End Property

Public Property Let PRECISIONN(ByVal strPRECISIONN As String)

On Error GoTo PRECISIONN_Error

    mvarstrPRECISIONN = strPRECISIONN

On Error GoTo 0
    Exit Property
PRECISIONN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PRECISIONN of Módulo de clase clsEquipo"

End Property

Public Property Get TEMPERATURA_MIN() As String

On Error GoTo TEMPERATURA_MIN_Error

    TEMPERATURA_MIN = mvarstrTEMPERATURA_MIN

On Error GoTo 0
    Exit Property
TEMPERATURA_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TEMPERATURA_MIN of Módulo de clase clsEquipo"

End Property

Public Property Let TEMPERATURA_MIN(ByVal strTEMPERATURA_MIN As String)

On Error GoTo TEMPERATURA_MIN_Error

    mvarstrTEMPERATURA_MIN = strTEMPERATURA_MIN

On Error GoTo 0
    Exit Property
TEMPERATURA_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TEMPERATURA_MIN of Módulo de clase clsEquipo"

End Property

Public Property Get TEMPERATURA_MAX() As String

On Error GoTo TEMPERATURA_MAX_Error

    TEMPERATURA_MAX = mvarstrTEMPERATURA_MAX

On Error GoTo 0
    Exit Property
TEMPERATURA_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TEMPERATURA_MAX of Módulo de clase clsEquipo"

End Property

Public Property Let TEMPERATURA_MAX(ByVal strTEMPERATURA_MAX As String)

On Error GoTo TEMPERATURA_MAX_Error

    mvarstrTEMPERATURA_MAX = strTEMPERATURA_MAX

On Error GoTo 0
    Exit Property
TEMPERATURA_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TEMPERATURA_MAX of Módulo de clase clsEquipo"

End Property

Public Property Get HUMEDAD_MIN() As String

On Error GoTo HUMEDAD_MIN_Error

    HUMEDAD_MIN = mvarstrHUMEDAD_MIN

On Error GoTo 0
    Exit Property
HUMEDAD_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HUMEDAD_MIN of Módulo de clase clsEquipo"

End Property

Public Property Let HUMEDAD_MIN(ByVal strHUMEDAD_MIN As String)

On Error GoTo HUMEDAD_MIN_Error

    mvarstrHUMEDAD_MIN = strHUMEDAD_MIN

On Error GoTo 0
    Exit Property
HUMEDAD_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HUMEDAD_MIN of Módulo de clase clsEquipo"

End Property

Public Property Get HUMEDAD_MAX() As String

On Error GoTo HUMEDAD_MAX_Error

    HUMEDAD_MAX = mvarstrHUMEDAD_MAX

On Error GoTo 0
    Exit Property
HUMEDAD_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HUMEDAD_MAX of Módulo de clase clsEquipo"

End Property

Public Property Let HUMEDAD_MAX(ByVal strHUMEDAD_MAX As String)

On Error GoTo HUMEDAD_MAX_Error

    mvarstrHUMEDAD_MAX = strHUMEDAD_MAX

On Error GoTo 0
    Exit Property
HUMEDAD_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HUMEDAD_MAX of Módulo de clase clsEquipo"

End Property

Public Property Get COND_AMBIENTALES_OTRAS() As String

On Error GoTo COND_AMBIENTALES_OTRAS_Error

    COND_AMBIENTALES_OTRAS = mvarstrCOND_AMBIENTALES_OTRAS

On Error GoTo 0
    Exit Property
COND_AMBIENTALES_OTRAS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure COND_AMBIENTALES_OTRAS of Módulo de clase clsEquipo"

End Property

Public Property Let COND_AMBIENTALES_OTRAS(ByVal strCOND_AMBIENTALES_OTRAS As String)

On Error GoTo COND_AMBIENTALES_OTRAS_Error

    mvarstrCOND_AMBIENTALES_OTRAS = strCOND_AMBIENTALES_OTRAS

On Error GoTo 0
    Exit Property
COND_AMBIENTALES_OTRAS_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure COND_AMBIENTALES_OTRAS of Módulo de clase clsEquipo"

End Property

Public Property Get RANGO_TRABAJO_MIN() As String

On Error GoTo RANGO_TRABAJO_MIN_Error

    RANGO_TRABAJO_MIN = mvarstrRANGO_TRABAJO_MIN

On Error GoTo 0
    Exit Property
RANGO_TRABAJO_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_TRABAJO_MIN of Módulo de clase clsEquipo"

End Property

Public Property Let RANGO_TRABAJO_MIN(ByVal strRANGO_TRABAJO_MIN As String)

On Error GoTo RANGO_TRABAJO_MIN_Error

    mvarstrRANGO_TRABAJO_MIN = strRANGO_TRABAJO_MIN

On Error GoTo 0
    Exit Property
RANGO_TRABAJO_MIN_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_TRABAJO_MIN of Módulo de clase clsEquipo"

End Property

Public Property Get RANGO_TRABAJO_MAX() As String

On Error GoTo RANGO_TRABAJO_MAX_Error

    RANGO_TRABAJO_MAX = mvarstrRANGO_TRABAJO_MAX

On Error GoTo 0
    Exit Property
RANGO_TRABAJO_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_TRABAJO_MAX of Módulo de clase clsEquipo"

End Property

Public Property Let RANGO_TRABAJO_MAX(ByVal strRANGO_TRABAJO_MAX As String)

On Error GoTo RANGO_TRABAJO_MAX_Error

    mvarstrRANGO_TRABAJO_MAX = strRANGO_TRABAJO_MAX

On Error GoTo 0
    Exit Property
RANGO_TRABAJO_MAX_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RANGO_TRABAJO_MAX of Módulo de clase clsEquipo"

End Property

Public Property Get LIMITACIONES_USO() As String

On Error GoTo LIMITACIONES_USO_Error

    LIMITACIONES_USO = mvarstrLIMITACIONES_USO

On Error GoTo 0
    Exit Property
LIMITACIONES_USO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LIMITACIONES_USO of Módulo de clase clsEquipo"

End Property

Public Property Let LIMITACIONES_USO(ByVal strLIMITACIONES_USO As String)

On Error GoTo LIMITACIONES_USO_Error

    mvarstrLIMITACIONES_USO = strLIMITACIONES_USO

On Error GoTo 0
    Exit Property
LIMITACIONES_USO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LIMITACIONES_USO of Módulo de clase clsEquipo"

End Property

Public Property Get PERIODICIDAD_CALIBRACION_ID() As Integer

On Error GoTo PERIODICIDAD_CALIBRACION_ID_Error

    PERIODICIDAD_CALIBRACION_ID = mvarintPERIODICIDAD_CALIBRACION_ID

On Error GoTo 0
    Exit Property
PERIODICIDAD_CALIBRACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PERIODICIDAD_CALIBRACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Let PERIODICIDAD_CALIBRACION_ID(ByVal intPERIODICIDAD_CALIBRACION_ID As Integer)

On Error GoTo PERIODICIDAD_CALIBRACION_ID_Error

    mvarintPERIODICIDAD_CALIBRACION_ID = intPERIODICIDAD_CALIBRACION_ID

On Error GoTo 0
    Exit Property
PERIODICIDAD_CALIBRACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PERIODICIDAD_CALIBRACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Get FECHA_PROX_CALIBRACION() As String

On Error GoTo FECHA_PROX_CALIBRACION_Error

    FECHA_PROX_CALIBRACION = mvarstrFECHA_PROX_CALIBRACION

On Error GoTo 0
    Exit Property
FECHA_PROX_CALIBRACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_PROX_CALIBRACION of Módulo de clase clsEquipo"

End Property

Public Property Let FECHA_PROX_CALIBRACION(ByVal strFECHA_PROX_CALIBRACION As String)

On Error GoTo FECHA_PROX_CALIBRACION_Error

    mvarstrFECHA_PROX_CALIBRACION = strFECHA_PROX_CALIBRACION

On Error GoTo 0
    Exit Property
FECHA_PROX_CALIBRACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_PROX_CALIBRACION of Módulo de clase clsEquipo"

End Property

Public Property Get PERIODICIDAD_VERIFICACION_ID() As Integer

On Error GoTo PERIODICIDAD_VERIFICACION_ID_Error

    PERIODICIDAD_VERIFICACION_ID = mvarintPERIODICIDAD_VERIFICACION_ID

On Error GoTo 0
    Exit Property
PERIODICIDAD_VERIFICACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PERIODICIDAD_VERIFICACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Let PERIODICIDAD_VERIFICACION_ID(ByVal intPERIODICIDAD_VERIFICACION_ID As Integer)

On Error GoTo PERIODICIDAD_VERIFICACION_ID_Error

    mvarintPERIODICIDAD_VERIFICACION_ID = intPERIODICIDAD_VERIFICACION_ID

On Error GoTo 0
    Exit Property
PERIODICIDAD_VERIFICACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PERIODICIDAD_VERIFICACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Get FECHA_PROX_VERIFICACION() As String

On Error GoTo FECHA_PROX_VERIFICACION_Error

    FECHA_PROX_VERIFICACION = mvarstrFECHA_PROX_VERIFICACION

On Error GoTo 0
    Exit Property
FECHA_PROX_VERIFICACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_PROX_VERIFICACION of Módulo de clase clsEquipo"

End Property

Public Property Let FECHA_PROX_VERIFICACION(ByVal strFECHA_PROX_VERIFICACION As String)

On Error GoTo FECHA_PROX_VERIFICACION_Error

    mvarstrFECHA_PROX_VERIFICACION = strFECHA_PROX_VERIFICACION

On Error GoTo 0
    Exit Property
FECHA_PROX_VERIFICACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_PROX_VERIFICACION of Módulo de clase clsEquipo"

End Property

Public Property Get FECHA_PROX_MANTENIMIENTO() As String

On Error GoTo FECHA_PROX_MANTENIMIENTO_Error

    FECHA_PROX_MANTENIMIENTO = mvarstrFECHA_PROX_MANTENIMIENTO

On Error GoTo 0
    Exit Property
FECHA_PROX_MANTENIMIENTO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_PROX_MANTENIMIENTO of Módulo de clase clsEquipo"

End Property

Public Property Let FECHA_PROX_MANTENIMIENTO(ByVal strFECHA_PROX_MANTENIMIENTO As String)

On Error GoTo FECHA_PROX_MANTENIMIENTO_Error

    mvarstrFECHA_PROX_MANTENIMIENTO = strFECHA_PROX_MANTENIMIENTO

On Error GoTo 0
    Exit Property
FECHA_PROX_MANTENIMIENTO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure FECHA_PROX_MANTENIMIENTO of Módulo de clase clsEquipo"

End Property

Public Property Get PERIODICIDAD_MANTENIMIENTO_ID() As Integer

On Error GoTo PERIODICIDAD_MANTENIMIENTO_ID_Error

    PERIODICIDAD_MANTENIMIENTO_ID = mvarintPERIODICIDAD_MANTENIMIENTO_ID

On Error GoTo 0
    Exit Property
PERIODICIDAD_MANTENIMIENTO_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PERIODICIDAD_MANTENIMIENTO_ID of Módulo de clase clsEquipo"

End Property

Public Property Let PERIODICIDAD_MANTENIMIENTO_ID(ByVal intPERIODICIDAD_MANTENIMIENTO_ID As Integer)

On Error GoTo PERIODICIDAD_MANTENIMIENTO_ID_Error

    mvarintPERIODICIDAD_MANTENIMIENTO_ID = intPERIODICIDAD_MANTENIMIENTO_ID

On Error GoTo 0
    Exit Property
PERIODICIDAD_MANTENIMIENTO_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PERIODICIDAD_MANTENIMIENTO_ID of Módulo de clase clsEquipo"

End Property

Public Property Get TIPO_CALIBRACION_ID() As Integer

On Error GoTo TIPO_CALIBRACION_ID_Error

    TIPO_CALIBRACION_ID = mvarintTIPO_CALIBRACION_ID

On Error GoTo 0
    Exit Property
TIPO_CALIBRACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TIPO_CALIBRACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Let TIPO_CALIBRACION_ID(ByVal intTIPO_CALIBRACION_ID As Integer)

On Error GoTo TIPO_CALIBRACION_ID_Error

    mvarintTIPO_CALIBRACION_ID = intTIPO_CALIBRACION_ID

On Error GoTo 0
    Exit Property
TIPO_CALIBRACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure TIPO_CALIBRACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Get CALIBRADOR_ID() As Integer

On Error GoTo CALIBRADOR_ID_Error

    CALIBRADOR_ID = mvarintCALIBRADOR_ID

On Error GoTo 0
    Exit Property
CALIBRADOR_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CALIBRADOR_ID of Módulo de clase clsEquipo"

End Property

Public Property Let CALIBRADOR_ID(ByVal intCALIBRADOR_ID As Integer)

On Error GoTo CALIBRADOR_ID_Error

    mvarintCALIBRADOR_ID = intCALIBRADOR_ID

On Error GoTo 0
    Exit Property
CALIBRADOR_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CALIBRADOR_ID of Módulo de clase clsEquipo"

End Property

Public Property Get CALIBRADOR_INTERNO_ID() As Integer

On Error GoTo CALIBRADOR_INTERNO_ID_Error

    CALIBRADOR_INTERNO_ID = mvarintCALIBRADOR_INTERNO_ID

On Error GoTo 0
    Exit Property
CALIBRADOR_INTERNO_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CALIBRADOR_INTERNO_ID of Módulo de clase clsEquipo"

End Property

Public Property Let CALIBRADOR_INTERNO_ID(ByVal intCALIBRADOR_INTERNO_ID As Integer)

On Error GoTo CALIBRADOR_INTERNO_ID_Error

    mvarintCALIBRADOR_INTERNO_ID = intCALIBRADOR_INTERNO_ID

On Error GoTo 0
    Exit Property
CALIBRADOR_INTERNO_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CALIBRADOR_INTERNO_ID of Módulo de clase clsEquipo"

End Property

Public Property Get SITUACION_ID() As Integer

On Error GoTo SITUACION_ID_Error

    SITUACION_ID = mvarintSITUACION_ID

On Error GoTo 0
    Exit Property
SITUACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SITUACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Let SITUACION_ID(ByVal intSITUACION_ID As Integer)

On Error GoTo SITUACION_ID_Error

    mvarintSITUACION_ID = intSITUACION_ID

On Error GoTo 0
    Exit Property
SITUACION_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SITUACION_ID of Módulo de clase clsEquipo"

End Property

Public Property Get CON_CALIBRACION() As Integer

On Error GoTo CON_CALIBRACION_Error

    CON_CALIBRACION = mvarintCON_CALIBRACION

On Error GoTo 0
    Exit Property
CON_CALIBRACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CON_CALIBRACION of Módulo de clase clsEquipo"

End Property

Public Property Let CON_CALIBRACION(ByVal intCON_CALIBRACION As Integer)

On Error GoTo CON_CALIBRACION_Error

    mvarintCON_CALIBRACION = intCON_CALIBRACION

On Error GoTo 0
    Exit Property
CON_CALIBRACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CON_CALIBRACION of Módulo de clase clsEquipo"

End Property

Public Property Get CON_VERIFICACION() As Integer

On Error GoTo CON_VERIFICACION_Error

    CON_VERIFICACION = mvarintCON_VERIFICACION

On Error GoTo 0
    Exit Property
CON_VERIFICACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CON_VERIFICACION of Módulo de clase clsEquipo"

End Property

Public Property Let CON_VERIFICACION(ByVal intCON_VERIFICACION As Integer)

On Error GoTo CON_VERIFICACION_Error

    mvarintCON_VERIFICACION = intCON_VERIFICACION

On Error GoTo 0
    Exit Property
CON_VERIFICACION_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CON_VERIFICACION of Módulo de clase clsEquipo"

End Property

Public Property Get CON_MANTENIMIENTO() As Integer

On Error GoTo CON_MANTENIMIENTO_Error

    CON_MANTENIMIENTO = mvarintCON_MANTENIMIENTO

On Error GoTo 0
    Exit Property
CON_MANTENIMIENTO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CON_MANTENIMIENTO of Módulo de clase clsEquipo"

End Property

Public Property Let CON_MANTENIMIENTO(ByVal intCON_MANTENIMIENTO As Integer)

On Error GoTo CON_MANTENIMIENTO_Error

    mvarintCON_MANTENIMIENTO = intCON_MANTENIMIENTO

On Error GoTo 0
    Exit Property
CON_MANTENIMIENTO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CON_MANTENIMIENTO of Módulo de clase clsEquipo"

End Property

Public Property Get REVISADO() As Integer

On Error GoTo REVISADO_Error

    REVISADO = mvarintREVISADO

On Error GoTo 0
    Exit Property
REVISADO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure REVISADO of Módulo de clase clsEquipo"

End Property

Public Property Let REVISADO(ByVal intREVISADO As Integer)

On Error GoTo REVISADO_Error

    mvarintREVISADO = intREVISADO

On Error GoTo 0
    Exit Property
REVISADO_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure REVISADO of Módulo de clase clsEquipo"

End Property

Public Property Get SUBFAMILIA_ID() As Long

On Error GoTo SUBFAMILIA_ID_Error

    SUBFAMILIA_ID = mvarlngSUBFAMILIA_ID

On Error GoTo 0
    Exit Property
SUBFAMILIA_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SUBFAMILIA_ID of Módulo de clase clsEquipo"

End Property

Public Property Let SUBFAMILIA_ID(ByVal lngSUBFAMILIA_ID As Long)

On Error GoTo SUBFAMILIA_ID_Error

    mvarlngSUBFAMILIA_ID = lngSUBFAMILIA_ID

On Error GoTo 0
    Exit Property
SUBFAMILIA_ID_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SUBFAMILIA_ID of Módulo de clase clsEquipo"

End Property

Public Property Get CoordY() As Long

On Error GoTo CoordY_Error

    CoordY = mvarlngCoordY

On Error GoTo 0
    Exit Property
CoordY_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CoordY of Módulo de clase clsEquipo"

End Property

Public Property Let CoordY(ByVal lngCoordY As Long)

On Error GoTo CoordY_Error

    mvarlngCoordY = lngCoordY

On Error GoTo 0
    Exit Property
CoordY_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CoordY of Módulo de clase clsEquipo"

End Property

Public Function ModificarPosicionEnPlano() As Boolean
Dim consulta As String

On Error GoTo ModificarPosicionEnPlano_Error

    If mvarlngID_EQUIPO >= 0 Then Exit Function
    
    consulta = "update equipos set coordx = " & CStr(mvarlngCoordX)
    consulta = consulta & ", coordy = " & CStr(mvarlngCoordY)
    consulta = consulta & " where id_equipo = " & mvarlngID_EQUIPO
    
    execute_bd consulta
    
    ModificarPosicionEnPlano = True

On Error GoTo 0
    Exit Function
ModificarPosicionEnPlano_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ModificarPosicionEnPlano of Módulo de clase clsEquipo"
    ModificarPosicionEnPlano = False
End Function
