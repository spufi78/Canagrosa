VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeterminaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_DETERMINACION As Long
Private MUESTRA_ID As Long
Private TIPO_DETERMINACION_ID As Long
Private FORMULA_ID As Integer
Private ORDEN As Integer
Private ESTADO As Integer
Private SITUACION As Integer
Private OBSERVACIONES As String
Private RESULTADO As String
Private DIF_DUPLICADOS As String
Private fecha As String
Private hora As String
Private EMPLEADO_ID As Integer
Private ES_DUPLICADO As Integer
Private METODO As String
Private GRADO As String
'E0132-I
Private ENVIADO_PAQUETE As Integer
Private USUARIO_ID_REVISION As Integer
Private FECHA_REVISION As String
'E0132-F
'M1371-I
Private REVISION_DUPLICADO As Integer
'M1371-F
Public Property Let setID_DETERMINACION(ByVal dato As Long)
    ID_DETERMINACION = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
    MUESTRA_ID = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
    TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setFORMULA_ID(ByVal dato As Integer)
    FORMULA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = dato
End Property
Public Property Let setESTADO(ByVal dato As Integer)
    ESTADO = dato
End Property
Public Property Let setSITUACION(ByVal dato As Integer)
    SITUACION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setRESULTADO(ByVal dato As String)
    RESULTADO = Trim(dato)
End Property
Public Property Let setDIF_DUPLICADOS(ByVal dato As String)
    DIF_DUPLICADOS = Trim(dato)
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = Trim(dato)
End Property
Public Property Let setHORA(ByVal dato As String)
    hora = Trim(dato)
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
    EMPLEADO_ID = dato
End Property
Public Property Let setES_DUPLICADO(ByVal dato As Integer)
    ES_DUPLICADO = dato
End Property
Public Property Let setMETODO(ByVal dato As String)
    METODO = dato
End Property
Public Property Let setGRADO(ByVal dato As String)
    GRADO = dato
End Property
'E0134-I
Public Property Let setENVIADO_PAQUETE(ByVal dato As Integer)
    ENVIADO_PAQUETE = dato
End Property
Public Property Let setUSUARIO_ID_REVISION(ByVal dato As Integer)
    USUARIO_ID_REVISION = dato
End Property
Public Property Let setFECHA_REVISION(ByVal dato As String)
    FECHA_REVISION = dato
End Property
'E0134-F
'M1371-I
Public Property Let setREVISION_DUPLICADO(ByVal dato As Integer)
    REVISION_DUPLICADO = dato
End Property
'M1371-F
Public Property Get getID_DETERMINACION() As Long
    getID_DETERMINACION = ID_DETERMINACION
End Property
Public Property Get getMUESTRA_ID() As Long
    getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Long
    getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getFORMULA_ID() As Integer
    getFORMULA_ID = FORMULA_ID
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Property Get getESTADO() As Integer
    getESTADO = ESTADO
End Property
Public Property Get getSITUACION() As Integer
    getSITUACION = SITUACION
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getRESULTADO() As String
    getRESULTADO = RESULTADO
End Property
Public Property Get getDIF_DUPLICADOS() As String
    getDIF_DUPLICADOS = DIF_DUPLICADOS
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getHORA() As String
    getHORA = hora
End Property
Public Property Get getEMPLEADO_ID() As Integer
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getES_DUPLICADO() As Integer
    getES_DUPLICADO = ES_DUPLICADO
End Property
Public Property Get getMETODO() As String
    getMETODO = METODO
End Property
Public Property Get getGRADO() As String
    getGRADO = GRADO
End Property
'E0133-I
Public Property Get getENVIADO_PAQUETE() As Integer
    getENVIADO_PAQUETE = ENVIADO_PAQUETE
End Property
Public Property Get getUSUARIO_ID_REVISION() As Integer
    getUSUARIO_ID_REVISION = USUARIO_ID_REVISION
End Property
Public Property Get getFECHA_REVISION() As String
    getFECHA_REVISION = FECHA_REVISION
End Property
'E0133-F
'M1371-I
Public Property Get getREVISION_DUPLICADO() As Integer
    getREVISION_DUPLICADO = REVISION_DUPLICADO
End Property
Public Function InsertarDeterminacion() As Long
    On Error GoTo fallo
    Dim consulta As String
    CrearId_DETERMINACION
    Me.setREVISION_DUPLICADO = 0
    consulta = "Insert into determinaciones " & _
               " values(" & _
               ID_DETERMINACION & "," & MUESTRA_ID & "," & TIPO_DETERMINACION_ID & "," & _
               FORMULA_ID & "," & ORDEN & "," & ESTADO & "," & SITUACION & ",'" & _
               OBSERVACIONES & "','" & RESULTADO & "','" & DIF_DUPLICADOS & "','" & fecha_bd(fecha) & "','" & hora & "'," & _
               EMPLEADO_ID & "," & ES_DUPLICADO & ",'" & METODO & "','" & GRADO & "'," & ENVIADO_PAQUETE & "," & _
               USUARIO_ID_REVISION & ",'" & fecha_bd(FECHA_REVISION) & "'" & "," & REVISION_DUPLICADO & ")"
    execute_bd consulta
    'MYYYY Insertar valores por defecto de configuración
    Dim oDC As New clsDeterminaciones_conf
    oDC.Insertar ID_DETERMINACION
    
    InsertarDeterminacion = ID_DETERMINACION
    Exit Function
fallo:
    InsertarDeterminacion = 0
    MsgBox "Error al insertar las determinaciones (clsDeterminaciones)", vbCritical, Err.Description
End Function

Private Sub CrearId_DETERMINACION()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(Id_DETERMINACION) FROM determinaciones"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_DETERMINACION = 1
    Else
        ID_DETERMINACION = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function CargaPorMuestraTipo(MUESTRA_ID As Long, TIPO_DETERMINACION_ID As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT ID_DETERMINACION FROM determinaciones WHERE MUESTRA_ID=" & MUESTRA_ID & " AND TIPO_DETERMINACION_ID = " & TIPO_DETERMINACION_ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        CargaPorMuestraTipo = False
    Else
        CargaPorMuestraTipo = True
        Me.CargarDeterminacion rs(0)
    End If
End Function
Public Function CargarDeterminacion(ID As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT * FROM determinaciones WHERE ID_DETERMINACION=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
         CargarDeterminacion = False
    Else
        ID_DETERMINACION = rs.Fields("id_determinacion")
        MUESTRA_ID = rs.Fields("muestra_id")
        TIPO_DETERMINACION_ID = rs.Fields("tipo_determinacion_id")
        FORMULA_ID = rs.Fields("FORMULA_ID")
        ORDEN = rs.Fields("ORDEN")
        ESTADO = rs.Fields("ESTADO")
        SITUACION = rs.Fields("SITUACION")
        If Not IsNull(rs.Fields("OBSERVACIONES")) Then
            OBSERVACIONES = rs.Fields("OBSERVACIONES")
        Else
            OBSERVACIONES = ""
        End If
        If Not IsNull(rs.Fields("RESULTADO")) Then
            RESULTADO = rs.Fields("RESULTADO")
        Else
            RESULTADO = ""
        End If
        If Not IsNull(rs.Fields("DIF_DUPLICADOS")) Then
            DIF_DUPLICADOS = rs.Fields("DIF_DUPLICADOS")
        Else
            DIF_DUPLICADOS = ""
        End If
        If Not IsNull(rs.Fields("FECHA")) Then
            fecha = rs.Fields("FECHA")
        Else
            fecha = ""
        End If
        If Not IsNull(rs.Fields("HORA")) Then
            hora = rs.Fields("HORA")
        Else
            hora = ""
        End If
        If Not IsNull(rs.Fields("EMPLEADO_ID")) Then
            EMPLEADO_ID = rs.Fields("EMPLEADO_ID")
        Else
            EMPLEADO_ID = 0
        End If
        ES_DUPLICADO = rs.Fields("ES_DUPLICADO")
        METODO = rs("METODO")
        GRADO = rs("GRADO")
        'E0135-I
        ENVIADO_PAQUETE = rs.Fields("ENVIADO_PAQUETE")
        'E0135-F
        USUARIO_ID_REVISION = rs.Fields("USUARIO_ID_REVISION")
        If Not IsNull(rs.Fields("FECHA_REVISION")) Then
            FECHA_REVISION = rs.Fields("FECHA_REVISION")
        Else
            FECHA_REVISION = 0
        End If
        'M1371-I
        If Not IsNull(rs.Fields("REVISION_DUPLICADO")) Then
            REVISION_DUPLICADO = rs.Fields("REVISION_DUPLICADO")
        Else
            REVISION_DUPLICADO = 0
        End If
        'M1371-F
    End If
    Set rs = Nothing
    CargarDeterminacion = True
    Exit Function
fallo:
    CargarDeterminacion = False
    MsgBox "Error al cargar la determinacion (CargarDeterminacion)", vbCritical, Err.Description
End Function
Function lista_determinaciones(MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT d.id_determinacion, td.nombre, td.id_tipo_determinacion, " & _
               "       d.resultado, td.pnt, d.empleado_id, d.tipo_determinacion_id, " & _
               "       td.descripcion,d.formula_id,'0' as tipo,td.tarifa_codigo_id, " & _
               "       td.particulas,td.proc_ref_eads,td.precio,'' as lc,td.metodo,d.metodo as metodo_deter, " & _
               "       d.ORDEN , d.GRADO, d.REVISION_DUPLICADO, d.situacion, " & _
               "       dc.rango,dc.dif_duplicados,dc.dif_duplicados_numerica,dc.dif_aviso,dc.incertidumbre,dc.dif_historico,dc.c_variacion " & _
               "  FROM determinaciones d " & _
               " INNER JOIN tipos_determinacion td on tipo_determinacion_id=id_tipo_determinacion " & _
               " LEFT JOIN determinaciones_conf dc on d.id_determinacion = dc.determinacion_id " & _
               " WHERE muestra_id=" & MUESTRA & _
               " order by orden"
    Set lista_determinaciones = datos_bd(consulta)
fallo:
End Function
Function lista_determinaciones_factura(MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT d.id_determinacion, td.nombre, td.id_tipo_determinacion,td.tarifa_codigo_id,td.proc_ref_eads " & _
               "  FROM determinaciones d, tipos_determinacion td " & _
               " WHERE tipo_determinacion_id=id_tipo_determinacion" & _
               "   AND muestra_id=" & MUESTRA & _
               "   AND d.resultado <> '--'" & _
               " order by orden"
    Set lista_determinaciones_factura = datos_bd(consulta)
fallo:
End Function

Function lista_determinaciones_vida(MUESTRA As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT C.NOMBRE,E.NOMBRE ,A.VALOR_1,A.VALOR_2,D.USUARIO, A.FECHA,A.HORA,B.TIPO_DETERMINACION_ID,B.ID_DETERMINACION,D.ID_EMPLEADO,A.DETERMINACION_ID,A.ORDEN,A.CAMPO_ID " & _
               "  FROM DETERMINACIONES B " & _
               "  LEFT JOIN DETERMINACIONES_HISTORICO A ON a.DETERMINACION_ID = b.ID_DETERMINACION " & _
               "  LEFT JOIN TIPOS_DETERMINACION C ON B.TIPO_DETERMINACION_ID = C.ID_TIPO_DETERMINACION " & _
               "  LEFT JOIN usuarios D ON A.EMPLEADO_ID = D.ID_EMPLEADO " & _
               "  LEFT JOIN FORMULAS_CAMPOS E ON A.CAMPO_ID = E.ID_CAMPO " & _
               " WHERE B.MUESTRA_ID = " & MUESTRA & _
               " ORDER BY B.ORDEN,A.CAMPO_ID,A.ORDEN"

    Set lista_determinaciones_vida = datos_bd(consulta)
End Function
Function lista_determinaciones_muestra(MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT td.pnt, td.nombre, d.resultado,td.id_tipo_determinacion " & _
               "  FROM determinaciones d, tipos_determinacion td " & _
               " WHERE tipo_determinacion_id=id_tipo_determinacion" & _
               "   AND muestra_id=" & MUESTRA & " order by orden"
    Set lista_determinaciones_muestra = datos_bd(consulta)
fallo:
End Function
Function lista_determinaciones_fuera_rango(MUESTRA As Long) As ADODB.Recordset
    Dim c As String
   On Error GoTo lista_determinaciones_fuera_rango_Error

    c = "SELECT c.NOMBRE, b.RESULTADO, f.NOMBRE " & _
        " FROM muestras a,determinaciones b, tipos_determinacion c, formulas d, formulas_campos e, unidades f " & _
        " where a.ID_MUESTRA = b.MUESTRA_ID " & _
        " and b.TIPO_DETERMINACION_ID = c.ID_TIPO_DETERMINACION " & _
        " and c.FORMULA_ID = d.ID_FORMULA " & _
        " and d.CAMPO_ID_RESULTADO = e.ID_CAMPO " & _
        " and e.UNIDAD_ID = f.ID_UNIDAD " & _
        " and b.SITUACION = 2 " & _
        " and a.ID_MUESTRA = " & MUESTRA
    Set lista_determinaciones_fuera_rango = datos_bd(c)

   On Error GoTo 0
   Exit Function

lista_determinaciones_fuera_rango_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_determinaciones_fuera_rango of Módulo de clase clsDeterminaciones"
End Function
Function lista_determinaciones_para_factura(MUESTRA As Long, TARIFA As Long) As ADODB.Recordset
   On Error GoTo lista_determinaciones_para_factura_Error

    On Error GoTo fallo
    Dim consulta As String
    Dim oMuestra As New clsMuestra
    If oMuestra.CargaMuestra(MUESTRA_ID) Then
        ' SELLANTE
        If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.SELLANTE Then
'            Dim oSR As New clsSellantes_resultados
'            Set lista_determinaciones_para_factura = oSR.Listado_Resultados(MUESTRA_ID)
        Else
        'DETERMINACIONES
            consulta = "SELECT td.pnt, td.nombre, td.proc_ref_eads, tc.CODIGO,td.PRECIO, tp.PRECIO,td.id_tipo_determinacion,d.id_determinacion " & _
                        "  FROM determinaciones d " & _
                        " inner join tipos_determinacion td on tipo_determinacion_id=id_tipo_determinacion " & _
                        "  left join tarifas_codigos tc on tc.ID_CODIGO = td.TARIFA_CODIGO_ID " & _
                        "  left join tarifas_precios tp on tp.TIPO_DETERMINACION_ID =d.TIPO_DETERMINACION_ID and tp.TARIFA_ID = " & TARIFA & _
                        " where MUESTRA_ID = " & MUESTRA & _
                        " order by orden"
            Set lista_determinaciones_para_factura = datos_bd(consulta)
        End If
    End If
fallo:

   On Error GoTo 0
   Exit Function

lista_determinaciones_para_factura_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_determinaciones_para_factura of Módulo de clase clsDeterminaciones"
End Function
Function lista_determinaciones_para_factura_determinacion(MUESTRA As Long, TARIFA As Long, TIPO_DETER As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT td.pnt, td.nombre, td.proc_ref_eads, tc.CODIGO,td.PRECIO, tp.PRECIO,td.id_tipo_determinacion " & _
                "  FROM determinaciones d " & _
                " inner join tipos_determinacion td on tipo_determinacion_id=id_tipo_determinacion " & _
                "  left join tarifas_codigos tc on tc.ID_CODIGO = td.TARIFA_CODIGO_ID " & _
                "  left join tarifas_precios tp on tp.TIPO_DETERMINACION_ID =d.TIPO_DETERMINACION_ID and tp.TARIFA_ID = " & TARIFA & _
                " where MUESTRA_ID = " & MUESTRA & _
                "   and d.TIPO_DETERMINACION_ID = " & TIPO_DETER & _
                " order by orden"
    Set lista_determinaciones_para_factura_determinacion = datos_bd(consulta)
fallo:
End Function
Function lista_determinaciones_bano(MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT td.nombre, td.nombre_ingles, td.proc_ref_eads, '0' as mensaje_cero, " & _
               "       det.id_determinacion, det.fecha, det.resultado, u.nombre,mue.bano_id, det.orden,rb.minimo,rb.maximo,td.id_tipo_determinacion,td.pnt,rb.minimo_texto,rb.maximo_texto " & _
               "         FROM determinaciones as det, tipos_determinacion as td, muestras as mue, " & _
               "              formulas as f,formulas_campos as cf, unidades as u,determinaciones_analisis as rb " & _
               "       WHERE det.tipo_determinacion_id = td.id_tipo_determinacion " & _
               "         AND rb.bano_id = mue.bano_id " & _
               "         AND rb.tipo_determinacion_id = det.tipo_determinacion_id " & _
               "         AND det.muestra_id = mue.id_muestra " & _
               "         AND det.formula_id = f.id_formula " & _
               "         AND f.campo_id_resultado = cf.id_campo " & _
               "         AND cf.unidad_id = u.id_unidad " & _
               "         AND mue.id_muestra = " & MUESTRA & _
               "      ORDER BY rb.ORDEN"
    Set lista_determinaciones_bano = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function
Function lista_determinaciones_tipo_analisis(MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT td.nombre, td.nombre_ingles, td.proc_ref_eads, '0' as mensaje_cero, " & _
               "       det.id_determinacion, det.fecha, det.resultado, u.nombre,mue.bano_id, det.orden,rb.minimo,rb.maximo,td.id_tipo_determinacion,td.pnt,'' as lc,rb.minimo_texto,rb.maximo_texto " & _
               "         FROM determinaciones as det, tipos_determinacion as td, muestras as mue, " & _
               "              formulas as f,formulas_campos as cf, unidades as u,determinaciones_analisis as rb " & _
               "       WHERE det.tipo_determinacion_id = td.id_tipo_determinacion " & _
               "         AND rb.bano_id = mue.bano_id " & _
               "         AND rb.tipo_determinacion_id = det.tipo_determinacion_id " & _
               "         AND rb.tipo_analisis_id = mue.tipo_analisis_id " & _
               "         AND det.muestra_id = mue.id_muestra " & _
               "         AND det.formula_id = f.id_formula " & _
               "         AND f.campo_id_resultado = cf.id_campo " & _
               "         AND cf.unidad_id = u.id_unidad " & _
               "         AND mue.id_muestra = " & MUESTRA & _
               "      ORDER BY rb.ORDEN"
    Set lista_determinaciones_tipo_analisis = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function


Public Function InsertarSolucion(deter As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim SITUACION As Integer
'               "   GRADO = '" & GRADO & "',"
    consulta = "update determinaciones SET " & _
               "   RESULTADO = '" & RESULTADO & "'," & _
               "   DIF_DUPLICADOS = '" & DIF_DUPLICADOS & "'," & _
               "   FECHA = '" & fecha & "'," & _
               "   HORA = '" & hora & "'," & _
               "   EMPLEADO_ID = " & EMPLEADO_ID & ", " & _
               "   REVISION_DUPLICADO = " & REVISION_DUPLICADO & _
               " WHERE ID_DETERMINACION = " & deter
    execute_bd consulta
    
    CalcularSituacion deter
    
    InsertarSolucion = True
    Exit Function
fallo:
    InsertarSolucion = False
    MsgBox "Error al insertar la Solucion (clsDeterminaciones)", vbCritical, Err.Description
End Function
Private Function CalcularSituacion(deter As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As ADODB.Recordset
    Dim SITUACION As Integer
    SITUACION = C_SITUACION.S_EN_RANGO
    consulta = "select muestra_id,tipo_determinacion_id,resultado,situacion,ifnull(b.DIF_AVISO,'') as DIF_AVISO,ifnull(b.INCERTIDUMBRE,'') as INCERTIDUMBRE " & _
               "  from determinaciones a left join determinaciones_conf b on a.ID_DETERMINACION = b.DETERMINACION_ID " & _
               " where id_determinacion = " & deter
    Set rs = datos_bd(consulta)
    Dim TA As Long
    Dim BANO As Long
    Dim RESULTADO As String
    If rs.RecordCount > 0 Then
        Dim DIF_AVISO As String
        Dim INCERTIDUMBRE As String
        RESULTADO = Trim(rs("resultado"))
        DIF_AVISO = Trim(rs("DIF_AVISO"))
        INCERTIDUMBRE = Trim(rs("INCERTIDUMBRE"))
        If RESULTADO = "" Then
            Exit Function
        Else
            RESULTADO = Replace(RESULTADO, "<", "")
            RESULTADO = Replace(RESULTADO, ">", "")
            If Not IsNumeric(RESULTADO) Then
                Exit Function
            End If
        End If
        Dim rs2 As ADODB.Recordset
        consulta = "select tipo_analisis_id, bano_id from muestras where id_muestra = " & rs("muestra_id")
        Set rs2 = datos_bd(consulta)
        If rs2.RecordCount = 0 Then
            Exit Function
        Else
            TA = rs2(0)
            BANO = rs2(1)
        End If
        Dim tipo As String
        If BANO <> 0 Then
            tipo = " AND DA.BANO_ID = " & BANO
        Else
            tipo = " AND DA.TIPO_ANALISIS_ID =  " & TA
        End If
        consulta = "select da.minimo,da.maximo " & _
                   "  from tipos_determinacion td " & _
                   "  left join determinaciones_analisis da ON td.ID_TIPO_DETERMINACION = da.TIPO_DETERMINACION_ID " & tipo & _
                   " where td.id_tipo_determinacion = " & rs("tipo_determinacion_id")
        Set rs2 = datos_bd(consulta)
        If rs2.RecordCount > 0 Then
            Dim min As Single
            Dim Max As Single
            min = 0
            Max = 0
            
            If Trim(rs2("MINIMO")) <> "" Then
             If IsNumeric(Trim(rs2("MINIMO"))) Then
              min = CSng(Replace(rs2("MINIMO"), ".", ","))
              
              If INCERTIDUMBRE <> "" Then
                If IsNumeric(INCERTIDUMBRE) Then
                    min = min - CSng(Replace(INCERTIDUMBRE, ".", ","))
                End If
              End If
              
              If CSng(Replace(RESULTADO, ".", ",")) < min Then
                SITUACION = C_SITUACION.S_FUERA_RANGO
              End If
             End If
            End If
            If Trim(rs2("MAXIMO")) <> "" Then
             If IsNumeric(Trim(rs2("MAXIMO"))) Then
              Max = CSng(Replace(rs2("MAXIMO"), ".", ","))
              
              If INCERTIDUMBRE <> "" Then
                If IsNumeric(INCERTIDUMBRE) Then
                    Max = Max + CSng(Replace(INCERTIDUMBRE, ".", ","))
                End If
              End If
              
              If CSng(Replace(RESULTADO, ".", ",")) > Max Then
                SITUACION = C_SITUACION.S_FUERA_RANGO
              End If
             End If
            End If
            ' Verificar alerta de resultado
            If SITUACION = C_SITUACION.S_EN_RANGO And (Trim(rs2("MINIMO")) <> "" Or Trim(rs2("MAXIMO")) <> "") Then
                If DIF_AVISO <> "" Then
                    If Max > min Then
                        Dim dif As Single
                        dif = ((Max - min) * DIF_AVISO / 100)
                        If Trim(rs2("MINIMO")) <> "" Then
                            If IsNumeric(Trim(rs2("MINIMO"))) Then
                                min = CSng(Replace(rs2("MINIMO"), ".", ",")) + dif
                                If CSng(Replace(RESULTADO, ".", ",")) < min Then
                                   SITUACION = C_SITUACION.S_LIMITES
                                End If
                            End If
                        End If
                        If Trim(rs2("MAXIMO")) <> "" Then
                            If IsNumeric(Trim(rs2("MAXIMO"))) Then
                                Max = CSng(Replace(rs2("MAXIMO"), ".", ",")) - dif
                                If CSng(Replace(RESULTADO, ".", ",")) > Max Then
                                   SITUACION = C_SITUACION.S_LIMITES
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
        If SITUACION <> rs(3) Then
            execute_bd "UPDATE DETERMINACIONES SET SITUACION = " & SITUACION & " WHERE ID_DETERMINACION = " & deter, False
        End If
    End If
    CalcularSituacion = True
    Exit Function
fallo:
    CalcularSituacion = False
    MsgBox "Error al insertar la Situacion (clsDeterminaciones->CalcularSituacion).ID_DETERMINACION: " & deter, vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    execute_bd "DELETE FROM DETERMINACIONES WHERE ID_DETERMINACION=" & ID
    execute_bd "DELETE FROM DETERMINACIONES_HISTORICO WHERE DETERMINACION_ID=" & ID
    execute_bd "DELETE FROM DETERMINACIONES_EQUIPOS WHERE DETERMINACION_ID=" & ID
    execute_bd "DELETE FROM DETERMINACIONES_REACTIVOS WHERE DETERMINACION_ID=" & ID
    execute_bd "DELETE FROM GRANULOMETRIAS WHERE DETERMINACION_ID=" & ID
    Eliminar = True
    Exit Function
fallo:
    Eliminar = False
End Function
Public Function Eliminar_Por_Muestra(ID_MUESTRA As Long) As Boolean
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "select id_determinacion from determinaciones where muestra_id = " & ID_MUESTRA
    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
       Do
          consulta = "delete from datos_determinaciones where determinacion_id = " & rs("id_determinacion")
          execute_bd consulta
          consulta = "delete from determinaciones_equipos where determinacion_id = " & rs("id_determinacion")
          execute_bd consulta
          consulta = "delete from determinaciones_historico where determinacion_id = " & rs("id_determinacion")
         execute_bd consulta
          consulta = "delete from determinaciones_reactivos where determinacion_id = " & rs("id_determinacion")
          execute_bd consulta

          rs.MoveNext
       Loop Until rs.EOF
    End If
    ' Borramos las determinaciones
    consulta = "delete from determinaciones where muestra_id = " & ID_MUESTRA
    execute_bd consulta
    Eliminar_Por_Muestra = True
    Exit Function
fallo:
    Eliminar_Por_Muestra = False
End Function
Public Function NumeroDeterminaciones(ID As Long) As Integer
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT COUNT(Id_DETERMINACION) FROM determinaciones where muestra_id=" & ID
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        NumeroDeterminaciones = 1
    Else
        NumeroDeterminaciones = rs.Fields(0)
    End If
    Set rs = Nothing
End Function

Function Insertar_determinaciones_por_defecto(ID_MUESTRA As Long) As Boolean
    Dim oMuestra As New clsMuestra
    Dim oDeter As New clsDeterminaciones
    Dim otipodet As New clsTipos_determinacion
    Dim ocampos As New clsFormulas_campos
    Dim oDatosDet As New clsDatos_determinaciones
    Dim DETERMINACION As Long
    Dim i As Integer
    Dim rscampos As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim oDET_Equipos As New clsDeterminaciones_equipos
    Dim oDET_Reactivos As New clsDeterminaciones_reactivos
   
   On Error GoTo Insertar_determinaciones_por_defecto_Error

    oMuestra.CargaMuestra (ID_MUESTRA)
    If oMuestra.getBANO_ID = 0 Then
        Set rs = otipodet.DeterminacionesPorDefecto(oMuestra.getTIPO_ANALISIS_ID)
    Else
        Set rs = otipodet.DeterminacionesPorDefectoBano(oMuestra.getBANO_ID)
    End If
    i = 1
    If rs.RecordCount <> 0 Then
       Do
          oDeter.setMUESTRA_ID = ID_MUESTRA
          oDeter.setTIPO_DETERMINACION_ID = rs("tipo")
          oDeter.setORDEN = i
          ' Recuperar tipos_determinacion para su FORMULA_ID
          otipodet.CargarTipoDeterminacion (rs("tipo"))
          oDeter.setFORMULA_ID = otipodet.getFORMULA_ID
          oDeter.setES_DUPLICADO = oMuestra.getANALISIS_DUPLICADO
          oDeter.setMETODO = otipodet.getMETODO
          oDeter.setSITUACION = C_SITUACION.S_EN_RANGO
          DETERMINACION = oDeter.InsertarDeterminacion
          ' Insertar Datos_Determinaciones
          Set rscampos = ocampos.ListaFormulas(otipodet.getFORMULA_ID)
          If rscampos.RecordCount <> 0 Then
            Do
                oDatosDet.setDETERMINACION_ID = DETERMINACION
                oDatosDet.setCAMPO_ID = rscampos("id_campo")
'                oDatosDet.setVALOR_1 = "I-1"
'                oDatosDet.setVALOR_2 = "I-1"
                oDatosDet.setVALOR_1 = ""
                oDatosDet.setVALOR_2 = ""
                oDatosDet.Insertar
                rscampos.MoveNext
            Loop Until rscampos.EOF
          End If
          
          
          oDET_Equipos.Insertar DETERMINACION, rs("tipo")
          oDET_Reactivos.Insertar DETERMINACION, rs("tipo")
          
          rs.MoveNext
          i = i + 1
        Loop Until rs.EOF
    End If

    Set oDET_Equipos = Nothing
    Set oDET_Reactivos = Nothing
    
   On Error GoTo 0
   Exit Function

Insertar_determinaciones_por_defecto_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Insertar_determinaciones_por_defecto of Módulo de clase clsDeterminaciones"
End Function
Function lista_contradictorio(MUESTRA As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT * FROM determinaciones " & _
               " WHERE muestra_id=" & MUESTRA & " order by orden"
    Set lista_contradictorio = datos_bd(consulta)
fallo:
End Function
Public Function tieneSC(MUESTRA As Long) As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo tieneSC_Error

    consulta = "select id_determinacion from determinaciones where muestra_id = " & MUESTRA & " and enviado_paquete = 1 limit 1"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        tieneSC = False
    Else
        tieneSC = True
    End If
    ' Si no es por determinaciones, mirar por  muestra
    If Not tieneSC Then
        consulta = "select paquete_id from sc_paquetes_detalle where muestra_id = " & MUESTRA & " limit 1"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            tieneSC = False
        Else
            tieneSC = True
        End If
    End If

   On Error GoTo 0
   Exit Function

tieneSC_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tieneSC of Módulo de clase clsDeterminaciones"

End Function

Public Function Contiene_Determinacion_Fluoruro(id_bano As Integer) As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
   On Error GoTo Contiene_Determinacion_Fluoruro_Error

    consulta = "SELECT * FROM DETERMINACIONES_ANALISIS " & _
               " WHERE BANO_ID = " & id_bano & _
               "   AND TIPO_DETERMINACION_ID = 226"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Contiene_Determinacion_Fluoruro = False
    Else
        Contiene_Determinacion_Fluoruro = True
    End If

   On Error GoTo 0
   Exit Function

Contiene_Determinacion_Fluoruro_Error:

    log ("Error " & Err.Number & " (" & Err.Description & ") in procedure Contiene_Determinacion_Fluoruro of Módulo de clase clsTipos_determinacion")
End Function


Public Function CargarResultado(ID As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT RESULTADO FROM DETERMINACIONES WHERE ID_DETERMINACION=" & ID
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
         CargarResultado = False
    Else
        If Not IsNull(rs.Fields("RESULTADO")) Then
            RESULTADO = rs.Fields("RESULTADO")
        Else
            RESULTADO = ""
        End If
    End If
    Set rs = Nothing
    CargarResultado = True
    Exit Function
fallo:
    CargarResultado = False
    MsgBox "Error al cargar la determinacion (CargarDeterminacion)", vbCritical, Err.Description
End Function

Function revisar_facturacion(MUESTRA As Long) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT td.revisar_factura " & _
               "  FROM determinaciones d, tipos_determinacion td " & _
               " WHERE tipo_determinacion_id=id_tipo_determinacion" & _
               "   AND muestra_id=" & MUESTRA & _
               "   AND td.revisar_factura = 1"
    Set rs = datos_bd(consulta)
    revisar_facturacion = False
    If rs.RecordCount > 0 Then
        revisar_facturacion = True
    End If
    Exit Function
fallo:
    revisar_facturacion = False
End Function
Function revisar_no_realizado(MUESTRA As Long) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT d.resultado " & _
               "  FROM determinaciones d, tipos_determinacion td " & _
               " WHERE tipo_determinacion_id=id_tipo_determinacion" & _
               "   AND muestra_id=" & MUESTRA & _
               "   AND d.resultado = '--'"
    Set rs = datos_bd(consulta)
    revisar_no_realizado = False
    If rs.RecordCount > 0 Then
        revisar_no_realizado = True
    End If
    Exit Function
fallo:
    revisar_no_realizado = False
End Function

'E0136-I
Public Function lista_determinaciones_subcontratables_no_enviadas(stranno As String, strCODIGO As String, strRefCliente As String, strDeter As String, strContrata As String) As ADODB.Recordset
    On Error GoTo fallo
    
    Dim consulta As String
    
    consulta = "SELECT CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) CODIGO " & _
               "     , M.REFERENCIA_CLIENTE, TD.NOMBRE DETERMINACION, PR.NOMBRE CONTRATA " & _
               "     , TD.VALOR_REFERENCIA, TD.NORMATIVA_APLICABLE, PR.ID_PROVEEDOR " & _
               "     , M.ID_MUESTRA, TD.ID_TIPO_DETERMINACION ID_TD, D.ID_DETERMINACION " & _
               "  FROM MUESTRAS M, TIPOS_MUESTRA TM " & _
               "     , DETERMINACIONES D, TIPOS_DETERMINACION TD, PROVEEDORES PR " & _
               " WHERE M.ID_MUESTRA = D.MUESTRA_ID " & _
               "   AND D.TIPO_DETERMINACION_ID = TD.ID_TIPO_DETERMINACION " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND TD.CONTRATA_ID = PR.ID_PROVEEDOR " & _
               "   AND M.CERRADA <> 1 " & _
               "   AND M.ANULADA = 0 " & _
               "   AND TD.ES_SUBCONTRATABLE = 1 " & _
               "   AND D.ENVIADO_PAQUETE = 0 " & _
               "   AND M.ANNO LIKE " & stranno & _
               "   AND CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) LIKE '%" & strCODIGO & "%' " & _
               "   AND M.REFERENCIA_CLIENTE LIKE '%" & strRefCliente & "%' " & _
               "   AND TD.NOMBRE LIKE '%" & strDeter & "%' " & _
               IIf(strContrata <> "", " AND PR.ID_PROVEEDOR = " & strContrata, "") & _
               " ORDER BY CODIGO "
    
    Set lista_determinaciones_subcontratables_no_enviadas = datos_bd(consulta)
    Exit Function
    
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function

Public Function lista_determinaciones_duplicadas(fdesde As Date, fhasta As Date, FUERA_RANGO As Boolean, revisadas As Boolean) As ADODB.Recordset
    Dim consulta As String
    consulta = " select td.ID_TIPO_DETERMINACION, CONCAT(td.nombre,' (',td.descripcion,')'), dc.dif_duplicados, IF (det.DIF_DUPLICADOS > dc.DIF_DUPLICADOS,1,0) AS FUERA_RANGO,count(*) " & _
               "  from muestras as mue, determinaciones as det, tipos_determinacion as td, determinaciones_conf dc " & _
               " Where mue.ID_MUESTRA = det.MUESTRA_ID " & _
               "   and det.TIPO_DETERMINACION_ID = td.ID_TIPO_DETERMINACION " & _
               "   and det.ID_DETERMINACION = dc.DETERMINACION_ID " & _
               "   and det.FECHA >= '" & fecha_bd(fdesde) & "' " & _
               "   and det.FECHA <= '" & fecha_bd(fhasta) & "' " & _
               "   and mue.ANALISIS_DUPLICADO = 1  " & _
               "   and mue.CERRADA <> 0  " & _
               "   and mue.ANULADA = 0  " & _
               "   and det.DIF_DUPLICADOS <> ''  " & _
               IIf(FUERA_RANGO, " and det.DIF_DUPLICADOS > dc.DIF_DUPLICADOS ", "") & _
               IIf(revisadas, " and det.usuario_id_revision = 0 ", "") & _
               "  group by td.ID_TIPO_DETERMINACION, td.nombre, IF (det.DIF_DUPLICADOS > dc.DIF_DUPLICADOS,1,0)  " & _
               "  order by td.nombre,td.ID_TIPO_DETERMINACION,FUERA_RANGO DESC"
    Set lista_determinaciones_duplicadas = datos_bd(consulta)
End Function

Public Function marcar_determinacion_enviado_paquete(lngDeter_ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    consulta = "UPDATE DETERMINACIONES " & _
               "   SET ENVIADO_PAQUETE = 1 " & _
               " WHERE ID_DETERMINACION = " & lngDeter_ID
    execute_bd consulta
    marcar_determinacion_enviado_paquete = True
    Exit Function
    
fallo:
    marcar_determinacion_enviado_paquete = False
    MsgBox "Error al marcar determinación como enviada en paquete.", vbCritical, Err.Description
End Function
Public Function marcar_como_revisada(lngDeter_ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    consulta = "UPDATE DETERMINACIONES " & _
               "   SET USUARIO_ID_REVISION = " & USUARIO.getID_EMPLEADO & _
               "       ,FECHA_REVISION = '" & fecha_bd(Date) & "'" & _
               " WHERE ID_DETERMINACION = " & lngDeter_ID & _
               "   AND USUARIO_ID_REVISION = 0"
    execute_bd consulta
    marcar_como_revisada = True
    Exit Function
    
fallo:
    marcar_como_revisada = False
    MsgBox "Error al marcar determinación como marcar_como_revisada.", vbCritical, Err.Description
End Function
Public Function modificar_metodo(lngDeter_ID As Long, lmetodo As String, lgrado As String) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    consulta = "UPDATE DETERMINACIONES " & _
               "   SET METODO = '" & lmetodo & "'" & _
               "     , GRADO  = '" & lgrado & "'" & _
               " WHERE ID_DETERMINACION = " & lngDeter_ID
    execute_bd consulta
    modificar_metodo = True
    Exit Function
    
fallo:
    modificar_metodo = False
    MsgBox "Error al modificar_metodo.", vbCritical, Err.Description
End Function
Public Function modificar_situacion(lngDeter_ID As Long, SITUACION As Integer) As Boolean
    On Error GoTo fallo
    execute_bd "UPDATE DETERMINACIONES SET SITUACION = " & SITUACION & " WHERE ID_DETERMINACION = " & lngDeter_ID
    modificar_situacion = True
    Exit Function
    
fallo:
    modificar_situacion = False
    MsgBox "Error al modificar_situacion.", vbCritical, Err.Description
End Function


Public Function informarGrados(idMuestra As Long) As Boolean
    Dim esTrigo As Boolean
    Dim consulta As String
   On Error GoTo informarGrados_Error

    esTrigo = False
    Dim oMuestra As New clsMuestra
    Dim oTipoAnalisis As New clsTipos_analisis
    oMuestra.CargaMuestra idMuestra
    If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TRIGO Or _
       oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.HARINA_BLANDA Or _
       oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.HARINA_DURO Then
        esTrigo = True
        oTipoAnalisis.CARGAR oMuestra.getTIPO_ANALISIS_ID
    End If
    
    If esTrigo Then
        Dim rsDeter As New ADODB.Recordset
        Set rsDeter = Me.lista_determinaciones(idMuestra)
        Dim GRADO As String
        If rsDeter.RecordCount <> 0 Then
            Do
                GRADO = calcularClasificacionTrigo(oTipoAnalisis.getTIPO_TRIGO, rsDeter("id_determinacion"), rsDeter("resultado"))
                consulta = "UPDATE DETERMINACIONES " & _
                           "   SET GRADO = '" & GRADO & "'" & _
                           " WHERE ID_DETERMINACION = " & rsDeter("id_determinacion")
                execute_bd consulta
                rsDeter.MoveNext
            Loop Until rsDeter.EOF
        End If
    End If
    informarGrados = True

   On Error GoTo 0
   Exit Function

informarGrados_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure informarGrados of Módulo de clase clsDeterminaciones"
End Function
Private Function calcularClasificacionTrigo(tipoTrigo As Integer, idDeterminacion As Long, RESULTADO As String) As String
    Dim salida As String
    salida = "--"
    If Not IsNumeric(RESULTADO) Then
        calcularClasificacionTrigo = salida
        Exit Function
    End If
    Dim res As Single
    res = CSng(Replace(RESULTADO, ".", ","))
    ' Recuperar el tipo de determinacion
    Dim oDeterminacion As New clsDeterminaciones
    oDeterminacion.CargarDeterminacion idDeterminacion
    Dim idTipoDeterminacion As Long
    idTipoDeterminacion = oDeterminacion.getTIPO_DETERMINACION_ID
    ' DUROS
    If tipoTrigo = TIPOS_TRIGO.TIPOS_TRIGO_DURO Then
        If idTipoDeterminacion = CLASIFICACION_TRIGOS_DUROS.CLASIFICACION_TRIGOS_DUROS_PROTEINA Or _
           idTipoDeterminacion = CLASIFICACION_TRIGOS_DUROS.CLASIFICACION_TRIGOS_DUROS_PROTEINA2 Then
            If res >= 13 Then
                salida = "GRUPO 1"
            ElseIf res >= 12 Then
                salida = "GRUPO 2"
            ElseIf res >= 11 Then
                salida = "GRUPO 3"
            Else
                salida = "GRUPO 4"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_DUROS.CLASIFICACION_TRIGOS_DUROS_PESO_ESPECIFICO Then
            If res >= 80 Then
                salida = "GRUPO 1"
            ElseIf res >= 78 Then
                salida = "GRUPO 2"
            ElseIf res >= 77 Then
                salida = "GRUPO 3"
            Else
                salida = "GRUPO 4"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_DUROS.CLASIFICACION_TRIGOS_DUROS_HUMEDAD Then
            If res <= 12 Then
                salida = "GRADO I"
            ElseIf res <= 12.5 Then
                salida = "GRADO II"
            ElseIf res <= 13 Then
                salida = "GRADO III"
            Else
                salida = "GRADO IV"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_DUROS.CLASIFICACION_TRIGOS_DUROS_CENIZAS Then
            If res < 1.75 Then
                salida = "GRADO I"
            ElseIf res < 1.85 Then
                salida = "GRADO II"
            ElseIf res < 2 Then
                salida = "GRADO III"
            Else
                salida = "GRADO IV"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_DUROS.CLASIFICACION_TRIGOS_DUROS_INDICE_CAIDA Then
            If res > 300 Then
                salida = "GRADO I"
            ElseIf res > 250 Then
                salida = "GRADO III"
            Else
                salida = "GRADO IV"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_DUROS.CLASIFICACION_TRIGOS_DUROS_VITROCIDAD Then
            If res > 80 Then
                salida = "GRUPO 1"
            ElseIf res > 75 Then
                salida = "GRUPO 2"
            ElseIf res > 60 Then
                salida = "GRUPO 3"
            Else
                salida = "GRUPO 4"
            End If
        End If
    Else ' BLANDOS
        If idTipoDeterminacion = CLASIFICACION_TRIGOS_BLANDOS.CLASIFICACION_TRIGOS_BLANDOS_PROTEINA Or _
           idTipoDeterminacion = CLASIFICACION_TRIGOS_BLANDOS.CLASIFICACION_TRIGOS_BLANDOS_PROTEINA2 Then
            If res >= 13 Then
                salida = "GRUPO 1"
            ElseIf res >= 12 Then
                salida = "GRUPO 2"
            ElseIf res >= 11 Then
                salida = "GRUPO 3"
            ElseIf res >= 10 Then
                salida = "GRUPO 4"
            Else
                salida = "GRUPO 5"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_BLANDOS.CLASIFICACION_TRIGOS_BLANDOS_INDICE_CAIDA Then
            If res >= 300 Then
                salida = "GRADO I"
            ElseIf res >= 280 Then
                salida = "GRADO II"
            ElseIf res >= 250 Then
                salida = "GRADO III"
            Else
                salida = "GRADO IV"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_BLANDOS.CLASIFICACION_TRIGOS_BLANDOS_HUMEDAD Then
            If res <= 12 Then
                salida = "GRADO I"
            ElseIf res <= 12.5 Then
                salida = "GRADO II"
            ElseIf res <= 13 Then
                salida = "GRADO III"
            Else
                salida = "GRADO IV"
            End If
        ElseIf idTipoDeterminacion = CLASIFICACION_TRIGOS_BLANDOS.CLASIFICACION_TRIGOS_BLANDOS_PESO_ESPECIFICO Then
            If res >= 80 Then
                salida = "GRADO I"
            ElseIf res >= 78 Then
                salida = "GRADO II"
            ElseIf res >= 75 Then
                salida = "GRADO III"
            Else
                salida = "GRADO IV"
            End If
        End If
    End If
    calcularClasificacionTrigo = salida
End Function

Public Function modificar_orden(lngDeter_ID As Long, lORDEN As Integer) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    
    consulta = "UPDATE DETERMINACIONES " & _
               "   SET ORDEN = " & lORDEN & _
               " WHERE ID_DETERMINACION = " & lngDeter_ID
    execute_bd consulta
    modificar_orden = True
    Exit Function
    
fallo:
    modificar_orden = False
    MsgBox "Error al modificar_orden.", vbCritical, Err.Description
End Function

'M0927-I
Public Function lista_determinaciones_subcontratables_parametro(stranno As String, strCODIGO As String, strRefCliente As String, strDeter As String, strContrata As Long, strEnviado As String) As ADODB.Recordset
    On Error GoTo fallo
    
    Dim consulta As String
    
    ' JGM : OPTIMIZACIÓN DE LA CONSULTA
'    consulta = "SELECT CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) CODIGO " & _
'               "     , M.REFERENCIA_CLIENTE, TD.NOMBRE DETERMINACION, PR.NOMBRE CONTRATA " & _
'               "     , TDC.VALOR_REFERENCIA, TDC.NORMATIVA_APLICABLE, TDC.PRECIO, PR.ID_PROVEEDOR " & _
'               "     , M.ID_MUESTRA, TD.ID_TIPO_DETERMINACION ID_TD, D.ID_DETERMINACION, D.ENVIADO_PAQUETE" & _
'               "  FROM MUESTRAS M, TIPOS_MUESTRA TM " & _
'               "     , DETERMINACIONES D, TIPOS_DETERMINACION TD, PROVEEDORES PR, TIPOS_DETERMINACION_CONTRATAS TDC" & _
'               " WHERE M.ID_MUESTRA = D.MUESTRA_ID " & _
'               "   AND D.TIPO_DETERMINACION_ID = TD.ID_TIPO_DETERMINACION " & _
'               "   AND TD.ID_TIPO_DETERMINACION = TDC.TIPO_DETERMINACION_ID " & _
'               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
'               "   AND TDC.CONTRATA_ID = PR.ID_PROVEEDOR " & _
'               "   AND M.CERRADA <> 1 " & _
'               "   AND M.ANULADA = 0 " & _
'               "   AND TD.ES_SUBCONTRATABLE = 1 " & _
'               "   AND D.ENVIADO_PAQUETE IN (" & strEnviado & ")" & _
'               "   AND M.ANNO LIKE " & stranno & _
'               "   AND CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) LIKE '%" & strCODIGO & "%' " & _
'               "   AND M.REFERENCIA_CLIENTE LIKE '%" & strRefCliente & "%' " & _
'               "   AND TD.NOMBRE LIKE '%" & strDeter & "%' " & _
'               IIf(strContrata <> "", " AND PR.ID_PROVEEDOR = " & strContrata, "") & _
'               " ORDER BY CODIGO "
    
    Dim filtro As String
    If Trim(strCODIGO) <> "" Then
        filtro = filtro & " AND CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) LIKE '%" & strCODIGO & "%' "
    End If
    If Trim(strRefCliente) <> "" Then
        filtro = filtro & " AND M.REFERENCIA_CLIENTE LIKE '%" & strRefCliente & "%' "
    End If
    If Trim(strDeter) <> "" Then
        filtro = filtro & " AND TD.NOMBRE LIKE '%" & strDeter & "%' "
    End If
    If strEnviado = "0" Then
        filtro = filtro & " AND D.ENVIADO_PAQUETE = 0 "
    End If
    If strContrata <> 0 Then
        filtro = filtro & " AND PR.ID_PROVEEDOR = " & strContrata
    End If

    consulta = "SELECT CONCAT(TM.CODIGO, '-', CAST(M.ID_PARTICULAR AS CHAR)) CODIGO " & _
               "     , M.REFERENCIA_CLIENTE, TD.NOMBRE DETERMINACION, PR.NOMBRE CONTRATA " & _
               "     , TDC.VALOR_REFERENCIA, TDC.NORMATIVA_APLICABLE, TDC.PRECIO, PR.ID_PROVEEDOR " & _
               "     , M.ID_MUESTRA, TD.ID_TIPO_DETERMINACION ID_TD, D.ID_DETERMINACION, D.ENVIADO_PAQUETE" & _
               "  FROM MUESTRAS M, TIPOS_MUESTRA TM " & _
               "     , DETERMINACIONES D, TIPOS_DETERMINACION TD, PROVEEDORES PR, TIPOS_DETERMINACION_CONTRATAS TDC" & _
               " WHERE M.ID_MUESTRA = D.MUESTRA_ID " & _
               "   AND D.TIPO_DETERMINACION_ID = TD.ID_TIPO_DETERMINACION " & _
               "   AND TD.ID_TIPO_DETERMINACION = TDC.TIPO_DETERMINACION_ID " & _
               "   AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA " & _
               "   AND TDC.CONTRATA_ID = PR.ID_PROVEEDOR " & _
               "   AND M.CERRADA <> 1 " & _
               "   AND M.ANULADA = 0 " & _
               "   AND TD.ES_SUBCONTRATABLE = 1 " & _
               "   AND M.ANNO = " & stranno & _
               filtro & _
               " ORDER BY CODIGO "
    
    Set lista_determinaciones_subcontratables_parametro = datos_bd(consulta)
    Exit Function
    
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function
'M0927-F

Public Function existePendiente(ID_MUESTRA As Long) As Boolean
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = " select ID_DETERMINACION from determinaciones where RESULTADO = 'PENDIENTE' and MUESTRA_ID = " & ID_MUESTRA
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        existePendiente = True
    Else
        existePendiente = False
    End If
    Set rs = Nothing
    
End Function

